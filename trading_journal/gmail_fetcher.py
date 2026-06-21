"""
Gmail 연동 모듈 (OAuth 2.0, 개인 Gmail 계정)

사전 준비 (최초 1회, 사용자가 직접 수행):
1. https://console.cloud.google.com 에서 프로젝트 생성
2. APIs & Services > Library 에서 "Gmail API" 검색 후 Enable
3. APIs & Services > OAuth consent screen 구성
   - Audience: External (개인 Gmail이면) 또는 Internal (Workspace 계정이면)
   - Test users에 본인 Gmail 주소 추가
4. APIs & Services > Credentials > Create Credentials > OAuth client ID
   - Application type: Desktop app
   - 다운받은 JSON을 이 파일과 같은 폴더에 credentials.json 으로 저장
5. 최초 실행 시 브라우저가 열리며 로그인 동의 -> token.json 자동 생성
   (이후 실행부터는 token.json으로 자동 인증, 만료시 refresh_token으로 자동 갱신)

스코프: gmail.readonly (읽기 전용 - 첨부파일 다운로드만 하므로 이걸로 충분,
        메일 발송/삭제 권한은 의도적으로 요청하지 않음)
"""
import base64
import os
from pathlib import Path
from datetime import datetime, timedelta

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

BASE_DIR = Path(__file__).parent
CREDENTIALS_PATH = BASE_DIR / "config" / "credentials.json.json"
TOKEN_PATH = BASE_DIR / "config" / "token.json"


def get_gmail_service():
    """OAuth 인증 후 Gmail API 서비스 객체 반환. 최초 1회 브라우저 동의 필요."""
    creds = None
    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDENTIALS_PATH.exists():
                raise FileNotFoundError(
                    f"{CREDENTIALS_PATH} 가 없습니다. Google Cloud Console에서 "
                    f"OAuth 클라이언트(Desktop app)를 만들고 JSON을 다운받아 "
                    f"이 경로에 저장하세요."
                )
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_PATH), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_PATH.parent.mkdir(exist_ok=True)
        TOKEN_PATH.write_text(creds.to_json())

    return build("gmail", "v1", credentials=creds)


def search_messages(service, query: str, max_results: int = 50) -> list[dict]:
    """
    Gmail 검색 쿼리로 메일 목록 조회.
    query 예시: 'subject:"가정산보고서" after:2026/06/20 has:attachment'
    """
    results = []
    page_token = None
    while True:
        resp = service.users().messages().list(
            userId="me", q=query, maxResults=min(max_results, 100), pageToken=page_token
        ).execute()
        results.extend(resp.get("messages", []))
        page_token = resp.get("nextPageToken")
        if not page_token or len(results) >= max_results:
            break
    return results[:max_results]


def get_message_detail(service, msg_id: str) -> dict:
    return service.users().messages().get(userId="me", id=msg_id, format="full").execute()


def get_subject(msg_detail: dict) -> str:
    headers = msg_detail.get("payload", {}).get("headers", [])
    for h in headers:
        if h["name"].lower() == "subject":
            return h["value"]
    return ""


def get_date(msg_detail: dict) -> str:
    headers = msg_detail.get("payload", {}).get("headers", [])
    for h in headers:
        if h["name"].lower() == "date":
            return h["value"]
    return ""


def list_attachments(msg_detail: dict) -> list[dict]:
    """첨부파일 메타정보(filename, attachmentId, mimeType) 리스트 반환"""
    attachments = []

    def _walk(parts):
        for part in parts:
            filename = part.get("filename", "")
            body = part.get("body", {})
            if filename and body.get("attachmentId"):
                attachments.append({
                    "filename": filename,
                    "attachment_id": body["attachmentId"],
                    "mime_type": part.get("mimeType", ""),
                })
            if "parts" in part:
                _walk(part["parts"])

    payload = msg_detail.get("payload", {})
    if "parts" in payload:
        _walk(payload["parts"])
    return attachments


def download_attachment(service, msg_id: str, attachment_id: str, save_path: str) -> str:
    att = service.users().messages().attachments().get(
        userId="me", messageId=msg_id, id=attachment_id
    ).execute()
    data = base64.urlsafe_b64decode(att["data"])
    Path(save_path).parent.mkdir(parents=True, exist_ok=True)
    with open(save_path, "wb") as f:
        f.write(data)
    return save_path


def fetch_attachments_for_subject_patterns(
    service, subject_query: str, save_dir: str, after_date: str | None = None,
    max_results: int = 30,
) -> list[dict]:
    """
    subject_query: Gmail 검색 쿼리 (subject: 조건 포함)
    after_date: 'YYYY/MM/DD' 형식, 지정하면 해당 날짜 이후 메일만 검색
    반환: [{"msg_id", "subject", "date", "saved_paths": [...]}]
    """
    query = subject_query
    if after_date:
        query += f" after:{after_date}"
    query += " has:attachment"

    messages = search_messages(service, query, max_results=max_results)
    results = []
    for m in messages:
        detail = get_message_detail(service, m["id"])
        subject = get_subject(detail)
        date = get_date(detail)
        attachments = list_attachments(detail)
        saved_paths = []
        for att in attachments:
            if not att["filename"].lower().endswith(".pdf"):
                continue
            safe_name = f"{m['id']}_{att['filename']}"
            save_path = os.path.join(save_dir, safe_name)
            download_attachment(service, m["id"], att["attachment_id"], save_path)
            saved_paths.append(save_path)
        results.append({
            "msg_id": m["id"], "subject": subject, "date": date,
            "saved_paths": saved_paths,
        })
    return results


if __name__ == "__main__":
    # 연동 테스트 (실제 실행은 로컬 Claude Code 환경에서)
    print("Gmail 인증 테스트는 credentials.json이 준비된 로컬 환경에서 실행하세요.")
    print(f"필요 파일 경로: {CREDENTIALS_PATH}")
    print(f"토큰 저장 경로: {TOKEN_PATH}")
