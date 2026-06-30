# -*- coding: utf-8 -*-
"""
Gmail에서 선물거래 확인서 PDF를 다운로드
- 발신자: NHfutures@futures.co.kr / master@ssfutures.com
- 처음 실행 시 브라우저 OAuth 인증 필요 (token.json 생성됨)
"""
import base64
import os
import pickle
from email.utils import parsedate_to_datetime
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

BASE_DIR = Path(__file__).parent
CREDS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE = BASE_DIR / "token.json"
DOWNLOAD_DIR = BASE_DIR / "다운로드"
DOWNLOAD_DIR.mkdir(exist_ok=True)

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

SENDERS = {
    "nhfutures@futures.co.kr": "NH선물",
    "master@ssfutures.com": "SS선물",
}

# NH선물은 하루에 01A101(체결시간 없음)/01A103(체결시간 포함,"국문가정산(체결시분)")/
# 02A101(예탁자산현황, 무관) 3종 메일을 보낸다. 체결시간이 꼭 필요하므로 검색 자체를
# "국문가정산(체결시분)" 제목 메일로 좁혀서, 01A101/02A101은 처음부터 받지 않는다.
SUBJECT_FILTERS = {
    "nhfutures@futures.co.kr": 'subject:"국문가정산(체결시분)"',
}


def get_gmail_service():
    creds = None
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)

    if not creds or not creds.valid:
        refreshed = False
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                refreshed = True
            except Exception:
                refreshed = False  # refresh token expired/revoked - fall back to interactive login
        if not refreshed:
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDS_FILE), SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())

    return build("gmail", "v1", credentials=creds)


def fetch_pdfs(target_date: str = None) -> list[dict]:
    """
    target_date: "YYYY/MM/DD" 형식 (없으면 오늘)
    반환: [{"source": "NH선물"|"SS선물", "path": Path, "date": str}, ...]
    """
    from datetime import date
    if not target_date:
        target_date = date.today().strftime("%Y/%m/%d")

    service = get_gmail_service()
    results = []

    for sender_email, source_name in SENDERS.items():
        # Gmail 검색 쿼리
        base_query = f"from:{sender_email} after:{target_date.replace('/', '/')} has:attachment filename:pdf"
        query = base_query
        if sender_email in SUBJECT_FILTERS:
            query += f" {SUBJECT_FILTERS[sender_email]}"
        resp = service.users().messages().list(userId="me", q=query).execute()
        messages = resp.get("messages", [])

        # 체결시분(01A103) 리포트가 그날 안 왔을 수 있음(예: 정정포함 리포트만 발송된 경우)
        # — 제목 필터 없이 재시도해서 다른 변형이라도 받아온다. _dedup_nh가 01A103 > 01A101/02A101
        # 순으로 최선의 변형을 골라내므로, 시간 없는 리포트라도 우선 기록할 수 있다.
        if not messages and sender_email in SUBJECT_FILTERS:
            print(f"[{source_name}] 체결시분 메일 없음 — 제목 필터 없이 재시도")
            resp = service.users().messages().list(userId="me", q=base_query).execute()
            messages = resp.get("messages", [])

        print(f"[{source_name}] {len(messages)}개 메일 발견")

        for msg_meta in messages:
            msg = service.users().messages().get(
                userId="me", id=msg_meta["id"], format="full"
            ).execute()

            # 날짜 헤더 파싱 (실제 메일 수신일 — target_date는 검색 기준일일 뿐이라 그대로 쓰면 안 됨)
            headers = {h["name"]: h["value"] for h in msg["payload"]["headers"]}
            try:
                received_date = parsedate_to_datetime(headers["Date"]).strftime("%Y-%m-%d")
            except Exception:
                received_date = target_date.replace("/", "-")

            # 첨부파일 탐색
            for part in _iter_parts(msg["payload"]):
                filename = part.get("filename", "")
                if not filename.lower().endswith(".pdf"):
                    continue

                att_id = part.get("body", {}).get("attachmentId")
                if not att_id:
                    continue

                att = service.users().messages().attachments().get(
                    userId="me", messageId=msg_meta["id"], id=att_id
                ).execute()

                data = base64.urlsafe_b64decode(att["data"])
                save_path = DOWNLOAD_DIR / f"{received_date}_{source_name}_{filename}"

                with open(save_path, "wb") as f:
                    f.write(data)

                print(f"  [다운로드] {save_path.name}")
                results.append({
                    "source": source_name,
                    "path": save_path,
                    "date": received_date,
                    "filename": filename,
                })

    # 중복 제거
    results = _dedup_nh(results)
    results = _dedup_ss(results)
    return results


def _dedup_nh(results: list) -> list:
    """
    같은 날짜의 NH선물 파일이 여러 개일 때:
    - 01A103 (체결시간 포함 확인서) 우선
    - 02A101 (위탁자산현황) 제외 — 동일 거래 데이터 중복
    - 01A101 (체결시간 없는 버전)은 01A103 없을 때만 포함
    """
    nh_by_date: dict[str, list] = {}
    others = []

    for item in results:
        if item["source"] == "NH선물":
            d = item["date"]
            nh_by_date.setdefault(d, []).append(item)
        else:
            others.append(item)

    deduped_nh = []
    for d, items in nh_by_date.items():
        # 02A101 제외 (위탁자산현황 — 거래 데이터 중복)
        filtered = [x for x in items if "02A101" not in x["filename"]]
        if not filtered:
            filtered = items

        # 01A103 우선 (체결시간 포함), 없으면 01A101 사용
        a103 = [x for x in filtered if "01A103" in x["filename"]]
        if a103:
            deduped_nh.extend(a103)
        else:
            deduped_nh.extend(filtered)

    return others + deduped_nh


def _dedup_ss(results: list) -> list:
    """
    같은 날짜 SS선물 파일이 여러 개일 때:
    - 체결내역 포함 파일(파일명 숫자 더 짧은 것, 보통 오전 발송) 우선
    - 두 번째 파일(오후 발송, 미결제 업데이트본)은 체결내역 없으면 스킵
    """
    import pdfplumber

    def has_exec_section(path) -> bool:
        try:
            with pdfplumber.open(path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text() or ""
                    if "체결내역" in text:
                        return True
        except Exception:
            pass
        return False

    ss_by_date: dict[str, list] = {}
    others = []

    for item in results:
        if item["source"] == "SS선물":
            d = item["date"]
            ss_by_date.setdefault(d, []).append(item)
        else:
            others.append(item)

    deduped_ss = []
    for d, items in ss_by_date.items():
        if len(items) == 1:
            deduped_ss.extend(items)
            continue
        # 체결내역 있는 파일 우선
        with_exec = [x for x in items if has_exec_section(x["path"])]
        if with_exec:
            deduped_ss.extend(with_exec)
        else:
            deduped_ss.extend(items)

    return others + deduped_ss


def _iter_parts(payload):
    """멀티파트 메일에서 모든 파트를 재귀 탐색"""
    yield payload
    for part in payload.get("parts", []):
        yield from _iter_parts(part)
