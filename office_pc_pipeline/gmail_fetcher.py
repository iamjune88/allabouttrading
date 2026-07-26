# -*- coding: utf-8 -*-
"""
Gmail에서 선물거래 자료를 다운로드 — 제목으로 정확히 특정한다.

각 자료는 발신자 + 제목으로 유일하게 지정되므로, "안 오면 재시도/여러 개면 골라내기"
같은 추측 로직 없이 필요한 파일만 정확히 받는다. 지정 메일이 없으면 조용히 넘어가지
않고 경고를 남긴다(누락을 크로스체크가 잡을 수 있도록).

수집 대상 (크로스체크에 필요한 최소 세트):
  SS 미결제약정 ← [삼성선물]선물옵션거래 및 예탁자산현황(가정산보고서)  (PDF)
  SS 체결(틱)   ← 매매내역_삼성선물                                    (CSV)
  NH 체결+미결  ← 국문가정산(체결시분)                                 (CSV)
"""
import base64
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

# 수집 스펙: 각 항목이 메일 1건을 유일하게 특정한다.
FETCH_SPECS = [
    {"role": "SS미결(PDF)", "source": "SS선물", "sender": "master@ssfutures.com",
     "subject": "선물옵션거래 및 예탁자산현황(가정산보고서)", "ext": "pdf", "required": True},
    {"role": "SS체결(CSV)", "source": "SS선물", "sender": "master@ssfutures.com",
     "subject": "매매내역_삼성선물", "ext": "csv", "required": True},
    {"role": "NH체결+미결(CSV)", "source": "NH선물", "sender": "nhfutures@futures.co.kr",
     "subject": "국문가정산(체결시분)", "ext": "csv", "required": True},
]


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
                refreshed = False
        if not refreshed:
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDS_FILE), SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())

    return build("gmail", "v1", credentials=creds)


def fetch_pdfs(target_date: str = None) -> list[dict]:
    """
    target_date: "YYYY/MM/DD" (없으면 오늘)
    반환: [{"source","path","date","filename","kind","role"}, ...]
    FETCH_SPECS의 (발신자+제목+확장자)로 각 자료를 특정. 추측 fallback·dedup 없음.
    """
    from datetime import date
    if not target_date:
        target_date = date.today().strftime("%Y/%m/%d")

    service = get_gmail_service()
    results = []

    for spec in FETCH_SPECS:
        q = (f'from:{spec["sender"]} after:{target_date} '
             f'has:attachment filename:{spec["ext"]} subject:"{spec["subject"]}"')
        resp = service.users().messages().list(userId="me", q=q).execute()
        messages = resp.get("messages", [])

        if not messages:
            level = "경고" if spec["required"] else "정보"
            print(f"[{level}] {spec['role']} 메일 없음 (제목: {spec['subject']})")
            continue

        got = _download_spec_attachments(service, messages, spec, target_date)
        if not got and spec["required"]:
            print(f"[경고] {spec['role']} 메일은 있으나 {spec['ext'].upper()} 첨부 없음")
        results.extend(got)

    return results


def _download_spec_attachments(service, messages, spec, target_date) -> list:
    """spec에 매칭된 메일들에서 해당 확장자 첨부를 내려받는다."""
    out = []
    for msg_meta in messages:
        msg = service.users().messages().get(
            userId="me", id=msg_meta["id"], format="full"
        ).execute()

        headers = {h["name"]: h["value"] for h in msg["payload"]["headers"]}
        try:
            received_date = parsedate_to_datetime(headers["Date"]).strftime("%Y-%m-%d")
        except Exception:
            received_date = target_date.replace("/", "-")

        for part in _iter_parts(msg["payload"]):
            filename = part.get("filename", "")
            if not filename.lower().endswith("." + spec["ext"]):
                continue
            att_id = part.get("body", {}).get("attachmentId")
            if not att_id:
                continue

            att = service.users().messages().attachments().get(
                userId="me", messageId=msg_meta["id"], id=att_id
            ).execute()
            data = base64.urlsafe_b64decode(att["data"])
            save_path = DOWNLOAD_DIR / f"{received_date}_{spec['source']}_{filename}"
            with open(save_path, "wb") as f:
                f.write(data)

            print(f"  [다운로드] {spec['role']}: {save_path.name}")
            out.append({
                "source": spec["source"],
                "path": save_path,
                "date": received_date,
                "filename": filename,
                "kind": spec["ext"],
                "role": spec["role"],
            })
    return out


def _iter_parts(payload):
    """멀티파트 메일에서 모든 파트를 재귀 탐색"""
    yield payload
    for part in payload.get("parts", []):
        yield from _iter_parts(part)
