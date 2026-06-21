"""
Gmail 재인증 스크립트
- 브라우저가 자동으로 열립니다
- iamjune88@gmail.com 로그인 후 허용 클릭
"""
from pathlib import Path
from google_auth_oauthlib.flow import InstalledAppFlow

BASE_DIR = Path(__file__).parent
CREDENTIALS_PATH = BASE_DIR / "config" / "credentials.json.json"
TOKEN_PATH = BASE_DIR / "config" / "token.json"
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_PATH), SCOPES)
creds = flow.run_local_server(port=0, open_browser=True, timeout_seconds=300)
TOKEN_PATH.write_text(creds.to_json())
print(f"토큰 저장 완료: {TOKEN_PATH}")
