# Trading Journal Automation

## 셋업 순서

1. 패키지 설치
   pip install -r requirements.txt

2. Google Cloud Console에서 발급받은 OAuth 클라이언트 JSON을
   config/credentials.json 으로 저장 (반드시 이 이름이어야 함)

3. 최초 실행 (브라우저 로그인 동의창이 뜸)
   python3 daily_run.py --after 2026/06/01

   이후부터는 config/token.json 으로 자동 인증됩니다.

4. 평소 실행 (당일 메일만 처리)
   python3 daily_run.py

5. 로컬 PDF로 테스트하고 싶을 때 (Gmail 연동 없이)
   data/attachments/ 폴더에 PDF를 넣고:
   python3 daily_run.py --dry-run

## 결과물
output/trading_journal.xlsx 에 Fills(체결내역)/Trades(진입청산매칭) 누적

## cron 등록 예시 (평일 16:30 자동 실행)
30 16 * * 1-5 cd /path/to/trading_journal && /usr/bin/python3 daily_run.py >> logs/daily_run.log 2>&1
