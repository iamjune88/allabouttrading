# 인포맥스 3807 일자별 스크래퍼

3807(투자자별 매매동향) 화면을 **날짜별로 자동 조회 → Shift+Z 엑셀보내기 → 파싱/누적**하여
장기 시계열(분봉 누적 + 일별 요약)을 만든다.

## 동작 원리
날짜마다: FROM 날짜칸에 대상일자 입력(+Enter) → 당일 데이터 재조회 → `Shift+Z`로 새 엑셀
워크북 생성 → win32com으로 `raw/3807_YYYYMMDD.xlsx` 저장/닫기 → 파싱 → CSV 누적.
`data/state.json`에 진행상황을 기록하므로 **중단해도 이어서 재실행** 가능.

## 실행 전 준비 (중요)
1. **3807 화면을 열어둔다.** 원하는 종목(예: KTB_03)으로 세팅. TO 날짜는 블락이라 건드리지 않음.
2. **다른 Excel 창을 모두 닫는다.** (캡처가 "저장 안 된 새 워크북"을 잡으므로 깨끗한 환경 필요)
3. 실행 중에는 **마우스/키보드를 건드리지 않는다.** (스크립트가 커서·입력을 조작함)
4. 장중 충돌을 피하려면 **장 마감 후**에 돌린다.

## 실행
```powershell
cd C:\Users\infomax\infomax_3807_scraper

# 대상일자만 미리 확인
python scrape_3807.py --start 2024-01-01 --end 2026-07-24 --dry-run

# 실제 수집 (오래된 날부터)
python scrape_3807.py --start 2024-01-01 --end 2026-07-24

# 최신일부터
python scrape_3807.py --start 2024-01-01 --end 2026-07-24 --descending

# 오류난 날짜만 재시도
python scrape_3807.py --start 2024-01-01 --end 2026-07-24 --retry-errors

# 수집 후 휴장/중복 의심일 점검
python verify_daily.py
```

## 산출물 (`data/`)
- `intraday.csv` — 분봉 **누적치** long-format: `date,time,datetime,index,{frgn,indi,inst}_{sell,buy,net}`
  (분당 순flow가 필요하면 각 net을 `date`별로 diff)
- `daily.csv` — 일별 최종 누적 요약(한 행 = 하루)
- `state.json` — 재실행 이어받기용
- `raw/3807_YYYYMMDD.xlsx` — 엑셀보내기 원본 보관

## 값 의미
- 컬럼: 외국인합/개인/기관계 × 매도/매수/순매수, 단위 (천주,백만,계약)
- **장중 누적**: 08:46 → 15:45(장 마감) 증가, 이후 시간은 최종값 반복. 마지막 시각행 = 당일 최종 = 일별요약.

## 튜닝 (scrape_3807.py 상단)
- `REQUERY_WAIT` — 날짜 입력 후 재조회 대기(네트워크 느리면 ↑)
- `CAPTURE_TIMEOUT` — 엑셀 워크북 등장 대기
- `MIN_BARS_TRADING` — 이보다 봉이 적으면 휴장 처리

## 주의: 휴장일 처리
후보일은 `holidays` 라이브러리로 주말·공휴일을 1차 제외한다. 임시 휴장(라이브러리 미반영)에
과거 자료가 딸려올 수 있으니, 수집 후 `python verify_daily.py`로 전일과 값이 완전히 같은 날을
점검하고 필요시 해당 raw 파일을 확인한다.

## 구성
- `scrape_3807.py` — 메인 드라이버(창 조작 + 루프)
- `parser.py` — 엑셀보내기 xlsx → 정제 DataFrame
- `excel_capture.py` — Shift+Z 결과 워크북 저장/닫기(win32com)
- `calendar_krx.py` — 후보 거래일 생성
- `storage.py` — CSV 누적 + state 관리
- `verify_daily.py` — 품질 점검
