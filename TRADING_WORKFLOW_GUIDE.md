# KTB 선물 트레이딩 워크플로우 가이드
> Claude Code + TradingView MCP 기반 | 2026-06-15 작성

---

## 1. 환경 셋업

### TradingView 실행 (MSIX/AppxPackage 방식)
```powershell
# 방법 1: 런치 스크립트 직접 실행 (추천)
& "C:\Users\iamju\Desktop\workingwithC\launch_tradingview.ps1"

# 방법 2: 수동 실행
$pkg = Get-AppxPackage | Where-Object { $_.Name -like "*TradingView*" } | Select-Object -First 1
$tvExe = Join-Path $pkg.InstallLocation "TradingView.exe"
$user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
icacls $pkg.InstallLocation /grant "${user}:(OI)(CI)RX" /T /Q 2>&1 | Out-Null
Get-Process -Name "TradingView" -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Process -FilePath $tvExe -ArgumentList "--remote-debugging-port=9222"
```

> ⚠️ TV는 Microsoft Store(AppxPackage) 설치. 일반 경로 탐색 실패함. 반드시 위 방법 사용.

### CDP 연결 확인
```
http://localhost:9222/json/version
```
또는 Claude Code에서: `tv_health_check` 툴 호출

---

## 2. 심볼 매핑

### 국내 KTB 선물
| 가정산보고서 코드 | 종목 | TradingView 심볼 |
|--|--|--|
| A6566 | KTB3년물 선물 | `KRX:BM31!` |
| A6766 | KTB10년물 선물 | `KRX:BMA1!` |
| A6966 | KTB30년물 선물 | `KRX:BML1!` |

### 해외 채권선물
| 종목 | TradingView 심볼 |
|--|--|
| US 2yr (ZT) | `ZT1!` |
| US 10yr (ZN) | `ZN1!` |
| German Schatz (2yr) | `FGBS1!` |
| German Bund (10yr) | `FGBL1!` |
| UK Long Gilt | `Z1!` (확인 필요) |

---

## 3. 가정산보고서 → TV 거래 증적 워크플로우

### 가정산보고서 파싱 항목
```
[거래내역] 에서 추출:
- 종목코드 → TV 심볼 변환
- 구분 (매도/매수)
- 수량
- 당일체결가 (평균)
- 체결 시간 (체결내역 탭에서 확인)

[미결제약정] 에서 확인:
- 현재 잔고
- 당일정산가
- 갱신차금 (MTM 손익)
```

### 타임스탬프 계산 (KST → Unix)

**기준값:**
```
2026-01-01 00:00 UTC = 1767225600
2026-06-15 00:00 UTC = 1781481600  ← 자주 쓰는 기준
하루 = 86400초
```

**KST → UTC 변환:**
```
UTC = KST - 9시간 (= KST - 32400초)

예시 (2026-06-15):
09:01 KST = 00:01 UTC → 5분봉 bar 시작: 00:00 UTC = 1781481600
09:03 KST = 00:03 UTC → 5분봉 bar 시작: 00:00 UTC = 1781481600  ← 09:01과 같은 봉!
09:26 KST = 00:26 UTC → 5분봉 bar 시작: 00:25 UTC = 1781481600 + (25×60) = 1781483100
10:41 KST = 01:41 UTC → 5분봉 bar 시작: 01:40 UTC = 1781481600 + (100×60) = 1781487600
```

**날짜별 00:00 UTC 계산:**
```
기준일 + N일 = 기준 timestamp + N × 86400
예: 2026-06-16 = 1781481600 + 86400 = 1781568000
    2026-06-17 = 1781481600 + 172800 = 1781654400
```

---

## 4. draw_shape 라벨링 컨벤션

### 기본 원칙
- **SELL 라벨** → 체결가 **위에** 배치 (체결가 + 0.06~0.08pt)
- **BUY 라벨** → 체결가 **아래에** 배치 (체결가 - 0.06~0.08pt)
- **같은 봉에 2개 이상** → 0.06pt 이상 위아래 간격 확보

### 라벨 포맷 (Claude Code 명령 기준)

**SELL 라벨:**
```python
draw_shape(
    shape="text",
    point={"time": {bar_timestamp}, "price": {체결가 + 0.07}},
    text="▼ {HH:MM}  SELL {수량} @ {가격}",
    overrides={"color": "#FF4444", "fontsize": 12, "bold": True}
)
```

**BUY 라벨:**
```python
draw_shape(
    shape="text",
    point={"time": {bar_timestamp}, "price": {체결가 - 0.07}},
    text="▲ {HH:MM}  BUY(커버) {수량} @ {가격}",
    overrides={"color": "#44BB44", "fontsize": 12, "bold": True}
)
```

**평균단가 수평선:**
```python
draw_shape(
    shape="horizontal_line",
    point={"time": {첫_체결_timestamp}, "price": {평균가}},
    text="SELL avg {평균가} (-{총수량}계약)",
    overrides={"linecolor": "#FF8888", "linewidth": 1, "linestyle": 1}
)
```

### 실행 순서
```
1. chart_set_symbol → 해당 종목
2. chart_set_timeframe → "5" (5분봉)
3. draw_clear → 기존 드로잉 제거
4. draw_shape × 각 체결건 (SELL/BUY)
5. draw_shape × avg 수평선
6. capture_screenshot → 확인
7. 다음 종목으로 전환 후 반복
```

---

## 5. 테크니컬 분석 워크플로우

### 세션 시작 시
```
1. tv_health_check → 연결 확인
2. chart_get_state → 현재 심볼/봉/지표 확인
3. 분석 종목 순서: KTB3(BM31!) → KTB10(BMA1!) → 필요시 해외
```

### 중장기 분석 (일봉 기준)
```
chart_set_symbol → chart_set_timeframe("D")
→ data_get_ohlcv(summary=true, count=100)
→ data_get_study_values
→ capture_screenshot
→ WebSearch (글로벌 금리/뉴스)
```

### 분석 필수 체크 항목
| 항목 | 도구 |
|--|--|
| 현재가 / OHLC | `quote_get` |
| MA 스택, RSI, VFI, CVD | `data_get_study_values` |
| VRVP 레벨, ATR 존 | `data_get_pine_lines` / `data_get_pine_labels` |
| 가격 요약 | `data_get_ohlcv(summary=true)` |
| 시각 확인 | `capture_screenshot` |

---

## 6. 매크로 뉴스 서치 쿼리 (WebSearch)

```
"Korea KTB bond market [월 연도]"
"Japan JGB super long yield [월 연도]"
"Fed rate cut expectations [월 연도]"
"BOJ interest rate decision [월 연도]"
"ECB interest rate decision [월 연도]"
"한국은행 기준금리 전망 [연도]"
```

---

## 7. 트레이딩 평가 기준 (테크니컬)

### SELL 진입 평가
- **A**: 장중 고점 ±0.03pt 이내 진입
- **B**: 당일 상위 25% 가격대 진입
- **C**: 중간값 이하 진입 (타이밍 아쉬움)

### BUY 커버 평가
- **A**: 장중 저점 ±0.03pt 이내 커버
- **B**: 당일 하위 25% 가격대 커버
- **C**: 중간값 이상 커버 (갭업 직후 패닉 커버 등)

### 분할매매 평가 기준
- 첫 트랜치: 방향성 확인 전 선진입
- 추가 트랜치: 방향 확인 후 추격 or 가격 개선 확인 후 추가
- 같은 방향 3번 이상 분할: 25~30분 간격 이내 집행이 이상적

---

## 8. 포지션 관리 원칙 (2026-06 현황 기반)

### 헷지 비율 기준
| 시장 리스크 수준 | 권장 헷지 비율 |
|--|--|
| 일반 (변동성 낮음) | 20~30% |
| 주의 (이벤트 앞) | 40~60% |
| 경계 (NFP 충격 등) | **50~70%** |
| 위기 (시장 붕괴 우려) | 70~90% |

### 손절 트리거 원칙
- 기술적 지지선 **이탈 확인 후** 차음 봉 시가에서 집행
- 갭다운 시작 → 당일 추가 손절 우선 고려
- 정책개입 시그널 나오면 → 헷지 먼저 축소 후 판단

---

## 9. 자주 쓰는 Claude Code 명령 패턴

```
# TV 켜기
launch_tradingview.ps1 실행 후 tv_health_check

# 전종목 테크니컬 스캔
chart_set_symbol → chart_set_timeframe("D") → data_get_ohlcv(summary=true) + capture_screenshot
→ 종목별 반복

# 거래 증적
draw_clear → draw_shape(text) × N건 → draw_shape(horizontal_line) × avg → capture_screenshot

# 뉴스 + 차트 동시
WebSearch × 3~4개 쿼리 (병렬) + chart_set_timeframe + data_get_ohlcv (병렬)
```

---

## 10. 주의사항 / 알려진 이슈

| 문제 | 해결책 |
|--|--|
| TV CDP 연결 실패 | `launch_tradingview.ps1` 재실행 |
| batch_run 전체 실패 | 개별 심볼 전환 후 순차 실행 |
| 라벨 겹침 | 같은 봉: 0.06pt 이상 가격 간격 / 또는 TV에서 직접 드래그 |
| 차트 전체뷰에서 라벨 눌림 | TV 직접 드래그로 미세조정 (API 한계) |
| 외부 HTTPS 차단 (회사 프록시) | `verify=False` 패치 (`__init__.py`) |
| MSIX앱 일반경로 탐색 실패 | AppxPackage 방식만 사용 |

---

*작성: 2026-06-15 | Claude Code (claude-sonnet-4-6) 기반*
