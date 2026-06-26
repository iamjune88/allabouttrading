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
| (참고) Claude Code/VSCode 셸에서 TV 실행 시 `bad option: --remote-debugging-port` 에러 + 즉시 종료 (exit code 9) | 다른 PC에서는 동일 스크립트가 문제없이 작동함 — 특정 환경(Claude Code가 띄운 셸)에 `ELECTRON_RUN_AS_NODE` 환경변수가 남아있어 Electron 앱이 Node CLI 모드로 부팅되는 경우에만 발생. 해당 환경이면 `Remove-Item Env:\ELECTRON_RUN_AS_NODE -ErrorAction SilentlyContinue` 후 재시도 |
| `tradingview-mcp` (tradesdontlie 원본) `draw clear`/`draw list`/`draw get`/`draw remove`가 `getChartApi is not defined` 에러 | 원본 저장소 `src/core/drawing.js`의 버그(`drawShape`만 `_resolve(_deps)` 호출, 나머지 4개 함수는 누락) — 새로 `git clone` 할 때마다 재발함. 4개 함수에 `const { evaluate, getChartApi } = _resolve(_deps);` 한 줄씩 추가해서 패치 (남의 저장소라 push 불가, 매번 로컬 패치 필요) |
| `chart_set_visible_range`로 빈 여백(마지막 봉 이후 등) 확보 안 됨 | 봉 인덱스 기반이라 실제 데이터 범위 밖은 항상 가장 가까운 실제 봉으로 clamp됨. `ui scroll`은 수직 줌이라 대안 안 됨 — 현재 도구 한계로 받아들이고 넘어갈 것 |
| `draw_shape`(text) 라벨이 `draw get`상으로는 정상인데 스크린샷엔 안 보임 | ① 지표 범례 패널이 차트 좌측 ~150px을 덮어서 그 구간 봉의 라벨이 가려짐 ② 라벨 가격 오프셋이 당일 고저폭 대비 너무 커서 자동확대범위 밖으로 밀려남(수평선은 동일 문제 없음 — 자동확대범위를 확장시킴). 오프셋을 당일 범위의 10~15% 수준으로 줄이고, 장초반 봉은 범례에 가려질 수 있음을 감안할 것 |
| CDP 탭 index가 매번 다른 탭을 가리킴 / 같은 차트인데 탭 2개로 보임 | `tab switch --index N`의 index는 안정적인 식별자가 아님 — 매번 `tab switch` 직후 `state`(심볼)와 `draw list`(도형 수)로 재확인 후 그리기. 두 CDP 타겟이 URL의 같은 chart_id를 가리키면 실제로 같은 화면일 수 있음 |
| NH선물 체결시간 누락 (01A101만 받았을 때) | NH는 하루에 `01A101`(시간없음)/`01A103`(체결시간 포함, "국문가정산(체결시분)" 메일)/`02A101`(예탁자산현황, 무관) 3종을 보냄. 항상 `01A103` 우선 선택 — `_dedup_nh()` 류 로직에 이 우선순위가 반대로 들어가 있던 적이 있으니 직접 확인할 것 |
| 일중 고저 기준 진입/청산 등급(A/B/C) 계산이 이상하게 나옴 | OHLCV를 기본 100봉 그대로 쓰면 여러 날짜가 섞여 고저가 부정확함 — 등급 계산 전에 반드시 **그날 UTC 윈도우로 필터링** 후 high/low 계산 |
| `refresh_excel.ps1`가 성공 로그를 남기는데 INFOMAX 데이터(`_xll.IMDH(...)`)가 실제로는 갱신 안 됨 | `New-Object -ComObject Excel.Application`(DCOM)으로 띄우면 `Visible=$true`여도 화면에 안 보이는 창 스테이션에서 실행됨 — INFOMAX 애드인이 실제 렌더링된 창이 있어야 데이터를 다시 받아옴. `Start-Process`로 첫 파일을 직접 열어 진짜 보이는 Excel을 띄운 뒤 `[Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")`로 그 인스턴스에 연결해서 나머지를 처리해야 함. 재계산 강제(`CalculateFullRebuild`)는 효과 없음 — 원인이 재계산이 아니라 창 가시성이기 때문 |
| PowerShell `.ps1` 파일의 한글 경로/문자열이 깨져서 `Out-File`/경로 관련 에러 발생 | 파일이 BOM 없이 저장되면 Windows PowerShell 5.1이 시스템 ANSI 코드페이지로 읽어서 한글이 깨짐. 에디터/툴로 `.ps1`을 다시 쓴 뒤에는 항상 UTF-8 BOM으로 재저장할 것 (`[System.IO.File]::WriteAllText(path, content, [Text.Encoding]::UTF8)` 또는 Python `open(path, 'w', encoding='utf-8-sig')`) |
| Gmail 거래증적 파이프라인이 며칠째 조용히 멈춰있음 | Google Testing 앱은 OAuth refresh token이 **7일**만에 만료됨. `creds.refresh()`가 예외를 던지면 잡아주지 않으면 매일 작업이 조용히 죽음 — refresh 실패 시 `flow.run_local_server()` 재인증으로 폴백하도록 try/except 추가. 며칠째 갱신이 안 된다면 이게 1순위 의심 대상 |
| 며칠 치를 한 번에 backfill했더니 날짜가 섞여서 엉뚱한 시트에 저장됨 | Gmail `after:X` 검색은 X 이후 전체를 반환함(X일만이 아님) — 첨부파일 저장 시 검색 기준일(target_date)을 그대로 파일명에 echo하면 실제 메일 날짜와 달라짐. 메일의 실제 `Date` 헤더로 파싱해서 사용하고, Excel 저장도 입력받은 날짜 하나가 아니라 **파싱된 각 PDF의 실제 거래일별로 그룹화**해서 호출할 것 |
| NH선물 01A101/02A101이 계속 같이 받아짐 | 다운로드 후 dedup 말고, Gmail 검색 쿼리 자체에 `subject:"국문가정산(체결시분)"`을 추가해서 처음부터 01A103만 받도록 좁히는 게 더 깔끔함 |
| SS선물 PDF가 "체결 0건"으로 파싱됨 | 버그가 아닐 수 있음 — SS는 거래 없었던 날 오후에 `[거래내역]`만 있고 `[체결내역]` 섹션 자체가 없는 정보성 메일을 보냄. 둘 다(NH/SS) 거래 없는 날은 메일 자체가 안 오는 게 정상 |

---

## 11. 다른 PC 셋업 체크리스트 (2026-06-22 정리)

새 PC에서 이 워크플로우를 다시 셋업할 때 위 "알려진 이슈"들을 순서대로 만나지 않으려면:

1. TradingView Desktop을 Claude Code(또는 VSCode 확장)가 띄운 셸에서 실행하기 전에 `ELECTRON_RUN_AS_NODE` 환경변수를 지운다.
2. `tradingview-mcp`를 새로 클론했다면 `src/core/drawing.js`의 4개 함수(`clearAll`/`listDrawings`/`getProperties`/`removeOne`) 패치를 다시 적용한다(원본 저장소 버그, 매번 재발).
3. KRX KTB 선물 정규장은 실제로 08:45 KST부터 봉이 시작한다(09:00은 시간축 그리드 라벨일 뿐).
4. NH선물 메일은 `01A103`(체결시분 포함) 첨부를 우선 선택하도록 dedup 로직을 확인한다.
5. 일중 등급(A/B/C) 계산 전에는 항상 그날 하루치로 필터링한 OHLCV에서 high/low를 다시 계산한다.
6. 차트 라벨은 당일 고저폭의 10~15% 수준 오프셋만 사용하고, 장초반 라벨은 지표 범례에 가려질 수 있음을 감안한다.
7. CDP 탭은 index로 믿지 말고 `tab switch` 후 `state`/`draw list`로 매번 재확인한다.
8. INFOMAX 연동 Excel 자동갱신은 `New-Object -ComObject`로 띄우지 말고, `시계열/refresh_excel.ps1`처럼 `Start-Process` + `GetActiveObject` 방식을 사용한다(보이지 않는 창에서는 데이터가 갱신되지 않음).

---

*작성: 2026-06-15 | Claude Code (claude-sonnet-4-6) 기반*
