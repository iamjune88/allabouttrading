# daily_pl_update.ps1
# 하루 한 번 실행: PDF 다운로드 → Claude Code 파싱 → Excel DB 업데이트
# 사용법: .\daily_pl_update.ps1 [-Date "2026-06-16"] [-PdfPath "C:\...\file.pdf"]

param(
    [string]$Date    = (Get-Date -Format "yyyy-MM-dd"),
    [string]$PdfPath = "",   # 직접 PDF 경로 지정 시 Gmail 다운로드 생략
    [string]$WorkDir = "C:\Users\iamju\Desktop\workingwithC"
)

$settleDir = Join-Path $WorkDir "settlements"
$dbScript  = Join-Path $WorkDir "trading_db_update.ps1"
$fetchScript = Join-Path $WorkDir "gmail_fetch_settlement.ps1"

Write-Host "===== $Date 일일 P&L 업데이트 ====="

# ── Step 1: PDF 확보 ─────────────────────────────────────
if ($PdfPath -ne "" -and (Test-Path $PdfPath)) {
    Write-Host "[1] PDF 지정 경로 사용: $PdfPath"
} else {
    Write-Host "[1] Gmail에서 가정산보고서 다운로드..."
    $downloaded = & $fetchScript -Date $Date
    if ($downloaded -and $downloaded.Count -gt 0) {
        # 가장 최신 파일 선택
        $PdfPath = $downloaded | Select-Object -Last 1
        Write-Host "    사용: $PdfPath"
    } else {
        Write-Host "    PDF 없음. 직접 경로를 -PdfPath 파라미터로 지정하거나"
        Write-Host "    settlements\ 폴더에 파일을 놓고 다시 실행하세요."
        exit 1
    }
}

# ── Step 2: JSON 파일 경로 ────────────────────────────────
$dateStr  = $Date -replace '-', ''
$jsonPath = Join-Path $WorkDir "parsed_${dateStr}.json"

Write-Host "[2] PDF 경로: $PdfPath"
Write-Host "    → Claude Code로 파싱 후 JSON 저장: $jsonPath"
Write-Host ""
Write-Host "*** 아래 명령을 Claude Code 채팅에 붙여넣기 ***"
Write-Host "--------------------------------------------"
Write-Host "가정산보고서 PDF 파싱해줘: $PdfPath"
Write-Host "파싱 결과를 $jsonPath 에 저장하고"
Write-Host "trading_db_update.ps1 실행해줘"
Write-Host "--------------------------------------------"
Write-Host ""

# ── Step 3: JSON 있으면 바로 DB 업데이트 ─────────────────
if (Test-Path $jsonPath) {
    Write-Host "[3] JSON 발견, DB 업데이트 실행..."
    $jsonData = Get-Content $jsonPath -Raw -Encoding UTF8
    & $dbScript -JsonData $jsonData
} else {
    Write-Host "[3] JSON 대기 중... Claude Code 파싱 완료 후 아래 실행:"
    Write-Host "    .\trading_db_update.ps1 -JsonData (Get-Content $jsonPath -Raw)"
}
