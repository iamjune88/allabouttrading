# Trading Journal 자동 실행 스크립트
# - 17:00 정기 실행 또는 로그온 시 미업데이트분 일괄 처리
$journalDir  = "C:\Users\iamju\Desktop\workingwithC\trading_journal"
$python      = "C:\Users\iamju\AppData\Local\Programs\Python\Python312\python.exe"
$lastRunFile = "$journalDir\last_run.txt"
$logFile     = "$journalDir\logs\scheduler.log"

New-Item -ItemType Directory -Force "$journalDir\logs" | Out-Null

function Write-Log($msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$ts  $msg" | Out-File -Append -Encoding utf8 $logFile
}

# 마지막 실행일 기준 --after 결정 (하루 겹치게 해서 누락 방지)
if (Test-Path $lastRunFile) {
    $lastRun   = (Get-Content $lastRunFile -Raw).Trim()
    $afterDate = [datetime]::ParseExact($lastRun, "yyyy/MM/dd", $null).AddDays(-1).ToString("yyyy/MM/dd")
} else {
    # 최초 실행 시 7일치 소급
    $afterDate = (Get-Date).AddDays(-7).ToString("yyyy/MM/dd")
}

Write-Log "=== START (after=$afterDate) ==="

$output = & $python "$journalDir\daily_run.py" --after $afterDate 2>&1
$output | ForEach-Object { Write-Log $_ }

if ($LASTEXITCODE -eq 0) {
    Get-Date -Format "yyyy/MM/dd" | Out-File -Encoding utf8 $lastRunFile
    Write-Log "=== DONE ==="
} else {
    Write-Log "=== ERROR (exit code $LASTEXITCODE) ==="
}
