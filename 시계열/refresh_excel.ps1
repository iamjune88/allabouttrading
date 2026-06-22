# Excel Data Refresh Script (with INFOMAX pre-check)
# NOTE: Excel must be launched via Start-Process (real window, visible desktop),
#       then attached to via GetActiveObject. Creating it with `New-Object -ComObject
#       Excel.Application` (DCOM out-of-process activation) starts Excel in a window
#       station the interactive desktop can't see, so the INFOMAX add-in (IMDH())
#       never actually refetches data even though the script reports success.
#       Confirmed 2026-06-22 — see TRADING_WORKFLOW_GUIDE.md known issues.
$targetFolder   = $PSScriptRoot
$waitSeconds    = 20
$logFile        = Join-Path $targetFolder "refresh_log.txt"
$infomaxLnk     = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\인포맥스\인포맥스.lnk"
$infomaxProcess = "infomaxmain"
$infomaxWait    = 60

function Write-Log {
    param([string]$msg)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$ts  $msg" | Out-File -Append -Encoding utf8 -FilePath $logFile
}

Write-Log "=== REFRESH START ==="

# 선결조건: INFOMAX 실행 여부 확인
$running = Get-Process -Name $infomaxProcess -ErrorAction SilentlyContinue
if (-not $running) {
    Write-Log "INFOMAX not running - launching via shortcut"
    Start-Process -FilePath $infomaxLnk
    Write-Log "Waiting ${infomaxWait}s for INFOMAX to initialize..."
    Start-Sleep -Seconds $infomaxWait

    $running = Get-Process -Name $infomaxProcess -ErrorAction SilentlyContinue
    if (-not $running) {
        Write-Log "ERROR: INFOMAX still not running after wait - aborting"
        exit 1
    }
    Write-Log "INFOMAX confirmed running"
} else {
    Write-Log "INFOMAX already running - OK"
}

# xlsx 파일 목록 (폴더 내 전체, 하드코딩 없음)
$files = Get-ChildItem -Path $targetFolder -Filter "*.xlsx" | Where-Object { -not $_.Name.StartsWith("~$") }

if ($files.Count -eq 0) {
    Write-Log "No xlsx files found"
    exit 0
}

# 첫 파일을 Start-Process로 열어 "진짜" 보이는 Excel 창을 생성
$firstFile = $files[0]
Write-Log "Launching visibly via Start-Process: $($firstFile.Name)"
Start-Process -FilePath $firstFile.FullName
Start-Sleep -Seconds 5

# 그 visible 인스턴스에 COM으로 연결
$excel = $null
for ($i = 0; $i -lt 10; $i++) {
    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        break
    } catch {
        Start-Sleep -Seconds 1
    }
}
if (-not $excel) {
    Write-Log "ERROR: could not attach to Excel via GetActiveObject"
    exit 1
}
$excel.DisplayAlerts    = $false
$excel.AskToUpdateLinks = $false

$openedWorkbooks = @()
$wb1 = $excel.Workbooks | Where-Object { $_.Name -eq $firstFile.Name }
if ($wb1) { $openedWorkbooks += $wb1 }

foreach ($file in $files | Select-Object -Skip 1) {
    try {
        Write-Log "Opening: $($file.Name)"
        $wb = $excel.Workbooks.Open($file.FullName, 2, $false)
        $openedWorkbooks += $wb
        Start-Sleep -Seconds 2
    } catch {
        Write-Log "Open failed ($($file.Name)): $_"
    }
}

Write-Log "$($openedWorkbooks.Count) file(s) open in visible Excel - waiting ${waitSeconds}s for IMDH() to refetch"
Start-Sleep -Seconds $waitSeconds

# 저장+닫기
foreach ($wb in $openedWorkbooks) {
    try {
        $name = $wb.Name
        $wb.Save()
        $wb.Close($false)
        Write-Log "Saved+closed: $name"
    } catch {
        Write-Log "Close failed: $_"
        try { $wb.Close($false) } catch {}
    }
}

try { $excel.Quit() } catch {}
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
Write-Log "=== REFRESH DONE ==="
