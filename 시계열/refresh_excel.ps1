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

# Zone.Identifier(인터넷/네트워크 출처 표시)가 남아있으면 Excel이 "보호된 보기"로
# 열어서 편집 사용을 누르기 전까지 INFOMAX 매크로/링크가 동작하지 않음
# (2026-06-22 확인됨). 매 실행마다 미리 차단 해제.
$files | Unblock-File

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

# 이름 매칭 대신 Workbooks.Count를 직접 폴링 - 인코딩/타이밍 차이로 이름이
# 안 맞아서 첫 파일이 조용히 누락되는 문제가 있었음 (2026-06-22 확인됨,
# refresh_log.txt에 Saved+closed 기록이 빠짐 + LastWriteTime 안 바뀜으로 발견).
$wb1 = $null
for ($i = 0; $i -lt 10; $i++) {
    if ($excel.Workbooks.Count -ge 1) {
        $wb1 = $excel.Workbooks.Item(1)
        break
    }
    Start-Sleep -Seconds 1
}
if ($wb1) {
    Write-Log "Attached to first workbook: $($wb1.Name)"
    $openedWorkbooks += $wb1
} else {
    Write-Log "ERROR: $($firstFile.Name) never appeared in Workbooks collection"
}

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
# 참고: 2026-06-22 테스트에서 Quit() 직후 한두 차례 EXCEL.EXE가 빈 창으로
# 잠깐 남는 걸 본 적 있음 (매크로 포함 파일 Save/Close 에러 직후). 별도 조치 없이도
# 위 GC 호출 후 몇 초 내 프로세스가 정상 종료됐음 - COM RCW 해제 지연일 뿐, 재발 시
# 무시해도 됨 (강제 종료 로직은 일부러 안 넣음 - 사용자가 따로 띄운 Excel을 죽일 위험).
Write-Log "=== REFRESH DONE ==="
