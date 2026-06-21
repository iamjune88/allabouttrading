# trading_db_update.ps1
# KTB 선물 일일 P&L DB 업데이트 스크립트 (PowerShell Excel COM)
# 사용법: .\trading_db_update.ps1 -JsonData '{...}'

param(
    [string]$JsonData = "",
    [string]$DbPath = "C:\Users\iamju\Desktop\workingwithC\trading_db.xlsx"
)

# ============================================================
# 색상 변환 (hex RGB -> Excel BGR long)
# ============================================================
function To-ExcelColor([string]$hex) {
    $r = [Convert]::ToInt32($hex.Substring(0,2), 16)
    $g = [Convert]::ToInt32($hex.Substring(2,2), 16)
    $b = [Convert]::ToInt32($hex.Substring(4,2), 16)
    return [long]($r + $g * 256 + $b * 65536)
}

$C_HEADER   = To-ExcelColor "1F4E79"  # 진한 파랑 헤더
$C_SELL     = To-ExcelColor "FFE0E0"  # 연빨강 (매도)
$C_BUY      = To-ExcelColor "E0FFE0"  # 연초록 (매수)
$C_POSITIVE = To-ExcelColor "C6EFCE"  # 수익 초록
$C_NEGATIVE = To-ExcelColor "FFC7CE"  # 손실 빨강
$C_WHITE    = [long]16777215           # 흰색 폰트 (0xFFFFFF)

# ============================================================
# 헤더 행 작성 공통 함수
# ============================================================
function Write-Headers {
    param($ws, [string[]]$headers, [double]$colWidth = 13)
    for ($c = 1; $c -le $headers.Count; $c++) {
        $cell = $ws.Cells.Item(1, $c)
        $cell.Value2 = $headers[$c-1]
        $cell.Font.Bold = $true
        $cell.Font.Color = $C_WHITE
        $cell.Interior.Color = $C_HEADER
        $cell.ColumnWidth = $colWidth
    }
}

# ============================================================
# 다음 빈 행 인덱스
# ============================================================
function Get-NextRow([object]$ws) {
    $row = 2
    while ($ws.Cells.Item($row, 1).Value2 -ne $null -and
           $ws.Cells.Item($row, 1).Value2 -ne "") {
        $row++
    }
    return $row
}

# ============================================================
# 행 전체 배경색
# ============================================================
function Set-RowColor {
    param($ws, [int]$row, [int]$colCount, [long]$color)
    for ($c = 1; $c -le $colCount; $c++) {
        $ws.Cells.Item($row, $c).Interior.Color = $color
    }
}

# ============================================================
# DB 파일 최초 생성
# ============================================================
function Create-TradingDB {
    param($excel, [string]$path)
    $wb = $excel.Workbooks.Add()

    while ($wb.Worksheets.Count -lt 4) { $wb.Worksheets.Add() | Out-Null }
    while ($wb.Worksheets.Count -gt 4) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }

    $names = @("일별PL", "거래내역", "포지션스냅샷", "포지션PL")
    for ($i = 1; $i -le 4; $i++) { $wb.Worksheets.Item($i).Name = $names[$i-1] }

    Write-Headers $wb.Worksheets.Item("일별PL") @(
        "날짜","선물정산차금","옵션매매대금","수수료","제세금","순손익","누적손익","예탁총액","비고"
    ) 14

    Write-Headers $wb.Worksheets.Item("거래내역") @(
        "날짜","종목코드","종목명","구분","수량","체결가","정산가","약정금액","수수료","당일차금","메모"
    )

    Write-Headers $wb.Worksheets.Item("포지션스냅샷") @(
        "날짜","종목코드","종목명","구분","잔고","전일잔고","전일정산가","당일정산가","갱신차금","미실현손익"
    )

    Write-Headers $wb.Worksheets.Item("포지션PL") @(
        "날짜","종목코드","종목명","구분","잔고","진입avg","당일정산가","미실현손익","당일실현손익","누적실현손익"
    )

    $wb.SaveAs($path)
    Write-Host "DB 생성: $path"
    return $wb
}

# ============================================================
# 데이터 업데이트
# ============================================================
function Update-TradingDB {
    param($wb, $data)

    $date      = $data.date
    $trades    = $data.trades
    $positions = $data.positions
    $pl        = $data.daily_pl

    # ── 거래내역 ─────────────────────────────────────────────
    $ws = $wb.Worksheets.Item("거래내역")
    foreach ($t in $trades) {
        $row = Get-NextRow $ws
        $ws.Cells.Item($row, 1).Value2  = $date
        $ws.Cells.Item($row, 2).Value2  = $t.code
        $ws.Cells.Item($row, 3).Value2  = $t.name
        $ws.Cells.Item($row, 4).Value2  = $t.side
        $ws.Cells.Item($row, 5).Value2  = [double]$t.qty
        $ws.Cells.Item($row, 6).Value2  = [double]$t.price
        $ws.Cells.Item($row, 7).Value2  = [double]$t.settlement
        $ws.Cells.Item($row, 8).Value2  = [double]$t.notional
        $ws.Cells.Item($row, 9).Value2  = [double]$t.fee
        $ws.Cells.Item($row, 10).Value2 = [double]$t.daily_pnl
        $ws.Cells.Item($row, 11).Value2 = if ($t.memo) { $t.memo } else { "" }
        $clr = if ($t.side -eq "매도") { $C_SELL } else { $C_BUY }
        Set-RowColor $ws $row 11 $clr
    }

    # ── 포지션스냅샷 ─────────────────────────────────────────
    $ws = $wb.Worksheets.Item("포지션스냅샷")
    foreach ($p in $positions) {
        $row = Get-NextRow $ws
        $ws.Cells.Item($row, 1).Value2  = $date
        $ws.Cells.Item($row, 2).Value2  = $p.code
        $ws.Cells.Item($row, 3).Value2  = $p.name
        $ws.Cells.Item($row, 4).Value2  = $p.side
        $ws.Cells.Item($row, 5).Value2  = [double]$p.qty
        $ws.Cells.Item($row, 6).Value2  = [double]$p.prev_qty
        $ws.Cells.Item($row, 7).Value2  = [double]$p.prev_settlement
        $ws.Cells.Item($row, 8).Value2  = [double]$p.curr_settlement
        $ws.Cells.Item($row, 9).Value2  = [double]$p.mtm_pnl
        $ws.Cells.Item($row, 10).Value2 = [double]$p.mtm_pnl
        $clr = if ($p.mtm_pnl -gt 0) { $C_POSITIVE } else { $C_NEGATIVE }
        Set-RowColor $ws $row 10 $clr
    }

    # ── 일별PL ───────────────────────────────────────────────
    $ws  = $wb.Worksheets.Item("일별PL")
    $row = Get-NextRow $ws

    $fut = [double]$pl.futures_settlement
    $opt = [double]$pl.option_pnl
    $fee = [double]$pl.fee
    $tax = [double]$pl.tax
    $net = $fut + $opt - $fee - $tax

    $cum = [double]0
    if ($row -gt 2) {
        $prev = $ws.Cells.Item($row - 1, 7).Value2
        if ($prev -ne $null) { $cum = [double]$prev }
    }
    $cum += $net

    $ws.Cells.Item($row, 1).Value2 = $date
    $ws.Cells.Item($row, 2).Value2 = $fut
    $ws.Cells.Item($row, 3).Value2 = $opt
    $ws.Cells.Item($row, 4).Value2 = $fee
    $ws.Cells.Item($row, 5).Value2 = $tax
    $ws.Cells.Item($row, 6).Value2 = $net
    $ws.Cells.Item($row, 7).Value2 = $cum
    $ws.Cells.Item($row, 8).Value2 = if ($pl.total_deposit) { [double]$pl.total_deposit } else { 0 }
    $ws.Cells.Item($row, 9).Value2 = if ($pl.memo) { $pl.memo } else { "" }

    $clr = if ($net -ge 0) { $C_POSITIVE } else { $C_NEGATIVE }
    Set-RowColor $ws $row 9 $clr

    $netFmt = "{0:N0}" -f $net
    $cumFmt = "{0:N0}" -f $cum
    Write-Host "$date 입력 완료 | 순손익: ${netFmt}원 | 누적: ${cumFmt}원"
}

# ============================================================
# 엔트리포인트
# ============================================================
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    if (-not (Test-Path $DbPath)) {
        Write-Host "DB 파일 없음 - 신규 생성"
        $wb = Create-TradingDB $excel $DbPath
    } else {
        $wb = $excel.Workbooks.Open($DbPath)
    }

    if ($JsonData -ne "") {
        $data = $JsonData | ConvertFrom-Json
        Update-TradingDB $wb $data
        $wb.Save()
        Write-Host "저장 완료: $DbPath"
    } else {
        Write-Host "DB 준비 완료 (데이터 없음)"
    }

    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
