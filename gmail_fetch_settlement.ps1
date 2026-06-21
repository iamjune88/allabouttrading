# gmail_fetch_settlement.ps1
# 삼성선물 가정산보고서 PDF 자동 다운로드

param(
    [string]$Date        = (Get-Date -Format "yyyy-MM-dd"),
    [string]$CredFile    = "C:\Users\iamju\Desktop\workingwithC\gmail_credentials.json",
    [string]$TokenFile   = "C:\Users\iamju\Desktop\workingwithC\gmail_token.json",
    [string]$DownloadDir = "C:\Users\iamju\Desktop\workingwithC\settlements"
)

$BASE = "https://gmail.googleapis.com/gmail/v1/users/me"

# ── 토큰 로드 & 갱신 ─────────────────────────────────────
if (-not (Test-Path $TokenFile)) { Write-Host "토큰 없음. gmail_auth.ps1 먼저 실행"; exit 1 }

$token = Get-Content $TokenFile -Raw -Encoding UTF8 | ConvertFrom-Json
$cred  = Get-Content $CredFile  -Raw -Encoding UTF8 | ConvertFrom-Json

$elapsed = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds() - [long]$token.created_at
if ($elapsed -ge ([long]$token.expires_in - 60)) {
    Write-Host "토큰 갱신..."
    $body = "client_id=$($cred.installed.client_id)&client_secret=$($cred.installed.client_secret)&refresh_token=$($token.refresh_token)&grant_type=refresh_token"
    $resp = Invoke-RestMethod -Uri "https://oauth2.googleapis.com/token" -Method POST -Body $body -ContentType "application/x-www-form-urlencoded"
    $token.access_token = $resp.access_token
    $token.expires_in   = $resp.expires_in
    $token.created_at   = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $token | ConvertTo-Json | Set-Content $TokenFile -Encoding UTF8
}

$headers = @{ Authorization = "Bearer $($token.access_token)" }

# ── 날짜 범위 계산 ────────────────────────────────────────
$dateObj     = [datetime]::ParseExact($Date, "yyyy-MM-dd", $null)
$afterEpoch  = [DateTimeOffset]::new($dateObj, [TimeSpan]::Zero).ToUnixTimeSeconds()
$beforeEpoch = [DateTimeOffset]::new($dateObj.AddDays(1), [TimeSpan]::Zero).ToUnixTimeSeconds()

if (-not (Test-Path $DownloadDir)) { New-Item -ItemType Directory -Path $DownloadDir | Out-Null }

# ── 메일 검색 (발신자별) ──────────────────────────────────
# 삼성선물: master@ssfutures.com
# NH선물:   NHfutures@futures.co.kr
$queries = @(
    "from:master@ssfutures.com has:attachment after:$afterEpoch before:$beforeEpoch",
    "from:NHfutures@futures.co.kr has:attachment after:$afterEpoch before:$beforeEpoch",
    "가정산보고서 has:attachment after:$afterEpoch before:$beforeEpoch"
)

$messages = @()
foreach ($q in $queries) {
    Write-Host "검색: $q"
    $qEnc    = [Uri]::EscapeDataString($q)
    $listUri = "${BASE}/messages?q=${qEnc}&maxResults=10"
    $listResp = Invoke-RestMethod -Uri $listUri -Headers $headers
    if ($listResp.messages -and $listResp.messages.Count -gt 0) {
        $messages += $listResp.messages
        Write-Host "  $($listResp.messages.Count)건 추가"
    }
}

# 중복 제거
$messages = $messages | Sort-Object id -Unique
Write-Host "총 $($messages.Count)건"

if (-not $messages) {
    Write-Host "메일 없음: $Date 가정산보고서 없음"
    exit 0
}

# ── 첨부파일 다운로드 ─────────────────────────────────────
$downloaded = @()

foreach ($msg in $messages) {
    $detailUri = "${BASE}/messages/$($msg.id)?format=full"
    $detail    = Invoke-RestMethod -Uri $detailUri -Headers $headers

    $fromHdr    = ($detail.payload.headers | Where-Object { $_.name -eq "From"    }).value
    $subjectHdr = ($detail.payload.headers | Where-Object { $_.name -eq "Subject" }).value
    Write-Host "  메일: $fromHdr | $subjectHdr"

    # 첨부파일 재귀 탐색
    function Find-Parts([object]$part) {
        $r = @()
        if ($part.filename -and $part.filename -match "\.pdf$" -and $part.body.attachmentId) { $r += $part }
        if ($part.parts) { foreach ($p in $part.parts) { $r += Find-Parts $p } }
        return $r
    }

    $atts = Find-Parts $detail.payload
    if ($atts.Count -eq 0) { Write-Host "  PDF 첨부 없음"; continue }

    foreach ($att in $atts) {
        $attUri  = "${BASE}/messages/$($msg.id)/attachments/$($att.body.attachmentId)"
        $attResp = Invoke-RestMethod -Uri $attUri -Headers $headers

        $dateStr = $Date -replace '-', ''
        $safeName = $att.filename -replace '[\\/:*?"<>|]', '_'
        $outPath = Join-Path $DownloadDir "${dateStr}_${safeName}"

        $dataFixed = $attResp.data -replace '-', '+' -replace '_', '/'
        $mod = $dataFixed.Length % 4
        if ($mod -eq 2) { $dataFixed += '==' }
        elseif ($mod -eq 3) { $dataFixed += '=' }
        $bytes = [Convert]::FromBase64String($dataFixed)
        [System.IO.File]::WriteAllBytes($outPath, $bytes)

        Write-Host "  저장: $outPath ($([math]::Round($bytes.Length/1024,1)) KB)"
        $downloaded += $outPath
    }
}

if ($downloaded.Count -gt 0) {
    Write-Host ""
    Write-Host "완료: $($downloaded.Count)개 다운로드"
    $downloaded | ForEach-Object { Write-Host "  $_" }
} else {
    Write-Host "다운로드 실패"
}

return $downloaded
