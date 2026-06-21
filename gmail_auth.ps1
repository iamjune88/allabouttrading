# gmail_auth.ps1
# Gmail OAuth 2.0 인증 (최초 1회 실행)
# 이후 토큰은 gmail_token.json에 캐싱됨

param(
    [string]$AuthCode  = "",
    [string]$CredFile  = "C:\Users\iamju\Desktop\workingwithC\gmail_credentials.json",
    [string]$TokenFile = "C:\Users\iamju\Desktop\workingwithC\gmail_token.json"
)

if (-not (Test-Path $CredFile)) {
    Write-Host "ERROR: $CredFile 없음"
    Write-Host ""
    Write-Host "=== Google Cloud Console 설정 방법 ==="
    Write-Host "1. https://console.cloud.google.com 접속"
    Write-Host "2. 새 프로젝트 생성 (예: trading-db)"
    Write-Host "3. API 및 서비스 > Gmail API 활성화"
    Write-Host "4. 사용자 인증 정보 > OAuth 2.0 클라이언트 ID 생성"
    Write-Host "   - 유형: 데스크톱 앱"
    Write-Host "5. JSON 다운로드 → 파일명을 gmail_credentials.json으로 변경"
    Write-Host "6. C:\Users\iamju\Desktop\workingwithC\ 에 저장"
    exit 1
}

$cred = Get-Content $CredFile | ConvertFrom-Json
$clientId     = $cred.installed.client_id
$clientSecret = $cred.installed.client_secret
$redirectUri  = "urn:ietf:wg:oauth:2.0:oob"
$scope        = "https://www.googleapis.com/auth/gmail.readonly"

# 인증 URL 생성
$authUrl = "https://accounts.google.com/o/oauth2/auth" +
    "?client_id=$clientId" +
    "&redirect_uri=$([Uri]::EscapeDataString($redirectUri))" +
    "&response_type=code" +
    "&scope=$([Uri]::EscapeDataString($scope))" +
    "&access_type=offline" +
    "&prompt=consent"

if ($AuthCode -eq "") {
    Write-Host "=== Step 1: 브라우저에서 아래 URL 열고 로그인 후 코드 복사 ==="
    Write-Host ""
    Write-Host $authUrl
    Write-Host ""
    Write-Host "=== Step 2: 표시된 코드를 -AuthCode 파라미터로 재실행 ==="
    Write-Host ".\gmail_auth.ps1 -AuthCode '여기에코드입력'"
    Start-Process $authUrl
    exit 0
}

$authCode = $AuthCode

# 토큰 교환
$body = @{
    code          = $authCode
    client_id     = $clientId
    client_secret = $clientSecret
    redirect_uri  = $redirectUri
    grant_type    = "authorization_code"
}

$resp = Invoke-RestMethod -Uri "https://oauth2.googleapis.com/token" -Method POST -Body $body
$resp | Add-Member -NotePropertyName "created_at" -NotePropertyValue ([DateTimeOffset]::UtcNow.ToUnixTimeSeconds())
$resp | ConvertTo-Json | Set-Content $TokenFile -Encoding UTF8
Write-Host "토큰 저장 완료: $TokenFile"
