# TradingView CDP Launcher — requires admin elevation
param([int]$Port = 9222)

# Auto-elevate if not admin
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $args = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Port $Port"
    Start-Process powershell -Verb RunAs -ArgumentList $args
    exit
}

# Find TradingView via AppxPackage
$pkg = Get-AppxPackage | Where-Object { $_.PackageFamilyName -like "*TradingView*" } | Select-Object -First 1
if (-not $pkg) {
    Write-Error "TradingView package not found. Make sure it is installed."
    pause; exit 1
}

$tvDir  = $pkg.InstallLocation
$tvExe  = Join-Path $tvDir "TradingView.exe"

if (-not (Test-Path $tvExe)) {
    # Some builds use a subdirectory — search one level deep
    $tvExe = Get-ChildItem $tvDir -Filter "TradingView.exe" -Recurse -Depth 2 -ErrorAction SilentlyContinue |
             Select-Object -First 1 -ExpandProperty FullName
}

if (-not $tvExe) {
    Write-Error "TradingView.exe not found inside $tvDir"
    pause; exit 1
}

Write-Host "Found: $tvExe"

# Kill existing TradingView if running
Get-Process -Name "TradingView" -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 1500

# Grant current user read+execute on the install dir (non-destructive — no takeown)
$user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
icacls $tvDir /grant "${user}:(OI)(CI)RX" /T /Q 2>&1 | Out-Null
Write-Host "Permissions granted for $user"

# Launch with CDP port
Write-Host "Launching TradingView with --remote-debugging-port=$Port ..."
Start-Process -FilePath $tvExe -ArgumentList "--remote-debugging-port=$Port"

# Wait for CDP to respond (up to 20s)
$ready = $false
for ($i = 0; $i -lt 20; $i++) {
    Start-Sleep -Seconds 1
    try {
        $r = Invoke-WebRequest -Uri "http://localhost:$Port/json/version" -UseBasicParsing -TimeoutSec 1 -ErrorAction Stop
        Write-Host "CDP ready: $($r.Content)"
        $ready = $true
        break
    } catch {}
    Write-Host "Waiting... ($($i+1)s)"
}

if ($ready) {
    Write-Host "`nTradingView is running with CDP on port $Port. You can now use tv_health_check in Claude Code."
} else {
    Write-Warning "TradingView launched but CDP not responding yet. Try tv_health_check in a few seconds."
}

pause
