# Phase 2: open the costing dashboard over Tailscale (phones and laptops anywhere, no open ports).
# Run through Setup-Tailscale.bat (asks for administrator rights). Needs the Tailscale app installed and signed in.
#   - makes sure a dashboard password is set (everyone coming over Tailscale has to sign in)
#   - publishes the dashboard on this PC's private Tailscale HTTPS address (tailscale serve)
#   - allows the plain address http://<Tailscale IP>:<port> through the firewall as a fallback
# Undo: Switch-Phase.bat (back to Phase 1) runs "tailscale serve reset".
$ErrorActionPreference = 'Continue'
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $p = Start-Process powershell.exe -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File',('"' + $PSCommandPath + '"')) -Verb RunAs -Wait -PassThru
    exit $p.ExitCode
}
$here = $PSScriptRoot
$root = (Resolve-Path (Join-Path $here '..')).Path
New-Item -ItemType Directory -Path (Join-Path $root 'logs') -Force | Out-Null
Start-Transcript -Path (Join-Path $root 'logs\setup-tailscale.log') -Append | Out-Null
try {
    $ts = $null
    $c = Get-Command tailscale.exe -ErrorAction SilentlyContinue
    if ($c) { $ts = $c.Source } elseif (Test-Path "$env:ProgramFiles\Tailscale\tailscale.exe") { $ts = "$env:ProgramFiles\Tailscale\tailscale.exe" }
    if (-not $ts) {
        Start-Process 'https://tailscale.com/download/windows'
        throw 'Tailscale is not installed. Install it from the page that just opened, sign in with the company Tailscale account, then run this again.'
    }

    $null = & $ts status 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Host 'Tailscale is not signed in. A browser window opens: sign in with the company Tailscale account.' -ForegroundColor Yellow
        & $ts up --unattended
        $null = & $ts status 2>&1
        if ($LASTEXITCODE -ne 0) { throw 'Tailscale is still not connected. Open the Tailscale app, sign in, then run this again.' }
    }

    # 1. password first: the Tailscale relay comes from inside this PC, and the dashboard must still ask
    $vpy = Join-Path $root 'venv\Scripts\python.exe'
    if (-not (Test-Path $vpy)) { $vpy = 'python' }
    $hasPassword = & $vpy -c "import sys; sys.path.insert(0, r'$root\SCRIPTS'); import costing_config as c; print(bool(c.load_settings().get('password_hash')))"
    if ($hasPassword -ne 'True') {
        Write-Host 'A dashboard password is needed before it can be opened over Tailscale.' -ForegroundColor Yellow
        & $vpy (Join-Path $here 'set_password.py')
        $hasPassword = & $vpy -c "import sys; sys.path.insert(0, r'$root\SCRIPTS'); import costing_config as c; print(bool(c.load_settings().get('password_hash')))"
        if ($hasPassword -ne 'True') { throw 'No password was set, so Tailscale access was not switched on.' }
    }
    $port = & $vpy -c "import sys; sys.path.insert(0, r'$root\SCRIPTS'); import costing_config as c; print(c.load_settings().get('port', 5000))"

    # 2. fallback: plain http on the Tailscale address (firewall allows Tailscale's 100.64.0.0/10 range only)
    $rule = 'Unitech Costing (Tailscale)'
    Remove-NetFirewallRule -DisplayName $rule -ErrorAction SilentlyContinue
    New-NetFirewallRule -DisplayName $rule -Direction Inbound -Action Allow -Protocol TCP -LocalPort ([int]$port) -RemoteAddress '100.64.0.0/10' -Profile Any | Out-Null

    # 3. HTTPS on the Tailscale name (needs MagicDNS + HTTPS certificates switched on for the tailnet)
    $out = & $ts serve --bg $port 2>&1
    $serveOk = ($LASTEXITCODE -eq 0)
    $out | ForEach-Object { Write-Host "  tailscale: $_" }

    $ip = ((& $ts ip -4) | Select-Object -First 1)
    $name = ''
    try { $name = ((& $ts status --json | Out-String | ConvertFrom-Json).Self.DNSName).TrimEnd('.') } catch {}

    Write-Host ''
    Write-Host 'Tailscale access is on. From any phone or laptop signed in to the same Tailscale account:' -ForegroundColor Green
    if ($serveOk -and $name) { Write-Host "  https://$name/" -ForegroundColor Cyan }
    if ($ip) { Write-Host "  http://${ip}:$port/" -ForegroundColor Cyan }
    if (-not $serveOk) {
        Write-Host ''
        Write-Host 'The https address is not available yet: in the Tailscale admin console, open DNS and switch on' -ForegroundColor Yellow
        Write-Host 'MagicDNS and HTTPS Certificates, then run this again. The http address above works meanwhile.' -ForegroundColor Yellow
    }
} catch {
    Write-Host ('FAILED: ' + $_.Exception.Message) -ForegroundColor Red
    Stop-Transcript | Out-Null
    exit 1
}
Stop-Transcript | Out-Null
