# One-time setup of the costing dashboard on the Windows server. Run through Setup.bat (asks for administrator rights).
# Needs Python 3.10 or newer (64-bit) from python.org. Libraries download from the internet.
$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
$root = (Resolve-Path (Join-Path $here '..')).Path          # the costing folder
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $p = Start-Process powershell.exe -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File',('"' + $PSCommandPath + '"')) -Verb RunAs -Wait -PassThru
    exit $p.ExitCode
}
New-Item -ItemType Directory -Path (Join-Path $root 'logs') -Force | Out-Null
Start-Transcript -Path (Join-Path $root 'logs\setup.log') -Append | Out-Null
try {
    if ($root.StartsWith('\\')) { throw 'Copy the costing folder to a local disk first (not a network share).' }

    # 1. Python
    $py = $null
    foreach ($c in @('py -3.13', 'py -3.12', 'py -3.11', 'py -3.10', 'python')) {
        try { $v = & cmd /c "$c -c ""import sys;print(sys.version_info>=(3,10) and sys.maxsize>2**32)""" 2>$null; if ($v -eq 'True') { $py = $c; break } } catch {}
    }
    if (-not $py) { throw 'Python 3.10 or newer (64-bit) is needed. Install it from python.org/downloads/windows, then run Setup again.' }

    $venv = Join-Path $root 'venv'
    $vpy = Join-Path $venv 'Scripts\python.exe'
    if (-not (Test-Path $vpy)) {
        Write-Host 'Making the private Python environment...'
        & cmd /c "$py -m venv ""$venv"""
        if ($LASTEXITCODE -ne 0) { throw 'Could not create the Python environment.' }
    }
    Write-Host 'Installing libraries (downloading)...'
    & $vpy -m pip install -r (Join-Path $root 'requirements.txt') -q --disable-pip-version-check
    if ($LASTEXITCODE -ne 0) { throw 'Installing libraries failed. See logs\setup.log.' }

    # 2. Where the databases are
    $configDir = Join-Path $root 'config'
    New-Item -ItemType Directory -Path $configDir -Force | Out-Null
    $settingsFile = Join-Path $configDir 'settings.json'
    if (-not (Test-Path $settingsFile)) {
        $default = Join-Path $root 'DATABASE'
        Write-Host ''
        Write-Host 'Where are the database (.db) files?' -ForegroundColor Yellow
        Write-Host "Press Enter for $default (copy the .db files from the laptop's DATABASE folder there),"
        Write-Host 'or type another folder, for example D:\Unitech Dashboard Database\DATABASE'
        $db = (Read-Host 'Database folder').Trim().Trim('"')
        if (-not $db) { $db = $default }
        New-Item -ItemType Directory -Path $db -Force | Out-Null
        $json = @{ database_dir = $db; port = 5000; host = '0.0.0.0'; auto_sync = $true; backup_keep_days = 7 } | ConvertTo-Json
        [IO.File]::WriteAllText($settingsFile, $json)          # UTF-8 without BOM
        Write-Host "Saved $settingsFile"
    }

    # 3. ERP login and dashboard password
    if (-not (Test-Path (Join-Path $configDir 'erp_login.dat'))) {
        Write-Host ''
        & $vpy (Join-Path $here 'set_erp_login.py')
    }
    $hasPassword = & $vpy -c "import sys; sys.path.insert(0, r'$root\SCRIPTS'); import costing_config as c; print(bool(c.load_settings().get('password_hash')))"
    if ($hasPassword -ne 'True') {
        Write-Host ''
        & $vpy (Join-Path $here 'set_password.py')
    }

    if ((Get-Date).ToString('zzz') -ne '+05:30') { Write-Host 'WARNING: this PC is not on India Standard Time. The sync times follow the PC clock.' -ForegroundColor Yellow }

    # 4. Start at boot, restart after a failure
    $task = 'Unitech Costing'
    if (Get-ScheduledTask -TaskName $task -ErrorAction SilentlyContinue) { Stop-ScheduledTask -TaskName $task; Start-Sleep 2 }
    $action = New-ScheduledTaskAction -Execute $vpy -Argument ('-u "' + (Join-Path $here 'run_unattended.py') + '"') -WorkingDirectory $root
    $trigger = New-ScheduledTaskTrigger -AtStartup
    $set = New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1) -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew
    $who = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    Register-ScheduledTask -TaskName $task -Action $action -Trigger $trigger -Settings $set -Principal $who -Description 'Unitech costing dashboard: starts before sign-in, runs the ERP syncs, recovers from failures.' -Force | Out-Null

    $port = & $vpy -c "import sys; sys.path.insert(0, r'$root\SCRIPTS'); import costing_config as c; print(c.load_settings().get('port', 5000))"
    Remove-NetFirewallRule -DisplayName $task -ErrorAction SilentlyContinue
    New-NetFirewallRule -DisplayName $task -Direction Inbound -Action Allow -Protocol TCP -LocalPort ([int]$port) -RemoteAddress @('10.0.0.0/8','172.16.0.0/12','192.168.0.0/16','100.64.0.0/10') -Profile Any | Out-Null
    & powercfg.exe /change standby-timeout-ac 0 | Out-Null
    & powercfg.exe /change hibernate-timeout-ac 0 | Out-Null

    Start-ScheduledTask -TaskName $task
    $ok = $false
    for ($i = 0; $i -lt 120; $i++) { try { if ((Invoke-RestMethod "http://127.0.0.1:$port/health" -TimeoutSec 2).ok) { $ok = $true; break } } catch {}; Start-Sleep 1 }
    if (-not $ok) { throw 'The dashboard did not start. See logs\dashboard.log.' }
    Write-Host ''
    Write-Host "Setup complete. Open http://localhost:$port/ here, or http://<this PC's address>:$port/ from another PC on the plant network." -ForegroundColor Green
} catch {
    Write-Host ('SETUP FAILED: ' + $_.Exception.Message) -ForegroundColor Red
    Stop-Transcript | Out-Null
    exit 1
}
Stop-Transcript | Out-Null
