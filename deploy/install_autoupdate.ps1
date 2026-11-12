# One-time: let this PC download costing dashboard updates from GitHub by itself, every night at 3 AM.
# Run through Install-AutoUpdate.bat (asks for administrator rights). Needs Git for Windows (git-scm.com).
#   1. makes a read-only key for this PC and shows it: you add it once on GitHub (Settings > Deploy keys)
#   2. links this folder to the GitHub repo (works on a folder copied from the zip too)
#   3. adds the scheduled task "Unitech Costing Update" (daily 03:00, runs update_windows.ps1)
# (GitHub allows one repository per deploy key, so this key is separate from the MES one.)
param([string]$Repo = 'Subash-1306/Costing', [string]$At = '03:00')
# Continue, not Stop: Windows PowerShell 5.1 turns git's messages on stderr into errors. Exit codes are checked.
$ErrorActionPreference = 'Continue'
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $p = Start-Process powershell.exe -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File',('"' + $PSCommandPath + '"')) -Verb RunAs -Wait -PassThru
    exit $p.ExitCode
}
$here = $PSScriptRoot
$root = (Resolve-Path (Join-Path $here '..')).Path
New-Item -ItemType Directory -Path (Join-Path $root 'logs') -Force | Out-Null
Start-Transcript -Path (Join-Path $root 'logs\install-autoupdate.log') -Append | Out-Null
try {
    $g = Get-Command git.exe -ErrorAction SilentlyContinue
    $git = if ($g) { $g.Source } elseif (Test-Path "$env:ProgramFiles\Git\cmd\git.exe") { "$env:ProgramFiles\Git\cmd\git.exe" } else { $null }
    if (-not $git) { throw 'Install Git for Windows first (git-scm.com, all the default options), then run this again.' }
    $gitRoot = Split-Path (Split-Path $git)
    $ssh = Join-Path $gitRoot 'usr\bin\ssh.exe'
    $keygen = Join-Path $gitRoot 'usr\bin\ssh-keygen.exe'
    if (-not (Test-Path $ssh)) { throw "Git's ssh.exe was not found at $ssh. Reinstall Git for Windows with the default options." }

    # 1. this PC's key (kept outside the program folder, readable by administrators and SYSTEM only)
    $dir = Join-Path $env:ProgramData 'UnitechCosting'
    New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null
    $key = Join-Path $dir 'github_deploy_key'
    if (-not (Test-Path $key)) {
        # Start-Process hands ssh-keygen this command line exactly as written (no cmd.exe, no PowerShell quote
        # rewriting), so the path with spaces and the empty passphrase ("") both arrive intact.
        $args_ = '-q -t ed25519 -N "" -C "unitech-costing ' + $env:COMPUTERNAME + '" -f "' + $key + '"'
        $kp = Start-Process -FilePath $keygen -ArgumentList $args_ -Wait -NoNewWindow -PassThru
        if ($kp.ExitCode -ne 0 -or -not (Test-Path "$key.pub")) { throw 'Could not make the key (ssh-keygen failed).' }
    }
    & icacls $dir /inheritance:r /grant:r 'SYSTEM:(OI)(CI)F' 'Administrators:(OI)(CI)F' | Out-Null
    $pub = (Get-Content "$key.pub" -Raw).Trim()
    try { Set-Clipboard -Value $pub } catch {}

    # Git for Windows runs this through its own shell, where its ssh is already on the PATH, so the command needs no
    # "C:\Program Files" path and no inner quotes (Windows PowerShell 5.1 mangles quotes inside native arguments).
    $k = $key -replace '\\', '/'
    $kh = (Join-Path $dir 'known_hosts') -replace '\\', '/'
    if ($k -match ' ') { throw "The key folder has a space in it ($k). ProgramData should not." }
    $sshCmd = "ssh -i $k -o IdentitiesOnly=yes -o StrictHostKeyChecking=accept-new -o UserKnownHostsFile=$kh"
    $url = "git@github.com:$Repo.git"

    # 2. link the folder to the repo
    Set-Location $root
    & $git config --system --add safe.directory ($root -replace '\\', '/')
    if (-not (Test-Path (Join-Path $root '.git'))) { & $git init -q -b main }
    & $git config core.sshCommand $sshCmd
    $has = (& $git remote) -contains 'origin'
    if ($has) { & $git remote set-url origin $url } else { & $git remote add origin $url }

    for ($try = 0; $try -lt 20; $try++) {
        $null = & $git ls-remote origin HEAD 2>&1
        if ($LASTEXITCODE -eq 0) { break }
        Write-Host ''
        Write-Host 'Add this PC on GitHub once (it can only download, never change anything):' -ForegroundColor Yellow
        Write-Host "  1. Open https://github.com/$Repo/settings/keys and click 'Add deploy key'"
        Write-Host "  2. Title: $env:COMPUTERNAME   Key: paste this (it is already copied):"
        Write-Host ''
        Write-Host "     $pub" -ForegroundColor Cyan
        Write-Host ''
        Write-Host "  3. Leave 'Allow write access' OFF and click 'Add key'"
        Read-Host 'Press Enter here once it is added'
    }
    if ($LASTEXITCODE -ne 0) { throw 'GitHub still says no. Check the deploy key, then run this again.' }
    Write-Host 'GitHub answers. Downloading the code...'
    & $git fetch -q origin main
    if ($LASTEXITCODE -ne 0) { throw 'Could not download from GitHub.' }
    $cur = (& $git rev-parse --verify -q HEAD) 2>$null
    if (-not $cur) {
        # a folder copied from the zip: take GitHub's version of every file (config, databases, logs stay as they are)
        & $git reset -q --hard origin/main
        if ($LASTEXITCODE -ne 0) { throw 'Could not take the GitHub version of the files.' }
        & $git branch -q --set-upstream-to=origin/main main
        Write-Host 'This folder now matches GitHub. Restarting the dashboard on that version...'
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $here 'restart_windows.ps1')
    }

    # 3. the nightly task
    $name = 'Unitech Costing Update'
    $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument ('-NoProfile -ExecutionPolicy Bypass -File "' + (Join-Path $here 'update_windows.ps1') + '"') -WorkingDirectory $root
    $trigger = New-ScheduledTaskTrigger -Daily -At $At
    $set = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 30) -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew
    $who = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    Register-ScheduledTask -ErrorAction Stop -TaskName $name -Action $action -Trigger $trigger -Settings $set -Principal $who -Description 'Unitech costing dashboard: download updates from GitHub every night and restart (goes back if the new version does not start).' -Force | Out-Null
    Write-Host ''
    Write-Host "Done. This PC checks GitHub every night at $At and updates itself. Log: logs\update.log" -ForegroundColor Green
    Write-Host 'To update straight away, run Update-Now.bat.'
} catch {
    Write-Host ('FAILED: ' + $_.Exception.Message) -ForegroundColor Red
    Stop-Transcript | Out-Null
    exit 1
}
Stop-Transcript | Out-Null
