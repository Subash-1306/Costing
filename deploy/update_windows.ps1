# Bring the costing dashboard up to date from GitHub, then restart it. Runs every night at 3 AM (task
# "Unitech Costing Update", set up by Install-AutoUpdate.bat) and from Update-Now.bat.
#   - nothing new on GitHub: nothing happens, the dashboard keeps running
#   - something new: download it, install any new libraries, restart, and check the dashboard answers
#   - the new version does not start within 2 minutes: go back to the version that was running, restart that
# config (settings, ERP login), the databases, backups and logs are never touched (they are not on GitHub).
# Log: logs\update.log
# (Continue, not Stop: Windows PowerShell 5.1 turns any git/pip message on stderr into an error; exit codes are checked)
$ErrorActionPreference = 'Continue'
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $p = Start-Process powershell.exe -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File',('"' + $PSCommandPath + '"')) -Verb RunAs -Wait -PassThru
    exit $p.ExitCode
}
$here = $PSScriptRoot
$root = (Resolve-Path (Join-Path $here '..')).Path
New-Item -ItemType Directory -Path (Join-Path $root 'logs') -Force | Out-Null
$logFile = Join-Path $root 'logs\update.log'
function Log($m) { $l = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '  ' + $m; Add-Content -Path $logFile -Value $l -Encoding UTF8; Write-Host $l }

function Find-Git {
    $g = Get-Command git.exe -ErrorAction SilentlyContinue
    if ($g) { return $g.Source }
    foreach ($p in @("$env:ProgramFiles\Git\cmd\git.exe", "${env:ProgramFiles(x86)}\Git\cmd\git.exe")) { if (Test-Path $p) { return $p } }
    throw 'Git for Windows is not installed (git-scm.com).'
}

function Setting($name, $default) {
    $f = Join-Path $root 'config\settings.json'
    if (Test-Path $f) {
        try { $v = (Get-Content $f -Raw | ConvertFrom-Json).$name; if ($v) { return $v } } catch {}
    }
    return $default
}

$task = 'Unitech Costing'
$port = Setting 'port' 5000
$branch = Setting 'update_branch' 'main'
$vpy = Join-Path $root 'venv\Scripts\python.exe'
$req = Join-Path $root 'requirements.txt'

function Install-Libraries {
    if (-not (Test-Path $vpy)) { Log '  No venv here (Setup.bat not run): using the PC''s own Python, libraries not checked.'; return $true }
    if (Get-ChildItem (Join-Path $here 'wheels') -Filter *.whl -ErrorAction SilentlyContinue) {
        & $vpy -m pip install --no-index --find-links (Join-Path $here 'wheels') -r $req -q 2>&1 | ForEach-Object { Log "  pip: $_" }
    } else {
        & $vpy -m pip install -r $req -q --disable-pip-version-check 2>&1 | ForEach-Object { Log "  pip: $_" }
    }
    return ($LASTEXITCODE -eq 0)
}

function Restart-Costing {
    $t = Get-ScheduledTask -TaskName $task -ErrorAction SilentlyContinue
    if (-not $t) { throw "The scheduled task '$task' is missing. The files are updated; restart the dashboard yourself (or run Setup.bat once)." }
    $exe = $t.Actions[0].Execute
    Stop-ScheduledTask -TaskName $task -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Get-CimInstance Win32_Process | Where-Object { $_.ExecutablePath -eq $exe } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    Start-Sleep -Seconds 1
    Start-ScheduledTask -TaskName $task -ErrorAction Stop
}

function Wait-Healthy {
    for ($i = 0; $i -lt 120; $i++) {
        try { if ((Invoke-RestMethod "http://127.0.0.1:$port/health" -TimeoutSec 3).ok) { return $true } } catch {}
        Start-Sleep -Seconds 1
    }
    return $false
}

try {
    $git = Find-Git
    Set-Location $root
    & $git fetch --quiet origin $branch 2>&1 | ForEach-Object { Log "  git: $_" }
    if ($LASTEXITCODE -ne 0) { throw 'Could not reach GitHub. The dashboard keeps running the version it has.' }
    $old = (& $git rev-parse HEAD).Trim()
    $new = (& $git rev-parse "origin/$branch").Trim()
    if ($old -eq $new) { Log ("Up to date (" + $old.Substring(0, 7) + "). Nothing to do."); exit 0 }

    Log ("Updating " + $old.Substring(0, 7) + " -> " + $new.Substring(0, 7) + ":")
    & $git log --oneline --no-decorate "$old..$new" | Select-Object -First 20 | ForEach-Object { Log "  $_" }
    & $git reset --hard --quiet $new
    if ($LASTEXITCODE -ne 0) { throw 'Could not switch to the new version.' }
    if (-not (Get-ScheduledTask -TaskName $task -ErrorAction SilentlyContinue)) {
        Install-Libraries | Out-Null
        Log ("Files updated to " + $new.Substring(0, 7) + ". The dashboard is not running as the '$task' task, so close its window and start it again to use the new version.")
        exit 0
    }
    if (-not (Install-Libraries)) { Log 'Installing libraries failed, going back.'; $bad = $true }
    else {
        Restart-Costing
        $bad = -not (Wait-Healthy)
        if ($bad) { Log 'The new version did not start within 2 minutes, going back. See logs\dashboard.log.' }
    }
    if ($bad) {
        & $git reset --hard --quiet $old
        Install-Libraries | Out-Null
        Restart-Costing
        if (Wait-Healthy) { Log ("Back on " + $old.Substring(0, 7) + ". The dashboard is running.") }
        else { Log 'Going back did not help either: the dashboard is not answering. Check logs\dashboard.log.' }
        exit 1
    }
    Log ("Done. Running " + $new.Substring(0, 7) + ".")
    exit 0
} catch {
    Log ('UPDATE FAILED: ' + $_.Exception.Message)
    exit 1
}
