# Restart the costing dashboard after changing settings, saving a new ERP login or copying in new program files.
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $p = Start-Process powershell.exe -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File',('"' + $PSCommandPath + '"')) -Verb RunAs -Wait -PassThru
    exit $p.ExitCode
}
$here = $PSScriptRoot
$root = (Resolve-Path (Join-Path $here '..')).Path
$task = 'Unitech Costing'
$t = Get-ScheduledTask -TaskName $task -ErrorAction SilentlyContinue
if (-not $t) { Write-Host 'The task is missing. Run Setup.bat first.' -ForegroundColor Red; exit 1 }
$vpy = $t.Actions[0].Execute
Stop-ScheduledTask -TaskName $task
Start-Sleep -Seconds 2
Get-CimInstance Win32_Process | Where-Object { $_.ExecutablePath -eq $vpy } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
& $vpy -m pip install -r (Join-Path $root 'requirements.txt') -q --disable-pip-version-check    # new libraries, if any
Start-ScheduledTask -TaskName $task
Write-Host 'Restarted. Give it a minute, then open the dashboard.'
