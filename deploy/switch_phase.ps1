# Switch the costing dashboard between Phase 1 and Phase 2. Run through Switch-Phase.bat (administrator).
#   Phase 1 = the version from before Tailscale access and part photos (GitHub branch "phase-1")
#   Phase 2 = the current version (GitHub branch "main")
# It sets "update_branch" in config\settings.json and runs the normal update, which downloads that version,
# restarts the dashboard and goes back by itself if it doesn't start. The nightly update then stays on it.
# Databases, settings, photos and backups are not touched. Phase 1 also switches off "tailscale serve".
# Needs Install-AutoUpdate.bat to have been run once (the folder must be linked to GitHub).
param([string]$Phase = '')
$ErrorActionPreference = 'Continue'
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $p = Start-Process powershell.exe -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File',('"' + $PSCommandPath + '"'),'-Phase',('"' + $Phase + '"')) -Verb RunAs -Wait -PassThru
    exit $p.ExitCode
}
# works from deploy\ and from the phase-switch\ copy kept outside the program files
$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
if (-not (Test-Path (Join-Path $root 'DASHBOARD'))) { Write-Host "Can't find the costing folder above $PSScriptRoot" -ForegroundColor Red; exit 1 }
if (-not (Test-Path (Join-Path $root '.git'))) { Write-Host 'This folder is not linked to GitHub yet. Run deploy\Install-AutoUpdate.bat once, then this again.' -ForegroundColor Red; exit 1 }

$settingsFile = Join-Path $root 'config\settings.json'
$settings = if (Test-Path $settingsFile) { Get-Content $settingsFile -Raw | ConvertFrom-Json } else { New-Object PSObject }
$current = if ($settings.update_branch -eq 'phase-1') { '1' } else { '2' }
Write-Host "Now on Phase $current."
if ($Phase -notin @('1', '2')) { $Phase = (Read-Host 'Switch to Phase 1 or 2? Type 1 or 2').Trim() }
if ($Phase -notin @('1', '2')) { Write-Host 'Nothing changed.'; exit 1 }
$branch = if ($Phase -eq '1') { 'phase-1' } else { 'main' }

# keep this tool outside the program files, so it is still there after going back to Phase 1
$keep = Join-Path $root 'phase-switch'
if ((Split-Path $PSScriptRoot -Leaf) -ne 'phase-switch') {
    New-Item -ItemType Directory -Path $keep -Force | Out-Null
    Copy-Item $PSCommandPath (Join-Path $keep 'switch_phase.ps1') -Force
    Copy-Item (Join-Path $PSScriptRoot 'Switch-Phase.bat') (Join-Path $keep 'Switch-Phase.bat') -Force
}

$settings | Add-Member -NotePropertyName update_branch -NotePropertyValue $branch -Force
New-Item -ItemType Directory -Path (Split-Path $settingsFile) -Force | Out-Null
[IO.File]::WriteAllText($settingsFile, ($settings | ConvertTo-Json -Depth 5))      # UTF-8 without BOM

if ($Phase -eq '1') {
    $ts = if (Get-Command tailscale.exe -ErrorAction SilentlyContinue) { 'tailscale.exe' } elseif (Test-Path "$env:ProgramFiles\Tailscale\tailscale.exe") { "$env:ProgramFiles\Tailscale\tailscale.exe" } else { $null }
    if ($ts) { & $ts serve reset 2>&1 | Out-Null; Write-Host 'Tailscale https address switched off.' }
    Remove-NetFirewallRule -DisplayName 'Unitech Costing (Tailscale)' -ErrorAction SilentlyContinue
}

Write-Host "Switching to Phase $Phase (GitHub branch '$branch')..."
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root 'deploy\update_windows.ps1')
if ($LASTEXITCODE -eq 0) {
    Write-Host "Done: Phase $Phase. To switch again later, run phase-switch\Switch-Phase.bat." -ForegroundColor Green
} else {
    Write-Host 'The switch did not finish - see logs\update.log.' -ForegroundColor Red
    exit 1
}
