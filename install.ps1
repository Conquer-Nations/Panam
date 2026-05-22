<#
.SYNOPSIS
  Deploys container-tasks.ps1 as a Windows scheduled task that runs every
  10 minutes while you are signed in. Run this ONCE on the machine where
  Outlook Desktop is installed.

  Usage:
    powershell -ExecutionPolicy Bypass -File .\install.ps1
#>

$ErrorActionPreference = 'Stop'
$ScriptPath = Join-Path $PSScriptRoot "container-tasks.ps1"
$TaskName   = "ContainerCalendarSync"

if (-not (Test-Path $ScriptPath)) {
    throw "container-tasks.ps1 not found next to install.ps1"
}

Write-Host "==== Step 1: Anthropic API key ====" -ForegroundColor Cyan
if (-not $env:ANTHROPIC_API_KEY) {
    $existing = [Environment]::GetEnvironmentVariable("ANTHROPIC_API_KEY","User")
    if ($existing) {
        $env:ANTHROPIC_API_KEY = $existing
        Write-Host "Found user env var ANTHROPIC_API_KEY."
    } else {
        $secure = Read-Host "Paste your ANTHROPIC_API_KEY (starts with sk-ant-)" -AsSecureString
        $bstr   = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        $key    = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        if (-not $key) { throw "No key provided." }
        [Environment]::SetEnvironmentVariable("ANTHROPIC_API_KEY", $key, "User")
        $env:ANTHROPIC_API_KEY = $key
        Write-Host "Saved ANTHROPIC_API_KEY to your user environment."
    }
} else {
    Write-Host "ANTHROPIC_API_KEY is already set in this session."
}

Write-Host ""
Write-Host "==== Step 2: Outlook reachable? ====" -ForegroundColor Cyan
try {
    $outlook = New-Object -ComObject Outlook.Application
    $ns      = $outlook.GetNamespace("MAPI")
    $defaultCal = $ns.GetDefaultFolder(9)
    Write-Host "Connected to Outlook. Default calendar store: $($defaultCal.Parent.Name)"
} catch {
    throw "Could not start Outlook via COM. Make sure Outlook Desktop is installed and you have signed in at least once. Error: $($_.Exception.Message)"
}

Write-Host ""
Write-Host "==== Step 3: Target calendar exists? ====" -ForegroundColor Cyan
# Read $CalendarName out of container-tasks.ps1 so install matches the script.
$calLine    = (Get-Content $ScriptPath) | Where-Object { $_ -match '^\s*\$CalendarName\s*=' } | Select-Object -First 1
$calName    = if ($calLine -match '"([^"]+)"') { $Matches[1] } else { "Containers" }
$found = $false
foreach ($c in $defaultCal.Folders) { if ($c.Name -eq $calName) { $found = $true; break } }
if (-not $found) {
    foreach ($store in $ns.Stores) {
        $root = $store.GetRootFolder()
        foreach ($f in $root.Folders) { if ($f.Name -eq $calName) { $found = $true; break } }
        if ($found) { break }
    }
}
if ($found) {
    Write-Host "Calendar '$calName' found."
} else {
    Write-Host "Calendar '$calName' NOT FOUND." -ForegroundColor Yellow
    Write-Host "Create it now in Outlook: right-click 'Calendar' in the folder pane -> New Calendar -> name it exactly: $calName"
    Read-Host "Press Enter once created"
}

# Release COM
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($defaultCal) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ns)         | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook)    | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()

Write-Host ""
Write-Host "==== Step 4: First run (dry test) ====" -ForegroundColor Cyan
Write-Host "Running container-tasks.ps1 once now to verify everything works..."
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $ScriptPath
$rc = $LASTEXITCODE
if ($rc -ne 0) {
    Write-Host "First run exited with code $rc. Check the logs folder before scheduling." -ForegroundColor Yellow
    Read-Host "Press Enter to continue anyway, or Ctrl+C to abort"
} else {
    Write-Host "First run completed. See .\logs\ for details."
}

Write-Host ""
Write-Host "==== Step 5: Register scheduled task ====" -ForegroundColor Cyan
$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument ("-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"{0}`"" -f $ScriptPath)
$startAt = (Get-Date).AddMinutes(2)
$trigger = New-ScheduledTaskTrigger -Once -At $startAt `
    -RepetitionInterval (New-TimeSpan -Minutes 10) `
    -RepetitionDuration (New-TimeSpan -Days 3650)
$userId    = "$env:USERDOMAIN\$env:USERNAME"
$principal = New-ScheduledTaskPrincipal -UserId $userId -LogonType Interactive -RunLevel Limited
$settings  = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -StartWhenAvailable -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 15)

if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Removed existing '$TaskName' task."
}
Register-ScheduledTask -TaskName $TaskName `
    -Action $action -Trigger $trigger -Principal $principal -Settings $settings `
    -Description "Auto-create Outlook calendar events from container pickup/dropoff emails." | Out-Null
Write-Host "Registered scheduled task '$TaskName'. It will run every 10 minutes while you are signed in."

Write-Host ""
Write-Host "==== Done ====" -ForegroundColor Green
Write-Host "Edit config knobs at the top of container-tasks.ps1 anytime (no need to re-run install)."
Write-Host "Check status:  Get-ScheduledTask -TaskName $TaskName"
Write-Host "Run manually:  Start-ScheduledTask -TaskName $TaskName"
Write-Host "Remove:        Unregister-ScheduledTask -TaskName $TaskName -Confirm:`$false"
Write-Host "Logs:          .\logs\run-YYYY-MM-DD.log"
