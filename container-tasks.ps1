<#
.SYNOPSIS
  Auto-create Outlook calendar appointments from container pickup/dropoff
  confirmation emails. Uses Outlook Desktop (COM) + Claude API for extraction.

.SETUP (one time)
  1. Get an Anthropic API key from https://console.anthropic.com
  2. Set it as a USER environment variable, then open a NEW PowerShell window:
       [Environment]::SetEnvironmentVariable("ANTHROPIC_API_KEY","sk-ant-...","User")
  3. In Outlook, create a calendar named exactly the value of $CalendarName below
     (default: "Containers"). Right-click your default Calendar -> New Calendar.
  4. First manual run (from this folder):
       powershell -ExecutionPolicy Bypass -File .\container-tasks.ps1
  5. Schedule it: Task Scheduler -> Create Task
       - Trigger: Daily, repeat task every 10 minutes for 1 day
       - Action:  powershell.exe
       - Args:    -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Users\LisaSandoval\GitHub\outlook-container-tasks\container-tasks.ps1"
       - Run only when user is logged on (so it can talk to Outlook)

.NOTES
  - Outlook does not need to be open; PowerShell launches it via COM.
  - State (processed email IDs) lives in .\state\processed.json.
  - Logs live in .\logs\run-YYYY-MM-DD.log.
  - Each emails costs one Claude call ONLY if it passes the keyword pre-filter.
#>

# ============================ CONFIG ============================
$CalendarName        = "Containers"          # Outlook calendar to write events to
$LookbackHours       = 24                    # First-run scan window; later runs use state
$ClaudeModel         = "claude-sonnet-4-6"   # Sonnet is plenty for this extraction
$AppointmentMins     = 60                    # Duration to block on the calendar
$ReminderMins        = 30                    # Reminder lead time
$EmailCategory       = "Container Confirmed" # Color category to tag the source email; "" to disable
$ProcessedMailFolder = "Containers"          # Inbox subfolder to move tagged email into; "" to disable
$AttachExtsDocs      = @('.pdf')
$AttachExtsImages    = @('.png','.jpg','.jpeg','.gif','.webp')
$AttachMaxBytesPerFile = 10MB
$AttachMaxBytesTotal   = 25MB
$KeywordPrefilter = @(
    'container','contianer',
    'pickup','pick up','pick-up',
    'dropoff','drop off','drop-off',
    'inbound','outbound',
    'confirmed','confirm','scheduled',
    'eta','arrival','booking','dispatch','delivery'
)
# ================================================================

$ErrorActionPreference = 'Stop'
$StateDir  = Join-Path $PSScriptRoot "state"
$LogDir    = Join-Path $PSScriptRoot "logs"
New-Item -ItemType Directory -Force -Path $StateDir | Out-Null
New-Item -ItemType Directory -Force -Path $LogDir   | Out-Null
$StateFile = Join-Path $StateDir "processed.json"
$LogFile   = Join-Path $LogDir   ("run-{0}.log" -f (Get-Date -Format "yyyy-MM-dd"))

function Write-Log {
    param([string]$Message)
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Add-Content -Path $LogFile -Value $line -Encoding utf8
}

# ---------- Load processed-EntryID set ----------
$processed = @{}
if (Test-Path $StateFile) {
    try {
        $raw = Get-Content $StateFile -Raw -Encoding utf8
        if ($raw) {
            (ConvertFrom-Json $raw).PSObject.Properties | ForEach-Object {
                $processed[$_.Name] = $_.Value
            }
        }
    } catch {
        Write-Log "WARN failed to read state file: $($_.Exception.Message)"
    }
}

# ---------- Connect to Outlook ----------
try {
    $outlook = New-Object -ComObject Outlook.Application
    $ns      = $outlook.GetNamespace("MAPI")
} catch {
    Write-Log "FATAL cannot connect to Outlook COM: $($_.Exception.Message)"
    throw
}
$inbox      = $ns.GetDefaultFolder(6)   # olFolderInbox
$defaultCal = $ns.GetDefaultFolder(9)   # olFolderCalendar

# ---------- Find the target calendar ----------
$targetCal = $null
foreach ($cal in $defaultCal.Folders) {
    if ($cal.Name -eq $CalendarName) { $targetCal = $cal; break }
}
if (-not $targetCal) {
    foreach ($store in $ns.Stores) {
        $root = $store.GetRootFolder()
        foreach ($f in $root.Folders) {
            if ($f.Name -eq $CalendarName) { $targetCal = $f; break }
        }
        if ($targetCal) { break }
    }
}
if (-not $targetCal) {
    $msg = "Calendar '$CalendarName' not found. Create it under your default Calendar in Outlook (right-click Calendar -> New Calendar)."
    Write-Log "ERROR $msg"
    throw $msg
}

# ---------- Restrict Inbox to recent items ----------
$cutoff = (Get-Date).AddHours(-$LookbackHours)
$filter = "[ReceivedTime] >= '{0}'" -f $cutoff.ToString("g")
$items  = $inbox.Items
$items.Sort("[ReceivedTime]", $true)   # newest first
$candidates = $items.Restrict($filter)
Write-Log "Scan started. Cutoff=$cutoff  Candidates=$($candidates.Count)"

function Get-AttachmentBlocks {
    param($Mail)
    $blocks    = @()
    $tempFiles = @()
    $filenames = @()
    $total     = 0
    foreach ($att in $Mail.Attachments) {
        $name = [string]$att.FileName
        $filenames += $name
        $ext  = [System.IO.Path]::GetExtension($name).ToLower()
        $isDoc = $AttachExtsDocs   -contains $ext
        $isImg = $AttachExtsImages -contains $ext
        if (-not ($isDoc -or $isImg)) { continue }
        if ($att.Size -gt $AttachMaxBytesPerFile) {
            Write-Log "SKIP attachment too big: $name ($($att.Size) bytes)"
            continue
        }
        if (($total + $att.Size) -gt $AttachMaxBytesTotal) {
            Write-Log "SKIP attachment over total budget: $name"
            continue
        }
        $safe = "ct_{0}{1}" -f [guid]::NewGuid().ToString('n'), $ext
        $tmp  = Join-Path $env:TEMP $safe
        try {
            $att.SaveAsFile($tmp)
            $bytes = [System.IO.File]::ReadAllBytes($tmp)
            $b64   = [Convert]::ToBase64String($bytes)
            $total += $bytes.Length
            $tempFiles += $tmp
            if ($isDoc) {
                $blocks += @{
                    type   = "document"
                    source = @{ type = "base64"; media_type = "application/pdf"; data = $b64 }
                }
            } else {
                $media = switch ($ext) {
                    '.png'  { 'image/png' }
                    '.gif'  { 'image/gif' }
                    '.webp' { 'image/webp' }
                    default { 'image/jpeg' }
                }
                $blocks += @{
                    type   = "image"
                    source = @{ type = "base64"; media_type = $media; data = $b64 }
                }
            }
        } catch {
            Write-Log "WARN reading attachment '$name' failed: $($_.Exception.Message)"
        }
    }
    return [pscustomobject]@{ Blocks = $blocks; TempFiles = $tempFiles; Filenames = $filenames }
}

function Remove-TempFiles {
    param([string[]]$Paths)
    foreach ($p in $Paths) {
        try { Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue } catch {}
    }
}

function Get-OrCreateMailSubfolder {
    param($Parent, [string]$Name)
    foreach ($f in $Parent.Folders) {
        if ($f.Name -eq $Name) { return $f }
    }
    return $Parent.Folders.Add($Name)
}

function Add-EmailCategory {
    param($Mail, [string]$Category)
    if (-not $Category) { return }
    $cats = [string]$Mail.Categories
    if ([string]::IsNullOrWhiteSpace($cats)) {
        $Mail.Categories = $Category
    } elseif ($cats -notmatch [regex]::Escape($Category)) {
        $Mail.Categories = "$cats, $Category"
    }
    $Mail.Save() | Out-Null
}

function Test-Keywords([string]$subj, [string]$body) {
    $hay = ($subj + ' ' + $body).ToLower()
    foreach ($k in $KeywordPrefilter) {
        if ($hay.Contains($k.ToLower())) { return $true }
    }
    return $false
}

# ---------- Claude extraction ----------
function Invoke-ClaudeExtract {
    param(
        [string]$Subject,
        [string]$From,
        [string]$Body,
        [datetime]$Received,
        [array]$AttachmentBlocks = @()
    )
    $apiKey = $env:ANTHROPIC_API_KEY
    if (-not $apiKey) { throw "ANTHROPIC_API_KEY env var is not set" }

    $systemPrompt = @'
You extract container pickup/dropoff confirmations from logistics emails. Return STRICT JSON only — no prose, no markdown, no backticks.

Schema:
{
  "is_confirmation": boolean,        // true ONLY if the client has confirmed a pickup or drop-off has happened or is scheduled at a specific date/time
  "direction": "inbound" | "outbound" | "unknown",
  "event_type": "pickup" | "dropoff" | "unknown",
  "container_number": string | null, // e.g. "MSCU1234567"; null if not present
  "datetime_iso": string | null,     // ISO 8601 LOCAL time, e.g. "2026-05-22T14:30:00"; null if no specific date+time
  "location": string | null,         // facility / address / port; null if missing
  "customer": string | null,         // client/company name; null if unclear
  "notes": string | null             // 1 short sentence of extra context, or null
}

Rules:
- "Inbound" = container coming TO our facility/destination. "Outbound" = container leaving our facility.
- If the email is a quote, RFQ, invoice, generic update, automated marketing, or anything other than an actual confirmed scheduling event, set is_confirmation=false and other fields to null/unknown.
- If only a date is given (no time), still produce datetime_iso with a sensible local time (default 09:00).
- Strip quoted reply chains; focus on the latest message.
- Attachments (PDFs/images such as BOLs, dock receipts, booking confirmations) often contain the actual details — extract from them when the email body is sparse.
'@

    $userPrompt = @"
From: $From
Received: $($Received.ToString("yyyy-MM-dd HH:mm"))
Subject: $Subject

Body:
$Body
"@

    $content = @()
    if ($AttachmentBlocks -and $AttachmentBlocks.Count -gt 0) {
        $content += $AttachmentBlocks
    }
    $content += @{ type = "text"; text = $userPrompt }

    $payload = @{
        model      = $ClaudeModel
        max_tokens = 600
        system     = $systemPrompt
        messages   = @(@{ role = "user"; content = $content })
    } | ConvertTo-Json -Depth 8 -Compress

    $headers = @{
        "x-api-key"         = $apiKey
        "anthropic-version" = "2023-06-01"
        "content-type"      = "application/json"
    }

    try {
        $resp = Invoke-RestMethod -Method Post `
            -Uri "https://api.anthropic.com/v1/messages" `
            -Headers $headers -Body $payload -TimeoutSec 60
    } catch {
        Write-Log "ERROR Claude API call failed: $($_.Exception.Message)"
        return $null
    }

    $text = $resp.content[0].text
    $text = $text -replace '^```(?:json)?\s*','' -replace '\s*```$',''
    try { return ConvertFrom-Json $text } catch {
        Write-Log "WARN failed to parse Claude JSON. Raw: $text"
        return $null
    }
}

# ---------- Resolve destination mail folder (optional) ----------
$destFolder = $null
if ($ProcessedMailFolder) {
    try {
        $destFolder = Get-OrCreateMailSubfolder -Parent $inbox -Name $ProcessedMailFolder
    } catch {
        Write-Log "WARN could not get/create mail subfolder '$ProcessedMailFolder': $($_.Exception.Message)"
    }
}

# Snapshot to an array — moving items out of Inbox during a live foreach skips siblings.
$snapshot = @()
foreach ($m in $candidates) { $snapshot += $m }

# ---------- Main loop ----------
$examined = 0
$created  = 0
foreach ($mail in $snapshot) {
    try {
        if ($mail.Class -ne 43) { continue }   # 43 = olMail
        $entryId = $mail.EntryID
        if ($processed.ContainsKey($entryId)) { continue }

        $subj = [string]$mail.Subject
        $body = [string]$mail.Body

        # Gather attachments first so their filenames participate in the keyword pre-filter.
        $attInfo = $null
        try { $attInfo = Get-AttachmentBlocks -Mail $mail } catch {
            Write-Log "WARN gathering attachments failed: $($_.Exception.Message)"
        }
        $attBlocks    = if ($attInfo) { $attInfo.Blocks }    else { @() }
        $attTempFiles = if ($attInfo) { $attInfo.TempFiles } else { @() }
        $attNames     = if ($attInfo) { ($attInfo.Filenames -join ' ') } else { '' }

        if (-not (Test-Keywords $subj ("$body $attNames"))) {
            $processed[$entryId] = "skipped-keyword"
            Remove-TempFiles $attTempFiles
            continue
        }

        $from = ''
        try { $from = [string]$mail.SenderEmailAddress } catch {}
        $received = $mail.ReceivedTime
        if ($body.Length -gt 8000) { $body = $body.Substring(0, 8000) }

        $extract = Invoke-ClaudeExtract -Subject $subj -From $from -Body $body -Received $received -AttachmentBlocks $attBlocks
        Remove-TempFiles $attTempFiles
        $examined++

        if (-not $extract -or -not $extract.is_confirmation) {
            $processed[$entryId] = "not-confirmation"
            continue
        }
        if (-not $extract.datetime_iso) {
            $processed[$entryId] = "no-datetime"
            Write-Log "SKIP no datetime: $subj"
            continue
        }

        try {
            $start = [datetime]::Parse($extract.datetime_iso)
        } catch {
            $processed[$entryId] = "bad-datetime"
            Write-Log "SKIP bad datetime '$($extract.datetime_iso)': $subj"
            continue
        }
        $end = $start.AddMinutes($AppointmentMins)

        $appt = $targetCal.Items.Add(1)   # olAppointmentItem
        $direction = if ($extract.direction)  { ([string]$extract.direction).ToUpper() }  else { 'UNKNOWN' }
        $etype     = if ($extract.event_type) { ([string]$extract.event_type).ToUpper() } else { 'UNKNOWN' }
        $cnum      = if ($extract.container_number) { [string]$extract.container_number } else { 'N/A' }

        $appt.Subject = "[$direction $etype] $cnum"
        if ($extract.location) { $appt.Location = [string]$extract.location }
        $appt.Start = $start
        $appt.End   = $end
        $appt.Body  = @"
Container: $cnum
Direction: $direction
Event:     $etype
Customer:  $($extract.customer)
Notes:     $($extract.notes)

---
Source email
From:     $from
Subject:  $subj
Received: $received
"@
        $appt.ReminderSet = $true
        $appt.ReminderMinutesBeforeStart = $ReminderMins
        $appt.Save() | Out-Null

        $processed[$entryId] = "created:$($start.ToString('s'))"
        $created++
        Write-Log "CREATED '$($appt.Subject)' at $start  (from: $from)"

        # Tag + move source email so it's visible which mail produced an event
        try { Add-EmailCategory -Mail $mail -Category $EmailCategory } catch {
            Write-Log "WARN failed to category email: $($_.Exception.Message)"
        }
        if ($destFolder) {
            try { [void]$mail.Move($destFolder) } catch {
                Write-Log "WARN failed to move email to '$ProcessedMailFolder': $($_.Exception.Message)"
            }
        }
    } catch {
        Write-Log "ERROR processing item: $($_.Exception.Message)"
        continue
    }
}

# ---------- Persist state ----------
$processed | ConvertTo-Json -Depth 3 | Set-Content -Path $StateFile -Encoding utf8
Write-Log "Scan complete. Examined(via Claude)=$examined  Created=$created"

# ---------- Release COM ----------
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($inbox)      | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($targetCal)  | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($defaultCal) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ns)         | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook)    | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
