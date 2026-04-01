# ============================================================================
# Policy Manager - Event Log List Provisioning
# Creates PM_EventLog list for persisting diagnostic events from the
# DWx Event Viewer. Idempotent — safe to run multiple times.
# Assumes: Already connected via Connect-PnPOnline
# ============================================================================

$ErrorActionPreference = "Continue"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Event Viewer - Event Log List Provisioning" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# ----------------------------------------------------------------------------
# 1. PM_EventLog
# ----------------------------------------------------------------------------
$listName = "PM_EventLog"
Write-Host "[1/2] Creating $listName..." -ForegroundColor Yellow

$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

# Event identification
Add-PnPField -List $listName -DisplayName "EventCode" -InternalName "EventCode" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Severity" -InternalName "Severity" -Type Choice -Choices "Verbose","Information","Warning","Error","Critical" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Channel" -InternalName "Channel" -Type Choice -Choices "Application","Console","Network","Audit","DLQ","System" -ErrorAction SilentlyContinue | Out-Null

# Event details
Add-PnPField -List $listName -DisplayName "Source" -InternalName "Source" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Message" -InternalName "Message" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "StackTrace" -InternalName "StackTrace" -Type Note -ErrorAction SilentlyContinue | Out-Null

# Correlation
Add-PnPField -List $listName -DisplayName "CorrelationId" -InternalName "CorrelationId" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "SessionId" -InternalName "SessionId" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "UserLogin" -InternalName "UserLogin" -Type Text -ErrorAction SilentlyContinue | Out-Null

# Timestamps and metrics
Add-PnPField -List $listName -DisplayName "EventTimestamp" -InternalName "EventTimestamp" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Duration" -InternalName "Duration" -Type Number -ErrorAction SilentlyContinue | Out-Null

# Network-specific
Add-PnPField -List $listName -DisplayName "Url" -InternalName "Url" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "HttpMethod" -InternalName "HttpMethod" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "HttpStatus" -InternalName "HttpStatus" -Type Number -ErrorAction SilentlyContinue | Out-Null

# Investigation
Add-PnPField -List $listName -DisplayName "InvestigationNotes" -InternalName "InvestigationNotes" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Classification" -InternalName "Classification" -Type Choice -Choices "Bug","Performance","Security","Configuration","External","Unknown" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsInvestigated" -InternalName "IsInvestigated" -Type Boolean -ErrorAction SilentlyContinue | Out-Null

# Persistence metadata
Add-PnPField -List $listName -DisplayName "AutoPersisted" -InternalName "AutoPersisted" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Metadata" -InternalName "Metadata" -Type Note -ErrorAction SilentlyContinue | Out-Null

Write-Host "    Fields added to $listName" -ForegroundColor Gray

# ----------------------------------------------------------------------------
# 2. Add indexes for query performance
# ----------------------------------------------------------------------------
Write-Host "[2/2] Adding indexes to $listName..." -ForegroundColor Yellow

try {
    # Index on Severity — most common filter
    $field = Get-PnPField -List $listName -Identity "Severity" -ErrorAction SilentlyContinue
    if ($field -and -not $field.Indexed) {
        Set-PnPField -List $listName -Identity "Severity" -Values @{Indexed=$true} -ErrorAction SilentlyContinue | Out-Null
        Write-Host "    Indexed: Severity" -ForegroundColor Gray
    }

    # Index on Channel
    $field = Get-PnPField -List $listName -Identity "Channel" -ErrorAction SilentlyContinue
    if ($field -and -not $field.Indexed) {
        Set-PnPField -List $listName -Identity "Channel" -Values @{Indexed=$true} -ErrorAction SilentlyContinue | Out-Null
        Write-Host "    Indexed: Channel" -ForegroundColor Gray
    }

    # Index on EventCode
    $field = Get-PnPField -List $listName -Identity "EventCode" -ErrorAction SilentlyContinue
    if ($field -and -not $field.Indexed) {
        Set-PnPField -List $listName -Identity "EventCode" -Values @{Indexed=$true} -ErrorAction SilentlyContinue | Out-Null
        Write-Host "    Indexed: EventCode" -ForegroundColor Gray
    }

    # Index on EventTimestamp
    $field = Get-PnPField -List $listName -Identity "EventTimestamp" -ErrorAction SilentlyContinue
    if ($field -and -not $field.Indexed) {
        Set-PnPField -List $listName -Identity "EventTimestamp" -Values @{Indexed=$true} -ErrorAction SilentlyContinue | Out-Null
        Write-Host "    Indexed: EventTimestamp" -ForegroundColor Gray
    }
} catch {
    Write-Host "    Warning: Some indexes could not be created (non-blocking)" -ForegroundColor Yellow
}

# ----------------------------------------------------------------------------
# Summary
# ----------------------------------------------------------------------------
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Provisioning Complete" -ForegroundColor Cyan
Write-Host "  List: PM_EventLog (20 columns, 4 indexes)" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
