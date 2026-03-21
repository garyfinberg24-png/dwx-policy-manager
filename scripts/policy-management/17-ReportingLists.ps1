# ============================================================================
# Policy Manager - Reporting Lists
# Creates 3 reporting lists: ReportDefinitions, ScheduledReports, ReportExecutions
# Assumes: Already connected via Connect-PnPOnline
# ============================================================================

$ErrorActionPreference = "Continue"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Reporting Lists Provisioning" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# ----------------------------------------------------------------------------
# 1. PM_ReportDefinitions
# ----------------------------------------------------------------------------
$listName = "PM_ReportDefinitions"
Write-Host "[1/3] Creating $listName..." -ForegroundColor Yellow

$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "ReportName" -InternalName "ReportName" -Type Text -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Description" -InternalName "Description" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Category" -InternalName "Category" -Type Choice -Choices "Executive","Operational","Compliance","Financial","HR","Custom" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "LayoutConfig" -InternalName "LayoutConfig" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "WidgetsJSON" -InternalName "WidgetsJSON" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "SettingsJSON" -InternalName "SettingsJSON" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "GlobalFiltersJSON" -InternalName "GlobalFiltersJSON" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsPublic" -InternalName "IsPublic" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Tags" -InternalName "Tags" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsTemplate" -InternalName "IsTemplate" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Status" -InternalName "Status" -Type Choice -Choices "Draft","Published","Archived" -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "Status" -Values @{DefaultValue="Draft"} -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Fields added to $listName" -ForegroundColor Gray

# ----------------------------------------------------------------------------
# 2. PM_ScheduledReports
# ----------------------------------------------------------------------------
$listName = "PM_ScheduledReports"
Write-Host "[2/3] Creating $listName..." -ForegroundColor Yellow

$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "ReportId" -InternalName "ReportId" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ReportType" -InternalName "ReportType" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Frequency" -InternalName "Frequency" -Type Choice -Choices "Daily","Weekly","Monthly","Quarterly" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Format" -InternalName "Format" -Type Choice -Choices "PDF","Excel","CSV" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Recipients" -InternalName "Recipients" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Filters" -InternalName "Filters" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Enabled" -InternalName "Enabled" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "Enabled" -Values @{DefaultValue="1"} -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "LastRun" -InternalName "LastRun" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "NextRun" -InternalName "NextRun" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Fields added to $listName" -ForegroundColor Gray

# ----------------------------------------------------------------------------
# 3. PM_ReportExecutions
# ----------------------------------------------------------------------------
$listName = "PM_ReportExecutions"
Write-Host "[3/3] Creating $listName..." -ForegroundColor Yellow

$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "ReportName" -InternalName "ReportName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ReportType" -InternalName "ReportType" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "GeneratedByEmail" -InternalName "GeneratedByEmail" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "GeneratedByName" -InternalName "GeneratedByName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Format" -InternalName "Format" -Type Choice -Choices "PDF","Excel","CSV" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "RecordCount" -InternalName "RecordCount" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "FileSize" -InternalName "FileSize" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ExecutionTime" -InternalName "ExecutionTime" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ExecutionStatus" -InternalName "ExecutionStatus" -Type Choice -Choices "Success","Failed" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ErrorMessage" -InternalName "ErrorMessage" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ExecutedAt" -InternalName "ExecutedAt" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Fields added to $listName" -ForegroundColor Gray

# ----------------------------------------------------------------------------
# Done
# ----------------------------------------------------------------------------
Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  Reporting Lists provisioning complete!" -ForegroundColor Green
Write-Host "  Lists: PM_ReportDefinitions, PM_ScheduledReports, PM_ReportExecutions" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
