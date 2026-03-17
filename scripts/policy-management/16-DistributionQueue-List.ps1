# ============================================================
# 16-DistributionQueue-List.ps1
# Creates PM_DistributionQueue list for bulk distribution jobs
# Assumes user is already connected to SharePoint
# ============================================================

$SiteUrl = "https://mf7m.sharepoint.com/sites/PolicyManager"
$ListName = "PM_DistributionQueue"

Write-Host "`n=== Creating $ListName ===" -ForegroundColor Cyan

# Create list if not exists
$list = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $ListName -Template GenericList -Url "Lists/$ListName"
    Write-Host "  Created list: $ListName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $ListName" -ForegroundColor Yellow
}

# ---- Fields ----

# PolicyId — which policy is being distributed
$field = Get-PnPField -List $ListName -Identity "PolicyId" -ErrorAction SilentlyContinue
if ($null -eq $field) {
    Add-PnPField -List $ListName -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -Required
    Write-Host "  Added field: PolicyId" -ForegroundColor Green
}

# PolicyName — display name for UI
$field = Get-PnPField -List $ListName -Identity "PolicyName" -ErrorAction SilentlyContinue
if ($null -eq $field) {
    Add-PnPField -List $ListName -DisplayName "PolicyName" -InternalName "PolicyName" -Type Text
    Write-Host "  Added field: PolicyName" -ForegroundColor Green
}

# TargetUserIds — JSON array of user IDs (can be large)
$field = Get-PnPField -List $ListName -Identity "TargetUserIds" -ErrorAction SilentlyContinue
if ($null -eq $field) {
    Add-PnPField -List $ListName -DisplayName "TargetUserIds" -InternalName "TargetUserIds" -Type Note
    Write-Host "  Added field: TargetUserIds" -ForegroundColor Green
}

# TotalUsers — total count for progress calculation
$field = Get-PnPField -List $ListName -Identity "TotalUsers" -ErrorAction SilentlyContinue
if ($null -eq $field) {
    Add-PnPField -List $ListName -DisplayName "TotalUsers" -InternalName "TotalUsers" -Type Number
    Write-Host "  Added field: TotalUsers" -ForegroundColor Green
}

# ProcessedUsers — how many have been processed so far
$field = Get-PnPField -List $ListName -Identity "ProcessedUsers" -ErrorAction SilentlyContinue
if ($null -eq $field) {
    Add-PnPField -List $ListName -DisplayName "ProcessedUsers" -InternalName "ProcessedUsers" -Type Number
    Set-PnPField -List $ListName -Identity "ProcessedUsers" -Values @{DefaultValue = "0"}
    Write-Host "  Added field: ProcessedUsers" -ForegroundColor Green
}

# FailedUsers — count of failures
$field = Get-PnPField -List $ListName -Identity "FailedUsers" -ErrorAction SilentlyContinue
if ($null -eq $field) {
    Add-PnPField -List $ListName -DisplayName "FailedUsers" -InternalName "FailedUsers" -Type Number
    Set-PnPField -List $ListName -Identity "FailedUsers" -Values @{DefaultValue = "0"}
    Write-Host "  Added field: FailedUsers" -ForegroundColor Green
}

# Status — Queued, Processing, Completed, Failed, Cancelled
$field = Get-PnPField -List $ListName -Identity "QueueStatus" -ErrorAction SilentlyContinue
if ($null -eq $field) {
    Add-PnPField -List $ListName -DisplayName "Status" -InternalName "QueueStatus" -Type Choice -Choices "Queued","Processing","Completed","Failed","Cancelled"
    Set-PnPField -List $ListName -Identity "QueueStatus" -Values @{DefaultValue = "Queued"}
    Write-Host "  Added field: QueueStatus" -ForegroundColor Green
}

# JobType — Publish, Redistribute, Reminder, Revoke
$field = Get-PnPField -List $ListName -Identity "JobType" -ErrorAction SilentlyContinue
if ($null -eq $field) {
    Add-PnPField -List $ListName -DisplayName "JobType" -InternalName "JobType" -Type Choice -Choices "Publish","Redistribute","Reminder","Revoke"
    Set-PnPField -List $ListName -Identity "JobType" -Values @{DefaultValue = "Publish"}
    Write-Host "  Added field: JobType" -ForegroundColor Green
}

# DueDate — acknowledgement deadline
$field = Get-PnPField -List $ListName -Identity "DueDate" -ErrorAction SilentlyContinue
if ($null -eq $field) {
    Add-PnPField -List $ListName -DisplayName "DueDate" -InternalName "DueDate" -Type DateTime
    Write-Host "  Added field: DueDate" -ForegroundColor Green
}

# SendNotifications — whether to send email notifications
$field = Get-PnPField -List $ListName -Identity "SendNotifications" -ErrorAction SilentlyContinue
if ($null -eq $field) {
    Add-PnPField -List $ListName -DisplayName "SendNotifications" -InternalName "SendNotifications" -Type Boolean
    Set-PnPField -List $ListName -Identity "SendNotifications" -Values @{DefaultValue = "1"}
    Write-Host "  Added field: SendNotifications" -ForegroundColor Green
}

# QueuedBy — who initiated the distribution
$field = Get-PnPField -List $ListName -Identity "QueuedBy" -ErrorAction SilentlyContinue
if ($null -eq $field) {
    Add-PnPField -List $ListName -DisplayName "QueuedBy" -InternalName "QueuedBy" -Type Text
    Write-Host "  Added field: QueuedBy" -ForegroundColor Green
}

# QueuedByEmail
$field = Get-PnPField -List $ListName -Identity "QueuedByEmail" -ErrorAction SilentlyContinue
if ($null -eq $field) {
    Add-PnPField -List $ListName -DisplayName "QueuedByEmail" -InternalName "QueuedByEmail" -Type Text
    Write-Host "  Added field: QueuedByEmail" -ForegroundColor Green
}

# StartedDate — when processing began
$field = Get-PnPField -List $ListName -Identity "StartedDate" -ErrorAction SilentlyContinue
if ($null -eq $field) {
    Add-PnPField -List $ListName -DisplayName "StartedDate" -InternalName "StartedDate" -Type DateTime
    Write-Host "  Added field: StartedDate" -ForegroundColor Green
}

# CompletedDate — when processing finished
$field = Get-PnPField -List $ListName -Identity "CompletedDate" -ErrorAction SilentlyContinue
if ($null -eq $field) {
    Add-PnPField -List $ListName -DisplayName "CompletedDate" -InternalName "CompletedDate" -Type DateTime
    Write-Host "  Added field: CompletedDate" -ForegroundColor Green
}

# ErrorLog — JSON array of error messages
$field = Get-PnPField -List $ListName -Identity "ErrorLog" -ErrorAction SilentlyContinue
if ($null -eq $field) {
    Add-PnPField -List $ListName -DisplayName "ErrorLog" -InternalName "ErrorLog" -Type Note
    Write-Host "  Added field: ErrorLog" -ForegroundColor Green
}

# PolicyVersionNumber — which version was distributed
$field = Get-PnPField -List $ListName -Identity "PolicyVersionNumber" -ErrorAction SilentlyContinue
if ($null -eq $field) {
    Add-PnPField -List $ListName -DisplayName "PolicyVersionNumber" -InternalName "PolicyVersionNumber" -Type Text
    Write-Host "  Added field: PolicyVersionNumber" -ForegroundColor Green
}

Write-Host "`n=== $ListName setup complete ===" -ForegroundColor Cyan
Write-Host "Fields: PolicyId, PolicyName, TargetUserIds, TotalUsers, ProcessedUsers, FailedUsers, QueueStatus, JobType, DueDate, SendNotifications, QueuedBy, QueuedByEmail, StartedDate, CompletedDate, ErrorLog, PolicyVersionNumber" -ForegroundColor Gray
