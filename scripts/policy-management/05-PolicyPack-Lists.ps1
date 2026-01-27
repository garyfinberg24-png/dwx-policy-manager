# ============================================================================
# DWx Policy Manager - Policy Pack Lists
# Part 5: PM_PolicyPacks, PM_PolicyPackAssignments
# ============================================================================

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,

    [Parameter(Mandatory=$false)]
    [switch]$UseWebLogin
)

# Connect to SharePoint
Write-Host "Connecting to SharePoint: $SiteUrl" -ForegroundColor Cyan
if ($UseWebLogin) {
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin
} else {
    Connect-PnPOnline -Url $SiteUrl -Interactive
}

# ============================================================================
# LIST 15: PM_PolicyPacks
# ============================================================================
Write-Host "`n Creating PM_PolicyPacks list..." -ForegroundColor Yellow

$listName = "PM_PolicyPacks"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Pack Name" -InternalName "PackName" -Type Text -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Pack Description" -InternalName "PackDescription" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Pack Category" -InternalName "PackCategory" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Pack Type" -InternalName "PackType" -Type Choice -Choices "Onboarding","Department","Role","Location","Custom" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Mandatory" -InternalName "IsMandatory" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Target Departments" -InternalName "TargetDepartments" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Target Roles" -InternalName "TargetRoles" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Target Locations" -InternalName "TargetLocations" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Policy IDs" -InternalName "PolicyIds" -Type Note -Required -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Policy Count" -InternalName "PolicyCount" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Require All Acknowledged" -InternalName "RequireAllAcknowledged" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Acknowledgement Deadline (Days)" -InternalName "AcknowledgementDeadlineDays" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Read Timeframe" -InternalName "ReadTimeframe" -Type Choice -Choices "Immediate","Day 1","Day 3","Week 1","Week 2","Month 1","Month 3","Month 6","Custom" -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Sequential" -InternalName "IsSequential" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Policy Sequence" -InternalName "PolicySequence" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Send Welcome Email" -InternalName "SendWelcomeEmail" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Send Teams Notification" -InternalName "SendTeamsNotification" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Welcome Email Template" -InternalName "WelcomeEmailTemplate" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Teams Message Template" -InternalName "TeamsMessageTemplate" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Total Assignments" -InternalName "TotalAssignments" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Total Completed" -InternalName "TotalCompleted" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Average Completion Days" -InternalName "AverageCompletionDays" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Completion Rate" -InternalName "CompletionRate" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Version" -InternalName "PackVersion" -Type Text -ErrorAction SilentlyContinue

Set-PnPField -List $listName -Identity "PackType" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "IsActive" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  PM_PolicyPacks list configured" -ForegroundColor Green

# ============================================================================
# LIST 16: PM_PolicyPackAssignments
# ============================================================================
Write-Host "`n Creating PM_PolicyPackAssignments list..." -ForegroundColor Yellow

$listName = "PM_PolicyPackAssignments"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Pack ID" -InternalName "PackId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User" -InternalName "AssignedUser" -Type User -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User Email" -InternalName "UserEmail" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User Department" -InternalName "UserDepartment" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User Role" -InternalName "UserRole" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Assigned Date" -InternalName "AssignedDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Assigned By" -InternalName "AssignedBy" -Type User -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Assignment Reason" -InternalName "AssignmentReason" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Onboarding Stage" -InternalName "OnboardingStage" -Type Choice -Choices "Pre-Start","Day 1","Week 1","Month 1","Month 3" -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Due Date" -InternalName "DueDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Read Timeframe" -InternalName "ReadTimeframe" -Type Choice -Choices "Immediate","Day 1","Day 3","Week 1","Week 2","Month 1","Month 3","Month 6","Custom" -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Total Policies" -InternalName "TotalPolicies" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Acknowledged Policies" -InternalName "AcknowledgedPolicies" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Pending Policies" -InternalName "PendingPolicies" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Overdue Policies" -InternalName "OverduePolicies" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Progress Percentage" -InternalName "ProgressPercentage" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Status" -InternalName "AssignmentStatus" -Type Choice -Choices "Not Started","In Progress","Completed","Overdue","Exempted" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Started Date" -InternalName "StartedDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Completed Date" -InternalName "CompletedDate" -Type DateTime -ErrorAction SilentlyContinue

Set-PnPField -List $listName -Identity "PackId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "UserEmail" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "AssignmentStatus" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  PM_PolicyPackAssignments list configured" -ForegroundColor Green

Write-Host "`n Policy Pack lists created successfully!" -ForegroundColor Green
Write-Host "   - PM_PolicyPacks" -ForegroundColor White
Write-Host "   - PM_PolicyPackAssignments" -ForegroundColor White
