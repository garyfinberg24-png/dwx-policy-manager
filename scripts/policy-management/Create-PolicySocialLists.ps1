# ============================================================================
# Create-PolicySocialLists.ps1
# Creates SharePoint lists for Policy Social Features and Policy Packs
# ============================================================================

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl
)

# Import PnP PowerShell
Import-Module PnP.PowerShell -ErrorAction Stop

Write-Host "============================================================================" -ForegroundColor Cyan
Write-Host "JML Policy Management - Social Features & Policy Packs Lists" -ForegroundColor Cyan
Write-Host "============================================================================" -ForegroundColor Cyan
Write-Host ""

# Connect to SharePoint
Write-Host "Connecting to SharePoint site: $SiteUrl" -ForegroundColor Yellow
Connect-PnPOnline -Url $SiteUrl -Interactive

Write-Host "Connected successfully!`n" -ForegroundColor Green

# ============================================================================
# 1. PM_PolicyRatings List
# ============================================================================

Write-Host "Creating PM_PolicyRatings list..." -ForegroundColor Yellow

$ratingsList = Get-PnPList -Identity "PM_PolicyRatings" -ErrorAction SilentlyContinue
if ($null -eq $ratingsList) {
    New-PnPList -Title "PM_PolicyRatings" -Template GenericList -OnQuickLaunch

    Add-PnPField -List "PM_PolicyRatings" -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyRatings" -DisplayName "UserId" -InternalName "UserId" -Type Number -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyRatings" -DisplayName "UserEmail" -InternalName "UserEmail" -Type Text -Required
    Add-PnPField -List "PM_PolicyRatings" -DisplayName "Rating" -InternalName "Rating" -Type Number -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyRatings" -DisplayName "RatingDate" -InternalName "RatingDate" -Type DateTime -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyRatings" -DisplayName "ReviewTitle" -InternalName "ReviewTitle" -Type Text
    Add-PnPField -List "PM_PolicyRatings" -DisplayName "ReviewText" -InternalName "ReviewText" -Type Note
    Add-PnPField -List "PM_PolicyRatings" -DisplayName "ReviewHelpfulCount" -InternalName "ReviewHelpfulCount" -Type Number
    Add-PnPField -List "PM_PolicyRatings" -DisplayName "IsVerifiedReader" -InternalName "IsVerifiedReader" -Type Boolean -AddToDefaultView
    Add-PnPField -List "PM_PolicyRatings" -DisplayName "UserRole" -InternalName "UserRole" -Type Text
    Add-PnPField -List "PM_PolicyRatings" -DisplayName "UserDepartment" -InternalName "UserDepartment" -Type Text

    Write-Host "✓ PM_PolicyRatings list created successfully" -ForegroundColor Green
} else {
    Write-Host "✓ PM_PolicyRatings list already exists" -ForegroundColor Cyan
}

# ============================================================================
# 2. PM_PolicyComments List
# ============================================================================

Write-Host "Creating PM_PolicyComments list..." -ForegroundColor Yellow

$commentsList = Get-PnPList -Identity "PM_PolicyComments" -ErrorAction SilentlyContinue
if ($null -eq $commentsList) {
    New-PnPList -Title "PM_PolicyComments" -Template GenericList -OnQuickLaunch

    Add-PnPField -List "PM_PolicyComments" -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyComments" -DisplayName "UserId" -InternalName "UserId" -Type Number -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyComments" -DisplayName "UserEmail" -InternalName "UserEmail" -Type Text -Required
    Add-PnPField -List "PM_PolicyComments" -DisplayName "CommentText" -InternalName "CommentText" -Type Note -Required
    Add-PnPField -List "PM_PolicyComments" -DisplayName "CommentDate" -InternalName "CommentDate" -Type DateTime -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyComments" -DisplayName "ModifiedDate" -InternalName "ModifiedDate" -Type DateTime
    Add-PnPField -List "PM_PolicyComments" -DisplayName "IsEdited" -InternalName "IsEdited" -Type Boolean
    Add-PnPField -List "PM_PolicyComments" -DisplayName "ParentCommentId" -InternalName "ParentCommentId" -Type Number
    Add-PnPField -List "PM_PolicyComments" -DisplayName "ReplyCount" -InternalName "ReplyCount" -Type Number
    Add-PnPField -List "PM_PolicyComments" -DisplayName "LikeCount" -InternalName "LikeCount" -Type Number -AddToDefaultView
    Add-PnPField -List "PM_PolicyComments" -DisplayName "IsStaffResponse" -InternalName "IsStaffResponse" -Type Boolean -AddToDefaultView
    Add-PnPField -List "PM_PolicyComments" -DisplayName "IsApproved" -InternalName "IsApproved" -Type Boolean -AddToDefaultView
    Add-PnPField -List "PM_PolicyComments" -DisplayName "IsDeleted" -InternalName "IsDeleted" -Type Boolean
    Add-PnPField -List "PM_PolicyComments" -DisplayName "DeletedReason" -InternalName "DeletedReason" -Type Note

    Write-Host "✓ PM_PolicyComments list created successfully" -ForegroundColor Green
} else {
    Write-Host "✓ PM_PolicyComments list already exists" -ForegroundColor Cyan
}

# ============================================================================
# 3. PM_PolicyCommentLikes List
# ============================================================================

Write-Host "Creating PM_PolicyCommentLikes list..." -ForegroundColor Yellow

$likesList = Get-PnPList -Identity "PM_PolicyCommentLikes" -ErrorAction SilentlyContinue
if ($null -eq $likesList) {
    New-PnPList -Title "PM_PolicyCommentLikes" -Template GenericList -OnQuickLaunch

    Add-PnPField -List "PM_PolicyCommentLikes" -DisplayName "CommentId" -InternalName "CommentId" -Type Number -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyCommentLikes" -DisplayName "UserId" -InternalName "UserId" -Type Number -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyCommentLikes" -DisplayName "LikedDate" -InternalName "LikedDate" -Type DateTime -Required -AddToDefaultView

    Write-Host "✓ PM_PolicyCommentLikes list created successfully" -ForegroundColor Green
} else {
    Write-Host "✓ PM_PolicyCommentLikes list already exists" -ForegroundColor Cyan
}

# ============================================================================
# 4. PM_PolicyShares List
# ============================================================================

Write-Host "Creating PM_PolicyShares list..." -ForegroundColor Yellow

$sharesList = Get-PnPList -Identity "PM_PolicyShares" -ErrorAction SilentlyContinue
if ($null -eq $sharesList) {
    New-PnPList -Title "PM_PolicyShares" -Template GenericList -OnQuickLaunch

    Add-PnPField -List "PM_PolicyShares" -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyShares" -DisplayName "SharedById" -InternalName "SharedById" -Type Number -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyShares" -DisplayName "SharedByEmail" -InternalName "SharedByEmail" -Type Text -Required
    Add-PnPField -List "PM_PolicyShares" -DisplayName "ShareMethod" -InternalName "ShareMethod" -Type Choice -Choices "Email","Teams","Link","QRCode","Download" -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyShares" -DisplayName "ShareDate" -InternalName "ShareDate" -Type DateTime -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyShares" -DisplayName "ShareMessage" -InternalName "ShareMessage" -Type Note
    Add-PnPField -List "PM_PolicyShares" -DisplayName "SharedWithUserIds" -InternalName "SharedWithUserIds" -Type Note
    Add-PnPField -List "PM_PolicyShares" -DisplayName "SharedWithEmails" -InternalName "SharedWithEmails" -Type Note
    Add-PnPField -List "PM_PolicyShares" -DisplayName "SharedWithTeamsChannelId" -InternalName "SharedWithTeamsChannelId" -Type Text
    Add-PnPField -List "PM_PolicyShares" -DisplayName "ViewCount" -InternalName "ViewCount" -Type Number
    Add-PnPField -List "PM_PolicyShares" -DisplayName "FirstViewedDate" -InternalName "FirstViewedDate" -Type DateTime
    Add-PnPField -List "PM_PolicyShares" -DisplayName "LastViewedDate" -InternalName "LastViewedDate" -Type DateTime

    Write-Host "✓ PM_PolicyShares list created successfully" -ForegroundColor Green
} else {
    Write-Host "✓ PM_PolicyShares list already exists" -ForegroundColor Cyan
}

# ============================================================================
# 5. PM_PolicyFollowers List
# ============================================================================

Write-Host "Creating PM_PolicyFollowers list..." -ForegroundColor Yellow

$followersList = Get-PnPList -Identity "PM_PolicyFollowers" -ErrorAction SilentlyContinue
if ($null -eq $followersList) {
    New-PnPList -Title "PM_PolicyFollowers" -Template GenericList -OnQuickLaunch

    Add-PnPField -List "PM_PolicyFollowers" -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyFollowers" -DisplayName "UserId" -InternalName "UserId" -Type Number -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyFollowers" -DisplayName "UserEmail" -InternalName "UserEmail" -Type Text -Required
    Add-PnPField -List "PM_PolicyFollowers" -DisplayName "FollowedDate" -InternalName "FollowedDate" -Type DateTime -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyFollowers" -DisplayName "NotifyOnUpdate" -InternalName "NotifyOnUpdate" -Type Boolean
    Add-PnPField -List "PM_PolicyFollowers" -DisplayName "NotifyOnComment" -InternalName "NotifyOnComment" -Type Boolean
    Add-PnPField -List "PM_PolicyFollowers" -DisplayName "NotifyOnNewVersion" -InternalName "NotifyOnNewVersion" -Type Boolean
    Add-PnPField -List "PM_PolicyFollowers" -DisplayName "EmailNotifications" -InternalName "EmailNotifications" -Type Boolean
    Add-PnPField -List "PM_PolicyFollowers" -DisplayName "TeamsNotifications" -InternalName "TeamsNotifications" -Type Boolean
    Add-PnPField -List "PM_PolicyFollowers" -DisplayName "InAppNotifications" -InternalName "InAppNotifications" -Type Boolean

    Write-Host "✓ PM_PolicyFollowers list created successfully" -ForegroundColor Green
} else {
    Write-Host "✓ PM_PolicyFollowers list already exists" -ForegroundColor Cyan
}

# ============================================================================
# 6. PM_PolicyPacks List
# ============================================================================

Write-Host "Creating PM_PolicyPacks list..." -ForegroundColor Yellow

$packsList = Get-PnPList -Identity "PM_PolicyPacks" -ErrorAction SilentlyContinue
if ($null -eq $packsList) {
    New-PnPList -Title "PM_PolicyPacks" -Template GenericList -OnQuickLaunch

    # Basic Info
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "PackName" -InternalName "PackName" -Type Text -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "PackDescription" -InternalName "PackDescription" -Type Note -Required
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "PackCategory" -InternalName "PackCategory" -Type Text
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "PackType" -InternalName "PackType" -Type Choice -Choices "Onboarding","Department","Role","Location","Custom" -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -AddToDefaultView
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "IsMandatory" -InternalName "IsMandatory" -Type Boolean

    # Targeting
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "TargetDepartments" -InternalName "TargetDepartments" -Type Note
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "TargetRoles" -InternalName "TargetRoles" -Type Note
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "TargetLocations" -InternalName "TargetLocations" -Type Note
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "TargetProcessType" -InternalName "TargetProcessType" -Type Choice -Choices "Joiner","Mover","Leaver"

    # Policies
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "PolicyIds" -InternalName "PolicyIds" -Type Note -Required
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "PolicyCount" -InternalName "PolicyCount" -Type Number -AddToDefaultView

    # Configuration
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "RequireAllAcknowledged" -InternalName "RequireAllAcknowledged" -Type Boolean
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "AcknowledgementDeadlineDays" -InternalName "AcknowledgementDeadlineDays" -Type Number
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "ReadTimeframe" -InternalName "ReadTimeframe" -Type Choice -Choices "Immediate","Day 1","Day 3","Week 1","Week 2","Month 1","Month 3","Month 6","Custom"
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "IsSequential" -InternalName "IsSequential" -Type Boolean
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "PolicySequence" -InternalName "PolicySequence" -Type Note

    # Notifications
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "SendWelcomeEmail" -InternalName "SendWelcomeEmail" -Type Boolean
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "SendTeamsNotification" -InternalName "SendTeamsNotification" -Type Boolean
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "WelcomeEmailTemplate" -InternalName "WelcomeEmailTemplate" -Type Note
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "TeamsMessageTemplate" -InternalName "TeamsMessageTemplate" -Type Note

    # Analytics
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "TotalAssignments" -InternalName "TotalAssignments" -Type Number
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "TotalCompleted" -InternalName "TotalCompleted" -Type Number
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "AverageCompletionDays" -InternalName "AverageCompletionDays" -Type Number
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "CompletionRate" -InternalName "CompletionRate" -Type Number

    # Metadata
    Add-PnPField -List "PM_PolicyPacks" -DisplayName "Version" -InternalName "Version" -Type Text

    Write-Host "✓ PM_PolicyPacks list created successfully" -ForegroundColor Green
} else {
    Write-Host "✓ PM_PolicyPacks list already exists" -ForegroundColor Cyan
}

# ============================================================================
# 7. PM_PolicyPackAssignments List
# ============================================================================

Write-Host "Creating PM_PolicyPackAssignments list..." -ForegroundColor Yellow

$assignmentsList = Get-PnPList -Identity "PM_PolicyPackAssignments" -ErrorAction SilentlyContinue
if ($null -eq $assignmentsList) {
    New-PnPList -Title "PM_PolicyPackAssignments" -Template GenericList -OnQuickLaunch

    # Assignment Info
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "PackId" -InternalName "PackId" -Type Number -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "UserId" -InternalName "UserId" -Type Number -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "UserEmail" -InternalName "UserEmail" -Type Text -Required
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "UserDepartment" -InternalName "UserDepartment" -Type Text
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "UserRole" -InternalName "UserRole" -Type Text

    # Assignment Details
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "AssignedDate" -InternalName "AssignedDate" -Type DateTime -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "AssignedById" -InternalName "AssignedById" -Type Number
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "AssignmentReason" -InternalName "AssignmentReason" -Type Text -Required

    # JML Integration
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "JMLProcessId" -InternalName "JMLProcessId" -Type Number
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "JMLProcessType" -InternalName "JMLProcessType" -Type Choice -Choices "Joiner","Mover","Leaver"
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "OnboardingStage" -InternalName "OnboardingStage" -Type Choice -Choices "Pre-Start","Day 1","Week 1","Month 1","Month 3"

    # Deadline
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "DueDate" -InternalName "DueDate" -Type DateTime -AddToDefaultView
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "ReadTimeframe" -InternalName "ReadTimeframe" -Type Choice -Choices "Immediate","Day 1","Day 3","Week 1","Week 2","Month 1","Month 3","Month 6","Custom"

    # Progress
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "TotalPolicies" -InternalName "TotalPolicies" -Type Number -AddToDefaultView
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "AcknowledgedPolicies" -InternalName "AcknowledgedPolicies" -Type Number -AddToDefaultView
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "PendingPolicies" -InternalName "PendingPolicies" -Type Number
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "OverduePolicies" -InternalName "OverduePolicies" -Type Number
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "ProgressPercentage" -InternalName "ProgressPercentage" -Type Number -AddToDefaultView

    # Status
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "Status" -InternalName "Status" -Type Choice -Choices "Not Started","In Progress","Completed","Overdue","Exempted" -Required -AddToDefaultView
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "StartedDate" -InternalName "StartedDate" -Type DateTime
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "CompletedDate" -InternalName "CompletedDate" -Type DateTime
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "CompletionDays" -InternalName "CompletionDays" -Type Number

    # Notifications
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "WelcomeEmailSent" -InternalName "WelcomeEmailSent" -Type Boolean
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "WelcomeEmailSentDate" -InternalName "WelcomeEmailSentDate" -Type DateTime
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "TeamsNotificationSent" -InternalName "TeamsNotificationSent" -Type Boolean
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "TeamsNotificationSentDate" -InternalName "TeamsNotificationSentDate" -Type DateTime
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "RemindersSent" -InternalName "RemindersSent" -Type Number
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "LastReminderDate" -InternalName "LastReminderDate" -Type DateTime

    # Link
    Add-PnPField -List "PM_PolicyPackAssignments" -DisplayName "PersonalViewURL" -InternalName "PersonalViewURL" -Type URL

    Write-Host "✓ PM_PolicyPackAssignments list created successfully" -ForegroundColor Green
} else {
    Write-Host "✓ PM_PolicyPackAssignments list already exists" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "============================================================================" -ForegroundColor Cyan
Write-Host "Social Features & Policy Packs Lists Created Successfully!" -ForegroundColor Green
Write-Host "============================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "The following lists have been created:" -ForegroundColor Yellow
Write-Host "  1. PM_PolicyRatings - Rate policies with reviews" -ForegroundColor White
Write-Host "  2. PM_PolicyComments - Comment and discuss policies" -ForegroundColor White
Write-Host "  3. PM_PolicyCommentLikes - Like comments" -ForegroundColor White
Write-Host "  4. PM_PolicyShares - Share policies via email/Teams" -ForegroundColor White
Write-Host "  5. PM_PolicyFollowers - Follow policies for updates" -ForegroundColor White
Write-Host "  6. PM_PolicyPacks - Bundled policy deployment" -ForegroundColor White
Write-Host "  7. PM_PolicyPackAssignments - Track pack assignments" -ForegroundColor White
Write-Host ""
Write-Host "Social Features:" -ForegroundColor Yellow
Write-Host "  ✓ 5-star ratings with reviews" -ForegroundColor Green
Write-Host "  ✓ Threaded comments and discussions" -ForegroundColor Green
Write-Host "  ✓ Comment likes and engagement" -ForegroundColor Green
Write-Host "  ✓ Share via email, Teams, link, QR code" -ForegroundColor Green
Write-Host "  ✓ Follow policies for updates" -ForegroundColor Green
Write-Host ""
Write-Host "Policy Packs:" -ForegroundColor Yellow
Write-Host "  ✓ Bundle policies for onboarding" -ForegroundColor Green
Write-Host "  ✓ Auto-assign by department/role/location" -ForegroundColor Green
Write-Host "  ✓ JML process integration" -ForegroundColor Green
Write-Host "  ✓ Progress tracking and analytics" -ForegroundColor Green
Write-Host "  ✓ Automated email and Teams notifications" -ForegroundColor Green
Write-Host ""

# Disconnect
# Disconnect-PnPOnline
