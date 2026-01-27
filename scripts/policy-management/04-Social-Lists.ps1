# ============================================================================
# DWx Policy Manager - Social Feature Lists
# Part 4: PM_PolicyRatings, PM_PolicyComments, PM_PolicyShares, PM_PolicyFollowers
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
# LIST 10: PM_PolicyRatings
# ============================================================================
Write-Host "`n Creating PM_PolicyRatings list..." -ForegroundColor Yellow

$listName = "PM_PolicyRatings"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User" -InternalName "RatingUser" -Type User -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User Email" -InternalName "UserEmail" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Rating" -InternalName "Rating" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Rating Date" -InternalName "RatingDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Review Title" -InternalName "ReviewTitle" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Review Text" -InternalName "ReviewText" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Review Helpful Count" -InternalName "ReviewHelpfulCount" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Verified Reader" -InternalName "IsVerifiedReader" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User Role" -InternalName "UserRole" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User Department" -InternalName "UserDepartment" -Type Text -ErrorAction SilentlyContinue

Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  PM_PolicyRatings list configured" -ForegroundColor Green

# ============================================================================
# LIST 11: PM_PolicyComments
# ============================================================================
Write-Host "`n Creating PM_PolicyComments list..." -ForegroundColor Yellow

$listName = "PM_PolicyComments"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User" -InternalName "CommentUser" -Type User -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User Email" -InternalName "UserEmail" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Comment Text" -InternalName "CommentText" -Type Note -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Comment Date" -InternalName "CommentDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Modified Date" -InternalName "CommentModifiedDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Edited" -InternalName "IsEdited" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Parent Comment ID" -InternalName "ParentCommentId" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Reply Count" -InternalName "ReplyCount" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Like Count" -InternalName "LikeCount" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Staff Response" -InternalName "IsStaffResponse" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Approved" -InternalName "IsApproved" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Deleted" -InternalName "IsDeleted" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Deleted Reason" -InternalName "DeletedReason" -Type Text -ErrorAction SilentlyContinue

Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "ParentCommentId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  PM_PolicyComments list configured" -ForegroundColor Green

# ============================================================================
# LIST 12: PM_PolicyCommentLikes
# ============================================================================
Write-Host "`n Creating PM_PolicyCommentLikes list..." -ForegroundColor Yellow

$listName = "PM_PolicyCommentLikes"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Comment ID" -InternalName "CommentId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User" -InternalName "LikeUser" -Type User -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Liked Date" -InternalName "LikedDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue

Set-PnPField -List $listName -Identity "CommentId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  PM_PolicyCommentLikes list configured" -ForegroundColor Green

# ============================================================================
# LIST 13: PM_PolicyShares
# ============================================================================
Write-Host "`n Creating PM_PolicyShares list..." -ForegroundColor Yellow

$listName = "PM_PolicyShares"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Shared By" -InternalName "SharedBy" -Type User -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Shared By Email" -InternalName "SharedByEmail" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Share Method" -InternalName "ShareMethod" -Type Choice -Choices "Email","Teams","Link","QRCode","Download" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Share Date" -InternalName "ShareDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Share Message" -InternalName "ShareMessage" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Shared With Emails" -InternalName "SharedWithEmails" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Teams Channel ID" -InternalName "SharedWithTeamsChannelId" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "View Count" -InternalName "ViewCount" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "First Viewed Date" -InternalName "FirstViewedDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Last Viewed Date" -InternalName "LastViewedDate" -Type DateTime -ErrorAction SilentlyContinue

Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  PM_PolicyShares list configured" -ForegroundColor Green

# ============================================================================
# LIST 14: PM_PolicyFollowers
# ============================================================================
Write-Host "`n Creating PM_PolicyFollowers list..." -ForegroundColor Yellow

$listName = "PM_PolicyFollowers"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User" -InternalName "FollowerUser" -Type User -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User Email" -InternalName "UserEmail" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Followed Date" -InternalName "FollowedDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Notify On Update" -InternalName "NotifyOnUpdate" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Notify On Comment" -InternalName "NotifyOnComment" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Notify On New Version" -InternalName "NotifyOnNewVersion" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Email Notifications" -InternalName "EmailNotifications" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Teams Notifications" -InternalName "TeamsNotifications" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "In-App Notifications" -InternalName "InAppNotifications" -Type Boolean -ErrorAction SilentlyContinue

Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "UserEmail" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  PM_PolicyFollowers list configured" -ForegroundColor Green

Write-Host "`n Social feature lists created successfully!" -ForegroundColor Green
Write-Host "   - PM_PolicyRatings" -ForegroundColor White
Write-Host "   - PM_PolicyComments" -ForegroundColor White
Write-Host "   - PM_PolicyCommentLikes" -ForegroundColor White
Write-Host "   - PM_PolicyShares" -ForegroundColor White
Write-Host "   - PM_PolicyFollowers" -ForegroundColor White
