# ============================================================================
# Policy Manager — Social Lists Provisioning
# Creates 5 social engagement lists: Ratings, Comments, CommentLikes, Shares, Followers
# Assumes: Already connected via Connect-PnPOnline
# ============================================================================

$ErrorActionPreference = "Continue"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Policy Manager — Social Lists Provisioning" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# ============================================================================
# 1. PM_PolicyRatings
# ============================================================================

$listName = "PM_PolicyRatings"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "User ID" -InternalName "UserId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "User Email" -InternalName "UserEmail" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "User Name" -InternalName "UserName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Rating" -InternalName "Rating" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Rating Date" -InternalName "RatingDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Review Title" -InternalName "ReviewTitle" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Review Text" -InternalName "ReviewText" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Review Helpful Count" -InternalName "ReviewHelpfulCount" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Is Verified Reader" -InternalName "IsVerifiedReader" -Type Boolean -ErrorAction SilentlyContinue | Out-Null

# Indexes
Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "UserId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  Provisioned: $listName (10 columns, 2 indexes)" -ForegroundColor Green

# ============================================================================
# 2. PM_PolicyComments
# ============================================================================

$listName = "PM_PolicyComments"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "User ID" -InternalName "UserId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "User Email" -InternalName "UserEmail" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "User Name" -InternalName "UserName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Comment Text" -InternalName "CommentText" -Type Note -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Comment Date" -InternalName "CommentDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Modified Date" -InternalName "ModifiedDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Is Edited" -InternalName "IsEdited" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "IsEdited" -Values @{DefaultValue="0"} -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Parent Comment ID" -InternalName "ParentCommentId" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Reply Count" -InternalName "ReplyCount" -Type Number -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "ReplyCount" -Values @{DefaultValue="0"} -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Like Count" -InternalName "LikeCount" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "LikeCount" -Values @{DefaultValue="0"} -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Staff Response" -InternalName "IsStaffResponse" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "IsStaffResponse" -Values @{DefaultValue="0"} -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Approved" -InternalName "IsApproved" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "IsApproved" -Values @{DefaultValue="1"} -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Deleted" -InternalName "IsDeleted" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "IsDeleted" -Values @{DefaultValue="0"} -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Deleted Reason" -InternalName "DeletedReason" -Type Text -ErrorAction SilentlyContinue | Out-Null

# Indexes
Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "UserId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "ParentCommentId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  Provisioned: $listName (16 columns, 3 indexes)" -ForegroundColor Green

# ============================================================================
# 3. PM_PolicyCommentLikes
# ============================================================================

$listName = "PM_PolicyCommentLikes"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Comment ID" -InternalName "CommentId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "User ID" -InternalName "UserId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "User Email" -InternalName "UserEmail" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Liked Date" -InternalName "LikedDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null

# Indexes
Set-PnPField -List $listName -Identity "CommentId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "UserId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  Provisioned: $listName (4 columns, 2 indexes)" -ForegroundColor Green

# ============================================================================
# 4. PM_PolicyShares
# ============================================================================

$listName = "PM_PolicyShares"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Shared By ID" -InternalName "SharedById" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Shared By Email" -InternalName "SharedByEmail" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Shared With Email" -InternalName "SharedWithEmail" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Shared With Emails" -InternalName "SharedWithEmails" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Shared With User IDs" -InternalName "SharedWithUserIds" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Share Method" -InternalName "ShareMethod" -Type Choice -Choices "Email","Teams","Link","QRCode","Download" -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Share Date" -InternalName "ShareDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Share Message" -InternalName "ShareMessage" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "View Count" -InternalName "ViewCount" -Type Number -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "ViewCount" -Values @{DefaultValue="0"} -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Teams Channel ID" -InternalName "SharedWithTeamsChannelId" -Type Text -ErrorAction SilentlyContinue | Out-Null

# Indexes
Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "SharedById" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  Provisioned: $listName (11 columns, 2 indexes)" -ForegroundColor Green

# ============================================================================
# 5. PM_PolicyFollowers
# ============================================================================

$listName = "PM_PolicyFollowers"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "User ID" -InternalName "UserId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "User Email" -InternalName "UserEmail" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "User Name" -InternalName "UserName" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Followed Date" -InternalName "FollowedDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Notify On Update" -InternalName "NotifyOnUpdate" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "NotifyOnUpdate" -Values @{DefaultValue="1"} -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Notify On Comment" -InternalName "NotifyOnComment" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "NotifyOnComment" -Values @{DefaultValue="1"} -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Notify On New Version" -InternalName "NotifyOnNewVersion" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "NotifyOnNewVersion" -Values @{DefaultValue="1"} -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Email Notifications" -InternalName "EmailNotifications" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "EmailNotifications" -Values @{DefaultValue="1"} -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Teams Notifications" -InternalName "TeamsNotifications" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "TeamsNotifications" -Values @{DefaultValue="1"} -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "In-App Notifications" -InternalName "InAppNotifications" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "InAppNotifications" -Values @{DefaultValue="1"} -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "IsActive" -Values @{DefaultValue="1"} -ErrorAction SilentlyContinue

# Indexes
Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "UserId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  Provisioned: $listName (14 columns, 2 indexes)" -ForegroundColor Green

# ============================================================================
# Summary
# ============================================================================

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  Social Lists provisioned successfully" -ForegroundColor Green
Write-Host "  - PM_PolicyRatings (10 columns)" -ForegroundColor Green
Write-Host "  - PM_PolicyComments (16 columns)" -ForegroundColor Green
Write-Host "  - PM_PolicyCommentLikes (4 columns)" -ForegroundColor Green
Write-Host "  - PM_PolicyShares (11 columns)" -ForegroundColor Green
Write-Host "  - PM_PolicyFollowers (14 columns)" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
