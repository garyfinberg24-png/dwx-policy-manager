# ============================================================================
# JML Policy Management - Sample Data: Social Features
# Creates ratings, comments, and feedback for policies
# ============================================================================

param(
    [Parameter(Mandatory=$false)]
    [string]$SiteUrl = "https://mf7m.sharepoint.com/sites/PolicyManager"
)

$clientId = "d91b5b78-de72-424e-898b-8b5c9512ebd9"
$tenantId = "03bbbdee-d78b-4613-9b99-c468398246b7"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  JML Policy Management - Social Data" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

Import-Module PnP.PowerShell -ErrorAction Stop
Connect-PnPOnline -Url $SiteUrl -DeviceLogin -ClientId $clientId -Tenant $tenantId
Write-Host "Connected!" -ForegroundColor Green

# ============================================================================
# POLICY RATINGS
# ============================================================================

$ratings = @(
    # Code of Conduct ratings
    @{ PolicyId = 1; Rating = 5; ReviewTitle = "Clear and comprehensive"; ReviewText = "Very well written policy that clearly explains what is expected of us. The examples really help illustrate the points."; UserDepartment = "Marketing"; IsVerifiedReader = $true },
    @{ PolicyId = 1; Rating = 4; ReviewTitle = "Good but lengthy"; ReviewText = "Good content but could be more concise. Took longer than expected to read through."; UserDepartment = "Sales"; IsVerifiedReader = $true },
    @{ PolicyId = 1; Rating = 5; ReviewTitle = "Essential reading"; ReviewText = "Everyone should read this carefully. It sets the tone for our workplace culture."; UserDepartment = "HR"; IsVerifiedReader = $true },
    @{ PolicyId = 1; Rating = 4; ReviewTitle = "Well structured"; ReviewText = "Easy to navigate and find specific topics. The Q&A section at the end is helpful."; UserDepartment = "Engineering"; IsVerifiedReader = $true },

    # Anti-Harassment ratings
    @{ PolicyId = 2; Rating = 5; ReviewTitle = "Important and clear"; ReviewText = "Glad we have such a strong stance on this. The reporting procedures are clearly laid out."; UserDepartment = "HR"; IsVerifiedReader = $true },
    @{ PolicyId = 2; Rating = 5; ReviewTitle = "Makes me feel safe"; ReviewText = "Knowing the company takes this seriously makes a real difference."; UserDepartment = "Customer Service"; IsVerifiedReader = $true },
    @{ PolicyId = 2; Rating = 4; ReviewTitle = "Good coverage"; ReviewText = "Covers all the important areas. Would like more real-world scenario examples."; UserDepartment = "Legal"; IsVerifiedReader = $true },

    # Info Security ratings
    @{ PolicyId = 6; Rating = 4; ReviewTitle = "Thorough but technical"; ReviewText = "Very comprehensive but some sections are quite technical. Could use a simpler summary."; UserDepartment = "Marketing"; IsVerifiedReader = $true },
    @{ PolicyId = 6; Rating = 5; ReviewTitle = "Well done IT team"; ReviewText = "Great balance between security requirements and practical usability."; UserDepartment = "IT"; IsVerifiedReader = $true },
    @{ PolicyId = 6; Rating = 3; ReviewTitle = "Hard to follow"; ReviewText = "The technical jargon makes this difficult for non-IT staff to understand fully."; UserDepartment = "Sales"; IsVerifiedReader = $true },
    @{ PolicyId = 6; Rating = 4; ReviewTitle = "Necessary reading"; ReviewText = "Important information, especially with increasing cyber threats."; UserDepartment = "Finance"; IsVerifiedReader = $true },

    # Remote Work ratings
    @{ PolicyId = 3; Rating = 5; ReviewTitle = "Finally!"; ReviewText = "Great to have clear guidelines on remote work. Very fair and flexible approach."; UserDepartment = "Engineering"; IsVerifiedReader = $true },
    @{ PolicyId = 3; Rating = 4; ReviewTitle = "Good flexibility"; ReviewText = "Appreciate the trust shown to employees. Equipment allowance is generous."; UserDepartment = "Product"; IsVerifiedReader = $true },
    @{ PolicyId = 3; Rating = 5; ReviewTitle = "Well balanced"; ReviewText = "Good balance between flexibility and ensuring team collaboration."; UserDepartment = "Design"; IsVerifiedReader = $true },

    # Health & Safety ratings
    @{ PolicyId = 11; Rating = 4; ReviewTitle = "Comprehensive coverage"; ReviewText = "Covers all the essentials. Good to see mental health included."; UserDepartment = "Operations"; IsVerifiedReader = $true },
    @{ PolicyId = 11; Rating = 5; ReviewTitle = "Safety first"; ReviewText = "Clear procedures for reporting hazards. Training requirements are reasonable."; UserDepartment = "Facilities"; IsVerifiedReader = $true },

    # Data Protection ratings
    @{ PolicyId = 16; Rating = 4; ReviewTitle = "GDPR made clearer"; ReviewText = "Helps understand our GDPR obligations in practical terms."; UserDepartment = "Marketing"; IsVerifiedReader = $true },
    @{ PolicyId = 16; Rating = 3; ReviewTitle = "Complex but necessary"; ReviewText = "Important policy but quite complex. The flowcharts help but still challenging."; UserDepartment = "Sales"; IsVerifiedReader = $true },
    @{ PolicyId = 16; Rating = 5; ReviewTitle = "Essential for our business"; ReviewText = "Critical policy for anyone handling customer data. Well written."; UserDepartment = "Customer Service"; IsVerifiedReader = $true }
)

Write-Host "`n[1/3] Creating policy ratings..." -ForegroundColor Yellow

$today = Get-Date
foreach ($rating in $ratings) {
    try {
        $ratingDate = $today.AddDays(-([int](Get-Random -Minimum 1 -Maximum 180)))
        Add-PnPListItem -List "PM_PolicyRatings" -Values @{
            Title = "Rating for Policy $($rating.PolicyId)"
            PolicyId = $rating.PolicyId
            Rating = $rating.Rating
            RatingDate = $ratingDate
            ReviewTitle = $rating.ReviewTitle
            ReviewText = $rating.ReviewText
            UserDepartment = $rating.UserDepartment
            IsVerifiedReader = $rating.IsVerifiedReader
            ReviewHelpfulCount = Get-Random -Minimum 0 -Maximum 15
        } | Out-Null
    }
    catch {
        Write-Host "  Failed rating for policy $($rating.PolicyId): $_" -ForegroundColor Red
    }
}
Write-Host "  Created $($ratings.Count) ratings" -ForegroundColor Green

# ============================================================================
# POLICY COMMENTS
# ============================================================================

$comments = @(
    # Code of Conduct comments
    @{ PolicyId = 1; CommentText = "Question: Does the gift policy apply to client entertainment as well, or just gifts we receive?"; IsStaffResponse = $false; LikeCount = 8 },
    @{ PolicyId = 1; CommentText = "Good question! Yes, it applies to both giving and receiving. For client entertainment over 100 GBP, please get pre-approval from your manager."; IsStaffResponse = $true; LikeCount = 12; ParentIdx = 0 },
    @{ PolicyId = 1; CommentText = "Thanks for clarifying!"; IsStaffResponse = $false; LikeCount = 2; ParentIdx = 1 },
    @{ PolicyId = 1; CommentText = "The section on social media use is particularly relevant. Good reminder to keep work and personal accounts separate."; IsStaffResponse = $false; LikeCount = 15 },

    # Info Security comments
    @{ PolicyId = 6; CommentText = "Is it okay to use public WiFi if Im connected to the VPN?"; IsStaffResponse = $false; LikeCount = 22 },
    @{ PolicyId = 6; CommentText = "Yes, using the company VPN on public WiFi is acceptable as it encrypts your traffic. However, avoid accessing highly sensitive data on public networks when possible."; IsStaffResponse = $true; LikeCount = 34 },
    @{ PolicyId = 6; CommentText = "The phishing examples were really helpful. I almost clicked on something similar last week!"; IsStaffResponse = $false; LikeCount = 18 },
    @{ PolicyId = 6; CommentText = "Can we get more frequent security awareness updates? The threat landscape changes so fast."; IsStaffResponse = $false; LikeCount = 25 },
    @{ PolicyId = 6; CommentText = "Great suggestion! We are planning quarterly security newsletters. Watch this space."; IsStaffResponse = $true; LikeCount = 19; ParentIdx = 7 },

    # Remote Work comments
    @{ PolicyId = 3; CommentText = "Love the flexibility! Does the equipment allowance cover standing desks?"; IsStaffResponse = $false; LikeCount = 31 },
    @{ PolicyId = 3; CommentText = "Yes, standing desks are covered under the home office equipment allowance. Submit your request through ServiceNow."; IsStaffResponse = $true; LikeCount = 28; ParentIdx = 9 },
    @{ PolicyId = 3; CommentText = "Appreciate the trust this policy shows in employees."; IsStaffResponse = $false; LikeCount = 42 },

    # Data Protection comments
    @{ PolicyId = 16; CommentText = "Whats the process if a customer requests all their data to be deleted?"; IsStaffResponse = $false; LikeCount = 14 },
    @{ PolicyId = 16; CommentText = "Right to erasure requests should be forwarded to privacy@company.com. We have 30 days to respond but aim to complete within 14 days."; IsStaffResponse = $true; LikeCount = 21; ParentIdx = 12 },
    @{ PolicyId = 16; CommentText = "The data classification matrix is really useful. Printed it out for quick reference!"; IsStaffResponse = $false; LikeCount = 16 },

    # Anti-Bribery comments
    @{ PolicyId = 14; CommentText = "What about cultural gift-giving expectations when working with international partners?"; IsStaffResponse = $false; LikeCount = 19 },
    @{ PolicyId = 14; CommentText = "Cultural considerations are addressed in the International Business Supplement. However, any gift over 50 GBP still requires approval regardless of cultural context."; IsStaffResponse = $true; LikeCount = 24; ParentIdx = 15 }
)

Write-Host "`n[2/3] Creating policy comments..." -ForegroundColor Yellow

$commentIds = @{}
$idx = 0
foreach ($comment in $comments) {
    try {
        $commentDate = $today.AddDays(-([int](Get-Random -Minimum 1 -Maximum 120)))
        $values = @{
            Title = "Comment on Policy $($comment.PolicyId)"
            PolicyId = $comment.PolicyId
            CommentText = $comment.CommentText
            CommentDate = $commentDate
            IsStaffResponse = $comment.IsStaffResponse
            LikeCount = $comment.LikeCount
            IsApproved = $true
            IsEdited = $false
            ReplyCount = 0
        }

        if ($comment.ContainsKey('ParentIdx') -and $commentIds.ContainsKey($comment.ParentIdx)) {
            $values.ParentCommentId = $commentIds[$comment.ParentIdx]
        }

        $item = Add-PnPListItem -List "PM_PolicyComments" -Values $values
        $commentIds[$idx] = $item.Id
        $idx++
    }
    catch {
        Write-Host "  Failed comment: $_" -ForegroundColor Red
    }
}
Write-Host "  Created $($comments.Count) comments" -ForegroundColor Green

# ============================================================================
# POLICY FEEDBACK
# ============================================================================

$feedback = @(
    @{ PolicyId = 1; FeedbackType = "Suggestion"; FeedbackText = "Could we add a quick reference card summarising the key points? Would be useful for new starters."; FeedbackStatus = "InProgress"; FeedbackPriority = "Medium"; IsPublic = $true; HelpfulCount = 8 },
    @{ PolicyId = 6; FeedbackType = "Question"; FeedbackText = "Is there a list of approved cloud storage providers we can use?"; FeedbackStatus = "Resolved"; FeedbackPriority = "Medium"; IsPublic = $true; HelpfulCount = 12; ResponseText = "Yes, approved providers are listed in the IT Service Catalog: OneDrive, SharePoint, and approved enterprise Box accounts." },
    @{ PolicyId = 3; FeedbackType = "Compliment"; FeedbackText = "This is one of the best remote work policies Ive seen. Clear, fair, and trusting."; FeedbackStatus = "Closed"; FeedbackPriority = "Low"; IsPublic = $true; HelpfulCount = 23 },
    @{ PolicyId = 16; FeedbackType = "Issue"; FeedbackText = "The link to the data classification tool on page 12 is broken."; FeedbackStatus = "Resolved"; FeedbackPriority = "High"; IsPublic = $false; HelpfulCount = 3; ResponseText = "Thank you for reporting. The link has been fixed." },
    @{ PolicyId = 11; FeedbackType = "Suggestion"; FeedbackText = "Would be helpful to include first aid kit locations for each floor in this policy."; FeedbackStatus = "Open"; FeedbackPriority = "Medium"; IsPublic = $true; HelpfulCount = 15 },
    @{ PolicyId = 2; FeedbackType = "Question"; FeedbackText = "Can harassment training be made available as a refresher outside of mandatory periods?"; FeedbackStatus = "Resolved"; FeedbackPriority = "Medium"; IsPublic = $true; HelpfulCount = 18; ResponseText = "Yes, the training module is available on-demand in the Learning Portal under Compliance Training." },
    @{ PolicyId = 14; FeedbackType = "Suggestion"; FeedbackText = "A flowchart for the gift approval process would be really helpful."; FeedbackStatus = "InProgress"; FeedbackPriority = "Medium"; IsPublic = $true; HelpfulCount = 11 }
)

Write-Host "`n[3/3] Creating policy feedback..." -ForegroundColor Yellow

foreach ($fb in $feedback) {
    try {
        $submitDate = $today.AddDays(-([int](Get-Random -Minimum 5 -Maximum 90)))
        $values = @{
            Title = "$($fb.FeedbackType) for Policy $($fb.PolicyId)"
            PolicyId = $fb.PolicyId
            FeedbackType = $fb.FeedbackType
            FeedbackText = $fb.FeedbackText
            FeedbackStatus = $fb.FeedbackStatus
            FeedbackPriority = $fb.FeedbackPriority
            IsPublic = $fb.IsPublic
            HelpfulCount = $fb.HelpfulCount
            SubmittedDate = $submitDate
            IsAnonymous = $false
        }

        if ($fb.ContainsKey('ResponseText')) {
            $values.ResponseText = $fb.ResponseText
            $values.RespondedDate = $submitDate.AddDays((Get-Random -Minimum 1 -Maximum 7))
        }

        if ($fb.FeedbackStatus -eq "Resolved" -or $fb.FeedbackStatus -eq "Closed") {
            $values.ResolvedDate = $submitDate.AddDays((Get-Random -Minimum 2 -Maximum 14))
        }

        Add-PnPListItem -List "PM_PolicyFeedback" -Values $values | Out-Null
    }
    catch {
        Write-Host "  Failed feedback: $_" -ForegroundColor Red
    }
}
Write-Host "  Created $($feedback.Count) feedback items" -ForegroundColor Green

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  Social data created successfully!" -ForegroundColor Green
Write-Host "  - $($ratings.Count) ratings" -ForegroundColor White
Write-Host "  - $($comments.Count) comments" -ForegroundColor White
Write-Host "  - $($feedback.Count) feedback items" -ForegroundColor White
Write-Host "============================================================" -ForegroundColor Cyan

Disconnect-PnPOnline
