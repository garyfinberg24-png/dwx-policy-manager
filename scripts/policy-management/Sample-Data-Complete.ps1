# ============================================================================
# DWx Policy Manager - Complete Sample Data for ALL Lists
# Creates realistic, real-world sample data for testing
# ============================================================================

param(
    [Parameter(Mandatory=$false)]
    [string]$SiteUrl = "https://mf7m.sharepoint.com/sites/PolicyManager",

    [Parameter(Mandatory=$false)]
    [switch]$UseWebLogin
)

# Connect to SharePoint
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "DWx Policy Manager - Complete Sample Data" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

Write-Host "Connecting to SharePoint: $SiteUrl" -ForegroundColor Yellow
if ($UseWebLogin) {
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin
} else {
    Connect-PnPOnline -Url $SiteUrl -Interactive
}
Write-Host "Connected successfully!`n" -ForegroundColor Green

# ============================================================================
# SAMPLE USERS (for realistic data - use actual tenant users)
# ============================================================================
$sampleUsers = @(
    @{ Name = "Sarah Mitchell"; Email = "sarah.mitchell@company.com"; Department = "Human Resources"; Role = "HR Manager" },
    @{ Name = "James Anderson"; Email = "james.anderson@company.com"; Department = "IT"; Role = "IT Director" },
    @{ Name = "Emma Thompson"; Email = "emma.thompson@company.com"; Department = "Finance"; Role = "CFO" },
    @{ Name = "Michael Chen"; Email = "michael.chen@company.com"; Department = "Engineering"; Role = "Software Engineer" },
    @{ Name = "Lisa Rodriguez"; Email = "lisa.rodriguez@company.com"; Department = "Marketing"; Role = "Marketing Manager" },
    @{ Name = "David Wilson"; Email = "david.wilson@company.com"; Department = "Operations"; Role = "Operations Lead" },
    @{ Name = "Jennifer Lee"; Email = "jennifer.lee@company.com"; Department = "Legal"; Role = "Legal Counsel" },
    @{ Name = "Robert Taylor"; Email = "robert.taylor@company.com"; Department = "Sales"; Role = "Sales Director" },
    @{ Name = "Amanda Garcia"; Email = "amanda.garcia@company.com"; Department = "Customer Service"; Role = "Support Lead" },
    @{ Name = "Christopher Brown"; Email = "chris.brown@company.com"; Department = "Compliance"; Role = "Compliance Officer" },
    @{ Name = "Jessica Martinez"; Email = "jessica.martinez@company.com"; Department = "Human Resources"; Role = "HR Specialist" },
    @{ Name = "Daniel Kim"; Email = "daniel.kim@company.com"; Department = "IT"; Role = "Security Analyst" },
    @{ Name = "Ashley Johnson"; Email = "ashley.johnson@company.com"; Department = "Finance"; Role = "Financial Analyst" },
    @{ Name = "Matthew Davis"; Email = "matthew.davis@company.com"; Department = "Engineering"; Role = "DevOps Engineer" },
    @{ Name = "Samantha White"; Email = "samantha.white@company.com"; Department = "Marketing"; Role = "Content Specialist" }
)

# Get current user for testing
$currentUser = Get-PnPProperty -ClientObject (Get-PnPWeb) -Property CurrentUser
$currentUserEmail = $currentUser.Email
Write-Host "Current User: $currentUserEmail`n" -ForegroundColor Gray

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================
function Get-RandomDate {
    param(
        [int]$DaysBack = 90,
        [int]$DaysForward = 30
    )
    $randomDays = Get-Random -Minimum (-$DaysBack) -Maximum $DaysForward
    return (Get-Date).AddDays($randomDays)
}

function Get-RandomElement {
    param([array]$Array)
    return $Array[(Get-Random -Maximum $Array.Count)]
}

# ============================================================================
# 1. PM_PolicyVersions - Version History
# ============================================================================
Write-Host "Creating PM_PolicyVersions sample data..." -ForegroundColor Yellow

$versionChanges = @(
    "Initial policy release",
    "Updated compliance requirements per regulatory changes",
    "Added clarification for remote work scenarios",
    "Revised approval process based on stakeholder feedback",
    "Minor grammatical corrections and formatting improvements",
    "Added new section on data classification",
    "Updated contact information and escalation paths",
    "Expanded scope to include contractor requirements",
    "Aligned with ISO 27001 certification requirements",
    "Simplified language for better comprehension",
    "Added FAQ section based on employee questions",
    "Updated to reflect organizational restructuring"
)

# Create version history for first 10 policies
for ($policyId = 1; $policyId -le 10; $policyId++) {
    $numVersions = Get-Random -Minimum 2 -Maximum 5

    for ($v = 1; $v -le $numVersions; $v++) {
        $isCurrentVersion = ($v -eq $numVersions)
        $versionType = if ($v -eq 1) { "Major" } elseif ((Get-Random -Maximum 3) -eq 0) { "Major" } else { "Minor" }
        $majorVer = if ($versionType -eq "Major") { [math]::Ceiling($v / 2) } else { [math]::Floor(($v + 1) / 2) }
        $minorVer = if ($versionType -eq "Minor") { ($v - 1) % 3 } else { 0 }

        $values = @{
            "Title" = "Policy $policyId - Version $majorVer.$minorVer"
            "PolicyId" = $policyId
            "VersionNumber" = "$majorVer.$minorVer"
            "VersionType" = $versionType
            "ChangeDescription" = Get-RandomElement -Array $versionChanges
            "ChangeSummary" = "Version $majorVer.$minorVer changes"
            "EffectiveDate" = (Get-Date).AddDays(-($numVersions - $v) * 90)
            "IsCurrentVersion" = $isCurrentVersion
        }

        try {
            Add-PnPListItem -List "PM_PolicyVersions" -Values $values -ErrorAction SilentlyContinue | Out-Null
        } catch { }
    }
}
Write-Host "  Created policy version history" -ForegroundColor Green

# ============================================================================
# 2. PM_PolicyAcknowledgements - User Acknowledgements
# ============================================================================
Write-Host "Creating PM_PolicyAcknowledgements sample data..." -ForegroundColor Yellow

$acknowledgementStatuses = @("Acknowledged", "Acknowledged", "Acknowledged", "Sent", "Opened", "Overdue", "In Progress")
$quizStatuses = @("Passed", "Passed", "Passed", "Failed", "In Progress", "Not Started")
$deviceTypes = @("Desktop - Windows", "Desktop - Mac", "Mobile - iOS", "Mobile - Android", "Tablet - iPad")

# Create acknowledgements for policies 1-15 with various users
for ($policyId = 1; $policyId -le 15; $policyId++) {
    $numAcks = Get-Random -Minimum 8 -Maximum 15
    $usedUsers = @()

    for ($a = 1; $a -le $numAcks; $a++) {
        $user = Get-RandomElement -Array $sampleUsers
        while ($usedUsers -contains $user.Email) {
            $user = Get-RandomElement -Array $sampleUsers
        }
        $usedUsers += $user.Email

        $status = Get-RandomElement -Array $acknowledgementStatuses
        $assignedDate = Get-RandomDate -DaysBack 60 -DaysForward 0
        $dueDate = $assignedDate.AddDays(14)
        $quizRequired = (Get-Random -Maximum 3) -eq 0

        $values = @{
            "Title" = "Ack-P$policyId-$($user.Name.Split(' ')[0])"
            "PolicyId" = $policyId
            "PolicyVersionNumber" = "1.0"
            "UserEmail" = $user.Email
            "UserDepartment" = $user.Department
            "UserRole" = $user.Role
            "AckStatus" = $status
            "AssignedDate" = $assignedDate
            "DueDate" = $dueDate
            "DeviceType" = Get-RandomElement -Array $deviceTypes
            "QuizRequired" = $quizRequired
        }

        # Add status-specific fields
        if ($status -in @("Opened", "In Progress", "Acknowledged")) {
            $values["FirstOpenedDate"] = $assignedDate.AddDays((Get-Random -Minimum 1 -Maximum 5))
            $values["DocumentOpenCount"] = Get-Random -Minimum 1 -Maximum 8
            $values["TotalReadTimeSeconds"] = Get-Random -Minimum 120 -Maximum 900
            $values["LastAccessedDate"] = Get-RandomDate -DaysBack 10 -DaysForward 0
        }

        if ($status -eq "Acknowledged") {
            $values["AcknowledgedDate"] = $values["FirstOpenedDate"].AddDays((Get-Random -Minimum 0 -Maximum 3))
            $values["AcknowledgementMethod"] = Get-RandomElement -Array @("Click", "Digital Signature", "Checkbox")
            $values["AcknowledgementText"] = "I have read and understood this policy and agree to comply with its requirements."

            if ($quizRequired) {
                $values["QuizStatus"] = Get-RandomElement -Array @("Passed", "Passed", "Passed", "Failed")
                $values["QuizScore"] = Get-Random -Minimum 60 -Maximum 100
                $values["QuizAttempts"] = Get-Random -Minimum 1 -Maximum 3
                $values["QuizCompletedDate"] = $values["AcknowledgedDate"]
            }
        }

        try {
            Add-PnPListItem -List "PM_PolicyAcknowledgements" -Values $values -ErrorAction SilentlyContinue | Out-Null
        } catch { }
    }
}
Write-Host "  Created policy acknowledgements" -ForegroundColor Green

# ============================================================================
# 3. PM_PolicyExemptions - Exemption Requests
# ============================================================================
Write-Host "Creating PM_PolicyExemptions sample data..." -ForegroundColor Yellow

$exemptionReasons = @(
    "Employee is on long-term medical leave and cannot complete the acknowledgement within the required timeframe.",
    "Contractor role does not require access to systems covered by this policy.",
    "Employee is based in a jurisdiction where this policy does not apply.",
    "Role-specific exemption approved by department head - alternative controls in place.",
    "Temporary exemption due to system access issues being resolved by IT.",
    "Employee transferring to a different department where policy does not apply.",
    "Executive exemption for overseas assignment with local compliance requirements.",
    "Accessibility accommodation required - alternative format being prepared."
)

$compensatingControls = @(
    "Employee will complete acknowledgement upon return from leave",
    "Direct supervisor oversight and quarterly compliance check",
    "Local compliance requirements supersede - documented in regional policy",
    "Alternative training completed and documented in HR system",
    "Manual process in place until system access restored",
    "Previous department acknowledgement still valid",
    "Local legal counsel review and approval obtained",
    "Audio version of policy being prepared"
)

$exemptionStatuses = @("Approved", "Approved", "Pending", "Denied", "Expired")

for ($e = 1; $e -le 12; $e++) {
    $user = Get-RandomElement -Array $sampleUsers
    $status = Get-RandomElement -Array $exemptionStatuses
    $requestDate = Get-RandomDate -DaysBack 45 -DaysForward 0
    $exemptionType = Get-RandomElement -Array @("Temporary", "Temporary", "Permanent", "Conditional")

    $values = @{
        "Title" = "Exemption-$e-$($user.Name.Split(' ')[0])"
        "PolicyId" = Get-Random -Minimum 1 -Maximum 15
        "ExemptionReason" = Get-RandomElement -Array $exemptionReasons
        "ExemptionType" = $exemptionType
        "ExemptionStatus" = $status
        "RequestDate" = $requestDate
    }

    if ($exemptionType -eq "Temporary") {
        $values["EffectiveDate"] = $requestDate.AddDays(1)
        $values["ExpiryDate"] = $requestDate.AddDays((Get-Random -Minimum 30 -Maximum 180))
    }

    if ($status -in @("Approved", "Denied")) {
        $values["ReviewedDate"] = $requestDate.AddDays((Get-Random -Minimum 1 -Maximum 5))
        $values["ReviewComments"] = if ($status -eq "Approved") { "Exemption approved with compensating controls." } else { "Exemption denied - employee must complete acknowledgement." }
    }

    if ($status -eq "Approved") {
        $values["ApprovedDate"] = $values["ReviewedDate"]
        $values["CompensatingControls"] = Get-RandomElement -Array $compensatingControls
    }

    try {
        Add-PnPListItem -List "PM_PolicyExemptions" -Values $values -ErrorAction SilentlyContinue | Out-Null
    } catch { }
}
Write-Host "  Created policy exemptions" -ForegroundColor Green

# ============================================================================
# 4. PM_PolicyDistributions - Distribution Campaigns
# ============================================================================
Write-Host "Creating PM_PolicyDistributions sample data..." -ForegroundColor Yellow

$distributionNames = @(
    "Q1 2026 Annual Compliance Refresh",
    "New Employee Onboarding - January Cohort",
    "IT Security Policy Update Rollout",
    "GDPR Annual Re-certification",
    "Health & Safety Policy Update",
    "Remote Work Policy - All Employees",
    "Code of Conduct Annual Refresh",
    "Data Privacy Training Rollout",
    "Anti-Harassment Policy Update",
    "Financial Controls Policy Distribution",
    "New Hire Policy Pack - February",
    "Department-Specific IT Policies - Engineering",
    "Contractor Onboarding Policies",
    "SOX Compliance Annual Review",
    "Cybersecurity Awareness Campaign"
)

$scopes = @("All Employees", "Department", "Role", "New Hires Only", "Custom")

for ($d = 1; $d -le 15; $d++) {
    $distDate = Get-RandomDate -DaysBack 60 -DaysForward 0
    $targetCount = Get-Random -Minimum 50 -Maximum 500
    $ackRate = Get-Random -Minimum 65 -Maximum 98
    $acknowledged = [math]::Floor($targetCount * $ackRate / 100)
    $opened = [math]::Floor($targetCount * (Get-Random -Minimum 80 -Maximum 100) / 100)
    $overdue = [math]::Floor(($targetCount - $acknowledged) * (Get-Random -Minimum 10 -Maximum 40) / 100)

    $values = @{
        "Title" = "Distribution $d"
        "PolicyId" = Get-Random -Minimum 1 -Maximum 22
        "DistributionName" = $distributionNames[$d - 1]
        "DistributionScope" = Get-RandomElement -Array $scopes
        "ScheduledDate" = $distDate.AddDays(-2)
        "DistributedDate" = $distDate
        "TargetCount" = $targetCount
        "TotalSent" = $targetCount
        "TotalDelivered" = $targetCount - (Get-Random -Maximum 5)
        "TotalOpened" = $opened
        "TotalAcknowledged" = $acknowledged
        "TotalOverdue" = $overdue
        "TotalExempted" = Get-Random -Maximum 5
        "TotalFailed" = Get-Random -Maximum 3
        "DueDate" = $distDate.AddDays(14)
        "EscalationEnabled" = $true
        "IsActive" = ($d -le 10)
    }

    if ($d -gt 10) {
        $values["CompletedDate"] = $distDate.AddDays((Get-Random -Minimum 10 -Maximum 20))
    }

    try {
        Add-PnPListItem -List "PM_PolicyDistributions" -Values $values -ErrorAction SilentlyContinue | Out-Null
    } catch { }
}
Write-Host "  Created policy distributions" -ForegroundColor Green

# ============================================================================
# 5. PM_PolicyQuizResults - Quiz Attempt Results
# ============================================================================
Write-Host "Creating PM_PolicyQuizResults sample data..." -ForegroundColor Yellow

for ($q = 1; $q -le 7; $q++) {  # For quizzes 1-7
    $numAttempts = Get-Random -Minimum 15 -Maximum 30

    for ($a = 1; $a -le $numAttempts; $a++) {
        $user = Get-RandomElement -Array $sampleUsers
        $score = Get-Random -Minimum 40 -Maximum 100
        $passed = $score -ge 80
        $startDate = Get-RandomDate -DaysBack 45 -DaysForward 0
        $timeSpent = Get-Random -Minimum 180 -Maximum 900
        $totalQuestions = Get-Random -Minimum 8 -Maximum 12
        $correctAnswers = [math]::Floor($totalQuestions * $score / 100)

        $values = @{
            "Title" = "Quiz $q - $($user.Name.Split(' ')[0]) - Attempt $((Get-Random -Minimum 1 -Maximum 3))"
            "QuizId" = $q
            "AcknowledgementId" = Get-Random -Minimum 1 -Maximum 100
            "AttemptNumber" = Get-Random -Minimum 1 -Maximum 3
            "Score" = $correctAnswers
            "Percentage" = $score
            "Passed" = $passed
            "StartedDate" = $startDate
            "CompletedDate" = $startDate.AddSeconds($timeSpent)
            "TimeSpentSeconds" = $timeSpent
            "CorrectAnswers" = $correctAnswers
            "IncorrectAnswers" = $totalQuestions - $correctAnswers
            "SkippedQuestions" = 0
        }

        # Create answers JSON
        $answers = @()
        for ($i = 1; $i -le $totalQuestions; $i++) {
            $answers += @{
                QuestionId = $i
                Answer = Get-RandomElement -Array @("A", "B", "C", "D")
                Correct = ($i -le $correctAnswers)
            }
        }
        $values["Answers"] = ($answers | ConvertTo-Json -Compress)

        try {
            Add-PnPListItem -List "PM_PolicyQuizResults" -Values $values -ErrorAction SilentlyContinue | Out-Null
        } catch { }
    }
}
Write-Host "  Created quiz results" -ForegroundColor Green

# ============================================================================
# 6. PM_PolicyRatings - Policy Ratings and Reviews
# ============================================================================
Write-Host "Creating PM_PolicyRatings sample data..." -ForegroundColor Yellow

$reviewTitles = @(
    "Clear and helpful",
    "Well written policy",
    "Easy to understand",
    "Comprehensive coverage",
    "Could use more examples",
    "Very informative",
    "Straightforward guidelines",
    "Good reference material",
    "Needs simplification",
    "Excellent resource"
)

$reviewTexts = @(
    "This policy clearly explains our obligations and I found it easy to follow. The examples provided were particularly helpful.",
    "Well-structured document that covers all the important points. Would recommend adding a quick reference guide.",
    "The policy is comprehensive but could benefit from simpler language in some sections.",
    "Excellent policy! The FAQ section answered all my questions.",
    "Good overall, but the approval process section could use more clarity.",
    "Very thorough coverage of the topic. The flowcharts really helped understand the process.",
    "Appreciated the real-world examples. Made it much easier to apply to daily work.",
    "The policy is clear, but it would help to have more visual aids or infographics.",
    "Great update to the previous version. The changes address the issues we had before.",
    "Professional and well-organized. Easy to find the information I needed."
)

for ($r = 1; $r -le 40; $r++) {
    $user = Get-RandomElement -Array $sampleUsers
    $rating = Get-Random -Minimum 3 -Maximum 6  # 3-5 stars
    $hasReview = (Get-Random -Maximum 3) -ne 0

    $values = @{
        "Title" = "Rating-P$((Get-Random -Minimum 1 -Maximum 22))-$r"
        "PolicyId" = Get-Random -Minimum 1 -Maximum 22
        "UserEmail" = $user.Email
        "Rating" = $rating
        "RatingDate" = Get-RandomDate -DaysBack 60 -DaysForward 0
        "IsVerifiedReader" = $true
        "UserRole" = $user.Role
        "UserDepartment" = $user.Department
    }

    if ($hasReview) {
        $values["ReviewTitle"] = Get-RandomElement -Array $reviewTitles
        $values["ReviewText"] = Get-RandomElement -Array $reviewTexts
        $values["ReviewHelpfulCount"] = Get-Random -Maximum 15
    }

    try {
        Add-PnPListItem -List "PM_PolicyRatings" -Values $values -ErrorAction SilentlyContinue | Out-Null
    } catch { }
}
Write-Host "  Created policy ratings" -ForegroundColor Green

# ============================================================================
# 7. PM_PolicyComments - Discussion Comments
# ============================================================================
Write-Host "Creating PM_PolicyComments sample data..." -ForegroundColor Yellow

$comments = @(
    "Does this policy apply to contractors as well, or just full-time employees?",
    "I think section 3.2 needs clarification regarding the escalation process.",
    "Great update! The new remote work section addresses exactly what we needed.",
    "Can someone confirm the approval threshold mentioned in section 4?",
    "This is much clearer than the previous version. Thank you!",
    "We should consider adding examples for the exception handling process.",
    "How does this interact with our regional data privacy requirements?",
    "The training requirements section is very helpful.",
    "Suggestion: Could we add a quick reference card for this policy?",
    "The compliance deadlines seem tight - is there any flexibility?",
    "Well written and easy to follow. Nice work on the revision!",
    "Question about section 5: Does this apply to part-time staff?",
    "The flowchart in appendix A really helps visualize the process.",
    "Minor typo in section 2.3 - 'their' should be 'there'.",
    "Can we get this policy translated for our international teams?"
)

$staffResponses = @(
    "Great question! Yes, this policy applies to all personnel including contractors. We'll add clarification in the next revision.",
    "Thank you for the feedback. We've noted this for the next policy review cycle.",
    "The approval threshold is $5,000 for standard requests. Anything above requires director approval.",
    "We're working on a quick reference card and will publish it next week.",
    "Regional variations are covered in Appendix B. Please let us know if you need additional guidance.",
    "Thank you for catching that typo! We'll correct it in the next update.",
    "Yes, this applies to part-time staff working more than 20 hours per week."
)

$commentId = 1
for ($c = 1; $c -le 30; $c++) {
    $user = Get-RandomElement -Array $sampleUsers
    $policyId = Get-Random -Minimum 1 -Maximum 15
    $isStaffResponse = ($c % 5 -eq 0)
    $commentDate = Get-RandomDate -DaysBack 45 -DaysForward 0

    $values = @{
        "Title" = "Comment-$commentId"
        "PolicyId" = $policyId
        "UserEmail" = $user.Email
        "CommentText" = if ($isStaffResponse) { Get-RandomElement -Array $staffResponses } else { Get-RandomElement -Array $comments }
        "CommentDate" = $commentDate
        "IsEdited" = $false
        "LikeCount" = Get-Random -Maximum 10
        "IsStaffResponse" = $isStaffResponse
        "IsApproved" = $true
        "IsDeleted" = $false
    }

    # Some comments are replies
    if ($c % 4 -eq 0 -and $commentId -gt 1) {
        $values["ParentCommentId"] = Get-Random -Minimum 1 -Maximum ($commentId - 1)
        $values["ReplyCount"] = 0
    } else {
        $values["ReplyCount"] = Get-Random -Maximum 4
    }

    try {
        Add-PnPListItem -List "PM_PolicyComments" -Values $values -ErrorAction SilentlyContinue | Out-Null
        $commentId++
    } catch { }
}
Write-Host "  Created policy comments" -ForegroundColor Green

# ============================================================================
# 8. PM_PolicyCommentLikes - Comment Likes
# ============================================================================
Write-Host "Creating PM_PolicyCommentLikes sample data..." -ForegroundColor Yellow

for ($l = 1; $l -le 50; $l++) {
    $user = Get-RandomElement -Array $sampleUsers

    $values = @{
        "Title" = "Like-$l"
        "CommentId" = Get-Random -Minimum 1 -Maximum 30
        "LikedDate" = Get-RandomDate -DaysBack 30 -DaysForward 0
    }

    try {
        Add-PnPListItem -List "PM_PolicyCommentLikes" -Values $values -ErrorAction SilentlyContinue | Out-Null
    } catch { }
}
Write-Host "  Created comment likes" -ForegroundColor Green

# ============================================================================
# 9. PM_PolicyShares - Policy Shares
# ============================================================================
Write-Host "Creating PM_PolicyShares sample data..." -ForegroundColor Yellow

$shareMethods = @("Email", "Teams", "Link", "Download")
$shareMessages = @(
    "Please review this policy before our team meeting.",
    "FYI - Important update to our department policies.",
    "Required reading for the new project kickoff.",
    "Sharing as discussed in our 1:1.",
    "Please ensure your team reviews this by end of week.",
    "New policy relevant to your role - please acknowledge.",
    "For your reference - updated guidelines.",
    "Action required: Please review and acknowledge."
)

for ($s = 1; $s -le 25; $s++) {
    $user = Get-RandomElement -Array $sampleUsers
    $shareDate = Get-RandomDate -DaysBack 45 -DaysForward 0

    $values = @{
        "Title" = "Share-$s"
        "PolicyId" = Get-Random -Minimum 1 -Maximum 22
        "SharedByEmail" = $user.Email
        "ShareMethod" = Get-RandomElement -Array $shareMethods
        "ShareDate" = $shareDate
        "ShareMessage" = Get-RandomElement -Array $shareMessages
        "ViewCount" = Get-Random -Maximum 20
    }

    if ($values["ViewCount"] -gt 0) {
        $values["FirstViewedDate"] = $shareDate.AddHours((Get-Random -Minimum 1 -Maximum 48))
        $values["LastViewedDate"] = Get-RandomDate -DaysBack 10 -DaysForward 0
    }

    try {
        Add-PnPListItem -List "PM_PolicyShares" -Values $values -ErrorAction SilentlyContinue | Out-Null
    } catch { }
}
Write-Host "  Created policy shares" -ForegroundColor Green

# ============================================================================
# 10. PM_PolicyFollowers - Policy Followers
# ============================================================================
Write-Host "Creating PM_PolicyFollowers sample data..." -ForegroundColor Yellow

for ($f = 1; $f -le 35; $f++) {
    $user = Get-RandomElement -Array $sampleUsers

    $values = @{
        "Title" = "Follower-$f"
        "PolicyId" = Get-Random -Minimum 1 -Maximum 22
        "UserEmail" = $user.Email
        "FollowedDate" = Get-RandomDate -DaysBack 60 -DaysForward 0
        "NotifyOnUpdate" = $true
        "NotifyOnComment" = (Get-Random -Maximum 2) -eq 0
    }

    try {
        Add-PnPListItem -List "PM_PolicyFollowers" -Values $values -ErrorAction SilentlyContinue | Out-Null
    } catch { }
}
Write-Host "  Created policy followers" -ForegroundColor Green

# ============================================================================
# 11. PM_PolicyPackAssignments - Pack Assignments
# ============================================================================
Write-Host "Creating PM_PolicyPackAssignments sample data..." -ForegroundColor Yellow

$assignmentReasons = @(
    "New hire onboarding",
    "Role change - promotion",
    "Department transfer",
    "Annual compliance refresh",
    "Project-specific requirements",
    "Contractor onboarding",
    "Return from leave",
    "Regulatory requirement"
)

$onboardingStages = @("Pre-Start", "Day 1", "Week 1", "Month 1", "Month 3")

for ($pa = 1; $pa -le 30; $pa++) {
    $user = Get-RandomElement -Array $sampleUsers
    $assignedDate = Get-RandomDate -DaysBack 60 -DaysForward 0
    $totalPolicies = Get-Random -Minimum 3 -Maximum 8
    $acknowledgedPolicies = Get-Random -Minimum 0 -Maximum ($totalPolicies + 1)
    $progress = [math]::Floor($acknowledgedPolicies / $totalPolicies * 100)

    $status = if ($acknowledgedPolicies -eq $totalPolicies) { "Completed" }
              elseif ($acknowledgedPolicies -eq 0) { "Not Started" }
              elseif ($assignedDate.AddDays(14) -lt (Get-Date)) { "Overdue" }
              else { "In Progress" }

    $values = @{
        "Title" = "PackAssign-$pa"
        "PackId" = Get-Random -Minimum 1 -Maximum 9
        "UserEmail" = $user.Email
        "UserDepartment" = $user.Department
        "UserRole" = $user.Role
        "AssignedDate" = $assignedDate
        "AssignmentReason" = Get-RandomElement -Array $assignmentReasons
        "OnboardingStage" = Get-RandomElement -Array $onboardingStages
        "DueDate" = $assignedDate.AddDays(14)
        "TotalPolicies" = $totalPolicies
        "AcknowledgedPolicies" = $acknowledgedPolicies
        "PendingPolicies" = $totalPolicies - $acknowledgedPolicies
        "OverduePolicies" = if ($status -eq "Overdue") { $totalPolicies - $acknowledgedPolicies } else { 0 }
        "ProgressPercentage" = $progress
        "AssignmentStatus" = $status
    }

    if ($status -ne "Not Started") {
        $values["StartedDate"] = $assignedDate.AddDays((Get-Random -Minimum 0 -Maximum 3))
    }

    if ($status -eq "Completed") {
        $values["CompletedDate"] = $assignedDate.AddDays((Get-Random -Minimum 5 -Maximum 14))
    }

    try {
        Add-PnPListItem -List "PM_PolicyPackAssignments" -Values $values -ErrorAction SilentlyContinue | Out-Null
    } catch { }
}
Write-Host "  Created policy pack assignments" -ForegroundColor Green

# ============================================================================
# 12. PM_PolicyAuditLog - Audit Trail
# ============================================================================
Write-Host "Creating PM_PolicyAuditLog sample data..." -ForegroundColor Yellow

$auditActions = @(
    @{ Action = "Policy.Created"; Description = "New policy created" },
    @{ Action = "Policy.Updated"; Description = "Policy content updated" },
    @{ Action = "Policy.Published"; Description = "Policy published for distribution" },
    @{ Action = "Policy.Archived"; Description = "Policy archived" },
    @{ Action = "Policy.Viewed"; Description = "Policy document accessed" },
    @{ Action = "Acknowledgement.Assigned"; Description = "Policy assigned to user" },
    @{ Action = "Acknowledgement.Completed"; Description = "User acknowledged policy" },
    @{ Action = "Acknowledgement.Overdue"; Description = "Acknowledgement deadline passed" },
    @{ Action = "Quiz.Started"; Description = "User started quiz attempt" },
    @{ Action = "Quiz.Completed"; Description = "User completed quiz" },
    @{ Action = "Quiz.Passed"; Description = "User passed quiz" },
    @{ Action = "Quiz.Failed"; Description = "User failed quiz" },
    @{ Action = "Exemption.Requested"; Description = "Exemption request submitted" },
    @{ Action = "Exemption.Approved"; Description = "Exemption request approved" },
    @{ Action = "Exemption.Denied"; Description = "Exemption request denied" },
    @{ Action = "Distribution.Started"; Description = "Policy distribution initiated" },
    @{ Action = "Distribution.Completed"; Description = "Policy distribution completed" },
    @{ Action = "Comment.Added"; Description = "Comment posted on policy" },
    @{ Action = "Rating.Submitted"; Description = "User rated policy" }
)

$entityTypes = @("Policy", "Acknowledgement", "Exemption", "Distribution", "Quiz")
$deviceTypes = @("Desktop - Windows", "Desktop - Mac", "Mobile - iOS", "Web Browser")

for ($al = 1; $al -le 100; $al++) {
    $user = Get-RandomElement -Array $sampleUsers
    $auditItem = Get-RandomElement -Array $auditActions
    $actionDate = Get-RandomDate -DaysBack 60 -DaysForward 0

    $entityType = $auditItem.Action.Split('.')[0]

    $values = @{
        "Title" = "Audit-$al"
        "EntityType" = $entityType
        "EntityId" = Get-Random -Minimum 1 -Maximum 50
        "PolicyId" = Get-Random -Minimum 1 -Maximum 22
        "AuditAction" = $auditItem.Action
        "ActionDescription" = $auditItem.Description
        "PerformedByEmail" = $user.Email
        "IPAddress" = "192.168." + (Get-Random -Minimum 1 -Maximum 255) + "." + (Get-Random -Minimum 1 -Maximum 255)
        "DeviceType" = Get-RandomElement -Array $deviceTypes
        "ActionDate" = $actionDate
        "ComplianceRelevant" = ($entityType -in @("Acknowledgement", "Exemption", "Quiz"))
    }

    # Add change details for updates
    if ($auditItem.Action -like "*Updated*" -or $auditItem.Action -like "*Completed*") {
        $values["ChangeDetails"] = "Status changed from previous state to current state"
    }

    try {
        Add-PnPListItem -List "PM_PolicyAuditLog" -Values $values -ErrorAction SilentlyContinue | Out-Null
    } catch { }
}
Write-Host "  Created audit log entries" -ForegroundColor Green

# ============================================================================
# 13. PM_PolicyAnalytics - Analytics Data
# ============================================================================
Write-Host "Creating PM_PolicyAnalytics sample data..." -ForegroundColor Yellow

$periodTypes = @("Daily", "Weekly", "Monthly")

for ($policyId = 1; $policyId -le 15; $policyId++) {
    # Create monthly analytics for last 3 months
    for ($month = 0; $month -le 2; $month++) {
        $analyticsDate = (Get-Date).AddMonths(-$month).Date
        $totalAssigned = Get-Random -Minimum 50 -Maximum 200
        $ackRate = Get-Random -Minimum 70 -Maximum 98
        $totalAck = [math]::Floor($totalAssigned * $ackRate / 100)

        $values = @{
            "Title" = "Analytics-P$policyId-M$month"
            "PolicyId" = $policyId
            "AnalyticsDate" = $analyticsDate
            "PeriodType" = "Monthly"
            "TotalViews" = Get-Random -Minimum 100 -Maximum 500
            "UniqueViewers" = Get-Random -Minimum 50 -Maximum 200
            "AverageReadTimeSeconds" = Get-Random -Minimum 180 -Maximum 600
            "TotalDownloads" = Get-Random -Minimum 10 -Maximum 50
            "TotalAssigned" = $totalAssigned
            "TotalAcknowledged" = $totalAck
            "TotalOverdue" = [math]::Floor(($totalAssigned - $totalAck) * 0.3)
            "ComplianceRate" = $ackRate
            "AverageTimeToAcknowledgeDays" = Get-Random -Minimum 2 -Maximum 10
            "TotalQuizAttempts" = Get-Random -Minimum 20 -Maximum 80
            "AverageQuizScore" = Get-Random -Minimum 75 -Maximum 95
            "QuizPassRate" = Get-Random -Minimum 80 -Maximum 98
            "TotalFeedback" = Get-Random -Minimum 5 -Maximum 20
            "PositiveFeedback" = Get-Random -Minimum 3 -Maximum 15
            "NegativeFeedback" = Get-Random -Maximum 5
            "HighRiskNonCompliance" = Get-Random -Maximum 3
            "EscalatedCases" = Get-Random -Maximum 2
        }

        try {
            Add-PnPListItem -List "PM_PolicyAnalytics" -Values $values -ErrorAction SilentlyContinue | Out-Null
        } catch { }
    }
}
Write-Host "  Created analytics data" -ForegroundColor Green

# ============================================================================
# 14. PM_PolicyFeedback - User Feedback
# ============================================================================
Write-Host "Creating PM_PolicyFeedback sample data..." -ForegroundColor Yellow

$feedbackItems = @(
    @{ Type = "Question"; Text = "Does section 3.2 about data retention apply to temporary files as well?" },
    @{ Type = "Question"; Text = "What is the process for requesting an extension on the acknowledgement deadline?" },
    @{ Type = "Suggestion"; Text = "It would be helpful to add a visual flowchart for the approval process." },
    @{ Type = "Suggestion"; Text = "Consider adding a FAQ section based on common questions from employees." },
    @{ Type = "Issue"; Text = "The link to the supporting document in section 4 is broken." },
    @{ Type = "Issue"; Text = "Some terminology in section 2 is inconsistent with our other policies." },
    @{ Type = "Compliment"; Text = "The new layout is much easier to navigate. Great improvement!" },
    @{ Type = "Compliment"; Text = "Thank you for simplifying the language. This is much clearer now." },
    @{ Type = "Question"; Text = "How does this policy interact with our existing contractor agreements?" },
    @{ Type = "Suggestion"; Text = "Would be useful to have a summary at the end for quick reference." },
    @{ Type = "Issue"; Text = "The mobile version of the policy document is difficult to read on smaller screens." },
    @{ Type = "Question"; Text = "Is there a grace period for first-time violations of this policy?" }
)

$feedbackStatuses = @("Open", "Open", "InProgress", "Resolved", "Closed")
$priorities = @("Medium", "Medium", "Low", "High", "Critical")

$responses = @(
    "Thank you for your question. Yes, section 3.2 applies to all files including temporary ones. We'll add clarification in the next revision.",
    "Extensions can be requested through your manager, who can approve up to 7 additional days.",
    "Great suggestion! We're working on adding visual aids and will include this in the next update.",
    "Thank you for the feedback. We've fixed the broken link.",
    "We appreciate the positive feedback! We're glad the new format is helpful."
)

for ($fb = 1; $fb -le 20; $fb++) {
    $user = Get-RandomElement -Array $sampleUsers
    $feedbackItem = Get-RandomElement -Array $feedbackItems
    $status = Get-RandomElement -Array $feedbackStatuses
    $submittedDate = Get-RandomDate -DaysBack 45 -DaysForward 0

    $values = @{
        "Title" = "Feedback-$fb"
        "PolicyId" = Get-Random -Minimum 1 -Maximum 22
        "FeedbackType" = $feedbackItem.Type
        "FeedbackText" = $feedbackItem.Text
        "IsAnonymous" = (Get-Random -Maximum 4) -eq 0
        "FeedbackStatus" = $status
        "FeedbackPriority" = Get-RandomElement -Array $priorities
        "IsPublic" = (Get-Random -Maximum 2) -eq 0
        "HelpfulCount" = Get-Random -Maximum 10
        "SubmittedDate" = $submittedDate
    }

    if ($status -in @("Resolved", "Closed", "InProgress")) {
        $values["ResponseText"] = Get-RandomElement -Array $responses
        $values["RespondedDate"] = $submittedDate.AddDays((Get-Random -Minimum 1 -Maximum 5))
    }

    if ($status -in @("Resolved", "Closed")) {
        $values["ResolvedDate"] = $values["RespondedDate"].AddDays((Get-Random -Minimum 1 -Maximum 3))
    }

    try {
        Add-PnPListItem -List "PM_PolicyFeedback" -Values $values -ErrorAction SilentlyContinue | Out-Null
    } catch { }
}
Write-Host "  Created feedback items" -ForegroundColor Green

# ============================================================================
# 15. PM_PolicyTemplates - Policy Templates
# ============================================================================
Write-Host "Creating PM_PolicyTemplates sample data..." -ForegroundColor Yellow

$templates = @(
    @{ Name = "Standard Corporate Policy"; Category = "Corporate"; Risk = "Medium" },
    @{ Name = "IT Security Policy Template"; Category = "IT & Security"; Risk = "High" },
    @{ Name = "HR Policy Template"; Category = "HR Policies"; Risk = "Medium" },
    @{ Name = "Data Privacy Policy Template"; Category = "Data Privacy"; Risk = "Critical" },
    @{ Name = "Health & Safety Template"; Category = "Health & Safety"; Risk = "High" },
    @{ Name = "Financial Controls Template"; Category = "Financial"; Risk = "High" },
    @{ Name = "Compliance Policy Template"; Category = "Compliance"; Risk = "Critical" },
    @{ Name = "Operational Procedure Template"; Category = "Operational"; Risk = "Low" },
    @{ Name = "Employee Handbook Section"; Category = "HR Policies"; Risk = "Medium" },
    @{ Name = "GDPR Compliance Template"; Category = "Data Privacy"; Risk = "Critical" }
)

for ($t = 0; $t -lt $templates.Count; $t++) {
    $template = $templates[$t]

    $values = @{
        "Title" = "Template-$($t + 1)"
        "TemplateName" = $template.Name
        "TemplateCategory" = $template.Category
        "TemplateDescription" = "Standard template for creating $($template.Category.ToLower()) policies with pre-defined sections and formatting."
        "DefaultAcknowledgementType" = Get-RandomElement -Array @("One-Time", "Periodic - Annual", "On Update")
        "DefaultDeadlineDays" = Get-RandomElement -Array @(7, 14, 30)
        "DefaultRequiresQuiz" = ($template.Risk -in @("High", "Critical"))
        "DefaultReviewCycleMonths" = Get-RandomElement -Array @(6, 12, 24)
        "DefaultComplianceRisk" = $template.Risk
        "UsageCount" = Get-Random -Minimum 5 -Maximum 50
        "IsActive" = $true
        "HTMLTemplate" = "<h1>Policy Title</h1><h2>1. Purpose</h2><p>[Enter purpose]</p><h2>2. Scope</h2><p>[Enter scope]</p><h2>3. Policy Statement</h2><p>[Enter policy]</p><h2>4. Responsibilities</h2><p>[Enter responsibilities]</p><h2>5. Compliance</h2><p>[Enter compliance requirements]</p>"
    }

    try {
        Add-PnPListItem -List "PM_PolicyTemplates" -Values $values -ErrorAction SilentlyContinue | Out-Null
    } catch { }
}
Write-Host "  Created policy templates" -ForegroundColor Green

# ============================================================================
# SUMMARY
# ============================================================================
Write-Host "`n========================================" -ForegroundColor Green
Write-Host "Sample Data Creation Complete!" -ForegroundColor Green
Write-Host "========================================`n" -ForegroundColor Green

Write-Host "Created sample data for:" -ForegroundColor White
Write-Host "  - PM_PolicyVersions (version history)" -ForegroundColor Gray
Write-Host "  - PM_PolicyAcknowledgements (user acknowledgements)" -ForegroundColor Gray
Write-Host "  - PM_PolicyExemptions (exemption requests)" -ForegroundColor Gray
Write-Host "  - PM_PolicyDistributions (distribution campaigns)" -ForegroundColor Gray
Write-Host "  - PM_PolicyQuizResults (quiz attempts)" -ForegroundColor Gray
Write-Host "  - PM_PolicyRatings (ratings and reviews)" -ForegroundColor Gray
Write-Host "  - PM_PolicyComments (discussion comments)" -ForegroundColor Gray
Write-Host "  - PM_PolicyCommentLikes (comment likes)" -ForegroundColor Gray
Write-Host "  - PM_PolicyShares (policy shares)" -ForegroundColor Gray
Write-Host "  - PM_PolicyFollowers (policy followers)" -ForegroundColor Gray
Write-Host "  - PM_PolicyPackAssignments (pack assignments)" -ForegroundColor Gray
Write-Host "  - PM_PolicyAuditLog (audit trail)" -ForegroundColor Gray
Write-Host "  - PM_PolicyAnalytics (analytics data)" -ForegroundColor Gray
Write-Host "  - PM_PolicyFeedback (user feedback)" -ForegroundColor Gray
Write-Host "  - PM_PolicyTemplates (policy templates)" -ForegroundColor Gray

Write-Host "`nNote: User fields require actual SharePoint users." -ForegroundColor Yellow
Write-Host "      Email addresses are placeholders for demonstration." -ForegroundColor Yellow
