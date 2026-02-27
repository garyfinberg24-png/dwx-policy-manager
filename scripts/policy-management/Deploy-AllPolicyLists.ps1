# ============================================================================
# Policy Manager - Master Deployment Script
# Deploys all 29 Policy Management lists to SharePoint
# Target: https://mf7m.sharepoint.com/sites/PolicyManager (Development)
# ============================================================================
#
# USAGE:
#   .\Deploy-AllPolicyLists.ps1
#
# This will open a browser for authentication.
# ============================================================================

$SiteUrl = "https://mf7m.sharepoint.com/sites/PolicyManager"
$clientId = "d91b5b78-de72-424e-898b-8b5c9512ebd9"
$tenantId = "03bbbdee-d78b-4613-9b99-c468398246b7"

$ErrorActionPreference = "Continue"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Policy Manager - List Provisioning" -ForegroundColor Cyan
Write-Host "  Target: $SiteUrl" -ForegroundColor Cyan
Write-Host "  Environment: DEVELOPMENT" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Check for PnP PowerShell module
$module = Get-Module -ListAvailable -Name "PnP.PowerShell"
if (-not $module) {
    Write-Host "PnP.PowerShell module not found. Installing..." -ForegroundColor Yellow
    Install-Module -Name PnP.PowerShell -Force -AllowClobber -Scope CurrentUser
}

Import-Module PnP.PowerShell -ErrorAction Stop
Write-Host "PnP.PowerShell module loaded" -ForegroundColor Green

# Connect to SharePoint using Device Login
Write-Host ""
Write-Host "Connecting to SharePoint using Device Login..." -ForegroundColor Cyan
Write-Host "Follow the instructions to authenticate in your browser." -ForegroundColor Yellow
try {
    Connect-PnPOnline -Url $SiteUrl -DeviceLogin -ClientId $clientId -Tenant $tenantId
    Write-Host "Connected successfully!" -ForegroundColor Green
} catch {
    Write-Host "Failed to connect to SharePoint: $_" -ForegroundColor Red
    exit 1
}

# Get current web to verify connection
$web = Get-PnPWeb
Write-Host "Connected to: $($web.Title)" -ForegroundColor Green

# ============================================================================
# CREATE LISTS - Inline to avoid script chaining issues
# ============================================================================

# ----------------------------------------------------------------------------
# PART 1: Core Policy Lists
# ----------------------------------------------------------------------------
Write-Host ""
Write-Host "[1/9] Creating Core Policy Lists..." -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Gray

# PM_Policies
$listName = "PM_Policies"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

# Add fields to PM_Policies
$fields = @(
    @{Name="PolicyNumber"; Type="Text"; Required=$true},
    @{Name="PolicyName"; Type="Text"; Required=$true},
    @{Name="PolicyCategory"; Type="Choice"; Choices="HR Policies,IT & Security,Health & Safety,Compliance,Financial,Operational,Legal,Environmental,Quality Assurance,Data Privacy,Custom"},
    @{Name="PolicyType"; Type="Choice"; Choices="Corporate,Departmental,Regional,Role-Specific,Project-Specific,Regulatory"},
    @{Name="PolicyDescription"; Type="Note"},
    @{Name="VersionNumber"; Type="Text"},
    @{Name="VersionType"; Type="Choice"; Choices="Major,Minor,Draft"},
    @{Name="DocumentFormat"; Type="Choice"; Choices="PDF,Word,HTML,Markdown,External Link"},
    @{Name="DocumentURL"; Type="URL"},
    @{Name="HTMLContent"; Type="Note"},
    @{Name="PolicyOwner"; Type="User"},
    @{Name="PolicyAuthors"; Type="UserMulti"},
    @{Name="DepartmentOwner"; Type="Text"},
    @{Name="PolicyStatus"; Type="Choice"; Choices="Draft,In Review,Pending Approval,Approved,Published,Archived,Retired,Expired"},
    @{Name="EffectiveDate"; Type="DateTime"},
    @{Name="ExpiryDate"; Type="DateTime"},
    @{Name="NextReviewDate"; Type="DateTime"},
    @{Name="ReviewCycleMonths"; Type="Number"},
    @{Name="IsActive"; Type="Boolean"},
    @{Name="IsMandatory"; Type="Boolean"},
    @{Name="ComplianceRisk"; Type="Choice"; Choices="Critical,High,Medium,Low,Informational"},
    @{Name="RequiresAcknowledgement"; Type="Boolean"},
    @{Name="AcknowledgementType"; Type="Choice"; Choices="One-Time,Periodic - Annual,Periodic - Quarterly,Periodic - Monthly,On Update,Conditional"},
    @{Name="AcknowledgementDeadlineDays"; Type="Number"},
    @{Name="ReadTimeframe"; Type="Choice"; Choices="Immediate,Day 1,Day 3,Week 1,Week 2,Month 1,Month 3,Month 6,Custom"},
    @{Name="RequiresQuiz"; Type="Boolean"},
    @{Name="QuizPassingScore"; Type="Number"},
    @{Name="DistributionScope"; Type="Choice"; Choices="All Employees,Department,Location,Role,Custom,New Hires Only"},
    @{Name="TotalDistributed"; Type="Number"},
    @{Name="TotalAcknowledged"; Type="Number"},
    @{Name="CompliancePercentage"; Type="Number"}
)

foreach ($field in $fields) {
    try {
        if ($field.Type -eq "Choice") {
            Add-PnPField -List $listName -DisplayName $field.Name -InternalName $field.Name -Type Choice -Choices $field.Choices.Split(",") -ErrorAction SilentlyContinue | Out-Null
        } elseif ($field.Type -eq "UserMulti") {
            Add-PnPField -List $listName -DisplayName $field.Name -InternalName $field.Name -Type UserMulti -ErrorAction SilentlyContinue | Out-Null
        } else {
            $params = @{
                List = $listName
                DisplayName = $field.Name
                InternalName = $field.Name
                Type = $field.Type
                ErrorAction = "SilentlyContinue"
            }
            if ($field.Required) { $params.Required = $true }
            Add-PnPField @params | Out-Null
        }
    } catch { }
}
Write-Host "    Fields added to PM_Policies" -ForegroundColor Gray

# PM_PolicyVersions
$listName = "PM_PolicyVersions"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "VersionNumber" -InternalName "VersionNumber" -Type Text -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "VersionType" -InternalName "VersionType" -Type Choice -Choices "Major","Minor","Draft" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ChangeDescription" -InternalName "ChangeDescription" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "DocumentURL" -InternalName "DocumentURL" -Type URL -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "EffectiveDate" -InternalName "EffectiveDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsCurrentVersion" -InternalName "IsCurrentVersion" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Fields added to PM_PolicyVersions" -ForegroundColor Gray

# PM_PolicyAcknowledgements
$listName = "PM_PolicyAcknowledgements"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "PolicyVersionNumber" -InternalName "PolicyVersionNumber" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "AckUser" -InternalName "AckUser" -Type User -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "UserEmail" -InternalName "UserEmail" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "UserDepartment" -InternalName "UserDepartment" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "AckStatus" -InternalName "AckStatus" -Type Choice -Choices "Not Sent","Sent","Opened","In Progress","Acknowledged","Overdue","Exempted","Failed" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "AssignedDate" -InternalName "AssignedDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "DueDate" -InternalName "DueDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "AcknowledgedDate" -InternalName "AcknowledgedDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "TotalReadTimeSeconds" -InternalName "TotalReadTimeSeconds" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "QuizRequired" -InternalName "QuizRequired" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "QuizStatus" -InternalName "QuizStatus" -Type Choice -Choices "Not Started","In Progress","Passed","Failed","Exempted" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "QuizScore" -InternalName "QuizScore" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsCompliant" -InternalName "IsCompliant" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "OverdueDays" -InternalName "OverdueDays" -Type Number -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Fields added to PM_PolicyAcknowledgements" -ForegroundColor Gray

Write-Host "  Core Policy Lists completed" -ForegroundColor Green

# ----------------------------------------------------------------------------
# PART 2: Quiz Lists
# ----------------------------------------------------------------------------
Write-Host ""
Write-Host "[2/9] Creating Quiz Lists..." -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Gray

# PM_PolicyQuizzes
$listName = "PM_PolicyQuizzes"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "QuizTitle" -InternalName "QuizTitle" -Type Text -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "QuizDescription" -InternalName "QuizDescription" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "PassingScore" -InternalName "PassingScore" -Type Number -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "AllowRetake" -InternalName "AllowRetake" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "MaxAttempts" -InternalName "MaxAttempts" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "TimeLimit" -InternalName "TimeLimit" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "RandomizeQuestions" -InternalName "RandomizeQuestions" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -ErrorAction SilentlyContinue | Out-Null

# PM_PolicyQuizQuestions
$listName = "PM_PolicyQuizQuestions"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "QuizId" -InternalName "QuizId" -Type Number -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "QuestionText" -InternalName "QuestionText" -Type Note -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "QuestionType" -InternalName "QuestionType" -Type Choice -Choices "MultipleChoice","TrueFalse","MultiSelect","ShortAnswer" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Options" -InternalName "Options" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "CorrectAnswer" -InternalName "CorrectAnswer" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Points" -InternalName "Points" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "OrderIndex" -InternalName "OrderIndex" -Type Number -ErrorAction SilentlyContinue | Out-Null

# PM_PolicyQuizResults
$listName = "PM_PolicyQuizResults"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "QuizId" -InternalName "QuizId" -Type Number -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "AcknowledgementId" -InternalName "AcknowledgementId" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "QuizUser" -InternalName "QuizUser" -Type User -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "AttemptNumber" -InternalName "AttemptNumber" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Score" -InternalName "Score" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Percentage" -InternalName "Percentage" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Passed" -InternalName "Passed" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "CompletedDate" -InternalName "CompletedDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null

Write-Host "  Quiz Lists completed" -ForegroundColor Green

# ----------------------------------------------------------------------------
# PART 3: Exemption & Distribution Lists
# ----------------------------------------------------------------------------
Write-Host ""
Write-Host "[3/9] Creating Exemption & Distribution Lists..." -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Gray

# PM_PolicyExemptions
$listName = "PM_PolicyExemptions"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ExemptUser" -InternalName "ExemptUser" -Type User -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ExemptionReason" -InternalName "ExemptionReason" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ExemptionType" -InternalName "ExemptionType" -Type Choice -Choices "Temporary","Permanent","Conditional" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ExemptionStatus" -InternalName "ExemptionStatus" -Type Choice -Choices "Pending","Approved","Denied","Expired","Revoked" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "RequestDate" -InternalName "RequestDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ExpiryDate" -InternalName "ExpiryDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ApprovedBy" -InternalName "ApprovedBy" -Type User -ErrorAction SilentlyContinue | Out-Null

# PM_PolicyDistributions
$listName = "PM_PolicyDistributions"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "DistributionName" -InternalName "DistributionName" -Type Text -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "DistributionScope" -InternalName "DistributionScope" -Type Choice -Choices "All Employees","Department","Location","Role","Custom","New Hires Only" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "DistributedDate" -InternalName "DistributedDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "TargetCount" -InternalName "TargetCount" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "TotalAcknowledged" -InternalName "TotalAcknowledged" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -ErrorAction SilentlyContinue | Out-Null

# PM_PolicyTemplates
$listName = "PM_PolicyTemplates"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "TemplateName" -InternalName "TemplateName" -Type Text -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "TemplateCategory" -InternalName "TemplateCategory" -Type Choice -Choices "HR Policies","IT & Security","Health & Safety","Compliance","Financial","Operational","Legal" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "TemplateDescription" -InternalName "TemplateDescription" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "HTMLTemplate" -InternalName "HTMLTemplate" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -ErrorAction SilentlyContinue | Out-Null

Write-Host "  Exemption & Distribution Lists completed" -ForegroundColor Green

# ----------------------------------------------------------------------------
# PART 4: Social Feature Lists
# ----------------------------------------------------------------------------
Write-Host ""
Write-Host "[4/9] Creating Social Feature Lists..." -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Gray

# PM_PolicyRatings
$listName = "PM_PolicyRatings"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "RatingUser" -InternalName "RatingUser" -Type User -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Rating" -InternalName "Rating" -Type Number -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "RatingDate" -InternalName "RatingDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ReviewText" -InternalName "ReviewText" -Type Note -ErrorAction SilentlyContinue | Out-Null

# PM_PolicyComments
$listName = "PM_PolicyComments"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "CommentUser" -InternalName "CommentUser" -Type User -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "CommentText" -InternalName "CommentText" -Type Note -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "CommentDate" -InternalName "CommentDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ParentCommentId" -InternalName "ParentCommentId" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "LikeCount" -InternalName "LikeCount" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsApproved" -InternalName "IsApproved" -Type Boolean -ErrorAction SilentlyContinue | Out-Null

# PM_PolicyCommentLikes
$listName = "PM_PolicyCommentLikes"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "CommentId" -InternalName "CommentId" -Type Number -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "LikeUser" -InternalName "LikeUser" -Type User -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "LikedDate" -InternalName "LikedDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null

# PM_PolicyShares
$listName = "PM_PolicyShares"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "SharedBy" -InternalName "SharedBy" -Type User -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ShareMethod" -InternalName "ShareMethod" -Type Choice -Choices "Email","Teams","Link","QRCode","Download" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ShareDate" -InternalName "ShareDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null

# PM_PolicyFollowers
$listName = "PM_PolicyFollowers"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "FollowerUser" -InternalName "FollowerUser" -Type User -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "FollowedDate" -InternalName "FollowedDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "NotifyOnUpdate" -InternalName "NotifyOnUpdate" -Type Boolean -ErrorAction SilentlyContinue | Out-Null

Write-Host "  Social Feature Lists completed" -ForegroundColor Green

# ----------------------------------------------------------------------------
# PART 5: Policy Pack Lists
# ----------------------------------------------------------------------------
Write-Host ""
Write-Host "[5/9] Creating Policy Pack Lists..." -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Gray

# PM_PolicyPacks
$listName = "PM_PolicyPacks"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "PackName" -InternalName "PackName" -Type Text -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "PackDescription" -InternalName "PackDescription" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "PackType" -InternalName "PackType" -Type Choice -Choices "Onboarding","Department","Role","Location","Custom" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "PolicyIds" -InternalName "PolicyIds" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "PolicyCount" -InternalName "PolicyCount" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsSequential" -InternalName "IsSequential" -Type Boolean -ErrorAction SilentlyContinue | Out-Null

# PM_PolicyPackAssignments
$listName = "PM_PolicyPackAssignments"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "PackId" -InternalName "PackId" -Type Number -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "AssignedUser" -InternalName "AssignedUser" -Type User -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "AssignedDate" -InternalName "AssignedDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "DueDate" -InternalName "DueDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "TotalPolicies" -InternalName "TotalPolicies" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "AcknowledgedPolicies" -InternalName "AcknowledgedPolicies" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ProgressPercentage" -InternalName "ProgressPercentage" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "AssignmentStatus" -InternalName "AssignmentStatus" -Type Choice -Choices "Not Started","In Progress","Completed","Overdue","Exempted" -ErrorAction SilentlyContinue | Out-Null

Write-Host "  Policy Pack Lists completed" -ForegroundColor Green

# ----------------------------------------------------------------------------
# PART 6: Analytics & Audit Lists
# ----------------------------------------------------------------------------
Write-Host ""
Write-Host "[6/9] Creating Analytics & Audit Lists..." -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Gray

# PM_PolicyAuditLog
$listName = "PM_PolicyAuditLog"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "EntityType" -InternalName "EntityType" -Type Choice -Choices "Policy","Acknowledgement","Exemption","Distribution","Quiz","Template" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "EntityId" -InternalName "EntityId" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "AuditAction" -InternalName "AuditAction" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ActionDescription" -InternalName "ActionDescription" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "PerformedBy" -InternalName "PerformedBy" -Type User -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ActionDate" -InternalName "ActionDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "OldValue" -InternalName "OldValue" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "NewValue" -InternalName "NewValue" -Type Note -ErrorAction SilentlyContinue | Out-Null

# PM_PolicyAnalytics
$listName = "PM_PolicyAnalytics"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "AnalyticsDate" -InternalName "AnalyticsDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "PeriodType" -InternalName "PeriodType" -Type Choice -Choices "Daily","Weekly","Monthly","Quarterly" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "TotalViews" -InternalName "TotalViews" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "TotalAcknowledged" -InternalName "TotalAcknowledged" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ComplianceRate" -InternalName "ComplianceRate" -Type Number -ErrorAction SilentlyContinue | Out-Null

# PM_PolicyFeedback
$listName = "PM_PolicyFeedback"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "FeedbackUser" -InternalName "FeedbackUser" -Type User -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "FeedbackType" -InternalName "FeedbackType" -Type Choice -Choices "Question","Suggestion","Issue","Compliment" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "FeedbackText" -InternalName "FeedbackText" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "FeedbackStatus" -InternalName "FeedbackStatus" -Type Choice -Choices "Open","InProgress","Resolved","Closed" -ErrorAction SilentlyContinue | Out-Null

# PM_PolicyDocuments
$listName = "PM_PolicyDocuments"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "DocumentType" -InternalName "DocumentType" -Type Choice -Choices "Primary","Appendix","Form","Template","Guide","Reference" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "FileName" -InternalName "FileName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "FileURL" -InternalName "FileURL" -Type URL -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "DocumentTitle" -InternalName "DocumentTitle" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -ErrorAction SilentlyContinue | Out-Null

Write-Host "  Analytics & Audit Lists completed" -ForegroundColor Green

# ----------------------------------------------------------------------------
# PART 7: Notification Lists
# ----------------------------------------------------------------------------
Write-Host ""
Write-Host "[7/9] Creating Notification Lists..." -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Gray

# PM_NotificationQueue
$listName = "PM_NotificationQueue"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "NotificationType" -InternalName "NotificationType" -Type Choice -Choices "PolicyShared","PolicyFollowed","PolicyUpdated","PolicyAcknowledgmentRequired","PolicyAcknowledged","PolicyExpiring","PolicyPublished","PolicyComment","Custom" -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "RecipientEmail" -InternalName "RecipientEmail" -Type Text -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "RecipientUserId" -InternalName "RecipientUserId" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "RecipientName" -InternalName "RecipientName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "SenderEmail" -InternalName "SenderEmail" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "SenderUserId" -InternalName "SenderUserId" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "SenderName" -InternalName "SenderName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "PolicyTitle" -InternalName "PolicyTitle" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "PolicyVersion" -InternalName "PolicyVersion" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Message" -InternalName "Message" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Channel" -InternalName "Channel" -Type Choice -Choices "Email","Teams","InApp","All" -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Priority" -InternalName "Priority" -Type Choice -Choices "Low","Normal","High","Urgent" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Status" -InternalName "Status" -Type Choice -Choices "Pending","Processing","Sent","Failed","Retry" -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "RetryCount" -InternalName "RetryCount" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "MaxRetries" -InternalName "MaxRetries" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "LastError" -InternalName "LastError" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ScheduledSendTime" -InternalName "ScheduledSendTime" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "SentTime" -InternalName "SentTime" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "RelatedShareId" -InternalName "RelatedShareId" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "RelatedFollowId" -InternalName "RelatedFollowId" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "TeamsChannelId" -InternalName "TeamsChannelId" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "TeamsTeamId" -InternalName "TeamsTeamId" -Type Text -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Fields added to PM_NotificationQueue" -ForegroundColor Gray

# PM_Notifications
$listName = "PM_Notifications"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Message" -InternalName "Message" -Type Note -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "RecipientId" -InternalName "RecipientId" -Type Number -Required -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Type" -InternalName "Type" -Type Choice -Choices "PolicyShare","PolicyFollow","PolicyUpdate","PolicyAcknowledgment","PolicyExpiring","Policy" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Priority" -InternalName "Priority" -Type Choice -Choices "Low","Normal","High","Urgent" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsRead" -InternalName "IsRead" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "RelatedItemType" -InternalName "RelatedItemType" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "RelatedItemId" -InternalName "RelatedItemId" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ActionUrl" -InternalName "ActionUrl" -Type Text -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Fields added to PM_Notifications" -ForegroundColor Gray

Write-Host "  Notification Lists completed" -ForegroundColor Green

# ----------------------------------------------------------------------------
# PART 8: Admin Configuration Lists
# ----------------------------------------------------------------------------
Write-Host ""
Write-Host "[8/9] Creating Admin Configuration Lists..." -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Gray

# PM_NamingRules
$listName = "PM_NamingRules"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Pattern" -InternalName "Pattern" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Segments" -InternalName "Segments" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "AppliesTo" -InternalName "AppliesTo" -Type Choice -Choices "All Policies","HR Policies","Compliance Policies","IT Policies","Finance Policies","Legal Policies","Operational Policies" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Example" -InternalName "Example" -Type Text -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Fields added to $listName" -ForegroundColor Gray

# PM_SLAConfigs
$listName = "PM_SLAConfigs"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "ProcessType" -InternalName "ProcessType" -Type Choice -Choices "Review","Acknowledgement","Approval","Authoring","Audit","Distribution","Escalation" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "TargetDays" -InternalName "TargetDays" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "WarningThresholdDays" -InternalName "WarningThresholdDays" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Description" -InternalName "Description" -Type Note -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Fields added to $listName" -ForegroundColor Gray

# PM_DataLifecyclePolicies
$listName = "PM_DataLifecyclePolicies"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "EntityType" -InternalName "EntityType" -Type Choice -Choices "Policies","Drafts","Acknowledgements","AuditLogs","Approvals","Quizzes","Notifications","Analytics" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "RetentionPeriodDays" -InternalName "RetentionPeriodDays" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "AutoDeleteEnabled" -InternalName "AutoDeleteEnabled" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ArchiveBeforeDelete" -InternalName "ArchiveBeforeDelete" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Description" -InternalName "Description" -Type Note -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Fields added to $listName" -ForegroundColor Gray

# PM_EmailTemplates
$listName = "PM_EmailTemplates"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "EventTrigger" -InternalName "EventTrigger" -Type Choice -Choices "Policy Published","Ack Overdue","Approval Needed","Policy Expiring","SLA Breached","Violation Found","Campaign Active","User Added","Policy Updated","Policy Retired" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Subject" -InternalName "Subject" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Body" -InternalName "Body" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Recipients" -InternalName "Recipients" -Type Choice -Choices "All Employees","Assigned Users","Approvers","Policy Owners","Managers","Compliance Officers","Target Groups","New Users","HR Team","IT Admins" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "MergeTags" -InternalName "MergeTags" -Type Note -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Fields added to $listName" -ForegroundColor Gray

Write-Host "  Admin Configuration Lists completed" -ForegroundColor Green

# ----------------------------------------------------------------------------
# PART 9: User Management Lists
# ----------------------------------------------------------------------------
Write-Host ""
Write-Host "[9/9] Creating User Management Lists..." -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Gray

# PM_Employees
$listName = "PM_Employees"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "FirstName" -InternalName "FirstName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "LastName" -InternalName "LastName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Email" -InternalName "Email" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "EmployeeNumber" -InternalName "EmployeeNumber" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "JobTitle" -InternalName "JobTitle" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Department" -InternalName "Department" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Location" -InternalName "Location" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "OfficePhone" -InternalName "OfficePhone" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "MobilePhone" -InternalName "MobilePhone" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ManagerEmail" -InternalName "ManagerEmail" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Status" -InternalName "Status" -Type Choice -Choices "Active","Inactive","PreHire","OnLeave","Terminated","Retired" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "EmploymentType" -InternalName "EmploymentType" -Type Choice -Choices "Full-Time","Part-Time","Contractor","Intern","Temporary" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "CostCenter" -InternalName "CostCenter" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "EntraObjectId" -InternalName "EntraObjectId" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "PMRole" -InternalName "PMRole" -Type Choice -Choices "User","Author","Manager","Admin" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "ProfilePhoto" -InternalName "ProfilePhoto" -Type URL -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "LastSyncedAt" -InternalName "LastSyncedAt" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Notes" -InternalName "Notes" -Type Note -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Fields added to $listName" -ForegroundColor Gray

# PM_Sync_Log
$listName = "PM_Sync_Log"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "SyncId" -InternalName "SyncId" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Status" -InternalName "Status" -Type Choice -Choices "Started","Running","Completed","CompletedWithErrors","Failed" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Message" -InternalName "Message" -Type Note -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Fields added to $listName" -ForegroundColor Gray

# PM_Audiences
$listName = "PM_Audiences"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Description" -InternalName "Description" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Criteria" -InternalName "Criteria" -Type Note -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "MemberCount" -InternalName "MemberCount" -Type Number -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "LastEvaluated" -InternalName "LastEvaluated" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Write-Host "    Fields added to $listName" -ForegroundColor Gray

Write-Host "  User Management Lists completed" -ForegroundColor Green

# ============================================================================
# SUMMARY
# ============================================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  DEPLOYMENT COMPLETE!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  29 Policy Management lists created:" -ForegroundColor White
Write-Host ""
Write-Host "  Core (3):        PM_Policies, PM_PolicyVersions, PM_PolicyAcknowledgements" -ForegroundColor Gray
Write-Host "  Quiz (3):        PM_PolicyQuizzes, PM_PolicyQuizQuestions, PM_PolicyQuizResults" -ForegroundColor Gray
Write-Host "  Workflow (3):    PM_PolicyExemptions, PM_PolicyDistributions, PM_PolicyTemplates" -ForegroundColor Gray
Write-Host "  Social (5):      PM_PolicyRatings, PM_PolicyComments, PM_PolicyCommentLikes," -ForegroundColor Gray
Write-Host "                   PM_PolicyShares, PM_PolicyFollowers" -ForegroundColor Gray
Write-Host "  Packs (2):       PM_PolicyPacks, PM_PolicyPackAssignments" -ForegroundColor Gray
Write-Host "  Analytics (4):   PM_PolicyAuditLog, PM_PolicyAnalytics, PM_PolicyFeedback," -ForegroundColor Gray
Write-Host "                   PM_PolicyDocuments" -ForegroundColor Gray
Write-Host "  Notifications (2): PM_NotificationQueue, PM_Notifications" -ForegroundColor Gray
Write-Host "  Admin Config (4): PM_NamingRules, PM_SLAConfigs," -ForegroundColor Gray
Write-Host "                   PM_DataLifecyclePolicies, PM_EmailTemplates" -ForegroundColor Gray
Write-Host "  User Mgmt (3):   PM_Employees, PM_Sync_Log, PM_Audiences" -ForegroundColor Gray
Write-Host ""
Write-Host "  Site: $SiteUrl" -ForegroundColor Yellow
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan

