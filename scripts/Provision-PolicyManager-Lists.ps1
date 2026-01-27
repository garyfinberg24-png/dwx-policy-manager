# ============================================================================
# JML Policy Management - Master Deployment Script
# Deploys all 20 Policy Management lists to SharePoint
# Target: https://mf7m.sharepoint.com/sites/JML (Development)
# ============================================================================
#
# USAGE:
#   .\Deploy-AllPolicyLists.ps1
#
# This will open a browser for authentication.
# ============================================================================

$SiteUrl = "https://mf7m.sharepoint.com/sites/JML"
$clientId = "d91b5b78-de72-424e-898b-8b5c9512ebd9"
$tenantId = "03bbbdee-d78b-4613-9b99-c468398246b7"

$ErrorActionPreference = "Continue"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  JML Policy Management - List Provisioning" -ForegroundColor Cyan
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
Write-Host "[1/6] Creating Core Policy Lists..." -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Gray

# JML_Policies
$listName = "JML_Policies"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  Created: $listName" -ForegroundColor Green
} else {
    Write-Host "  Exists: $listName" -ForegroundColor Gray
}

# Add fields to JML_Policies
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
Write-Host "    Fields added to JML_Policies" -ForegroundColor Gray

# JML_PolicyVersions
$listName = "JML_PolicyVersions"
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
Write-Host "    Fields added to JML_PolicyVersions" -ForegroundColor Gray

# JML_PolicyAcknowledgements
$listName = "JML_PolicyAcknowledgements"
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
Write-Host "    Fields added to JML_PolicyAcknowledgements" -ForegroundColor Gray

Write-Host "  Core Policy Lists completed" -ForegroundColor Green

# ----------------------------------------------------------------------------
# PART 2: Quiz Lists
# ----------------------------------------------------------------------------
Write-Host ""
Write-Host "[2/6] Creating Quiz Lists..." -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Gray

# JML_PolicyQuizzes
$listName = "JML_PolicyQuizzes"
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

# JML_PolicyQuizQuestions
$listName = "JML_PolicyQuizQuestions"
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

# JML_PolicyQuizResults
$listName = "JML_PolicyQuizResults"
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
Write-Host "[3/6] Creating Exemption & Distribution Lists..." -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Gray

# JML_PolicyExemptions
$listName = "JML_PolicyExemptions"
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

# JML_PolicyDistributions
$listName = "JML_PolicyDistributions"
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

# JML_PolicyTemplates
$listName = "JML_PolicyTemplates"
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
Write-Host "[4/6] Creating Social Feature Lists..." -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Gray

# JML_PolicyRatings
$listName = "JML_PolicyRatings"
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

# JML_PolicyComments
$listName = "JML_PolicyComments"
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

# JML_PolicyCommentLikes
$listName = "JML_PolicyCommentLikes"
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

# JML_PolicyShares
$listName = "JML_PolicyShares"
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

# JML_PolicyFollowers
$listName = "JML_PolicyFollowers"
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
Write-Host "[5/6] Creating Policy Pack Lists..." -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Gray

# JML_PolicyPacks
$listName = "JML_PolicyPacks"
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
Add-PnPField -List $listName -DisplayName "TargetProcessType" -InternalName "TargetProcessType" -Type Choice -Choices "Joiner","Mover","Leaver" -ErrorAction SilentlyContinue | Out-Null

# JML_PolicyPackAssignments
$listName = "JML_PolicyPackAssignments"
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
Add-PnPField -List $listName -DisplayName "JMLProcessId" -InternalName "JMLProcessId" -Type Number -ErrorAction SilentlyContinue | Out-Null

Write-Host "  Policy Pack Lists completed" -ForegroundColor Green

# ----------------------------------------------------------------------------
# PART 6: Analytics & Audit Lists
# ----------------------------------------------------------------------------
Write-Host ""
Write-Host "[6/6] Creating Analytics & Audit Lists..." -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Gray

# JML_PolicyAuditLog
$listName = "JML_PolicyAuditLog"
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

# JML_PolicyAnalytics
$listName = "JML_PolicyAnalytics"
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

# JML_PolicyFeedback
$listName = "JML_PolicyFeedback"
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

# JML_PolicyDocuments
$listName = "JML_PolicyDocuments"
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

# ============================================================================
# SUMMARY
# ============================================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  DEPLOYMENT COMPLETE!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  20 Policy Management lists created:" -ForegroundColor White
Write-Host ""
Write-Host "  Core (3):        JML_Policies, JML_PolicyVersions, JML_PolicyAcknowledgements" -ForegroundColor Gray
Write-Host "  Quiz (3):        JML_PolicyQuizzes, JML_PolicyQuizQuestions, JML_PolicyQuizResults" -ForegroundColor Gray
Write-Host "  Workflow (3):    JML_PolicyExemptions, JML_PolicyDistributions, JML_PolicyTemplates" -ForegroundColor Gray
Write-Host "  Social (5):      JML_PolicyRatings, JML_PolicyComments, JML_PolicyCommentLikes," -ForegroundColor Gray
Write-Host "                   JML_PolicyShares, JML_PolicyFollowers" -ForegroundColor Gray
Write-Host "  Packs (2):       JML_PolicyPacks, JML_PolicyPackAssignments" -ForegroundColor Gray
Write-Host "  Analytics (4):   JML_PolicyAuditLog, JML_PolicyAnalytics, JML_PolicyFeedback," -ForegroundColor Gray
Write-Host "                   JML_PolicyDocuments" -ForegroundColor Gray
Write-Host ""
Write-Host "  Site: $SiteUrl" -ForegroundColor Yellow
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan

Disconnect-PnPOnline
