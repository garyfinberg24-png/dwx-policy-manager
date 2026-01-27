# ============================================================================
# DWx Policy Manager - Quiz Lists Provisioning
# Part 2: PM_PolicyQuizzes, PM_PolicyQuizQuestions, PM_PolicyQuizResults
# ============================================================================

param(
    [Parameter(Mandatory=$false)]
    [string]$SiteUrl = "https://mf7m.sharepoint.com/sites/JML"
)

$clientId = "d91b5b78-de72-424e-898b-8b5c9512ebd9"
$tenantId = "03bbbdee-d78b-4613-9b99-c468398246b7"

# Connect to SharePoint
Write-Host "Connecting to SharePoint: $SiteUrl" -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -DeviceLogin -ClientId $clientId -Tenant $tenantId

# ============================================================================
# LIST 4: PM_PolicyQuizzes
# ============================================================================
Write-Host "`n Creating PM_PolicyQuizzes list..." -ForegroundColor Yellow

$listName = "PM_PolicyQuizzes"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Policy ID" -InternalName "PolicyId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Quiz Title" -InternalName "QuizTitle" -Type Text -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Quiz Description" -InternalName "QuizDescription" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Passing Score" -InternalName "PassingScore" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Allow Retake" -InternalName "AllowRetake" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Max Attempts" -InternalName "MaxAttempts" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Time Limit (Minutes)" -InternalName "TimeLimit" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Randomize Questions" -InternalName "RandomizeQuestions" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Show Correct Answers" -InternalName "ShowCorrectAnswers" -Type Boolean -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue

Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  PM_PolicyQuizzes list configured" -ForegroundColor Green

# ============================================================================
# LIST 5: PM_PolicyQuizQuestions
# ============================================================================
Write-Host "`n Creating PM_PolicyQuizQuestions list..." -ForegroundColor Yellow

$listName = "PM_PolicyQuizQuestions"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Quiz ID" -InternalName "QuizId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Question Text" -InternalName "QuestionText" -Type Note -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Question Type" -InternalName "QuestionType" -Type Choice -Choices "MultipleChoice","TrueFalse","MultiSelect","ShortAnswer" -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Options" -InternalName "Options" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Correct Answer" -InternalName "CorrectAnswer" -Type Note -Required -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Points" -InternalName "Points" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Explanation" -InternalName "Explanation" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Order Index" -InternalName "OrderIndex" -Type Number -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Mandatory" -InternalName "IsMandatory" -Type Boolean -ErrorAction SilentlyContinue

Set-PnPField -List $listName -Identity "QuizId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  PM_PolicyQuizQuestions list configured" -ForegroundColor Green

# ============================================================================
# LIST 6: PM_PolicyQuizResults
# ============================================================================
Write-Host "`n Creating PM_PolicyQuizResults list..." -ForegroundColor Yellow

$listName = "PM_PolicyQuizResults"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue

if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Add-PnPField -List $listName -DisplayName "Quiz ID" -InternalName "QuizId" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Acknowledgement ID" -InternalName "AcknowledgementId" -Type Number -Required -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "User" -InternalName "QuizUser" -Type User -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Attempt Number" -InternalName "AttemptNumber" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Score" -InternalName "Score" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Percentage" -InternalName "Percentage" -Type Number -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Passed" -InternalName "Passed" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Started Date" -InternalName "StartedDate" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Completed Date" -InternalName "CompletedDate" -Type DateTime -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Time Spent (Seconds)" -InternalName "TimeSpentSeconds" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Answers" -InternalName "Answers" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Correct Answers" -InternalName "CorrectAnswers" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Incorrect Answers" -InternalName "IncorrectAnswers" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Skipped Questions" -InternalName "SkippedQuestions" -Type Number -ErrorAction SilentlyContinue

Set-PnPField -List $listName -Identity "QuizId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "AcknowledgementId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  PM_PolicyQuizResults list configured" -ForegroundColor Green

Write-Host "`n✅ Quiz lists created successfully!" -ForegroundColor Green
Write-Host "   - PM_PolicyQuizzes" -ForegroundColor White
Write-Host "   - PM_PolicyQuizQuestions" -ForegroundColor White
Write-Host "   - PM_PolicyQuizResults" -ForegroundColor White
