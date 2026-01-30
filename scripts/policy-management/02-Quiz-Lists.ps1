# ============================================================================
# DWx Policy Manager — Quiz Lists Provisioning (Complete)
# Provisions all 9 quiz-related lists with full field sets matching QuizService.ts
# ============================================================================
# Usage:
#   Connect-PnPOnline -Url "https://mf7m.sharepoint.com/sites/PolicyManager" -Interactive
#   .\02-Quiz-Lists.ps1
# ============================================================================

$ErrorActionPreference = "Stop"

function Ensure-List {
    param([string]$ListName, [string]$Template = "GenericList")
    $list = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
    if ($null -eq $list) {
        if ($Template -eq "GenericList") {
            New-PnPList -Title $ListName -Template GenericList -EnableVersioning | Out-Null
        } else {
            New-PnPList -Title $ListName -Template $Template | Out-Null
        }
        Write-Host "  [CREATED] $ListName" -ForegroundColor Green
    } else {
        Write-Host "  [EXISTS]  $ListName — adding missing fields" -ForegroundColor Yellow
    }
}

function Add-Field {
    param(
        [string]$List,
        [string]$DisplayName,
        [string]$InternalName,
        [string]$Type,
        [switch]$Required,
        [switch]$AddToDefaultView,
        [switch]$Indexed,
        [string[]]$Choices
    )
    $existing = Get-PnPField -List $List -Identity $InternalName -ErrorAction SilentlyContinue
    if ($existing) {
        return
    }
    $params = @{
        List         = $List
        DisplayName  = $DisplayName
        InternalName = $InternalName
        Type         = $Type
        ErrorAction  = "SilentlyContinue"
    }
    if ($Required) { $params.Required = $true }
    if ($AddToDefaultView) { $params.AddToDefaultView = $true }
    if ($Choices) { $params.Choices = $Choices }

    Add-PnPField @params | Out-Null

    if ($Indexed) {
        Set-PnPField -List $List -Identity $InternalName -Values @{Indexed=$true} -ErrorAction SilentlyContinue
    }
    Write-Host "    [+] $InternalName ($Type)" -ForegroundColor DarkGreen
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Quiz Lists Provisioning (9 lists)" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

# ============================================================================
# LIST 1: PM_PolicyQuizzes — Quiz definitions
# Maps to IQuiz interface in QuizService.ts
# ============================================================================
Write-Host "`n  [1/9] PM_PolicyQuizzes..." -ForegroundColor Yellow
Ensure-List -ListName "PM_PolicyQuizzes"

$L = "PM_PolicyQuizzes"
Add-Field -List $L -DisplayName "Policy ID"               -InternalName "PolicyId"               -Type Number   -Required -AddToDefaultView -Indexed
Add-Field -List $L -DisplayName "Policy Title"             -InternalName "PolicyTitle"             -Type Text
Add-Field -List $L -DisplayName "Quiz Description"         -InternalName "QuizDescription"         -Type Note
Add-Field -List $L -DisplayName "Passing Score"            -InternalName "PassingScore"            -Type Number   -Required -AddToDefaultView
Add-Field -List $L -DisplayName "Time Limit (Minutes)"     -InternalName "TimeLimit"               -Type Number
Add-Field -List $L -DisplayName "Max Attempts"             -InternalName "MaxAttempts"             -Type Number
Add-Field -List $L -DisplayName "Is Active"                -InternalName "IsActive"                -Type Boolean  -AddToDefaultView
Add-Field -List $L -DisplayName "Question Count"           -InternalName "QuestionCount"           -Type Number
Add-Field -List $L -DisplayName "Average Score"            -InternalName "AverageScore"            -Type Number
Add-Field -List $L -DisplayName "Completion Rate"          -InternalName "CompletionRate"          -Type Number
Add-Field -List $L -DisplayName "Quiz Category"            -InternalName "QuizCategory"            -Type Text
Add-Field -List $L -DisplayName "Difficulty Level"         -InternalName "DifficultyLevel"         -Type Choice   -Choices "Easy","Medium","Hard","Expert"
Add-Field -List $L -DisplayName "Randomize Questions"      -InternalName "RandomizeQuestions"      -Type Boolean
Add-Field -List $L -DisplayName "Randomize Options"        -InternalName "RandomizeOptions"        -Type Boolean
Add-Field -List $L -DisplayName "Show Correct Answers"     -InternalName "ShowCorrectAnswers"      -Type Boolean
Add-Field -List $L -DisplayName "Show Explanations"        -InternalName "ShowExplanations"        -Type Boolean
Add-Field -List $L -DisplayName "Allow Review"             -InternalName "AllowReview"             -Type Boolean
Add-Field -List $L -DisplayName "Status"                   -InternalName "Status"                  -Type Choice   -Choices "Draft","Published","Scheduled","Archived" -AddToDefaultView
Add-Field -List $L -DisplayName "Grading Type"             -InternalName "GradingType"             -Type Choice   -Choices "Automatic","Manual","Hybrid"
Add-Field -List $L -DisplayName "Scheduled Start Date"     -InternalName "ScheduledStartDate"      -Type DateTime
Add-Field -List $L -DisplayName "Scheduled End Date"       -InternalName "ScheduledEndDate"        -Type DateTime
Add-Field -List $L -DisplayName "Question Bank ID"         -InternalName "QuestionBankId"          -Type Number
Add-Field -List $L -DisplayName "Question Pool Size"       -InternalName "QuestionPoolSize"        -Type Number
Add-Field -List $L -DisplayName "Generate Certificate"     -InternalName "GenerateCertificate"     -Type Boolean
Add-Field -List $L -DisplayName "Certificate Template ID"  -InternalName "CertificateTemplateId"   -Type Number
Add-Field -List $L -DisplayName "Per Question Time Limit"  -InternalName "PerQuestionTimeLimit"    -Type Number
Add-Field -List $L -DisplayName "Allow Partial Credit"     -InternalName "AllowPartialCredit"      -Type Boolean
Add-Field -List $L -DisplayName "Shuffle Within Sections"  -InternalName "ShuffleWithinSections"   -Type Boolean
Add-Field -List $L -DisplayName "Require Sequential"       -InternalName "RequireSequentialCompletion" -Type Boolean
Add-Field -List $L -DisplayName "Tags"                     -InternalName "Tags"                    -Type Note

Write-Host "  PM_PolicyQuizzes configured" -ForegroundColor Green

# ============================================================================
# LIST 2: PM_PolicyQuizQuestions — Question definitions
# Maps to IQuizQuestion interface in QuizService.ts
# ============================================================================
Write-Host "`n  [2/9] PM_PolicyQuizQuestions..." -ForegroundColor Yellow
Ensure-List -ListName "PM_PolicyQuizQuestions"

$L = "PM_PolicyQuizQuestions"
# Core
Add-Field -List $L -DisplayName "Quiz ID"              -InternalName "QuizId"              -Type Number -Required -AddToDefaultView -Indexed
Add-Field -List $L -DisplayName "Question Bank ID"     -InternalName "QuestionBankId"      -Type Number -Indexed
Add-Field -List $L -DisplayName "Question Text"        -InternalName "QuestionText"        -Type Note   -Required -AddToDefaultView
Add-Field -List $L -DisplayName "Question Type"        -InternalName "QuestionType"        -Type Choice -Required -AddToDefaultView -Choices "Multiple Choice","True/False","Multiple Select","Short Answer","Fill in the Blank","Matching","Ordering","Rating Scale","Essay","Image Choice","Hotspot"
Add-Field -List $L -DisplayName "Question HTML"        -InternalName "QuestionHtml"        -Type Note
# Options A-F
Add-Field -List $L -DisplayName "Option A"             -InternalName "OptionA"             -Type Note
Add-Field -List $L -DisplayName "Option B"             -InternalName "OptionB"             -Type Note
Add-Field -List $L -DisplayName "Option C"             -InternalName "OptionC"             -Type Note
Add-Field -List $L -DisplayName "Option D"             -InternalName "OptionD"             -Type Note
Add-Field -List $L -DisplayName "Option E"             -InternalName "OptionE"             -Type Note
Add-Field -List $L -DisplayName "Option F"             -InternalName "OptionF"             -Type Note
# Image options
Add-Field -List $L -DisplayName "Option A Image"       -InternalName "OptionAImage"        -Type Text
Add-Field -List $L -DisplayName "Option B Image"       -InternalName "OptionBImage"        -Type Text
Add-Field -List $L -DisplayName "Option C Image"       -InternalName "OptionCImage"        -Type Text
Add-Field -List $L -DisplayName "Option D Image"       -InternalName "OptionDImage"        -Type Text
Add-Field -List $L -DisplayName "Question Image"       -InternalName "QuestionImage"       -Type Text
# Type-specific JSON fields
Add-Field -List $L -DisplayName "Hotspot Data"         -InternalName "HotspotData"         -Type Note
Add-Field -List $L -DisplayName "Matching Pairs"       -InternalName "MatchingPairs"       -Type Note
Add-Field -List $L -DisplayName "Ordering Items"       -InternalName "OrderingItems"       -Type Note
Add-Field -List $L -DisplayName "Blank Answers"        -InternalName "BlankAnswers"        -Type Note
Add-Field -List $L -DisplayName "Case Sensitive"       -InternalName "CaseSensitive"       -Type Boolean
# Rating scale
Add-Field -List $L -DisplayName "Scale Min"            -InternalName "ScaleMin"            -Type Number
Add-Field -List $L -DisplayName "Scale Max"            -InternalName "ScaleMax"            -Type Number
Add-Field -List $L -DisplayName "Scale Labels"         -InternalName "ScaleLabels"         -Type Note
Add-Field -List $L -DisplayName "Correct Rating"       -InternalName "CorrectRating"       -Type Number
Add-Field -List $L -DisplayName "Rating Tolerance"     -InternalName "RatingTolerance"     -Type Number
# Essay
Add-Field -List $L -DisplayName "Min Word Count"       -InternalName "MinWordCount"        -Type Number
Add-Field -List $L -DisplayName "Max Word Count"       -InternalName "MaxWordCount"        -Type Number
Add-Field -List $L -DisplayName "Rubric ID"            -InternalName "RubricId"            -Type Number
# Answers
Add-Field -List $L -DisplayName "Correct Answer"       -InternalName "CorrectAnswer"       -Type Note
Add-Field -List $L -DisplayName "Correct Answers"      -InternalName "CorrectAnswers"      -Type Note
Add-Field -List $L -DisplayName "Accepted Answers"     -InternalName "AcceptedAnswers"     -Type Note
# Feedback
Add-Field -List $L -DisplayName "Explanation"          -InternalName "Explanation"          -Type Note
Add-Field -List $L -DisplayName "Correct Feedback"     -InternalName "CorrectFeedback"     -Type Note
Add-Field -List $L -DisplayName "Incorrect Feedback"   -InternalName "IncorrectFeedback"   -Type Note
Add-Field -List $L -DisplayName "Partial Feedback"     -InternalName "PartialFeedback"     -Type Note
Add-Field -List $L -DisplayName "Hint"                 -InternalName "Hint"                -Type Note
# Scoring
Add-Field -List $L -DisplayName "Points"               -InternalName "Points"              -Type Number -Required -AddToDefaultView
Add-Field -List $L -DisplayName "Partial Credit"       -InternalName "PartialCreditEnabled" -Type Boolean
Add-Field -List $L -DisplayName "Partial Credit %"     -InternalName "PartialCreditPercentages" -Type Note
Add-Field -List $L -DisplayName "Negative Marking"     -InternalName "NegativeMarking"     -Type Boolean
Add-Field -List $L -DisplayName "Negative Points"      -InternalName "NegativePoints"      -Type Number
# Organization
Add-Field -List $L -DisplayName "Question Order"       -InternalName "QuestionOrder"       -Type Number -AddToDefaultView
Add-Field -List $L -DisplayName "Section ID"           -InternalName "SectionId"           -Type Number
Add-Field -List $L -DisplayName "Section Name"         -InternalName "SectionName"         -Type Text
Add-Field -List $L -DisplayName "Difficulty Level"     -InternalName "DifficultyLevel"     -Type Choice -Choices "Easy","Medium","Hard","Expert"
Add-Field -List $L -DisplayName "Tags"                 -InternalName "Tags"                -Type Note
Add-Field -List $L -DisplayName "Category"             -InternalName "Category"            -Type Text
Add-Field -List $L -DisplayName "Time Limit"           -InternalName "TimeLimit"           -Type Number
# Status
Add-Field -List $L -DisplayName "Is Active"            -InternalName "IsActive"            -Type Boolean
Add-Field -List $L -DisplayName "Is Required"          -InternalName "IsRequired"          -Type Boolean
# Analytics
Add-Field -List $L -DisplayName "Times Answered"       -InternalName "TimesAnswered"       -Type Number
Add-Field -List $L -DisplayName "Times Correct"        -InternalName "TimesCorrect"        -Type Number
Add-Field -List $L -DisplayName "Average Time"         -InternalName "AverageTime"         -Type Number
Add-Field -List $L -DisplayName "Discrimination Index" -InternalName "DiscriminationIndex" -Type Number

Write-Host "  PM_PolicyQuizQuestions configured" -ForegroundColor Green

# ============================================================================
# LIST 3: PM_PolicyQuizResults — Legacy results (kept for backward compat)
# ============================================================================
Write-Host "`n  [3/9] PM_PolicyQuizResults..." -ForegroundColor Yellow
Ensure-List -ListName "PM_PolicyQuizResults"

$L = "PM_PolicyQuizResults"
Add-Field -List $L -DisplayName "Quiz ID"              -InternalName "QuizId"              -Type Number -Required -AddToDefaultView -Indexed
Add-Field -List $L -DisplayName "Acknowledgement ID"   -InternalName "AcknowledgementId"   -Type Number -Required -Indexed
Add-Field -List $L -DisplayName "User ID"              -InternalName "UserId"              -Type Number -Required -AddToDefaultView
Add-Field -List $L -DisplayName "User Name"            -InternalName "UserName"            -Type Text
Add-Field -List $L -DisplayName "Attempt Number"       -InternalName "AttemptNumber"       -Type Number -Required -AddToDefaultView
Add-Field -List $L -DisplayName "Score"                -InternalName "Score"               -Type Number -Required -AddToDefaultView
Add-Field -List $L -DisplayName "Percentage"           -InternalName "Percentage"          -Type Number -Required -AddToDefaultView
Add-Field -List $L -DisplayName "Passed"               -InternalName "Passed"              -Type Boolean -AddToDefaultView
Add-Field -List $L -DisplayName "Started Date"         -InternalName "StartedDate"         -Type DateTime
Add-Field -List $L -DisplayName "Completed Date"       -InternalName "CompletedDate"       -Type DateTime -AddToDefaultView
Add-Field -List $L -DisplayName "Time Spent (Seconds)" -InternalName "TimeSpentSeconds"    -Type Number
Add-Field -List $L -DisplayName "Answers JSON"         -InternalName "Answers"             -Type Note
Add-Field -List $L -DisplayName "Correct Answers"      -InternalName "CorrectAnswers"      -Type Number
Add-Field -List $L -DisplayName "Incorrect Answers"    -InternalName "IncorrectAnswers"    -Type Number
Add-Field -List $L -DisplayName "Skipped Questions"    -InternalName "SkippedQuestions"    -Type Number

Write-Host "  PM_PolicyQuizResults configured" -ForegroundColor Green

# ============================================================================
# LIST 4: PM_QuizAttempts — Full attempt tracking (primary)
# Maps to IQuizAttempt interface in QuizService.ts
# ============================================================================
Write-Host "`n  [4/9] PM_QuizAttempts..." -ForegroundColor Yellow
Ensure-List -ListName "PM_QuizAttempts"

$L = "PM_QuizAttempts"
Add-Field -List $L -DisplayName "Quiz ID"                -InternalName "QuizId"              -Type Number -Required -AddToDefaultView -Indexed
Add-Field -List $L -DisplayName "Policy ID"              -InternalName "PolicyId"            -Type Number -Required -Indexed
Add-Field -List $L -DisplayName "User ID"                -InternalName "UserId"              -Type Number -Required -AddToDefaultView -Indexed
Add-Field -List $L -DisplayName "User Name"              -InternalName "UserName"            -Type Text   -AddToDefaultView
Add-Field -List $L -DisplayName "User Email"             -InternalName "UserEmail"           -Type Text
Add-Field -List $L -DisplayName "Attempt Number"         -InternalName "AttemptNumber"       -Type Number -Required -AddToDefaultView
Add-Field -List $L -DisplayName "Start Time"             -InternalName "StartTime"           -Type DateTime -Required
Add-Field -List $L -DisplayName "End Time"               -InternalName "EndTime"             -Type DateTime
Add-Field -List $L -DisplayName "Score"                  -InternalName "Score"               -Type Number -AddToDefaultView
Add-Field -List $L -DisplayName "Max Score"              -InternalName "MaxScore"            -Type Number
Add-Field -List $L -DisplayName "Percentage"             -InternalName "Percentage"          -Type Number -AddToDefaultView
Add-Field -List $L -DisplayName "Passed"                 -InternalName "Passed"              -Type Boolean -AddToDefaultView
Add-Field -List $L -DisplayName "Time Spent"             -InternalName "TimeSpent"           -Type Number
Add-Field -List $L -DisplayName "Answers JSON"           -InternalName "AnswersJson"         -Type Note
Add-Field -List $L -DisplayName "Status"                 -InternalName "Status"              -Type Choice -AddToDefaultView -Choices "In Progress","Completed","Abandoned","Expired","Pending Review"
Add-Field -List $L -DisplayName "Points Earned"          -InternalName "PointsEarned"        -Type Number
Add-Field -List $L -DisplayName "Reviewed By ID"         -InternalName "ReviewedById"        -Type Number
Add-Field -List $L -DisplayName "Reviewed Date"          -InternalName "ReviewedDate"        -Type DateTime
Add-Field -List $L -DisplayName "Review Notes"           -InternalName "ReviewNotes"         -Type Note
Add-Field -List $L -DisplayName "Certificate Generated"  -InternalName "CertificateGenerated" -Type Boolean
Add-Field -List $L -DisplayName "Certificate URL"        -InternalName "CertificateUrl"      -Type Text
Add-Field -List $L -DisplayName "Questions Answered"     -InternalName "QuestionsAnswered"    -Type Number
Add-Field -List $L -DisplayName "Questions Correct"      -InternalName "QuestionsCorrect"     -Type Number
Add-Field -List $L -DisplayName "Questions Partial"      -InternalName "QuestionsPartial"     -Type Number
Add-Field -List $L -DisplayName "Questions Incorrect"    -InternalName "QuestionsIncorrect"   -Type Number
Add-Field -List $L -DisplayName "Questions Skipped"      -InternalName "QuestionsSkipped"     -Type Number

Write-Host "  PM_QuizAttempts configured" -ForegroundColor Green

# ============================================================================
# LIST 5: PM_QuestionBanks — Reusable question repositories
# Maps to IQuestionBank interface
# ============================================================================
Write-Host "`n  [5/9] PM_QuestionBanks..." -ForegroundColor Yellow
Ensure-List -ListName "PM_QuestionBanks"

$L = "PM_QuestionBanks"
Add-Field -List $L -DisplayName "Description"      -InternalName "Description"   -Type Note
Add-Field -List $L -DisplayName "Category"         -InternalName "Category"      -Type Text   -AddToDefaultView
Add-Field -List $L -DisplayName "Tags"             -InternalName "Tags"          -Type Note
Add-Field -List $L -DisplayName "Question Count"   -InternalName "QuestionCount" -Type Number -AddToDefaultView
Add-Field -List $L -DisplayName "Is Public"        -InternalName "IsPublic"      -Type Boolean -AddToDefaultView

Write-Host "  PM_QuestionBanks configured" -ForegroundColor Green

# ============================================================================
# LIST 6: PM_QuizSections — Question organization within quizzes
# Maps to IQuizSection interface
# ============================================================================
Write-Host "`n  [6/9] PM_QuizSections..." -ForegroundColor Yellow
Ensure-List -ListName "PM_QuizSections"

$L = "PM_QuizSections"
Add-Field -List $L -DisplayName "Quiz ID"                    -InternalName "QuizId"                    -Type Number -Required -Indexed
Add-Field -List $L -DisplayName "Description"                -InternalName "Description"                -Type Note
Add-Field -List $L -DisplayName "Order"                      -InternalName "Order"                      -Type Number -AddToDefaultView
Add-Field -List $L -DisplayName "Randomize Within Section"   -InternalName "RandomizeWithinSection"     -Type Boolean
Add-Field -List $L -DisplayName "Questions Required"         -InternalName "QuestionsRequired"          -Type Number

Write-Host "  PM_QuizSections configured" -ForegroundColor Green

# ============================================================================
# LIST 7: PM_GradingRubrics — Essay grading criteria
# Maps to IGradingRubric interface
# ============================================================================
Write-Host "`n  [7/9] PM_GradingRubrics..." -ForegroundColor Yellow
Ensure-List -ListName "PM_GradingRubrics"

$L = "PM_GradingRubrics"
Add-Field -List $L -DisplayName "Description"  -InternalName "Description" -Type Note
Add-Field -List $L -DisplayName "Criteria"     -InternalName "Criteria"    -Type Note
Add-Field -List $L -DisplayName "Max Score"    -InternalName "MaxScore"    -Type Number -AddToDefaultView

Write-Host "  PM_GradingRubrics configured" -ForegroundColor Green

# ============================================================================
# LIST 8: PM_QuizCertificates — Generated certificate records
# Maps to ICertificate interface
# ============================================================================
Write-Host "`n  [8/9] PM_QuizCertificates..." -ForegroundColor Yellow
Ensure-List -ListName "PM_QuizCertificates"

$L = "PM_QuizCertificates"
Add-Field -List $L -DisplayName "Attempt ID"          -InternalName "AttemptId"          -Type Number -Required -Indexed
Add-Field -List $L -DisplayName "User ID"             -InternalName "UserId"             -Type Number -Required -Indexed
Add-Field -List $L -DisplayName "User Name"           -InternalName "UserName"           -Type Text   -AddToDefaultView
Add-Field -List $L -DisplayName "Quiz Title"          -InternalName "QuizTitle"          -Type Text   -AddToDefaultView
Add-Field -List $L -DisplayName "Score"               -InternalName "Score"              -Type Number -AddToDefaultView
Add-Field -List $L -DisplayName "Passed Date"         -InternalName "PassedDate"         -Type DateTime -AddToDefaultView
Add-Field -List $L -DisplayName "Certificate Number"  -InternalName "CertificateNumber"  -Type Text   -AddToDefaultView
Add-Field -List $L -DisplayName "Certificate URL"     -InternalName "CertificateUrl"     -Type Text
Add-Field -List $L -DisplayName "Expiry Date"         -InternalName "ExpiryDate"         -Type DateTime

Write-Host "  PM_QuizCertificates configured" -ForegroundColor Green

# ============================================================================
# LIST 9: PM_CertificateTemplates — Certificate HTML templates
# Maps to ICertificateTemplate interface
# ============================================================================
Write-Host "`n  [9/9] PM_CertificateTemplates..." -ForegroundColor Yellow
Ensure-List -ListName "PM_CertificateTemplates"

$L = "PM_CertificateTemplates"
Add-Field -List $L -DisplayName "Template HTML"   -InternalName "TemplateHtml"   -Type Note
Add-Field -List $L -DisplayName "Template Styles" -InternalName "TemplateStyles" -Type Note
Add-Field -List $L -DisplayName "Placeholders"    -InternalName "Placeholders"   -Type Note
Add-Field -List $L -DisplayName "Is Default"      -InternalName "IsDefault"      -Type Boolean -AddToDefaultView

Write-Host "  PM_CertificateTemplates configured" -ForegroundColor Green

# ============================================================================
# DONE
# ============================================================================
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  All 9 quiz lists provisioned!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Lists:" -ForegroundColor White
Write-Host "    1. PM_PolicyQuizzes          (quiz definitions)"
Write-Host "    2. PM_PolicyQuizQuestions     (questions with all 11 types)"
Write-Host "    3. PM_PolicyQuizResults       (legacy results)"
Write-Host "    4. PM_QuizAttempts            (full attempt tracking)"
Write-Host "    5. PM_QuestionBanks           (reusable question repos)"
Write-Host "    6. PM_QuizSections            (quiz organization)"
Write-Host "    7. PM_GradingRubrics          (essay grading criteria)"
Write-Host "    8. PM_QuizCertificates        (generated certificates)"
Write-Host "    9. PM_CertificateTemplates    (certificate HTML templates)"
Write-Host ""
