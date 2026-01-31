# ============================================================================
# DWx Policy Manager — Upgrade PM_PolicyQuizQuestions List
# ============================================================================
# This script adds all columns required by the Quiz Builder's AI question
# generator and advanced question types. It is idempotent — columns that
# already exist are silently skipped.
#
# Prerequisites:
#   - PnP.PowerShell module installed
#   - Already connected to SharePoint via Connect-PnPOnline
#
# Usage (run after connecting):
#   .\upgrade-quiz-questions-list.ps1
# ============================================================================

$ErrorActionPreference = "Stop"
$ListName = "PM_PolicyQuizQuestions"

# ============================================================================
# HELPER
# ============================================================================
function Ensure-Field {
    param(
        [string]$FieldName,
        [string]$Type = "Text",
        [string]$DisplayName = "",
        [bool]$Required = $false,
        [string]$DefaultValue = "",
        [string[]]$Choices = @(),
        [string]$Group = "PM Quiz Builder"
    )
    if (-not $DisplayName) { $DisplayName = $FieldName }

    $existingField = Get-PnPField -List $ListName -Identity $FieldName -ErrorAction SilentlyContinue
    if ($existingField) {
        Write-Host "  [EXISTS] $FieldName" -ForegroundColor Gray
        return
    }

    $params = @{
        List         = $ListName
        InternalName = $FieldName
        DisplayName  = $DisplayName
        Type         = $Type
        Group        = $Group
        Required     = $Required
        ErrorAction  = "Stop"
    }

    if ($Type -eq "Choice" -and $Choices.Count -gt 0) {
        $params.Choices = $Choices
    }

    try {
        Add-PnPField @params | Out-Null
        Write-Host "  [ADDED]  $FieldName ($Type)" -ForegroundColor Green

        if ($DefaultValue) {
            Set-PnPField -List $ListName -Identity $FieldName -Values @{DefaultValue = $DefaultValue}
        }
    } catch {
        Write-Host "  [ERROR]  $FieldName — $_" -ForegroundColor Red
    }
}

# ============================================================================
# VERIFY LIST EXISTS
# ============================================================================
Write-Host "`n=== Upgrading $ListName ===" -ForegroundColor Cyan
$list = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
if (-not $list) {
    Write-Host "  List '$ListName' does not exist! Run provision-policy-lists.ps1 first." -ForegroundColor Red
    exit 1
}
Write-Host "  List found: $($list.Title) (Items: $($list.ItemCount))" -ForegroundColor Green

# ============================================================================
# CORE FIELDS (may already exist from original provisioning)
# ============================================================================
Write-Host "`n--- Core Fields ---" -ForegroundColor Yellow

Ensure-Field -FieldName "QuizId"          -Type "Number"  -Required $true -DisplayName "Quiz Id"
Ensure-Field -FieldName "QuestionText"    -Type "Note"    -Required $true -DisplayName "Question Text"
Ensure-Field -FieldName "QuestionType"    -Type "Choice"  -DisplayName "Question Type" -Choices @(
    "Multiple Choice", "True/False", "Multiple Select", "Short Answer",
    "Fill in the Blank", "Matching", "Ordering", "Rating Scale", "Essay",
    "Image Choice", "Hotspot"
)
Ensure-Field -FieldName "CorrectAnswer"   -Type "Note"    -DisplayName "Correct Answer"
Ensure-Field -FieldName "Points"          -Type "Number"  -DisplayName "Points"
Ensure-Field -FieldName "QuestionOrder"   -Type "Number"  -DisplayName "Question Order"

# ============================================================================
# OPTION FIELDS (Multiple Choice / True-False / Image Choice)
# ============================================================================
Write-Host "`n--- Option Fields ---" -ForegroundColor Yellow

Ensure-Field -FieldName "OptionA"         -Type "Note"    -DisplayName "Option A"
Ensure-Field -FieldName "OptionB"         -Type "Note"    -DisplayName "Option B"
Ensure-Field -FieldName "OptionC"         -Type "Note"    -DisplayName "Option C"
Ensure-Field -FieldName "OptionD"         -Type "Note"    -DisplayName "Option D"
Ensure-Field -FieldName "OptionE"         -Type "Note"    -DisplayName "Option E"
Ensure-Field -FieldName "OptionF"         -Type "Note"    -DisplayName "Option F"

# Image-based options
Ensure-Field -FieldName "OptionAImage"    -Type "URL"     -DisplayName "Option A Image"
Ensure-Field -FieldName "OptionBImage"    -Type "URL"     -DisplayName "Option B Image"
Ensure-Field -FieldName "OptionCImage"    -Type "URL"     -DisplayName "Option C Image"
Ensure-Field -FieldName "OptionDImage"    -Type "URL"     -DisplayName "Option D Image"
Ensure-Field -FieldName "QuestionImage"   -Type "URL"     -DisplayName "Question Image"

# ============================================================================
# ADVANCED QUESTION TYPE FIELDS
# ============================================================================
Write-Host "`n--- Advanced Question Type Fields ---" -ForegroundColor Yellow

# Hotspot
Ensure-Field -FieldName "HotspotData"    -Type "Note"    -DisplayName "Hotspot Data (JSON)"

# Matching
Ensure-Field -FieldName "MatchingPairs"  -Type "Note"    -DisplayName "Matching Pairs (JSON)"

# Ordering
Ensure-Field -FieldName "OrderingItems"  -Type "Note"    -DisplayName "Ordering Items (JSON)"

# Fill in the Blank
Ensure-Field -FieldName "BlankAnswers"   -Type "Note"    -DisplayName "Blank Answers (JSON)"

# Multiple Select
Ensure-Field -FieldName "CorrectAnswers" -Type "Note"    -DisplayName "Correct Answers (semicolon-separated)"
Ensure-Field -FieldName "AcceptedAnswers"-Type "Note"    -DisplayName "Accepted Answers (JSON)"

# Rating Scale
Ensure-Field -FieldName "ScaleMin"       -Type "Number"  -DisplayName "Scale Min"
Ensure-Field -FieldName "ScaleMax"       -Type "Number"  -DisplayName "Scale Max"
Ensure-Field -FieldName "ScaleLabels"    -Type "Note"    -DisplayName "Scale Labels (JSON)"
Ensure-Field -FieldName "CorrectRating"  -Type "Number"  -DisplayName "Correct Rating"
Ensure-Field -FieldName "RatingTolerance"-Type "Number"  -DisplayName "Rating Tolerance"

# Essay
Ensure-Field -FieldName "MinWordCount"   -Type "Number"  -DisplayName "Min Word Count"
Ensure-Field -FieldName "MaxWordCount"   -Type "Number"  -DisplayName "Max Word Count"
Ensure-Field -FieldName "RubricId"       -Type "Number"  -DisplayName "Rubric Id"

# ============================================================================
# FEEDBACK & EXPLANATION FIELDS
# ============================================================================
Write-Host "`n--- Feedback & Explanation Fields ---" -ForegroundColor Yellow

Ensure-Field -FieldName "Explanation"        -Type "Note" -DisplayName "Explanation"
Ensure-Field -FieldName "CorrectFeedback"    -Type "Note" -DisplayName "Correct Feedback"
Ensure-Field -FieldName "IncorrectFeedback"  -Type "Note" -DisplayName "Incorrect Feedback"
Ensure-Field -FieldName "PartialFeedback"    -Type "Note" -DisplayName "Partial Feedback"
Ensure-Field -FieldName "Hint"               -Type "Note" -DisplayName "Hint"

# ============================================================================
# SCORING FIELDS
# ============================================================================
Write-Host "`n--- Scoring Fields ---" -ForegroundColor Yellow

Ensure-Field -FieldName "DifficultyLevel"           -Type "Choice"  -DisplayName "Difficulty Level" -Choices @("Easy","Medium","Hard","Expert")
Ensure-Field -FieldName "PartialCreditEnabled"       -Type "Boolean" -DisplayName "Partial Credit Enabled" -DefaultValue "0"
Ensure-Field -FieldName "PartialCreditPercentages"   -Type "Note"    -DisplayName "Partial Credit Percentages (JSON)"
Ensure-Field -FieldName "NegativeMarking"            -Type "Boolean" -DisplayName "Negative Marking" -DefaultValue "0"
Ensure-Field -FieldName "NegativePoints"             -Type "Number"  -DisplayName "Negative Points"

# ============================================================================
# ORGANIZATION FIELDS
# ============================================================================
Write-Host "`n--- Organization Fields ---" -ForegroundColor Yellow

Ensure-Field -FieldName "QuestionBankId"  -Type "Number"  -DisplayName "Question Bank Id"
Ensure-Field -FieldName "QuestionHtml"    -Type "Note"    -DisplayName "Question HTML"
Ensure-Field -FieldName "SectionId"       -Type "Number"  -DisplayName "Section Id"
Ensure-Field -FieldName "SectionName"     -Type "Text"    -DisplayName "Section Name"
Ensure-Field -FieldName "Tags"            -Type "Note"    -DisplayName "Tags"
Ensure-Field -FieldName "Category"        -Type "Text"    -DisplayName "Category"
Ensure-Field -FieldName "TimeLimit"       -Type "Number"  -DisplayName "Time Limit (seconds)"
Ensure-Field -FieldName "DocumentExcerpt" -Type "Note"    -DisplayName "Document Excerpt"

# ============================================================================
# ALSO UPGRADE PM_PolicyQuizzes (add missing fields if needed)
# ============================================================================
Write-Host "`n=== Checking PM_PolicyQuizzes for missing fields ===" -ForegroundColor Cyan
$quizList = Get-PnPList -Identity "PM_PolicyQuizzes" -ErrorAction SilentlyContinue
if ($quizList) {
    $ListName = "PM_PolicyQuizzes"

    # These fields may be missing from the original provisioning
    Ensure-Field -FieldName "QuizCategory"      -Type "Text"    -DisplayName "Quiz Category"
    Ensure-Field -FieldName "DifficultyLevel"    -Type "Choice"  -DisplayName "Difficulty Level" -Choices @("Easy","Medium","Hard","Expert")
    Ensure-Field -FieldName "ScheduleEnabled"    -Type "Boolean" -DisplayName "Schedule Enabled" -DefaultValue "0"
    Ensure-Field -FieldName "ScheduleStartDate"  -Type "DateTime" -DisplayName "Schedule Start Date"
    Ensure-Field -FieldName "ScheduleEndDate"    -Type "DateTime" -DisplayName "Schedule End Date"
    Ensure-Field -FieldName "ScheduleRecurrence" -Type "Text"    -DisplayName "Schedule Recurrence"

    Write-Host "  PM_PolicyQuizzes checked." -ForegroundColor Green
} else {
    Write-Host "  PM_PolicyQuizzes not found — skipping." -ForegroundColor Yellow
}

# ============================================================================
# ENSURE PM_Configuration LIST EXISTS
# ============================================================================
Write-Host "`n=== Ensuring PM_Configuration list ===" -ForegroundColor Cyan
$configList = Get-PnPList -Identity "PM_Configuration" -ErrorAction SilentlyContinue
if (-not $configList) {
    Write-Host "  Creating PM_Configuration list..." -ForegroundColor Yellow
    try {
        New-PnPList -Title "PM_Configuration" -Template GenericList -Url "Lists/PM_Configuration" -ErrorAction Stop | Out-Null
        Write-Host "  [CREATED] PM_Configuration" -ForegroundColor Green

        $ListName = "PM_Configuration"
        Ensure-Field -FieldName "ConfigKey"      -Type "Text"    -DisplayName "Config Key"   -Required $true
        Ensure-Field -FieldName "ConfigValue"     -Type "Note"    -DisplayName "Config Value"
        Ensure-Field -FieldName "Category"        -Type "Text"    -DisplayName "Category"
        Ensure-Field -FieldName "IsActive"        -Type "Boolean" -DisplayName "Is Active"    -DefaultValue "1"
        Ensure-Field -FieldName "IsSystemConfig"  -Type "Boolean" -DisplayName "Is System Config" -DefaultValue "1"

        Write-Host "  PM_Configuration list created with all fields." -ForegroundColor Green
    } catch {
        Write-Host "  [ERROR] Failed to create PM_Configuration: $_" -ForegroundColor Red
    }
} else {
    Write-Host "  [EXISTS] PM_Configuration — skipping." -ForegroundColor Gray
}

# ============================================================================
# SUMMARY
# ============================================================================
Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host "  Upgrade Complete!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Fields added to PM_PolicyQuizQuestions:" -ForegroundColor White
Write-Host "    - Core: QuizId, QuestionText, QuestionType, CorrectAnswer, Points, QuestionOrder"
Write-Host "    - Options: OptionA-F, Option images, QuestionImage"
Write-Host "    - Advanced: HotspotData, MatchingPairs, OrderingItems, BlankAnswers"
Write-Host "    - Multi-select: CorrectAnswers, AcceptedAnswers"
Write-Host "    - Rating: ScaleMin/Max/Labels, CorrectRating, RatingTolerance"
Write-Host "    - Essay: MinWordCount, MaxWordCount, RubricId"
Write-Host "    - Feedback: Explanation, CorrectFeedback, IncorrectFeedback, PartialFeedback, Hint"
Write-Host "    - Scoring: DifficultyLevel, PartialCreditEnabled, NegativeMarking"
Write-Host "    - Organization: QuestionBankId, SectionId, SectionName, Tags, Category, TimeLimit"
Write-Host ""
Write-Host "    - PM_Configuration: ConfigKey, ConfigValue, Category, IsActive, IsSystemConfig"
Write-Host ""
Write-Host "  Next: Re-deploy the .sppkg and test AI question generation." -ForegroundColor Yellow
Write-Host ""
