# ============================================================================
# DWx Policy Manager — Quiz List Diagnostic & Fix
# Assumes you are already connected via Connect-PnPOnline
# ============================================================================
# Usage:
#   .\Fix-QuizLists.ps1
# ============================================================================

$ErrorActionPreference = "Continue"

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Quiz List Diagnostic & Fix" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# --- Step 1: Check if PM_PolicyQuizzes exists ---
Write-Host "[1] Checking PM_PolicyQuizzes list..." -ForegroundColor Yellow
$list = Get-PnPList -Identity "PM_PolicyQuizzes" -ErrorAction SilentlyContinue

if ($null -eq $list) {
    Write-Host "  LIST DOES NOT EXIST — creating..." -ForegroundColor Red
    New-PnPList -Title "PM_PolicyQuizzes" -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  [CREATED] PM_PolicyQuizzes" -ForegroundColor Green
} else {
    Write-Host "  [EXISTS] PM_PolicyQuizzes (Items: $($list.ItemCount))" -ForegroundColor Green
}

# --- Step 2: Check existing fields ---
Write-Host ""
Write-Host "[2] Checking existing fields on PM_PolicyQuizzes..." -ForegroundColor Yellow
$existingFields = Get-PnPField -List "PM_PolicyQuizzes" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty InternalName
Write-Host "  Found $($existingFields.Count) fields:" -ForegroundColor Gray
$existingFields | ForEach-Object { Write-Host "    - $_" -ForegroundColor DarkGray }

# --- Step 3: Define required fields ---
$requiredFields = @(
    @{ DisplayName = "Policy ID";               InternalName = "PolicyId";               Type = "Number" }
    @{ DisplayName = "Policy Title";             InternalName = "PolicyTitle";             Type = "Text" }
    @{ DisplayName = "Quiz Description";         InternalName = "QuizDescription";         Type = "Note" }
    @{ DisplayName = "Passing Score";            InternalName = "PassingScore";            Type = "Number" }
    @{ DisplayName = "Time Limit (Minutes)";     InternalName = "TimeLimit";               Type = "Number" }
    @{ DisplayName = "Max Attempts";             InternalName = "MaxAttempts";             Type = "Number" }
    @{ DisplayName = "Is Active";                InternalName = "IsActive";                Type = "Boolean" }
    @{ DisplayName = "Question Count";           InternalName = "QuestionCount";           Type = "Number" }
    @{ DisplayName = "Average Score";            InternalName = "AverageScore";            Type = "Number" }
    @{ DisplayName = "Completion Rate";          InternalName = "CompletionRate";          Type = "Number" }
    @{ DisplayName = "Quiz Category";            InternalName = "QuizCategory";            Type = "Text" }
    @{ DisplayName = "Difficulty Level";         InternalName = "DifficultyLevel";         Type = "Choice"; Choices = @("Easy","Medium","Hard","Expert") }
    @{ DisplayName = "Randomize Questions";      InternalName = "RandomizeQuestions";      Type = "Boolean" }
    @{ DisplayName = "Randomize Options";        InternalName = "RandomizeOptions";        Type = "Boolean" }
    @{ DisplayName = "Show Correct Answers";     InternalName = "ShowCorrectAnswers";      Type = "Boolean" }
    @{ DisplayName = "Show Explanations";        InternalName = "ShowExplanations";        Type = "Boolean" }
    @{ DisplayName = "Allow Review";             InternalName = "AllowReview";             Type = "Boolean" }
    @{ DisplayName = "Status";                   InternalName = "Status";                  Type = "Choice"; Choices = @("Draft","Published","Scheduled","Archived") }
    @{ DisplayName = "Grading Type";             InternalName = "GradingType";             Type = "Choice"; Choices = @("Automatic","Manual","Hybrid") }
    @{ DisplayName = "Scheduled Start Date";     InternalName = "ScheduledStartDate";      Type = "DateTime" }
    @{ DisplayName = "Scheduled End Date";       InternalName = "ScheduledEndDate";        Type = "DateTime" }
    @{ DisplayName = "Question Bank ID";         InternalName = "QuestionBankId";          Type = "Number" }
    @{ DisplayName = "Question Pool Size";       InternalName = "QuestionPoolSize";        Type = "Number" }
    @{ DisplayName = "Generate Certificate";     InternalName = "GenerateCertificate";     Type = "Boolean" }
    @{ DisplayName = "Certificate Template ID";  InternalName = "CertificateTemplateId";   Type = "Number" }
    @{ DisplayName = "Per Question Time Limit";  InternalName = "PerQuestionTimeLimit";    Type = "Number" }
    @{ DisplayName = "Allow Partial Credit";     InternalName = "AllowPartialCredit";      Type = "Boolean" }
    @{ DisplayName = "Shuffle Within Sections";  InternalName = "ShuffleWithinSections";   Type = "Boolean" }
    @{ DisplayName = "Require Sequential";       InternalName = "RequireSequentialCompletion"; Type = "Boolean" }
    @{ DisplayName = "Tags";                     InternalName = "Tags";                    Type = "Note" }
)

# --- Step 4: Add missing fields ---
Write-Host ""
Write-Host "[3] Adding missing fields..." -ForegroundColor Yellow
$added = 0
$skipped = 0

foreach ($field in $requiredFields) {
    $exists = Get-PnPField -List "PM_PolicyQuizzes" -Identity $field.InternalName -ErrorAction SilentlyContinue
    if ($exists) {
        $skipped++
        continue
    }

    $params = @{
        List         = "PM_PolicyQuizzes"
        DisplayName  = $field.DisplayName
        InternalName = $field.InternalName
        Type         = $field.Type
    }
    if ($field.Choices) { $params.Choices = $field.Choices }

    try {
        Add-PnPField @params | Out-Null
        Write-Host "  [ADDED]   $($field.InternalName) ($($field.Type))" -ForegroundColor Green
        $added++
    } catch {
        Write-Host "  [FAILED]  $($field.InternalName) — $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "  Fields added: $added, already existed: $skipped" -ForegroundColor Cyan

# --- Step 5: Test creating an item ---
Write-Host ""
Write-Host "[4] Testing item creation with just Title..." -ForegroundColor Yellow
try {
    $testItem = Add-PnPListItem -List "PM_PolicyQuizzes" -Values @{
        Title = "__TEST_QUIZ_DELETE_ME__"
    }
    Write-Host "  [SUCCESS] Item created with ID: $($testItem.Id)" -ForegroundColor Green

    # Clean up test item
    Remove-PnPListItem -List "PM_PolicyQuizzes" -Identity $testItem.Id -Force
    Write-Host "  [CLEANED] Test item removed" -ForegroundColor Gray
} catch {
    Write-Host "  [FAILED]  Could not create item: $($_.Exception.Message)" -ForegroundColor Red
}

# --- Step 6: Test creating with full fields ---
Write-Host ""
Write-Host "[5] Testing item creation with full fields..." -ForegroundColor Yellow
try {
    $fullItem = Add-PnPListItem -List "PM_PolicyQuizzes" -Values @{
        Title              = "__TEST_QUIZ_FULL_DELETE_ME__"
        PassingScore       = 70
        TimeLimit          = 30
        MaxAttempts        = 3
        IsActive           = $true
        QuestionCount      = 0
        QuizCategory       = "General"
        RandomizeQuestions  = $true
        ShowCorrectAnswers = $true
        Status             = "Draft"
    }
    Write-Host "  [SUCCESS] Full item created with ID: $($fullItem.Id)" -ForegroundColor Green

    # Clean up
    Remove-PnPListItem -List "PM_PolicyQuizzes" -Identity $fullItem.Id -Force
    Write-Host "  [CLEANED] Test item removed" -ForegroundColor Gray
} catch {
    Write-Host "  [FAILED]  Could not create full item: $($_.Exception.Message)" -ForegroundColor Red
}

# --- Step 7: Also provision PM_PolicyQuizQuestions if missing ---
Write-Host ""
Write-Host "[6] Checking PM_PolicyQuizQuestions list..." -ForegroundColor Yellow
$qList = Get-PnPList -Identity "PM_PolicyQuizQuestions" -ErrorAction SilentlyContinue
if ($null -eq $qList) {
    Write-Host "  LIST DOES NOT EXIST — creating..." -ForegroundColor Red
    New-PnPList -Title "PM_PolicyQuizQuestions" -Template GenericList -EnableVersioning | Out-Null
    Write-Host "  [CREATED] PM_PolicyQuizQuestions" -ForegroundColor Green

    $QL = "PM_PolicyQuizQuestions"
    $questionFields = @(
        @{ DisplayName = "Quiz ID";              InternalName = "QuizId";              Type = "Number" }
        @{ DisplayName = "Question Text";        InternalName = "QuestionText";        Type = "Note" }
        @{ DisplayName = "Question Type";        InternalName = "QuestionType";        Type = "Choice"; Choices = @("Multiple Choice","True/False","Multiple Select","Short Answer","Fill in the Blank","Matching","Ordering","Rating Scale","Essay","Image Choice","Hotspot") }
        @{ DisplayName = "Question HTML";        InternalName = "QuestionHtml";        Type = "Note" }
        @{ DisplayName = "Option A";             InternalName = "OptionA";             Type = "Note" }
        @{ DisplayName = "Option B";             InternalName = "OptionB";             Type = "Note" }
        @{ DisplayName = "Option C";             InternalName = "OptionC";             Type = "Note" }
        @{ DisplayName = "Option D";             InternalName = "OptionD";             Type = "Note" }
        @{ DisplayName = "Option E";             InternalName = "OptionE";             Type = "Note" }
        @{ DisplayName = "Option F";             InternalName = "OptionF";             Type = "Note" }
        @{ DisplayName = "Correct Answer";       InternalName = "CorrectAnswer";       Type = "Note" }
        @{ DisplayName = "Correct Answers";      InternalName = "CorrectAnswers";      Type = "Note" }
        @{ DisplayName = "Explanation";          InternalName = "Explanation";          Type = "Note" }
        @{ DisplayName = "Points";              InternalName = "Points";              Type = "Number" }
        @{ DisplayName = "Order Index";         InternalName = "OrderIndex";          Type = "Number" }
        @{ DisplayName = "Difficulty Level";     InternalName = "DifficultyLevel";     Type = "Choice"; Choices = @("Easy","Medium","Hard","Expert") }
        @{ DisplayName = "Is Active";           InternalName = "IsActive";            Type = "Boolean" }
        @{ DisplayName = "Is Required";         InternalName = "IsRequired";          Type = "Boolean" }
        @{ DisplayName = "Matching Pairs";      InternalName = "MatchingPairs";       Type = "Note" }
        @{ DisplayName = "Ordering Items";      InternalName = "OrderingItems";       Type = "Note" }
        @{ DisplayName = "Blank Answers";       InternalName = "BlankAnswers";        Type = "Note" }
        @{ DisplayName = "Hotspot Data";        InternalName = "HotspotData";         Type = "Note" }
        @{ DisplayName = "Scale Min";           InternalName = "ScaleMin";            Type = "Number" }
        @{ DisplayName = "Scale Max";           InternalName = "ScaleMax";            Type = "Number" }
        @{ DisplayName = "Correct Rating";      InternalName = "CorrectRating";       Type = "Number" }
        @{ DisplayName = "Rating Tolerance";    InternalName = "RatingTolerance";     Type = "Number" }
        @{ DisplayName = "Min Word Count";      InternalName = "MinWordCount";        Type = "Number" }
        @{ DisplayName = "Max Word Count";      InternalName = "MaxWordCount";        Type = "Number" }
        @{ DisplayName = "Partial Credit Enabled"; InternalName = "PartialCreditEnabled"; Type = "Boolean" }
        @{ DisplayName = "Negative Marking";    InternalName = "NegativeMarking";     Type = "Boolean" }
    )

    foreach ($f in $questionFields) {
        $params = @{
            List         = $QL
            DisplayName  = $f.DisplayName
            InternalName = $f.InternalName
            Type         = $f.Type
        }
        if ($f.Choices) { $params.Choices = $f.Choices }
        try {
            Add-PnPField @params | Out-Null
            Write-Host "    [+] $($f.InternalName)" -ForegroundColor DarkGreen
        } catch {
            Write-Host "    [!] $($f.InternalName) — $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    Write-Host "  PM_PolicyQuizQuestions configured" -ForegroundColor Green
} else {
    Write-Host "  [EXISTS] PM_PolicyQuizQuestions (Items: $($qList.ItemCount))" -ForegroundColor Green
}

# --- Done ---
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Done! If tests passed, Quiz Builder should" -ForegroundColor Cyan
Write-Host "  now be able to create and save quizzes." -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
