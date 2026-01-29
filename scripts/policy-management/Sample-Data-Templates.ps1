# ============================================================================
# JML Policy Management - Sample Data: Templates & Quizzes
# ============================================================================

param(
    [Parameter(Mandatory=$false)]
    [string]$SiteUrl = "https://mf7m.sharepoint.com/sites/PolicyManager"
)

$clientId = "d91b5b78-de72-424e-898b-8b5c9512ebd9"
$tenantId = "03bbbdee-d78b-4613-9b99-c468398246b7"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  JML Policy Management - Templates & Quizzes" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

Import-Module PnP.PowerShell -ErrorAction Stop
Connect-PnPOnline -Url $SiteUrl -DeviceLogin -ClientId $clientId -Tenant $tenantId
Write-Host "Connected!" -ForegroundColor Green

# ============================================================================
# POLICY TEMPLATES
# ============================================================================

# Simplified templates using only fields that exist in the list
$templates = @(
    @{
        Title = "HR Policy Template"
        TemplateDescription = "Standard template for HR policies including sections for purpose, scope, responsibilities, procedures, and compliance."
    },
    @{
        Title = "IT Security Policy Template"
        TemplateDescription = "Template for IT and security policies with sections for technical requirements, access controls, and incident response."
    },
    @{
        Title = "Compliance Policy Template"
        TemplateDescription = "Template for regulatory compliance policies with sections for legal requirements, controls, and audit provisions."
    },
    @{
        Title = "Health & Safety Policy Template"
        TemplateDescription = "Template for H&S policies covering hazards, risk assessments, and emergency procedures."
    }
)

Write-Host "`n[1/2] Creating policy templates..." -ForegroundColor Yellow

foreach ($template in $templates) {
    try {
        Add-PnPListItem -List "PM_PolicyTemplates" -Values $template | Out-Null
        Write-Host "  Created: $($template.TemplateName)" -ForegroundColor Green
    }
    catch {
        Write-Host "  Failed: $($template.TemplateName) - $_" -ForegroundColor Red
    }
}

# ============================================================================
# QUIZZES (linked to policies by ID - assuming policies are ID 1-22)
# ============================================================================

Write-Host "`n[2/2] Creating quizzes and questions..." -ForegroundColor Yellow

# Get policy IDs that require quizzes
$quizPolicies = @(
    @{ PolicyId = 1; Title = "Code of Conduct Quiz"; Passing = 80 },
    @{ PolicyId = 2; Title = "Anti-Harassment Quiz"; Passing = 85 },
    @{ PolicyId = 6; Title = "Information Security Quiz"; Passing = 90 },
    @{ PolicyId = 7; Title = "Acceptable Use Quiz"; Passing = 75 },
    @{ PolicyId = 8; Title = "Password Security Quiz"; Passing = 80 },
    @{ PolicyId = 11; Title = "Workplace Safety Quiz"; Passing = 80 },
    @{ PolicyId = 12; Title = "Emergency Evacuation Quiz"; Passing = 100 },
    @{ PolicyId = 14; Title = "Anti-Bribery Quiz"; Passing = 85 },
    @{ PolicyId = 16; Title = "Data Protection Quiz"; Passing = 80 }
)

# Quiz questions stored as array with PolicyId reference
$quizQuestionsList = @(
    # Policy 1 - Code of Conduct
    @{ PolicyId = 1; Q = "What should you do if you witness a colleague violating the code of conduct?"; Type = "MultipleChoice"; Options = '["Ignore it","Report to your manager or HR","Confront them publicly","Post about it on social media"]'; Answer = "Report to your manager or HR"; Points = 10 },
    @{ PolicyId = 1; Q = "Company resources should only be used for business purposes."; Type = "TrueFalse"; Options = '["True","False"]'; Answer = "True"; Points = 10 },
    @{ PolicyId = 1; Q = "Which of the following constitutes a conflict of interest?"; Type = "MultipleChoice"; Options = '["Working overtime","Having lunch with a competitor","Hiring a family member without disclosure","Attending industry conferences"]'; Answer = "Hiring a family member without disclosure"; Points = 10 },
    @{ PolicyId = 1; Q = "You can accept gifts from vendors without any restrictions."; Type = "TrueFalse"; Options = '["True","False"]'; Answer = "False"; Points = 10 },
    @{ PolicyId = 1; Q = "The code of conduct applies to:"; Type = "MultipleChoice"; Options = '["Only senior management","Only customer-facing staff","All employees","Only new hires"]'; Answer = "All employees"; Points = 10 },
    # Policy 2 - Anti-Harassment
    @{ PolicyId = 2; Q = "Harassment can only occur between a manager and their direct report."; Type = "TrueFalse"; Options = '["True","False"]'; Answer = "False"; Points = 10 },
    @{ PolicyId = 2; Q = "What is the first step if you experience harassment?"; Type = "MultipleChoice"; Options = '["Retaliate","Document the incident and report it","Quit your job","Ignore it"]'; Answer = "Document the incident and report it"; Points = 10 },
    @{ PolicyId = 2; Q = "Retaliation against someone who reports harassment is prohibited."; Type = "TrueFalse"; Options = '["True","False"]'; Answer = "True"; Points = 10 },
    # Policy 6 - Info Security
    @{ PolicyId = 6; Q = "What is the minimum classification for customer personal data?"; Type = "MultipleChoice"; Options = '["Public","Internal","Confidential","Restricted"]'; Answer = "Confidential"; Points = 10 },
    @{ PolicyId = 6; Q = "You can share your password with IT support to help troubleshoot."; Type = "TrueFalse"; Options = '["True","False"]'; Answer = "False"; Points = 10 },
    @{ PolicyId = 6; Q = "What should you do if you suspect a security breach?"; Type = "MultipleChoice"; Options = '["Wait and see if it resolves","Report immediately to IT Security","Try to fix it yourself","Tell a colleague"]'; Answer = "Report immediately to IT Security"; Points = 10 },
    @{ PolicyId = 6; Q = "Encryption is required for all data at rest and in transit."; Type = "TrueFalse"; Options = '["True","False"]'; Answer = "True"; Points = 10 },
    # Policy 8 - Password
    @{ PolicyId = 8; Q = "What is the minimum password length required?"; Type = "MultipleChoice"; Options = '["6 characters","8 characters","12 characters","16 characters"]'; Answer = "12 characters"; Points = 10 },
    @{ PolicyId = 8; Q = "Multi-factor authentication (MFA) is optional for all systems."; Type = "TrueFalse"; Options = '["True","False"]'; Answer = "False"; Points = 10 },
    @{ PolicyId = 8; Q = "How often should you change your password?"; Type = "MultipleChoice"; Options = '["Never","Every 30 days","Every 90 days or when compromised","Every year"]'; Answer = "Every 90 days or when compromised"; Points = 10 },
    # Policy 11 - H&S
    @{ PolicyId = 11; Q = "Who is responsible for workplace safety?"; Type = "MultipleChoice"; Options = '["Only the safety officer","Only management","Everyone","Only HR"]'; Answer = "Everyone"; Points = 10 },
    @{ PolicyId = 11; Q = "You should report near-misses as well as actual incidents."; Type = "TrueFalse"; Options = '["True","False"]'; Answer = "True"; Points = 10 },
    @{ PolicyId = 11; Q = "What is the correct action if you identify a hazard?"; Type = "MultipleChoice"; Options = '["Ignore if minor","Report immediately","Wait for the next safety meeting","Fix it yourself"]'; Answer = "Report immediately"; Points = 10 },
    # Policy 12 - Emergency
    @{ PolicyId = 12; Q = "What should you do when you hear the fire alarm?"; Type = "MultipleChoice"; Options = '["Finish your current task","Evacuate immediately via nearest exit","Wait for instructions","Collect personal belongings"]'; Answer = "Evacuate immediately via nearest exit"; Points = 20 },
    @{ PolicyId = 12; Q = "You should use the lift during a fire evacuation."; Type = "TrueFalse"; Options = '["True","False"]'; Answer = "False"; Points = 20 },
    @{ PolicyId = 12; Q = "Fire wardens wear which colour vest?"; Type = "MultipleChoice"; Options = '["Red","Yellow","Green","Orange"]'; Answer = "Orange"; Points = 20 },
    # Policy 14 - Anti-Bribery
    @{ PolicyId = 14; Q = "Facilitation payments are acceptable if they are small amounts."; Type = "TrueFalse"; Options = '["True","False"]'; Answer = "False"; Points = 10 },
    @{ PolicyId = 14; Q = "What is the maximum value of gifts you can accept without approval?"; Type = "MultipleChoice"; Options = '["No limit","50 GBP","100 GBP","250 GBP"]'; Answer = "50 GBP"; Points = 10 },
    @{ PolicyId = 14; Q = "Third parties acting on our behalf must also comply with anti-bribery laws."; Type = "TrueFalse"; Options = '["True","False"]'; Answer = "True"; Points = 10 },
    # Policy 16 - Data Protection
    @{ PolicyId = 16; Q = "Personal data can be processed without consent if there is a legitimate business interest."; Type = "TrueFalse"; Options = '["True","False"]'; Answer = "True"; Points = 10 },
    @{ PolicyId = 16; Q = "How quickly must data breaches be reported to the ICO?"; Type = "MultipleChoice"; Options = '["24 hours","72 hours","7 days","30 days"]'; Answer = "72 hours"; Points = 10 },
    @{ PolicyId = 16; Q = "Privacy by design means considering data protection from the start of any project."; Type = "TrueFalse"; Options = '["True","False"]'; Answer = "True"; Points = 10 }
)

foreach ($quiz in $quizPolicies) {
    try {
        # Create quiz
        $quizItem = Add-PnPListItem -List "PM_PolicyQuizzes" -Values @{
            Title = $quiz.Title
            PolicyId = $quiz.PolicyId
            QuizTitle = $quiz.Title
            QuizDescription = "Assessment to verify understanding of the policy requirements."
            PassingScore = $quiz.Passing
            AllowRetake = $true
            MaxAttempts = 3
            TimeLimit = 15
            RandomizeQuestions = $true
            IsActive = $true
        }
        Write-Host "  Created quiz: $($quiz.Title)" -ForegroundColor Green

        # Add questions for this policy
        $policyQuestions = $quizQuestionsList | Where-Object { $_.PolicyId -eq $quiz.PolicyId }
        $order = 1
        foreach ($q in $policyQuestions) {
            Add-PnPListItem -List "PM_PolicyQuizQuestions" -Values @{
                Title = "Q$order - $($quiz.Title)"
                QuizId = $quizItem.Id
                QuestionText = $q.Q
                QuestionType = $q.Type
                Options = $q.Options
                CorrectAnswer = $q.Answer
                Points = $q.Points
                OrderIndex = $order
            } | Out-Null
            $order++
        }
        if ($policyQuestions.Count -gt 0) {
            Write-Host "    Added $($policyQuestions.Count) questions" -ForegroundColor Gray
        }
    }
    catch {
        Write-Host "  Failed: $($quiz.Title) - $_" -ForegroundColor Red
    }
}

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  Templates and Quizzes created!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan

Disconnect-PnPOnline
