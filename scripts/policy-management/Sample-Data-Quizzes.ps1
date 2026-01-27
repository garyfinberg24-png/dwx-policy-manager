# ============================================================================
# JML Policy Management - Sample Data: Policy Quizzes
# Creates realistic quiz questions for policies requiring knowledge verification
# ============================================================================

param(
    [Parameter(Mandatory=$false)]
    [string]$SiteUrl = "https://mf7m.sharepoint.com/sites/JML"
)

$clientId = "d91b5b78-de72-424e-898b-8b5c9512ebd9"
$tenantId = "03bbbdee-d78b-4613-9b99-c468398246b7"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  JML Policy Management - Quiz Sample Data Loader" -ForegroundColor Cyan
Write-Host "  Target: $SiteUrl" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

Import-Module PnP.PowerShell -ErrorAction Stop

# Connect
Write-Host "`nConnecting to SharePoint..." -ForegroundColor Yellow
Connect-PnPOnline -Url $SiteUrl -DeviceLogin -ClientId $clientId -Tenant $tenantId
Write-Host "Connected!" -ForegroundColor Green

# ============================================================================
# HELPER FUNCTION: Convert options array to JSON
# ============================================================================
function ConvertTo-OptionsJson {
    param([string[]]$Options)
    return ($Options | ConvertTo-Json -Compress)
}

# ============================================================================
# STEP 1: Get Policy IDs from JML_Policies list
# ============================================================================
Write-Host "`nFetching policy IDs from JML_Policies..." -ForegroundColor Yellow

$policyItems = Get-PnPListItem -List "JML_Policies" -Fields "ID","PolicyNumber","PolicyName","RequiresQuiz" |
    Where-Object { $_["RequiresQuiz"] -eq $true }

$policyIdMap = @{}
foreach ($item in $policyItems) {
    $policyIdMap[$item["PolicyNumber"]] = $item.Id
    Write-Host "  Found quiz-enabled policy: $($item["PolicyNumber"]) (ID: $($item.Id))" -ForegroundColor Gray
}

if ($policyIdMap.Count -eq 0) {
    Write-Host "  No quiz-enabled policies found. Creating quizzes with placeholder IDs..." -ForegroundColor Yellow
    # Use placeholder IDs - these will need to be updated after policies are created
    $policyIdMap = @{
        "POL-HR-001" = 1
        "POL-HR-002" = 2
        "POL-IT-001" = 6
        "POL-IT-002" = 7
        "POL-HS-001" = 11
        "POL-CO-001" = 14
        "POL-DP-001" = 16
    }
}

# ============================================================================
# SAMPLE QUIZZES DATA
# ============================================================================

$quizzes = @(
    # -------------------------------------------------------------------------
    # QUIZ 1: Code of Conduct (POL-HR-001)
    # -------------------------------------------------------------------------
    @{
        PolicyNumber = "POL-HR-001"
        Title = "Employee Code of Conduct Knowledge Check"
        QuizTitle = "Code of Conduct Quiz"
        QuizDescription = "This quiz assesses your understanding of the Employee Code of Conduct, including ethical behavior, workplace standards, and professional responsibilities. You must score at least 80% to pass."
        PassingScore = 80
        AllowRetake = $true
        MaxAttempts = 3
        TimeLimit = 15
        RandomizeQuestions = $true
        ShowCorrectAnswers = $true
        IsActive = $true
        Questions = @(
            @{
                QuestionText = "According to the Code of Conduct, which of the following best describes a conflict of interest?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "When personal interests could influence your professional judgment",
                    "When two colleagues disagree about a project approach",
                    "When the company competes with other businesses",
                    "When you work overtime without approval"
                )
                CorrectAnswer = "0"
                Points = 10
                Explanation = "A conflict of interest occurs when your personal interests could improperly influence your professional decisions or loyalty to the company."
                OrderIndex = 1
                IsMandatory = $true
            },
            @{
                QuestionText = "You discover that a colleague is using company resources for personal business. According to the Code of Conduct, what should you do?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Ignore it since it's not your concern",
                    "Report it through the appropriate channels (manager or ethics hotline)",
                    "Confront the colleague publicly in a team meeting",
                    "Start using company resources yourself since others are doing it"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "The Code of Conduct requires employees to report violations through proper channels, such as your manager or the ethics hotline."
                OrderIndex = 2
                IsMandatory = $true
            },
            @{
                QuestionText = "Which of the following is an acceptable use of company email?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Sending personal chain letters to colleagues",
                    "Business-related communications with clients and colleagues",
                    "Running a side business using your company email address",
                    "Forwarding confidential company information to personal email"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Company email should primarily be used for business purposes. Personal use should be minimal and never involve confidential information."
                OrderIndex = 3
                IsMandatory = $true
            },
            @{
                QuestionText = "A vendor offers you expensive gifts and an invitation to an exclusive event. What should you do?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Accept everything as a gesture of good business relations",
                    "Decline and report the offer to your manager",
                    "Accept only if no one else knows about it",
                    "Accept the gifts but decline the event invitation"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Accepting lavish gifts from vendors could create real or perceived conflicts of interest. Always decline and report such offers."
                OrderIndex = 4
                IsMandatory = $true
            },
            @{
                QuestionText = "True or False: It is acceptable to share your company login credentials with a trusted colleague to help them complete urgent work."
                QuestionType = "TrueFalse"
                Options = @("True", "False")
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Sharing login credentials is never acceptable, regardless of circumstances. It violates security policies and can lead to audit issues."
                OrderIndex = 5
                IsMandatory = $true
            },
            @{
                QuestionText = "What is the primary purpose of the Employee Code of Conduct?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "To punish employees who make mistakes",
                    "To establish standards for ethical behavior and professional conduct",
                    "To increase bureaucratic procedures",
                    "To limit employee creativity and autonomy"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "The Code of Conduct establishes clear standards for professional and ethical behavior, creating a positive work environment."
                OrderIndex = 6
                IsMandatory = $true
            },
            @{
                QuestionText = "If you witness workplace bullying, which action is NOT appropriate?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Report it to HR or your manager",
                    "Document what you witnessed",
                    "Offer support to the person being bullied",
                    "Join in because it seems harmless"
                )
                CorrectAnswer = "3"
                Points = 10
                Explanation = "Workplace bullying is never acceptable. Witnesses should report incidents and support affected colleagues."
                OrderIndex = 7
                IsMandatory = $true
            },
            @{
                QuestionText = "Which statement about social media use is correct according to our policy?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "You may share confidential company information if it makes the company look good",
                    "Personal opinions about the company should never be posted",
                    "You must clearly identify personal views as your own, not the company's",
                    "Company social media accounts can be used for personal posts"
                )
                CorrectAnswer = "2"
                Points = 10
                Explanation = "When posting about work-related topics personally, you should make clear these are your own views, not official company positions."
                OrderIndex = 8
                IsMandatory = $true
            },
            @{
                QuestionText = "What should you do if you're unsure whether an action violates the Code of Conduct?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Proceed with the action if no one is watching",
                    "Ask your manager or HR for guidance before proceeding",
                    "Assume it's fine if the policy doesn't specifically prohibit it",
                    "Do nothing and hope the situation resolves itself"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "When in doubt, always seek guidance from your manager, HR, or the ethics hotline before taking action."
                OrderIndex = 9
                IsMandatory = $true
            },
            @{
                QuestionText = "Retaliation against someone who reports a Code of Conduct violation is:"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Acceptable if the report was inaccurate",
                    "Prohibited and is itself a violation of the Code",
                    "Allowed if done subtly",
                    "Only prohibited for managers"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Retaliation is strictly prohibited. Anyone who retaliates against a reporter will face disciplinary action."
                OrderIndex = 10
                IsMandatory = $true
            }
        )
    },

    # -------------------------------------------------------------------------
    # QUIZ 2: Anti-Harassment Policy (POL-HR-002)
    # -------------------------------------------------------------------------
    @{
        PolicyNumber = "POL-HR-002"
        Title = "Anti-Harassment and Discrimination Assessment"
        QuizTitle = "Workplace Harassment Prevention Quiz"
        QuizDescription = "This assessment tests your understanding of harassment and discrimination prevention in the workplace. A passing score of 85% is required to demonstrate competency."
        PassingScore = 85
        AllowRetake = $true
        MaxAttempts = 3
        TimeLimit = 20
        RandomizeQuestions = $true
        ShowCorrectAnswers = $true
        IsActive = $true
        Questions = @(
            @{
                QuestionText = "Which of the following is considered sexual harassment?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "A single genuine compliment about a colleague's presentation skills",
                    "Repeated unwelcome comments about someone's physical appearance",
                    "A respectful disagreement about work approaches",
                    "Declining a colleague's invitation to lunch"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Repeated unwelcome comments about physical appearance constitute sexual harassment, regardless of the intent behind them."
                OrderIndex = 1
                IsMandatory = $true
            },
            @{
                QuestionText = "What does 'hostile work environment' harassment mean?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Any workplace where employees sometimes disagree",
                    "Severe or pervasive conduct that interferes with work performance",
                    "A workplace with strict deadlines and high expectations",
                    "Any workplace without air conditioning"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "A hostile work environment exists when harassment is severe or pervasive enough to interfere with an employee's ability to perform their job."
                OrderIndex = 2
                IsMandatory = $true
            },
            @{
                QuestionText = "If you witness harassment, you should:"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Wait to see if it happens again before reporting",
                    "Only act if the victim asks for help",
                    "Report it even if you're not the direct target",
                    "Handle it yourself by confronting the harasser privately"
                )
                CorrectAnswer = "2"
                Points = 10
                Explanation = "All employees have a responsibility to report harassment they witness, regardless of whether they are directly affected."
                OrderIndex = 3
                IsMandatory = $true
            },
            @{
                QuestionText = "Which of the following could constitute discrimination?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Promoting the most qualified candidate regardless of background",
                    "Excluding someone from meetings based on their age",
                    "Providing constructive feedback on work performance",
                    "Assigning tasks based on skill sets and experience"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Excluding someone from work activities based on age, gender, race, or other protected characteristics is discrimination."
                OrderIndex = 4
                IsMandatory = $true
            },
            @{
                QuestionText = "True or False: Harassment can only occur between a supervisor and subordinate."
                QuestionType = "TrueFalse"
                Options = @("True", "False")
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Harassment can occur between any individuals in the workplace, including peer-to-peer and even subordinate-to-supervisor."
                OrderIndex = 5
                IsMandatory = $true
            },
            @{
                QuestionText = "Which action is protected under our anti-retaliation policy?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Filing a false report to get a colleague in trouble",
                    "Reporting suspected harassment in good faith",
                    "Refusing to cooperate with an investigation",
                    "Discussing confidential investigation details with coworkers"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Good-faith reporting of suspected harassment is protected. False reports or interference with investigations is not protected."
                OrderIndex = 6
                IsMandatory = $true
            },
            @{
                QuestionText = "What constitutes 'quid pro quo' harassment?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Trading office supplies with colleagues",
                    "Conditioning job benefits on sexual favors",
                    "Negotiating salary during hiring",
                    "Asking for a favor from a coworker"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Quid pro quo harassment occurs when job benefits (promotion, raise, continued employment) are conditioned on accepting sexual advances."
                OrderIndex = 7
                IsMandatory = $true
            },
            @{
                QuestionText = "Which statement about the investigation process is true?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Investigations are optional if the accused denies the allegation",
                    "All reports will be thoroughly investigated by trained personnel",
                    "The accuser must have witnesses for an investigation to proceed",
                    "Only HR can conduct investigations"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "All harassment reports are investigated thoroughly, regardless of whether there are witnesses or the accused person's response."
                OrderIndex = 8
                IsMandatory = $true
            },
            @{
                QuestionText = "Microaggressions are:"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Always intentional and malicious acts",
                    "Subtle, often unintentional comments or behaviors that can cause harm",
                    "Not covered by the harassment policy",
                    "Only harmful if repeated more than five times"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Microaggressions are often unintentional but can still cause significant harm and contribute to a hostile environment."
                OrderIndex = 9
                IsMandatory = $true
            },
            @{
                QuestionText = "What is the best first step when you feel you're being harassed?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Immediately file a lawsuit",
                    "Post about it on social media",
                    "Document the incidents and report to HR or your manager",
                    "Confront the harasser aggressively in public"
                )
                CorrectAnswer = "2"
                Points = 10
                Explanation = "Documenting incidents and reporting through proper channels allows the company to address the situation appropriately."
                OrderIndex = 10
                IsMandatory = $true
            }
        )
    },

    # -------------------------------------------------------------------------
    # QUIZ 3: Information Security Policy (POL-IT-001)
    # -------------------------------------------------------------------------
    @{
        PolicyNumber = "POL-IT-001"
        Title = "Information Security Fundamentals Assessment"
        QuizTitle = "Cybersecurity Awareness Quiz"
        QuizDescription = "This quiz tests your knowledge of information security best practices, including data protection, password security, and threat recognition. Passing score is 80%."
        PassingScore = 80
        AllowRetake = $true
        MaxAttempts = 3
        TimeLimit = 15
        RandomizeQuestions = $true
        ShowCorrectAnswers = $true
        IsActive = $true
        Questions = @(
            @{
                QuestionText = "What is phishing?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "A type of fishing technique used in oceans",
                    "Fraudulent attempts to obtain sensitive information by disguising as trustworthy entities",
                    "A legitimate IT security testing method",
                    "A way to recover forgotten passwords"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Phishing is a cybercrime where attackers impersonate legitimate organizations to trick victims into revealing sensitive information."
                OrderIndex = 1
                IsMandatory = $true
            },
            @{
                QuestionText = "Which password is the most secure?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "password123",
                    "MyDogSpot2023",
                    "P@ss$w0rd!X7#mK9vL2",
                    "12345678"
                )
                CorrectAnswer = "2"
                Points = 10
                Explanation = "Strong passwords are long (12+ characters), include mixed case, numbers, and special characters, and avoid common words or patterns."
                OrderIndex = 2
                IsMandatory = $true
            },
            @{
                QuestionText = "You receive an urgent email from 'IT Support' asking you to click a link and verify your password. What should you do?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Click the link immediately since IT Support is trustworthy",
                    "Forward it to all colleagues so they can verify too",
                    "Report it as a potential phishing attempt without clicking any links",
                    "Reply with your current password to verify your account"
                )
                CorrectAnswer = "2"
                Points = 10
                Explanation = "Legitimate IT departments never ask for passwords via email. Report suspicious emails to your security team immediately."
                OrderIndex = 3
                IsMandatory = $true
            },
            @{
                QuestionText = "Which of the following is a secure way to share confidential documents?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Email them to personal Gmail accounts for convenience",
                    "Use approved company file sharing platforms with proper permissions",
                    "Post them on public cloud storage for easy access",
                    "Print and leave copies in common areas"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Confidential documents should only be shared through approved company platforms with appropriate access controls."
                OrderIndex = 4
                IsMandatory = $true
            },
            @{
                QuestionText = "True or False: It's safe to use the same password for your work accounts and personal accounts."
                QuestionType = "TrueFalse"
                Options = @("True", "False")
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Using the same password across accounts is dangerous. If one account is compromised, all accounts with that password are at risk."
                OrderIndex = 5
                IsMandatory = $true
            },
            @{
                QuestionText = "What should you do if you find a USB drive in the parking lot?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Plug it into your computer to see whose it is",
                    "Give it to a colleague to check",
                    "Turn it in to IT security without plugging it in",
                    "Keep it and format it for personal use"
                )
                CorrectAnswer = "2"
                Points = 10
                Explanation = "Unknown USB drives can contain malware. Never plug in unidentified devices - turn them in to IT security."
                OrderIndex = 6
                IsMandatory = $true
            },
            @{
                QuestionText = "Multi-factor authentication (MFA) provides security by:"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Making passwords easier to remember",
                    "Requiring multiple forms of verification beyond just a password",
                    "Automatically changing your password every day",
                    "Encrypting all your emails"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "MFA adds security layers by requiring additional verification (like a phone code) beyond just your password."
                OrderIndex = 7
                IsMandatory = $true
            },
            @{
                QuestionText = "When working remotely on public WiFi, you should:"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Connect freely since most public WiFi is secure",
                    "Use the company VPN before accessing any work resources",
                    "Only check personal email, not work email",
                    "Disable your firewall for better connection speed"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Always use VPN when accessing work resources on public networks to encrypt your connection and protect data."
                OrderIndex = 8
                IsMandatory = $true
            },
            @{
                QuestionText = "What is ransomware?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Software that speeds up your computer",
                    "Malware that encrypts files and demands payment for decryption",
                    "A type of firewall protection",
                    "An antivirus program"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Ransomware is malicious software that encrypts your files and demands payment (usually cryptocurrency) to restore access."
                OrderIndex = 9
                IsMandatory = $true
            },
            @{
                QuestionText = "If you suspect your computer has been compromised, you should:"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Try to fix it yourself to avoid bothering IT",
                    "Disconnect from the network and contact IT security immediately",
                    "Continue working and monitor for unusual behavior",
                    "Restart the computer and hope the problem goes away"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Immediately disconnect from the network to prevent spread and contact IT security for proper incident response."
                OrderIndex = 10
                IsMandatory = $true
            }
        )
    },

    # -------------------------------------------------------------------------
    # QUIZ 4: Data Protection Policy (POL-IT-002)
    # -------------------------------------------------------------------------
    @{
        PolicyNumber = "POL-IT-002"
        Title = "Data Protection and Privacy Assessment"
        QuizTitle = "Data Protection Quiz"
        QuizDescription = "Test your understanding of data protection requirements, including handling personal data, GDPR principles, and data subject rights. Minimum passing score: 80%."
        PassingScore = 80
        AllowRetake = $true
        MaxAttempts = 3
        TimeLimit = 20
        RandomizeQuestions = $true
        ShowCorrectAnswers = $true
        IsActive = $true
        Questions = @(
            @{
                QuestionText = "What does GDPR stand for?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "General Data Protection Regulation",
                    "Global Data Privacy Rules",
                    "Government Data Processing Requirements",
                    "Generic Data Protection Registry"
                )
                CorrectAnswer = "0"
                Points = 10
                Explanation = "GDPR stands for General Data Protection Regulation, the EU's comprehensive data protection law."
                OrderIndex = 1
                IsMandatory = $true
            },
            @{
                QuestionText = "Which of the following is considered personal data under GDPR?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Only names and addresses",
                    "Any information that can identify a person directly or indirectly",
                    "Only financial information",
                    "Only medical records"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Personal data includes any information that can identify a person, including names, IDs, location data, IP addresses, and more."
                OrderIndex = 2
                IsMandatory = $true
            },
            @{
                QuestionText = "What is a data subject access request (DSAR)?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "A request to delete all company data",
                    "An individual's right to obtain their personal data held by an organization",
                    "A request for IT support access",
                    "A requirement to share data with regulators"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "A DSAR is an individual's right to request access to the personal data an organization holds about them."
                OrderIndex = 3
                IsMandatory = $true
            },
            @{
                QuestionText = "How quickly must a data breach be reported to the relevant authority under GDPR?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Within 24 hours",
                    "Within 72 hours",
                    "Within 7 days",
                    "Within 30 days"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Under GDPR, data breaches that pose risks to individuals must be reported to the supervisory authority within 72 hours."
                OrderIndex = 4
                IsMandatory = $true
            },
            @{
                QuestionText = "True or False: Personal data can be processed without consent if there is a legitimate business interest."
                QuestionType = "TrueFalse"
                Options = @("True", "False")
                CorrectAnswer = "0"
                Points = 10
                Explanation = "Legitimate interest is one of the lawful bases for processing, but it must be balanced against the individual's rights and interests."
                OrderIndex = 5
                IsMandatory = $true
            },
            @{
                QuestionText = "Which principle requires that personal data be accurate and kept up to date?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Data minimization",
                    "Purpose limitation",
                    "Accuracy",
                    "Storage limitation"
                )
                CorrectAnswer = "2"
                Points = 10
                Explanation = "The accuracy principle requires organizations to ensure personal data is accurate and, where necessary, kept up to date."
                OrderIndex = 6
                IsMandatory = $true
            },
            @{
                QuestionText = "What is 'data minimization'?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Compressing data to save storage space",
                    "Collecting only the data that is necessary for the specified purpose",
                    "Deleting data automatically after 30 days",
                    "Using smaller fonts in documents"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Data minimization means only collecting personal data that is adequate, relevant, and necessary for the intended purpose."
                OrderIndex = 7
                IsMandatory = $true
            },
            @{
                QuestionText = "The 'right to be forgotten' allows individuals to:"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Delete their own work from company systems",
                    "Request erasure of their personal data under certain conditions",
                    "Forget their passwords without consequences",
                    "Remove negative performance reviews"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "The right to erasure (right to be forgotten) allows individuals to request deletion of their personal data in certain circumstances."
                OrderIndex = 8
                IsMandatory = $true
            },
            @{
                QuestionText = "Special category data includes:"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Business financial records",
                    "Health data, biometric data, and information about race or religion",
                    "Customer order histories",
                    "Employee work schedules"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Special category data includes sensitive information like health, biometric, genetic, race/ethnicity, political opinions, and religious beliefs."
                OrderIndex = 9
                IsMandatory = $true
            },
            @{
                QuestionText = "If you accidentally send personal data to the wrong recipient, you should:"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Hope they don't notice and delete the email",
                    "Report it immediately to your Data Protection Officer",
                    "Ask the recipient to delete it and forget about it",
                    "Send a follow-up email marking it as confidential"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Accidental disclosure of personal data is a potential breach and must be reported to your DPO immediately for assessment."
                OrderIndex = 10
                IsMandatory = $true
            }
        )
    },

    # -------------------------------------------------------------------------
    # QUIZ 5: Health and Safety Policy (POL-HS-001)
    # -------------------------------------------------------------------------
    @{
        PolicyNumber = "POL-HS-001"
        Title = "Workplace Health and Safety Assessment"
        QuizTitle = "Health & Safety Awareness Quiz"
        QuizDescription = "This quiz covers essential workplace health and safety knowledge, including emergency procedures, hazard identification, and injury prevention. Pass mark: 80%."
        PassingScore = 80
        AllowRetake = $true
        MaxAttempts = 3
        TimeLimit = 15
        RandomizeQuestions = $true
        ShowCorrectAnswers = $true
        IsActive = $true
        Questions = @(
            @{
                QuestionText = "In case of a fire alarm, you should:"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Finish your current task then leave",
                    "Leave immediately via the nearest safe exit",
                    "Use the elevator to exit quickly",
                    "Gather your personal belongings before leaving"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "When a fire alarm sounds, leave immediately via the nearest safe exit. Never use elevators during a fire."
                OrderIndex = 1
                IsMandatory = $true
            },
            @{
                QuestionText = "What should you do if you discover a workplace hazard?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Ignore it if it doesn't affect your work area",
                    "Fix it yourself regardless of your training",
                    "Report it immediately to your supervisor or safety officer",
                    "Wait until the next safety meeting to mention it"
                )
                CorrectAnswer = "2"
                Points = 10
                Explanation = "All hazards should be reported immediately so they can be properly assessed and addressed."
                OrderIndex = 2
                IsMandatory = $true
            },
            @{
                QuestionText = "Proper ergonomic desk setup includes:"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Screen at arm's length, top of screen at or slightly below eye level",
                    "Screen as close as possible to reduce eye strain",
                    "Chair at its lowest setting for stability",
                    "Keyboard on your lap for comfortable typing"
                )
                CorrectAnswer = "0"
                Points = 10
                Explanation = "Proper ergonomics: monitor at arm's length, top at eye level, feet flat, arms at 90 degrees."
                OrderIndex = 3
                IsMandatory = $true
            },
            @{
                QuestionText = "True or False: Only employees trained in first aid should attempt to help an injured colleague."
                QuestionType = "TrueFalse"
                Options = @("True", "False")
                CorrectAnswer = "1"
                Points = 10
                Explanation = "While first aid training is valuable, anyone can call for help, comfort the injured person, and perform basic assistance safely."
                OrderIndex = 4
                IsMandatory = $true
            },
            @{
                QuestionText = "What is the purpose of a risk assessment?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "To create paperwork for regulatory compliance only",
                    "To identify hazards and determine appropriate control measures",
                    "To blame employees when accidents occur",
                    "To limit the activities employees can perform"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Risk assessments identify hazards, evaluate risks, and determine control measures to prevent harm."
                OrderIndex = 5
                IsMandatory = $true
            },
            @{
                QuestionText = "Personal Protective Equipment (PPE) should be:"
                QuestionType = "MultipleChoice"
                Options = @(
                    "The first line of defense against hazards",
                    "The last line of defense after other controls fail",
                    "Only worn when supervisors are watching",
                    "Shared among colleagues to save costs"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "PPE is the last resort - other controls (elimination, substitution, engineering controls) should be prioritized."
                OrderIndex = 6
                IsMandatory = $true
            },
            @{
                QuestionText = "How often should you take breaks when working at a computer?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Only when you feel tired",
                    "Every 5-10 minutes",
                    "Every 50-60 minutes, take a 5-10 minute break",
                    "Breaks are not necessary for office work"
                )
                CorrectAnswer = "2"
                Points = 10
                Explanation = "Regular breaks (every 50-60 minutes) reduce eye strain, prevent repetitive strain injuries, and improve productivity."
                OrderIndex = 7
                IsMandatory = $true
            },
            @{
                QuestionText = "Who is responsible for workplace health and safety?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Only the health and safety department",
                    "Only managers and supervisors",
                    "Everyone has a shared responsibility",
                    "Only employees who work in hazardous areas"
                )
                CorrectAnswer = "2"
                Points = 10
                Explanation = "Workplace safety is a shared responsibility - employers must provide a safe workplace, and employees must follow safety procedures."
                OrderIndex = 8
                IsMandatory = $true
            },
            @{
                QuestionText = "If you see a blocked fire exit, you should:"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Walk around the obstruction if you can fit through",
                    "Report it immediately so it can be cleared",
                    "Only report it if there's an actual fire",
                    "Move the obstruction yourself"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Blocked fire exits are serious safety hazards and must be reported immediately for clearance."
                OrderIndex = 9
                IsMandatory = $true
            },
            @{
                QuestionText = "Where are first aid kits typically located in the workplace?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "In locked cabinets only managers can access",
                    "In clearly marked, accessible locations throughout the building",
                    "Only in the medical room",
                    "Employees must bring their own"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "First aid kits should be clearly marked and easily accessible to all employees throughout the workplace."
                OrderIndex = 10
                IsMandatory = $true
            }
        )
    },

    # -------------------------------------------------------------------------
    # QUIZ 6: Corporate Governance Policy (POL-CO-001)
    # -------------------------------------------------------------------------
    @{
        PolicyNumber = "POL-CO-001"
        Title = "Corporate Governance Fundamentals"
        QuizTitle = "Corporate Governance Quiz"
        QuizDescription = "Test your knowledge of corporate governance principles, including ethical business practices, compliance requirements, and stakeholder responsibilities."
        PassingScore = 75
        AllowRetake = $true
        MaxAttempts = 3
        TimeLimit = 15
        RandomizeQuestions = $true
        ShowCorrectAnswers = $true
        IsActive = $true
        Questions = @(
            @{
                QuestionText = "What is corporate governance primarily concerned with?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Maximizing short-term profits at any cost",
                    "The system of rules, practices, and processes by which a company is directed and controlled",
                    "Government regulation of all business activities",
                    "Managing employee vacation schedules"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Corporate governance involves the structures and processes for decision-making, accountability, and control in an organization."
                OrderIndex = 1
                IsMandatory = $true
            },
            @{
                QuestionText = "Which of the following is a key stakeholder in corporate governance?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Only shareholders",
                    "Shareholders, employees, customers, suppliers, and the community",
                    "Only the board of directors",
                    "Only executive management"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Good governance considers all stakeholders who are affected by or can affect the organization's decisions."
                OrderIndex = 2
                IsMandatory = $true
            },
            @{
                QuestionText = "Transparency in corporate governance means:"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Sharing all confidential information publicly",
                    "Clear and honest communication of material information to stakeholders",
                    "Having glass walls in office buildings",
                    "Publishing all employee salaries"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Transparency involves open and honest communication about material matters that affect stakeholder decisions."
                OrderIndex = 3
                IsMandatory = $true
            },
            @{
                QuestionText = "True or False: Compliance with laws and regulations is optional if it impacts profitability."
                QuestionType = "TrueFalse"
                Options = @("True", "False")
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Compliance with applicable laws and regulations is mandatory, regardless of any impact on profitability."
                OrderIndex = 4
                IsMandatory = $true
            },
            @{
                QuestionText = "What role does the board of directors play in governance?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Day-to-day operational management",
                    "Strategic oversight and ensuring management acts in shareholders' interests",
                    "Approving all employee expense reports",
                    "Writing company policies"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "The board provides strategic direction and oversight, ensuring management acts responsibly and in stakeholders' interests."
                OrderIndex = 5
                IsMandatory = $true
            },
            @{
                QuestionText = "What is a conflict of interest in corporate governance?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "When two departments compete for budget",
                    "When personal interests could improperly influence business decisions",
                    "When employees disagree about priorities",
                    "When competitors offer lower prices"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "A conflict of interest occurs when personal interests could compromise objectivity in business decisions."
                OrderIndex = 6
                IsMandatory = $true
            },
            @{
                QuestionText = "Accountability in governance refers to:"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Counting company assets annually",
                    "Being responsible and answerable for decisions and their outcomes",
                    "Keeping accurate financial records only",
                    "Having a customer service department"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Accountability means decision-makers are responsible for their actions and must answer for the outcomes."
                OrderIndex = 7
                IsMandatory = $true
            },
            @{
                QuestionText = "Internal controls are designed to:"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Control employee behavior strictly",
                    "Provide reasonable assurance regarding reliability, compliance, and operations",
                    "Prevent all possible losses",
                    "Replace external auditors"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Internal controls provide reasonable (not absolute) assurance about reliability, compliance, and operational effectiveness."
                OrderIndex = 8
                IsMandatory = $true
            }
        )
    },

    # -------------------------------------------------------------------------
    # QUIZ 7: Data Privacy Awareness (POL-DP-001)
    # -------------------------------------------------------------------------
    @{
        PolicyNumber = "POL-DP-001"
        Title = "Data Privacy Awareness Assessment"
        QuizTitle = "Privacy Awareness Quiz"
        QuizDescription = "Assess your understanding of data privacy principles, personal information handling, and privacy rights. Required passing score: 80%."
        PassingScore = 80
        AllowRetake = $true
        MaxAttempts = 3
        TimeLimit = 15
        RandomizeQuestions = $true
        ShowCorrectAnswers = $true
        IsActive = $true
        Questions = @(
            @{
                QuestionText = "What is personally identifiable information (PII)?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Only names and social security numbers",
                    "Any information that can be used to identify, contact, or locate an individual",
                    "Only data stored in databases",
                    "Public information available on the internet"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "PII includes any data that could potentially identify a specific individual, directly or indirectly."
                OrderIndex = 1
                IsMandatory = $true
            },
            @{
                QuestionText = "Why is data privacy important?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Only to comply with regulations",
                    "To protect individuals from identity theft, maintain trust, and comply with legal obligations",
                    "To make data harder to access",
                    "Privacy is only important for celebrities"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Privacy protects individuals' rights, prevents harm, maintains organizational trust, and ensures legal compliance."
                OrderIndex = 2
                IsMandatory = $true
            },
            @{
                QuestionText = "Before collecting personal data, you should:"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Collect as much as possible for future needs",
                    "Ensure you have a valid legal basis and clear purpose for collection",
                    "Get approval from IT only",
                    "Check if competitors collect similar data"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Data collection requires a lawful basis, a clear purpose, and collection of only necessary information."
                OrderIndex = 3
                IsMandatory = $true
            },
            @{
                QuestionText = "True or False: Once consent is given for data processing, it cannot be withdrawn."
                QuestionType = "TrueFalse"
                Options = @("True", "False")
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Individuals have the right to withdraw consent at any time, and organizations must make withdrawal as easy as giving consent."
                OrderIndex = 4
                IsMandatory = $true
            },
            @{
                QuestionText = "How should personal data be disposed of when no longer needed?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Throw paper documents in regular trash",
                    "Securely destroy or anonymize data to prevent reconstruction",
                    "Keep it indefinitely just in case",
                    "Give it to marketing for future campaigns"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Personal data must be securely destroyed when no longer needed, preventing any possibility of reconstruction or recovery."
                OrderIndex = 5
                IsMandatory = $true
            },
            @{
                QuestionText = "Which of the following requires explicit consent under most privacy laws?"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Processing public directory information",
                    "Processing sensitive personal data like health or biometric information",
                    "Using business contact details for work purposes",
                    "Storing data required for contract fulfillment"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Sensitive personal data (health, biometric, genetic, religious beliefs) typically requires explicit consent for processing."
                OrderIndex = 6
                IsMandatory = $true
            },
            @{
                QuestionText = "Privacy by design means:"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Adding privacy features after a system is built",
                    "Incorporating privacy protections into systems from the earliest design stage",
                    "Designing attractive privacy policies",
                    "Using privacy-themed graphics"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "Privacy by design means building privacy protections into systems, products, and processes from the very beginning."
                OrderIndex = 7
                IsMandatory = $true
            },
            @{
                QuestionText = "When transferring personal data internationally, you must:"
                QuestionType = "MultipleChoice"
                Options = @(
                    "Send it by encrypted email only",
                    "Ensure adequate protections exist in the receiving country or use approved mechanisms",
                    "Just notify the individual after transfer",
                    "International transfers don't require special consideration"
                )
                CorrectAnswer = "1"
                Points = 10
                Explanation = "International data transfers require ensuring adequate protection through approved mechanisms like adequacy decisions or standard clauses."
                OrderIndex = 8
                IsMandatory = $true
            }
        )
    }
)

# ============================================================================
# STEP 2: Create Quizzes and Questions
# ============================================================================

Write-Host "`nCreating quizzes and questions..." -ForegroundColor Yellow

$quizCounter = 0
$questionCounter = 0

foreach ($quiz in $quizzes) {
    $policyNumber = $quiz.PolicyNumber
    $policyId = $policyIdMap[$policyNumber]

    if ($null -eq $policyId) {
        Write-Host "  Skipping quiz for $policyNumber - Policy not found" -ForegroundColor Red
        continue
    }

    Write-Host "`n  Creating quiz for $policyNumber..." -ForegroundColor Cyan

    # Create the quiz
    $quizValues = @{
        "Title" = $quiz.Title
        "PolicyId" = $policyId
        "QuizTitle" = $quiz.QuizTitle
        "QuizDescription" = $quiz.QuizDescription
        "PassingScore" = $quiz.PassingScore
        "AllowRetake" = $quiz.AllowRetake
        "MaxAttempts" = $quiz.MaxAttempts
        "TimeLimit" = $quiz.TimeLimit
        "RandomizeQuestions" = $quiz.RandomizeQuestions
        "ShowCorrectAnswers" = $quiz.ShowCorrectAnswers
        "IsActive" = $quiz.IsActive
    }

    $newQuiz = Add-PnPListItem -List "JML_PolicyQuizzes" -Values $quizValues
    $quizId = $newQuiz.Id
    $quizCounter++

    Write-Host "    Quiz created: $($quiz.QuizTitle) (ID: $quizId)" -ForegroundColor Green

    # Create questions for this quiz
    foreach ($question in $quiz.Questions) {
        $optionsJson = $question.Options | ConvertTo-Json -Compress

        $questionValues = @{
            "Title" = "Q$($question.OrderIndex): $($question.QuestionText.Substring(0, [Math]::Min(50, $question.QuestionText.Length)))..."
            "QuizId" = $quizId
            "QuestionText" = $question.QuestionText
            "QuestionType" = $question.QuestionType
            "Options" = $optionsJson
            "CorrectAnswer" = $question.CorrectAnswer
            "Points" = $question.Points
            "Explanation" = $question.Explanation
            "OrderIndex" = $question.OrderIndex
            "IsMandatory" = $question.IsMandatory
        }

        $null = Add-PnPListItem -List "JML_PolicyQuizQuestions" -Values $questionValues
        $questionCounter++
    }

    Write-Host "    Created $($quiz.Questions.Count) questions" -ForegroundColor Gray
}

# ============================================================================
# SUMMARY
# ============================================================================

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  Quiz Sample Data Load Complete!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Quizzes created: $quizCounter" -ForegroundColor White
Write-Host "  Questions created: $questionCounter" -ForegroundColor White
Write-Host "`n  Quiz topics covered:" -ForegroundColor Yellow
Write-Host "    - Employee Code of Conduct (10 questions)" -ForegroundColor Gray
Write-Host "    - Anti-Harassment & Discrimination (10 questions)" -ForegroundColor Gray
Write-Host "    - Information Security (10 questions)" -ForegroundColor Gray
Write-Host "    - Data Protection & GDPR (10 questions)" -ForegroundColor Gray
Write-Host "    - Health & Safety (10 questions)" -ForegroundColor Gray
Write-Host "    - Corporate Governance (8 questions)" -ForegroundColor Gray
Write-Host "    - Data Privacy Awareness (8 questions)" -ForegroundColor Gray
Write-Host "`n" -ForegroundColor White
