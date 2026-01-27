# ============================================================================
# JML Policy Management - Sample Data: Core Policies
# Creates realistic enterprise policies across all categories
# ============================================================================

param(
    [Parameter(Mandatory=$false)]
    [string]$SiteUrl = "https://mf7m.sharepoint.com/sites/JML"
)

$clientId = "d91b5b78-de72-424e-898b-8b5c9512ebd9"
$tenantId = "03bbbdee-d78b-4613-9b99-c468398246b7"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  JML Policy Management - Sample Data Loader" -ForegroundColor Cyan
Write-Host "  Target: $SiteUrl" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

Import-Module PnP.PowerShell -ErrorAction Stop

# Connect
Write-Host "`nConnecting to SharePoint..." -ForegroundColor Yellow
Connect-PnPOnline -Url $SiteUrl -DeviceLogin -ClientId $clientId -Tenant $tenantId
Write-Host "Connected!" -ForegroundColor Green

# ============================================================================
# SAMPLE POLICIES DATA
# ============================================================================

$policies = @(
    # HR POLICIES
    @{
        Title = "POL-HR-001 Employee Code of Conduct"
        PolicyNumber = "POL-HR-001"
        PolicyName = "Employee Code of Conduct"
        PolicyCategory = "HR Policies"
        PolicyType = "Corporate"
        PolicyDescription = "This policy establishes the standards of professional conduct expected from all employees. It covers ethical behavior, workplace interactions, conflicts of interest, and the responsible use of company resources. All employees must acknowledge this policy within their first week of employment."
        VersionNumber = "3.0"
        VersionType = "Major"
        PolicyStatus = "Published"
        ComplianceRisk = "Critical"
        IsMandatory = $true
        IsActive = $true
        RequiresAcknowledgement = $true
        AcknowledgementType = "Periodic - Annual"
        AcknowledgementDeadlineDays = 14
        ReadTimeframe = "Week 1"
        RequiresQuiz = $true
        QuizPassingScore = 80
        DistributionScope = "All Employees"
        EstimatedReadTimeMinutes = 25
        ReviewCycleMonths = 12
    },
    @{
        Title = "POL-HR-002 Anti-Harassment and Discrimination"
        PolicyNumber = "POL-HR-002"
        PolicyName = "Anti-Harassment and Discrimination Policy"
        PolicyCategory = "HR Policies"
        PolicyType = "Corporate"
        PolicyDescription = "This policy prohibits all forms of harassment and discrimination in the workplace, including but not limited to discrimination based on race, gender, age, religion, disability, or sexual orientation. It outlines reporting procedures and the investigation process."
        VersionNumber = "2.1"
        VersionType = "Minor"
        PolicyStatus = "Published"
        ComplianceRisk = "Critical"
        IsMandatory = $true
        IsActive = $true
        RequiresAcknowledgement = $true
        AcknowledgementType = "Periodic - Annual"
        AcknowledgementDeadlineDays = 7
        ReadTimeframe = "Day 3"
        RequiresQuiz = $true
        QuizPassingScore = 85
        DistributionScope = "All Employees"
        EstimatedReadTimeMinutes = 20
        ReviewCycleMonths = 12
    },
    @{
        Title = "POL-HR-003 Remote Work Policy"
        PolicyNumber = "POL-HR-003"
        PolicyName = "Remote Work and Flexible Working Policy"
        PolicyCategory = "HR Policies"
        PolicyType = "Corporate"
        PolicyDescription = "This policy defines the guidelines for remote work arrangements, including eligibility criteria, equipment requirements, communication expectations, and performance monitoring. It covers both regular remote work and temporary arrangements."
        VersionNumber = "2.0"
        VersionType = "Major"
        PolicyStatus = "Published"
        ComplianceRisk = "Medium"
        IsMandatory = $false
        IsActive = $true
        RequiresAcknowledgement = $true
        AcknowledgementType = "One-Time"
        AcknowledgementDeadlineDays = 30
        ReadTimeframe = "Month 1"
        RequiresQuiz = $false
        DistributionScope = "All Employees"
        EstimatedReadTimeMinutes = 15
        ReviewCycleMonths = 24
    },
    @{
        Title = "POL-HR-004 Leave and Time Off"
        PolicyNumber = "POL-HR-004"
        PolicyName = "Leave and Time Off Policy"
        PolicyCategory = "HR Policies"
        PolicyType = "Corporate"
        PolicyDescription = "Comprehensive policy covering all types of leave including annual leave, sick leave, parental leave, bereavement leave, and special circumstances leave. Includes procedures for requesting and approving leave."
        VersionNumber = "4.2"
        VersionType = "Minor"
        PolicyStatus = "Published"
        ComplianceRisk = "Low"
        IsMandatory = $true
        IsActive = $true
        RequiresAcknowledgement = $true
        AcknowledgementType = "One-Time"
        AcknowledgementDeadlineDays = 14
        ReadTimeframe = "Week 2"
        RequiresQuiz = $false
        DistributionScope = "All Employees"
        EstimatedReadTimeMinutes = 20
        ReviewCycleMonths = 12
    },
    @{
        Title = "POL-HR-005 Performance Management"
        PolicyNumber = "POL-HR-005"
        PolicyName = "Performance Management Policy"
        PolicyCategory = "HR Policies"
        PolicyType = "Corporate"
        PolicyDescription = "This policy outlines the performance management framework including goal setting, regular feedback, performance reviews, and development planning. It applies to all employees and their managers."
        VersionNumber = "1.5"
        VersionType = "Minor"
        PolicyStatus = "Published"
        ComplianceRisk = "Medium"
        IsMandatory = $true
        IsActive = $true
        RequiresAcknowledgement = $true
        AcknowledgementType = "On Update"
        AcknowledgementDeadlineDays = 21
        ReadTimeframe = "Month 1"
        RequiresQuiz = $false
        DistributionScope = "All Employees"
        EstimatedReadTimeMinutes = 18
        ReviewCycleMonths = 12
    },

    # IT & SECURITY POLICIES
    @{
        Title = "POL-IT-001 Information Security Policy"
        PolicyNumber = "POL-IT-001"
        PolicyName = "Information Security Policy"
        PolicyCategory = "IT & Security"
        PolicyType = "Corporate"
        PolicyDescription = "This policy establishes the framework for protecting company information assets. It covers data classification, access controls, encryption requirements, and incident response procedures. Compliance is mandatory for all employees handling company data."
        VersionNumber = "5.0"
        VersionType = "Major"
        PolicyStatus = "Published"
        ComplianceRisk = "Critical"
        IsMandatory = $true
        IsActive = $true
        RequiresAcknowledgement = $true
        AcknowledgementType = "Periodic - Annual"
        AcknowledgementDeadlineDays = 7
        ReadTimeframe = "Day 3"
        RequiresQuiz = $true
        QuizPassingScore = 90
        DistributionScope = "All Employees"
        EstimatedReadTimeMinutes = 30
        ReviewCycleMonths = 12
    },
    @{
        Title = "POL-IT-002 Acceptable Use Policy"
        PolicyNumber = "POL-IT-002"
        PolicyName = "Acceptable Use of Technology Policy"
        PolicyCategory = "IT & Security"
        PolicyType = "Corporate"
        PolicyDescription = "Defines acceptable use of company technology resources including computers, mobile devices, email, internet, and software. Covers personal use guidelines, prohibited activities, and monitoring practices."
        VersionNumber = "3.2"
        VersionType = "Minor"
        PolicyStatus = "Published"
        ComplianceRisk = "High"
        IsMandatory = $true
        IsActive = $true
        RequiresAcknowledgement = $true
        AcknowledgementType = "Periodic - Annual"
        AcknowledgementDeadlineDays = 14
        ReadTimeframe = "Week 1"
        RequiresQuiz = $true
        QuizPassingScore = 75
        DistributionScope = "All Employees"
        EstimatedReadTimeMinutes = 15
        ReviewCycleMonths = 12
    },
    @{
        Title = "POL-IT-003 Password and Authentication"
        PolicyNumber = "POL-IT-003"
        PolicyName = "Password and Authentication Policy"
        PolicyCategory = "IT & Security"
        PolicyType = "Corporate"
        PolicyDescription = "Establishes requirements for password complexity, multi-factor authentication, and secure credential management. Includes guidelines for password managers and biometric authentication."
        VersionNumber = "2.0"
        VersionType = "Major"
        PolicyStatus = "Published"
        ComplianceRisk = "High"
        IsMandatory = $true
        IsActive = $true
        RequiresAcknowledgement = $true
        AcknowledgementType = "On Update"
        AcknowledgementDeadlineDays = 7
        ReadTimeframe = "Day 1"
        RequiresQuiz = $true
        QuizPassingScore = 80
        DistributionScope = "All Employees"
        EstimatedReadTimeMinutes = 10
        ReviewCycleMonths = 12
    },
    @{
        Title = "POL-IT-004 Data Backup and Recovery"
        PolicyNumber = "POL-IT-004"
        PolicyName = "Data Backup and Recovery Policy"
        PolicyCategory = "IT & Security"
        PolicyType = "Departmental"
        PolicyDescription = "Defines backup schedules, retention periods, and recovery procedures for all company data. Includes responsibilities for IT staff and end users in maintaining data integrity."
        VersionNumber = "1.3"
        VersionType = "Minor"
        PolicyStatus = "Published"
        ComplianceRisk = "High"
        IsMandatory = $true
        IsActive = $true
        RequiresAcknowledgement = $true
        AcknowledgementType = "One-Time"
        AcknowledgementDeadlineDays = 30
        ReadTimeframe = "Month 1"
        RequiresQuiz = $false
        DistributionScope = "Department"
        EstimatedReadTimeMinutes = 12
        ReviewCycleMonths = 24
    },
    @{
        Title = "POL-IT-005 BYOD Policy"
        PolicyNumber = "POL-IT-005"
        PolicyName = "Bring Your Own Device (BYOD) Policy"
        PolicyCategory = "IT & Security"
        PolicyType = "Corporate"
        PolicyDescription = "Guidelines for employees using personal devices for work purposes. Covers security requirements, MDM enrollment, data separation, and support limitations."
        VersionNumber = "2.1"
        VersionType = "Minor"
        PolicyStatus = "Published"
        ComplianceRisk = "Medium"
        IsMandatory = $false
        IsActive = $true
        RequiresAcknowledgement = $true
        AcknowledgementType = "One-Time"
        AcknowledgementDeadlineDays = 14
        ReadTimeframe = "Week 2"
        RequiresQuiz = $false
        DistributionScope = "All Employees"
        EstimatedReadTimeMinutes = 10
        ReviewCycleMonths = 12
    },

    # HEALTH & SAFETY POLICIES
    @{
        Title = "POL-HS-001 Workplace Health and Safety"
        PolicyNumber = "POL-HS-001"
        PolicyName = "Workplace Health and Safety Policy"
        PolicyCategory = "Health & Safety"
        PolicyType = "Corporate"
        PolicyDescription = "Comprehensive health and safety policy covering workplace hazards, emergency procedures, reporting requirements, and employee responsibilities. Includes specific guidance for office and remote work environments."
        VersionNumber = "4.0"
        VersionType = "Major"
        PolicyStatus = "Published"
        ComplianceRisk = "Critical"
        IsMandatory = $true
        IsActive = $true
        RequiresAcknowledgement = $true
        AcknowledgementType = "Periodic - Annual"
        AcknowledgementDeadlineDays = 7
        ReadTimeframe = "Day 1"
        RequiresQuiz = $true
        QuizPassingScore = 80
        DistributionScope = "All Employees"
        EstimatedReadTimeMinutes = 25
        ReviewCycleMonths = 12
    },
    @{
        Title = "POL-HS-002 Emergency Evacuation Procedures"
        PolicyNumber = "POL-HS-002"
        PolicyName = "Emergency Evacuation Procedures"
        PolicyCategory = "Health & Safety"
        PolicyType = "Regional"
        PolicyDescription = "Site-specific emergency evacuation procedures including assembly points, fire warden responsibilities, and procedures for assisting persons with disabilities."
        VersionNumber = "2.0"
        VersionType = "Major"
        PolicyStatus = "Published"
        ComplianceRisk = "Critical"
        IsMandatory = $true
        IsActive = $true
        RequiresAcknowledgement = $true
        AcknowledgementType = "One-Time"
        AcknowledgementDeadlineDays = 3
        ReadTimeframe = "Day 1"
        RequiresQuiz = $true
        QuizPassingScore = 100
        DistributionScope = "Location"
        EstimatedReadTimeMinutes = 10
        ReviewCycleMonths = 12
    },
    @{
        Title = "POL-HS-003 Mental Health and Wellbeing"
        PolicyNumber = "POL-HS-003"
        PolicyName = "Mental Health and Wellbeing Policy"
        PolicyCategory = "Health & Safety"
        PolicyType = "Corporate"
        PolicyDescription = "Policy supporting employee mental health and wellbeing. Covers available resources, manager responsibilities, workplace adjustments, and return-to-work support after mental health-related absence."
        VersionNumber = "1.2"
        VersionType = "Minor"
        PolicyStatus = "Published"
        ComplianceRisk = "Medium"
        IsMandatory = $false
        IsActive = $true
        RequiresAcknowledgement = $true
        AcknowledgementType = "One-Time"
        AcknowledgementDeadlineDays = 30
        ReadTimeframe = "Month 1"
        RequiresQuiz = $false
        DistributionScope = "All Employees"
        EstimatedReadTimeMinutes = 15
        ReviewCycleMonths = 24
    },

    # COMPLIANCE POLICIES
    @{
        Title = "POL-CO-001 Anti-Bribery and Corruption"
        PolicyNumber = "POL-CO-001"
        PolicyName = "Anti-Bribery and Corruption Policy"
        PolicyCategory = "Compliance"
        PolicyType = "Corporate"
        PolicyDescription = "Zero-tolerance policy on bribery and corruption in accordance with UK Bribery Act and international anti-corruption laws. Covers gifts, hospitality, facilitation payments, and third-party due diligence."
        VersionNumber = "3.0"
        VersionType = "Major"
        PolicyStatus = "Published"
        ComplianceRisk = "Critical"
        IsMandatory = $true
        IsActive = $true
        RequiresAcknowledgement = $true
        AcknowledgementType = "Periodic - Annual"
        AcknowledgementDeadlineDays = 14
        ReadTimeframe = "Week 1"
        RequiresQuiz = $true
        QuizPassingScore = 85
        DistributionScope = "All Employees"
        EstimatedReadTimeMinutes = 20
        ReviewCycleMonths = 12
    },
    @{
        Title = "POL-CO-002 Whistleblowing Policy"
        PolicyNumber = "POL-CO-002"
        PolicyName = "Whistleblowing Policy"
        PolicyCategory = "Compliance"
        PolicyType = "Corporate"
        PolicyDescription = "Establishes procedures for reporting suspected wrongdoing, including fraud, safety violations, and legal breaches. Guarantees protection from retaliation for good-faith reports."
        VersionNumber = "2.1"
        VersionType = "Minor"
        PolicyStatus = "Published"
        ComplianceRisk = "High"
        IsMandatory = $true
        IsActive = $true
        RequiresAcknowledgement = $true
        AcknowledgementType = "One-Time"
        AcknowledgementDeadlineDays = 14
        ReadTimeframe = "Week 2"
        RequiresQuiz = $false
        DistributionScope = "All Employees"
        EstimatedReadTimeMinutes = 12
        ReviewCycleMonths = 24
    },

    # DATA PRIVACY POLICIES
    @{
        Title = "POL-DP-001 Data Protection and Privacy"
        PolicyNumber = "POL-DP-001"
        PolicyName = "Data Protection and Privacy Policy"
        PolicyCategory = "Data Privacy"
        PolicyType = "Corporate"
        PolicyDescription = "Comprehensive GDPR-compliant policy covering personal data processing, data subject rights, breach notification, international transfers, and privacy by design principles."
        VersionNumber = "4.1"
        VersionType = "Minor"
        PolicyStatus = "Published"
        ComplianceRisk = "Critical"
        IsMandatory = $true
        IsActive = $true
        RequiresAcknowledgement = $true
        AcknowledgementType = "Periodic - Annual"
        AcknowledgementDeadlineDays = 14
        ReadTimeframe = "Week 1"
        RequiresQuiz = $true
        QuizPassingScore = 80
        DistributionScope = "All Employees"
        EstimatedReadTimeMinutes = 30
        ReviewCycleMonths = 12
    },
    @{
        Title = "POL-DP-002 Data Retention Policy"
        PolicyNumber = "POL-DP-002"
        PolicyName = "Data Retention and Disposal Policy"
        PolicyCategory = "Data Privacy"
        PolicyType = "Corporate"
        PolicyDescription = "Defines retention periods for all categories of company data, including HR records, financial data, and customer information. Includes secure disposal procedures."
        VersionNumber = "2.0"
        VersionType = "Major"
        PolicyStatus = "Published"
        ComplianceRisk = "High"
        IsMandatory = $true
        IsActive = $true
        RequiresAcknowledgement = $true
        AcknowledgementType = "One-Time"
        AcknowledgementDeadlineDays = 30
        ReadTimeframe = "Month 1"
        RequiresQuiz = $false
        DistributionScope = "All Employees"
        EstimatedReadTimeMinutes = 15
        ReviewCycleMonths = 24
    },

    # FINANCIAL POLICIES
    @{
        Title = "POL-FI-001 Expense Reimbursement"
        PolicyNumber = "POL-FI-001"
        PolicyName = "Expense Reimbursement Policy"
        PolicyCategory = "Financial"
        PolicyType = "Corporate"
        PolicyDescription = "Guidelines for business expense claims including eligible expenses, approval limits, receipt requirements, and reimbursement timelines. Covers travel, meals, accommodation, and equipment."
        VersionNumber = "3.5"
        VersionType = "Minor"
        PolicyStatus = "Published"
        ComplianceRisk = "Medium"
        IsMandatory = $true
        IsActive = $true
        RequiresAcknowledgement = $true
        AcknowledgementType = "One-Time"
        AcknowledgementDeadlineDays = 14
        ReadTimeframe = "Week 2"
        RequiresQuiz = $false
        DistributionScope = "All Employees"
        EstimatedReadTimeMinutes = 12
        ReviewCycleMonths = 12
    },
    @{
        Title = "POL-FI-002 Procurement Policy"
        PolicyNumber = "POL-FI-002"
        PolicyName = "Procurement and Purchasing Policy"
        PolicyCategory = "Financial"
        PolicyType = "Corporate"
        PolicyDescription = "Establishes purchasing procedures, approval thresholds, preferred supplier requirements, and competitive bidding processes. Includes guidance on contract management."
        VersionNumber = "2.2"
        VersionType = "Minor"
        PolicyStatus = "Published"
        ComplianceRisk = "High"
        IsMandatory = $true
        IsActive = $true
        RequiresAcknowledgement = $true
        AcknowledgementType = "One-Time"
        AcknowledgementDeadlineDays = 30
        ReadTimeframe = "Month 1"
        RequiresQuiz = $false
        DistributionScope = "Role"
        EstimatedReadTimeMinutes = 20
        ReviewCycleMonths = 24
    },

    # DRAFT/REVIEW POLICIES (to show different statuses)
    @{
        Title = "POL-HR-006 Sabbatical Leave"
        PolicyNumber = "POL-HR-006"
        PolicyName = "Sabbatical Leave Policy"
        PolicyCategory = "HR Policies"
        PolicyType = "Corporate"
        PolicyDescription = "Draft policy for extended unpaid leave for long-serving employees. Covers eligibility (5+ years service), duration options, job protection, and return-to-work arrangements."
        VersionNumber = "0.3"
        VersionType = "Draft"
        PolicyStatus = "Draft"
        ComplianceRisk = "Low"
        IsMandatory = $false
        IsActive = $false
        RequiresAcknowledgement = $true
        AcknowledgementType = "One-Time"
        AcknowledgementDeadlineDays = 30
        ReadTimeframe = "Month 1"
        RequiresQuiz = $false
        DistributionScope = "All Employees"
        EstimatedReadTimeMinutes = 10
        ReviewCycleMonths = 24
    },
    @{
        Title = "POL-IT-006 AI and Machine Learning Usage"
        PolicyNumber = "POL-IT-006"
        PolicyName = "AI and Machine Learning Acceptable Use Policy"
        PolicyCategory = "IT & Security"
        PolicyType = "Corporate"
        PolicyDescription = "Guidelines for the responsible use of AI tools including ChatGPT, Copilot, and other generative AI. Covers data confidentiality, accuracy verification, and disclosure requirements."
        VersionNumber = "1.0"
        VersionType = "Major"
        PolicyStatus = "In Review"
        ComplianceRisk = "High"
        IsMandatory = $true
        IsActive = $false
        RequiresAcknowledgement = $true
        AcknowledgementType = "One-Time"
        AcknowledgementDeadlineDays = 14
        ReadTimeframe = "Week 1"
        RequiresQuiz = $true
        QuizPassingScore = 75
        DistributionScope = "All Employees"
        EstimatedReadTimeMinutes = 15
        ReviewCycleMonths = 6
    }
)

# ============================================================================
# CREATE POLICIES
# ============================================================================

Write-Host "`n[1/1] Creating sample policies..." -ForegroundColor Yellow
Write-Host "------------------------------------------------------------" -ForegroundColor Gray

$listName = "JML_Policies"
$createdCount = 0
$today = Get-Date

foreach ($policy in $policies) {
    try {
        # Calculate dates
        $effectiveDate = $today.AddDays(-([int](Get-Random -Minimum 30 -Maximum 365)))
        $publishedDate = $effectiveDate
        $nextReviewDate = $effectiveDate.AddMonths($policy.ReviewCycleMonths)

        # Build values hashtable
        $values = @{
            Title = $policy.Title
            PolicyNumber = $policy.PolicyNumber
            PolicyName = $policy.PolicyName
            PolicyCategory = $policy.PolicyCategory
            PolicyType = $policy.PolicyType
            PolicyDescription = $policy.PolicyDescription
            VersionNumber = $policy.VersionNumber
            VersionType = $policy.VersionType
            PolicyStatus = $policy.PolicyStatus
            ComplianceRisk = $policy.ComplianceRisk
            IsMandatory = $policy.IsMandatory
            IsActive = $policy.IsActive
            RequiresAcknowledgement = $policy.RequiresAcknowledgement
            AcknowledgementType = $policy.AcknowledgementType
            AcknowledgementDeadlineDays = $policy.AcknowledgementDeadlineDays
            ReadTimeframe = $policy.ReadTimeframe
            RequiresQuiz = $policy.RequiresQuiz
            DistributionScope = $policy.DistributionScope
            ReviewCycleMonths = $policy.ReviewCycleMonths
            EffectiveDate = $effectiveDate
            NextReviewDate = $nextReviewDate
        }

        # Add optional fields
        if ($policy.QuizPassingScore) {
            $values.QuizPassingScore = $policy.QuizPassingScore
        }

        if ($policy.PolicyStatus -eq "Published") {
            $values.PublishedDate = $publishedDate
            $values.TotalDistributed = Get-Random -Minimum 50 -Maximum 500
            $values.TotalAcknowledged = [int]($values.TotalDistributed * (Get-Random -Minimum 60 -Maximum 98) / 100)
            $values.CompliancePercentage = [int]($values.TotalAcknowledged / $values.TotalDistributed * 100)
        }

        Add-PnPListItem -List $listName -Values $values | Out-Null
        $createdCount++
        Write-Host "  Created: $($policy.PolicyNumber) - $($policy.PolicyName)" -ForegroundColor Green
    }
    catch {
        Write-Host "  Failed: $($policy.PolicyNumber) - $_" -ForegroundColor Red
    }
}

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  Policies Created: $createdCount / $($policies.Count)" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan

Disconnect-PnPOnline
