# ============================================================================
# Policy Manager — Comprehensive Demo Data Seed
# South African Business Context
# Seeds ALL core lists for a realistic enterprise demo
# Assumes: Already connected via Connect-PnPOnline
# ============================================================================

$ErrorActionPreference = "Continue"
$now = Get-Date
$siteUrl = "https://mf7m.sharepoint.com/sites/PolicyManager"

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Policy Manager — Comprehensive Demo Data Seed" -ForegroundColor Cyan
Write-Host "  South African Enterprise Context" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# ============================================================================
# HELPER: Get current user for Created/Modified fields
# ============================================================================
$currentUser = Get-PnPUser -Identity (Get-PnPProperty -ClientObject (Get-PnPWeb) -Property CurrentUser).LoginName -ErrorAction SilentlyContinue

# South African names for realistic data
$saNames = @(
    @{ First = "Thabo"; Last = "Mokoena"; Dept = "IT"; Title = "IT Manager" },
    @{ First = "Lindiwe"; Last = "Nkosi"; Dept = "HR"; Title = "HR Director" },
    @{ First = "Sipho"; Last = "Dlamini"; Dept = "Finance"; Title = "CFO" },
    @{ First = "Naledi"; Last = "Mahlangu"; Dept = "Compliance"; Title = "Compliance Officer" },
    @{ First = "Bongani"; Last = "Zulu"; Dept = "Operations"; Title = "Operations Manager" },
    @{ First = "Nomsa"; Last = "Khumalo"; Dept = "Legal"; Title = "Legal Counsel" },
    @{ First = "Mandla"; Last = "Ndlovu"; Dept = "IT"; Title = "Security Analyst" },
    @{ First = "Zanele"; Last = "Mthembu"; Dept = "HR"; Title = "Talent Manager" },
    @{ First = "Pieter"; Last = "van der Merwe"; Dept = "Finance"; Title = "Financial Controller" },
    @{ First = "Fatima"; Last = "Patel"; Dept = "Compliance"; Title = "Risk Analyst" },
    @{ First = "Dumisani"; Last = "Ngcobo"; Dept = "Operations"; Title = "Facilities Manager" },
    @{ First = "Priya"; Last = "Naidoo"; Dept = "IT"; Title = "Developer Lead" },
    @{ First = "Kobus"; Last = "Botha"; Dept = "HR"; Title = "L&D Specialist" },
    @{ First = "Ayanda"; Last = "Cele"; Dept = "Legal"; Title = "Contract Specialist" },
    @{ First = "Rashid"; Last = "Khan"; Dept = "Finance"; Title = "Audit Manager" }
)

# ============================================================================
# 1. PM_Configuration — App settings
# ============================================================================
$listName = "PM_Configuration"
Write-Host "[1] Seeding $listName..." -ForegroundColor Yellow
$configItems = @(
    @{ ConfigKey = "App.Version"; ConfigValue = "1.2.5"; Category = "System"; IsActive = $true; IsSystemConfig = $true },
    @{ ConfigKey = "App.CompanyName"; ConfigValue = "First Digital (Pty) Ltd"; Category = "Branding"; IsActive = $true; IsSystemConfig = $false },
    @{ ConfigKey = "App.ProductName"; ConfigValue = "Policy Manager"; Category = "Branding"; IsActive = $true; IsSystemConfig = $false },
    @{ ConfigKey = "Integration.AI.Chat.Enabled"; ConfigValue = "true"; Category = "AI"; IsActive = $true; IsSystemConfig = $false },
    @{ ConfigKey = "Integration.AI.Chat.MaxTokens"; ConfigValue = "1500"; Category = "AI"; IsActive = $true; IsSystemConfig = $false },
    @{ ConfigKey = "Compliance.DefaultAckDeadlineDays"; ConfigValue = "14"; Category = "Compliance"; IsActive = $true; IsSystemConfig = $false },
    @{ ConfigKey = "Compliance.RequireAcknowledgement"; ConfigValue = "true"; Category = "Compliance"; IsActive = $true; IsSystemConfig = $false },
    @{ ConfigKey = "Compliance.DefaultReviewFrequency"; ConfigValue = "Annual"; Category = "Compliance"; IsActive = $true; IsSystemConfig = $false },
    @{ ConfigKey = "Notifications.NewPolicy"; ConfigValue = "true"; Category = "Notifications"; IsActive = $true; IsSystemConfig = $false },
    @{ ConfigKey = "Notifications.DailyDigest"; ConfigValue = "true"; Category = "Notifications"; IsActive = $true; IsSystemConfig = $false },
    @{ ConfigKey = "Upload.MaxDocSizeMB"; ConfigValue = "25"; Category = "Upload"; IsActive = $true; IsSystemConfig = $false },
    @{ ConfigKey = "Upload.MaxVideoSizeMB"; ConfigValue = "100"; Category = "Upload"; IsActive = $true; IsSystemConfig = $false },
    @{ ConfigKey = "Quiz.DefaultPassingScore"; ConfigValue = "75"; Category = "Quiz"; IsActive = $true; IsSystemConfig = $false }
)
$count = 0
foreach ($item in $configItems) {
    try {
        $existing = Get-PnPListItem -List $listName -Query "<View><Query><Where><Eq><FieldRef Name='ConfigKey'/><Value Type='Text'>$($item.ConfigKey)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
        if ($null -eq $existing -or $existing.Count -eq 0) {
            Add-PnPListItem -List $listName -Values @{ Title = $item.ConfigKey; ConfigKey = $item.ConfigKey; ConfigValue = $item.ConfigValue; Category = $item.Category; IsActive = $item.IsActive; IsSystemConfig = $item.IsSystemConfig } | Out-Null
            $count++
        }
    } catch { }
}
Write-Host "    Seeded $count config items" -ForegroundColor Gray

# ============================================================================
# 2. PM_UserProfiles — Team members
# ============================================================================
$listName = "PM_UserProfiles"
Write-Host "[2] Seeding $listName..." -ForegroundColor Yellow
$count = 0
foreach ($person in $saNames) {
    try {
        $email = "$($person.First.ToLower()).$($person.Last.ToLower().Replace(' ', ''))@firstdigital.co.za"
        $roles = switch ($person.Dept) {
            "IT" { "Author;Manager" }
            "HR" { "Author;Manager" }
            "Compliance" { "Admin" }
            "Finance" { "Manager" }
            "Legal" { "Author" }
            "Operations" { "Manager" }
            default { "User" }
        }
        Add-PnPListItem -List $listName -Values @{
            Title = "$($person.First) $($person.Last)"
            Email = $email
            Department = $person.Dept
            JobTitle = $person.Title
            PMRole = $roles.Split(';')[0]
            PMRoles = $roles
            IsActive = $true
        } | Out-Null
        $count++
    } catch { }
}
Write-Host "    Seeded $count user profiles" -ForegroundColor Gray

# ============================================================================
# 3. PM_PolicyAcknowledgements — Realistic ack data
# ============================================================================
$listName = "PM_PolicyAcknowledgements"
Write-Host "[3] Seeding $listName..." -ForegroundColor Yellow

# Get existing policies
$policies = Get-PnPListItem -List "PM_Policies" -Fields "Id","Title","PolicyName","PolicyNumber" -PageSize 100 -ErrorAction SilentlyContinue
$count = 0
if ($policies) {
    $statuses = @("Acknowledged", "Acknowledged", "Acknowledged", "Acknowledged", "Pending", "Pending", "Overdue")
    foreach ($policy in ($policies | Select-Object -First 15)) {
        foreach ($person in ($saNames | Get-Random -Count (Get-Random -Minimum 3 -Maximum 8))) {
            $status = $statuses | Get-Random
            $assignedDate = $now.AddDays(-(Get-Random -Minimum 5 -Maximum 60))
            $dueDate = $assignedDate.AddDays(14)
            $ackDate = if ($status -eq "Acknowledged") { $assignedDate.AddDays((Get-Random -Minimum 1 -Maximum 10)) } else { $null }
            try {
                $values = @{
                    Title = "$($person.First) $($person.Last) - $($policy.FieldValues.PolicyName)"
                    PolicyId = $policy.Id
                    PolicyTitle = $policy.FieldValues.PolicyName
                    UserId = "$($person.First.ToLower()).$($person.Last.ToLower().Replace(' ', ''))@firstdigital.co.za"
                    UserDisplayName = "$($person.First) $($person.Last)"
                    Department = $person.Dept
                    AckStatus = $status
                    AssignedDate = $assignedDate.ToString("o")
                    DueDate = $dueDate.ToString("o")
                }
                if ($ackDate) {
                    $values["AcknowledgedDate"] = $ackDate.ToString("o")
                    $values["AcknowledgedTime"] = $ackDate.ToString("HH:mm:ss")
                }
                Add-PnPListItem -List $listName -Values $values | Out-Null
                $count++
            } catch { }
        }
    }
}
Write-Host "    Seeded $count acknowledgement records" -ForegroundColor Gray

# ============================================================================
# 4. PM_PolicyDistributions — Distribution campaigns
# ============================================================================
$listName = "PM_PolicyDistributions"
Write-Host "[4] Seeding $listName..." -ForegroundColor Yellow
$distributions = @(
    @{ Title = "Q1 2026 — POPIA Refresh"; DistributionType = "Policy"; TargetAudience = "All Employees"; RecipientCount = 342; AcknowledgedCount = 245; Status = "Active"; DueDate = $now.AddDays(30).ToString("o"); CreatedDate = $now.AddDays(-45).ToString("o") },
    @{ Title = "GDPR Annual Refresher"; DistributionType = "Policy"; TargetAudience = "Department"; RecipientCount = 85; AcknowledgedCount = 82; Status = "Completed"; DueDate = $now.AddDays(-15).ToString("o"); CreatedDate = $now.AddDays(-90).ToString("o") },
    @{ Title = "New Hire — Health & Safety Onboarding"; DistributionType = "PolicyPack"; TargetAudience = "New Hires Only"; RecipientCount = 18; AcknowledgedCount = 10; Status = "Active"; DueDate = $now.AddDays(14).ToString("o"); CreatedDate = $now.AddDays(-30).ToString("o") },
    @{ Title = "Code of Conduct 2026"; DistributionType = "Policy"; TargetAudience = "All Employees"; RecipientCount = 342; AcknowledgedCount = 0; Status = "Scheduled"; DueDate = $now.AddDays(60).ToString("o"); CreatedDate = $now.AddDays(-7).ToString("o") },
    @{ Title = "Finance Team — Anti-Fraud Policy"; DistributionType = "Policy"; TargetAudience = "Role"; RecipientCount = 45; AcknowledgedCount = 38; Status = "Active"; DueDate = $now.AddDays(7).ToString("o"); CreatedDate = $now.AddDays(-21).ToString("o") }
)
$count = 0
foreach ($dist in $distributions) {
    try {
        Add-PnPListItem -List $listName -Values $dist | Out-Null
        $count++
    } catch { }
}
Write-Host "    Seeded $count distribution campaigns" -ForegroundColor Gray

# ============================================================================
# 5. PM_PolicyAuditLog — Audit trail entries
# ============================================================================
$listName = "PM_PolicyAuditLog"
Write-Host "[5] Seeding $listName..." -ForegroundColor Yellow
$auditActions = @(
    @{ ActionType = "PolicyPublished"; ActionCategory = "Policy"; ResourceTitle = "POPIA Compliance Framework"; PerformedBy = "Naledi Mahlangu"; Department = "Compliance" },
    @{ ActionType = "PolicyApproved"; ActionCategory = "Approval"; ResourceTitle = "Anti-Harassment & Workplace Dignity"; PerformedBy = "Lindiwe Nkosi"; Department = "HR" },
    @{ ActionType = "AcknowledgementCompleted"; ActionCategory = "Acknowledgement"; ResourceTitle = "IT Acceptable Use Policy"; PerformedBy = "Thabo Mokoena"; Department = "IT" },
    @{ ActionType = "PolicyCreated"; ActionCategory = "Policy"; ResourceTitle = "Remote Work Policy v2.0"; PerformedBy = "Zanele Mthembu"; Department = "HR" },
    @{ ActionType = "QuizCompleted"; ActionCategory = "Quiz"; ResourceTitle = "POPIA Quiz"; PerformedBy = "Mandla Ndlovu"; Department = "IT" },
    @{ ActionType = "PolicyUpdated"; ActionCategory = "Policy"; ResourceTitle = "King IV Corporate Governance"; PerformedBy = "Sipho Dlamini"; Department = "Finance" },
    @{ ActionType = "DelegationCreated"; ActionCategory = "Delegation"; ResourceTitle = "BCEA Leave Management Review"; PerformedBy = "Bongani Zulu"; Department = "Operations" },
    @{ ActionType = "ApprovalSubmitted"; ActionCategory = "Approval"; ResourceTitle = "Cybersecurity Incident Response Plan"; PerformedBy = "Priya Naidoo"; Department = "IT" },
    @{ ActionType = "PolicyArchived"; ActionCategory = "Policy"; ResourceTitle = "Legacy Travel Policy v1.0"; PerformedBy = "Nomsa Khumalo"; Department = "Legal" },
    @{ ActionType = "AcknowledgementOverdue"; ActionCategory = "Acknowledgement"; ResourceTitle = "Information Security Policy"; PerformedBy = "System"; Department = "System" },
    @{ ActionType = "PolicyPublished"; ActionCategory = "Policy"; ResourceTitle = "Employment Equity Act Compliance"; PerformedBy = "Fatima Patel"; Department = "Compliance" },
    @{ ActionType = "ReviewCompleted"; ActionCategory = "Review"; ResourceTitle = "Whistleblower Protection Policy"; PerformedBy = "Ayanda Cele"; Department = "Legal" },
    @{ ActionType = "PolicyPublished"; ActionCategory = "Policy"; ResourceTitle = "BYOD Policy v2.1"; PerformedBy = "Thabo Mokoena"; Department = "IT" },
    @{ ActionType = "AcknowledgementCompleted"; ActionCategory = "Acknowledgement"; ResourceTitle = "Code of Conduct & Ethics"; PerformedBy = "Pieter van der Merwe"; Department = "Finance" },
    @{ ActionType = "ApprovalRejected"; ActionCategory = "Approval"; ResourceTitle = "Social Media Policy Draft"; PerformedBy = "Lindiwe Nkosi"; Department = "HR" },
    @{ ActionType = "PolicyCreated"; ActionCategory = "Policy"; ResourceTitle = "AI and Machine Learning Acceptable Use"; PerformedBy = "Priya Naidoo"; Department = "IT" },
    @{ ActionType = "SLABreach"; ActionCategory = "SLA"; ResourceTitle = "Finance team acknowledgement SLA"; PerformedBy = "System"; Department = "System" },
    @{ ActionType = "EscalationTriggered"; ActionCategory = "Escalation"; ResourceTitle = "Overdue: POPIA Quiz — Finance dept"; PerformedBy = "System"; Department = "System" },
    @{ ActionType = "AcknowledgementCompleted"; ActionCategory = "Acknowledgement"; ResourceTitle = "BCEA Leave Management"; PerformedBy = "Dumisani Ngcobo"; Department = "Operations" },
    @{ ActionType = "PolicyPublished"; ActionCategory = "Policy"; ResourceTitle = "Procurement and Purchasing Policy"; PerformedBy = "Rashid Khan"; Department = "Finance" }
)
$count = 0
$dayOffset = 0
foreach ($audit in $auditActions) {
    try {
        $perfDate = $now.AddDays(-$dayOffset).AddHours(-(Get-Random -Minimum 1 -Maximum 12))
        Add-PnPListItem -List $listName -Values @{
            Title = "$($audit.ActionType) — $($audit.ResourceTitle)"
            ActionType = $audit.ActionType
            ActionCategory = $audit.ActionCategory
            ResourceTitle = $audit.ResourceTitle
            PerformedBy = $audit.PerformedBy
            PerformedDate = $perfDate.ToString("o")
            Department = $audit.Department
        } | Out-Null
        $count++
        $dayOffset += (Get-Random -Minimum 1 -Maximum 4)
    } catch { }
}
Write-Host "    Seeded $count audit log entries" -ForegroundColor Gray

# ============================================================================
# 6. PM_ReportExecutions — Report generation history
# ============================================================================
$listName = "PM_ReportExecutions"
Write-Host "[6] Seeding $listName..." -ForegroundColor Yellow
$reportExecs = @(
    @{ ReportName = "Department Compliance Report"; ReportType = "dept-compliance"; GeneratedByName = "Thabo Mokoena"; Format = "PDF"; RecordCount = 342; FileSize = "2.4 MB"; ExecutionTime = 4500; ExecutionStatus = "Success" },
    @{ ReportName = "Acknowledgement Status Report"; ReportType = "ack-status"; GeneratedByName = "Lindiwe Nkosi"; Format = "Excel"; RecordCount = 189; FileSize = "1.8 MB"; ExecutionTime = 3200; ExecutionStatus = "Success" },
    @{ ReportName = "SLA Performance Report"; ReportType = "sla-performance"; GeneratedByName = "Sipho Dlamini"; Format = "PDF"; RecordCount = 56; FileSize = "3.1 MB"; ExecutionTime = 5100; ExecutionStatus = "Success" },
    @{ ReportName = "Risk & Violations Report"; ReportType = "risk-violations"; GeneratedByName = "Naledi Mahlangu"; Format = "PDF"; RecordCount = 28; FileSize = "4.2 MB"; ExecutionTime = 6800; ExecutionStatus = "Success" },
    @{ ReportName = "Audit Trail Export"; ReportType = "audit-trail"; GeneratedByName = "Thabo Mokoena"; Format = "CSV"; RecordCount = 1247; FileSize = "890 KB"; ExecutionTime = 2100; ExecutionStatus = "Success" },
    @{ ReportName = "Training Completion Report"; ReportType = "training-completion"; GeneratedByName = "Kobus Botha"; Format = "Excel"; RecordCount = 78; FileSize = "1.5 MB"; ExecutionTime = 3800; ExecutionStatus = "Success" },
    @{ ReportName = "Delegation Summary"; ReportType = "delegation-summary"; GeneratedByName = "Bongani Zulu"; Format = "Excel"; RecordCount = 34; FileSize = "720 KB"; ExecutionTime = 1900; ExecutionStatus = "Success" },
    @{ ReportName = "Department Compliance Report"; ReportType = "dept-compliance"; GeneratedByName = "System (Scheduled)"; Format = "PDF"; RecordCount = 342; FileSize = "2.3 MB"; ExecutionTime = 4200; ExecutionStatus = "Success" },
    @{ ReportName = "Policy Review Schedule"; ReportType = "review-schedule"; GeneratedByName = "Fatima Patel"; Format = "PDF"; RecordCount = 22; FileSize = "1.1 MB"; ExecutionTime = 2800; ExecutionStatus = "Success" },
    @{ ReportName = "Acknowledgement Status Report"; ReportType = "ack-status"; GeneratedByName = "System (Scheduled)"; Format = "Excel"; RecordCount = 195; FileSize = "1.9 MB"; ExecutionTime = 3400; ExecutionStatus = "Success" }
)
$count = 0
$dayOffset = 0
foreach ($exec in $reportExecs) {
    try {
        $execDate = $now.AddDays(-$dayOffset).AddHours(-(Get-Random -Minimum 1 -Maximum 10))
        Add-PnPListItem -List $listName -Values @{
            Title = $exec.ReportName
            ReportName = $exec.ReportName
            ReportType = $exec.ReportType
            GeneratedByName = $exec.GeneratedByName
            Format = $exec.Format
            RecordCount = $exec.RecordCount
            FileSize = $exec.FileSize
            ExecutionTime = $exec.ExecutionTime
            ExecutionStatus = $exec.ExecutionStatus
            ExecutedAt = $execDate.ToString("o")
        } | Out-Null
        $count++
        $dayOffset += (Get-Random -Minimum 1 -Maximum 3)
    } catch { }
}
Write-Host "    Seeded $count report executions" -ForegroundColor Gray

# ============================================================================
# 7. PM_ScheduledReports — Scheduled report configs
# ============================================================================
$listName = "PM_ScheduledReports"
Write-Host "[7] Seeding $listName..." -ForegroundColor Yellow
$schedules = @(
    @{ Title = "Department Compliance Report"; ReportId = "dept-compliance"; ReportType = "dept-compliance"; Frequency = "Weekly"; Format = "PDF"; Recipients = "thabo.mokoena@firstdigital.co.za, lindiwe.nkosi@firstdigital.co.za"; Enabled = $true; NextRun = $now.AddDays(3).ToString("o") },
    @{ Title = "Acknowledgement Status Report"; ReportId = "ack-status"; ReportType = "ack-status"; Frequency = "Daily"; Format = "Excel"; Recipients = "naledi.mahlangu@firstdigital.co.za"; Enabled = $true; NextRun = $now.AddDays(1).ToString("o") },
    @{ Title = "SLA Performance Report"; ReportId = "sla-performance"; ReportType = "sla-performance"; Frequency = "Monthly"; Format = "PDF"; Recipients = "sipho.dlamini@firstdigital.co.za, rashid.khan@firstdigital.co.za"; Enabled = $true; NextRun = $now.AddDays(15).ToString("o") },
    @{ Title = "Risk & Violations Report"; ReportId = "risk-violations"; ReportType = "risk-violations"; Frequency = "Weekly"; Format = "PDF"; Recipients = "fatima.patel@firstdigital.co.za"; Enabled = $false; NextRun = $now.AddDays(5).ToString("o") },
    @{ Title = "Training Completion Report"; ReportId = "training-completion"; ReportType = "training-completion"; Frequency = "Monthly"; Format = "Excel"; Recipients = "kobus.botha@firstdigital.co.za, lindiwe.nkosi@firstdigital.co.za"; Enabled = $true; NextRun = $now.AddDays(20).ToString("o") }
)
$count = 0
foreach ($sched in $schedules) {
    try {
        Add-PnPListItem -List $listName -Values $sched | Out-Null
        $count++
    } catch { }
}
Write-Host "    Seeded $count scheduled reports" -ForegroundColor Gray

# ============================================================================
# 8. PM_PolicyFeedback — User feedback
# ============================================================================
$listName = "PM_PolicyFeedback"
Write-Host "[8] Seeding $listName..." -ForegroundColor Yellow
$feedbacks = @(
    @{ Title = "POPIA framework very comprehensive"; FeedbackType = "Positive"; PolicyTitle = "POPIA Compliance Framework"; UserName = "Mandla Ndlovu"; Department = "IT"; Rating = 5; Comment = "The updated POPIA framework is excellent. Very clear on data breach notification procedures. The SA-specific examples really help." },
    @{ Title = "Code of Conduct needs examples"; FeedbackType = "Suggestion"; PolicyTitle = "Code of Conduct & Ethics"; UserName = "Zanele Mthembu"; Department = "HR"; Rating = 4; Comment = "Good policy overall but could benefit from more real-world examples of ethical dilemmas specific to our industry." },
    @{ Title = "BCEA leave policy unclear on family responsibility"; FeedbackType = "Issue"; PolicyTitle = "BCEA Leave Management"; UserName = "Dumisani Ngcobo"; Department = "Operations"; Rating = 3; Comment = "Section 3.4 on family responsibility leave is confusing. Does it apply to extended family as per customary law?" },
    @{ Title = "IT security quiz too difficult"; FeedbackType = "Suggestion"; PolicyTitle = "Information Security Policy"; UserName = "Pieter van der Merwe"; Department = "Finance"; Rating = 3; Comment = "The quiz questions on network segmentation are too technical for non-IT staff. Consider separate quizzes per audience." },
    @{ Title = "Whistleblower policy builds confidence"; FeedbackType = "Positive"; PolicyTitle = "Whistleblower Protection"; UserName = "Ayanda Cele"; Department = "Legal"; Rating = 5; Comment = "Excellent alignment with the Protected Disclosures Act. The anonymous reporting channels are clearly explained." },
    @{ Title = "BYOD policy needs update for tablets"; FeedbackType = "Suggestion"; PolicyTitle = "BYOD Policy"; UserName = "Priya Naidoo"; Department = "IT"; Rating = 4; Comment = "Policy covers smartphones and laptops but doesn't address tablets and wearables. Also need clarity on personal hotspot usage." },
    @{ Title = "Health & Safety induction very helpful"; FeedbackType = "Positive"; PolicyTitle = "Workplace Health and Safety"; UserName = "Bongani Zulu"; Department = "Operations"; Rating = 5; Comment = "The OHS Act references are spot-on. The emergency evacuation procedures are clear and well-illustrated." },
    @{ Title = "Employment Equity reporting unclear"; FeedbackType = "Issue"; PolicyTitle = "Employment Equity Act Compliance"; UserName = "Fatima Patel"; Department = "Compliance"; Rating = 3; Comment = "The annual EE reporting section needs more detail on designated employer obligations under Section 19-26." }
)
$count = 0
foreach ($fb in $feedbacks) {
    try {
        Add-PnPListItem -List $listName -Values @{
            Title = $fb.Title
            FeedbackType = $fb.FeedbackType
            PolicyTitle = $fb.PolicyTitle
            UserName = $fb.UserName
            Department = $fb.Department
            Rating = $fb.Rating
            Comment = $fb.Comment
            SubmittedDate = $now.AddDays(-(Get-Random -Minimum 1 -Maximum 30)).ToString("o")
            Status = "Open"
        } | Out-Null
        $count++
    } catch { }
}
Write-Host "    Seeded $count feedback items" -ForegroundColor Gray

# ============================================================================
# 9. PM_PolicyRequests — Policy creation requests
# ============================================================================
$listName = "PM_PolicyRequests"
Write-Host "[9] Seeding $listName..." -ForegroundColor Yellow
$requests = @(
    @{ Title = "Social Media Usage Policy"; RequestedBy = "Zanele Mthembu"; Department = "HR"; Priority = "Medium"; Status = "Pending"; BusinessJustification = "Employees are using personal social media during work hours and some have posted company-confidential information. We need a clear policy aligned with the ECTA (Electronic Communications and Transactions Act)." },
    @{ Title = "Environmental Sustainability Policy"; RequestedBy = "Bongani Zulu"; Department = "Operations"; Priority = "Low"; Status = "Pending"; BusinessJustification = "With the new Carbon Tax Act and growing ESG requirements from investors, we need a formal environmental policy covering waste management, energy efficiency, and carbon reporting." },
    @{ Title = "Third-Party Risk Management Policy"; RequestedBy = "Fatima Patel"; Department = "Compliance"; Priority = "High"; Status = "In Progress"; BusinessJustification = "Recent SARB (South African Reserve Bank) guidance requires financial services firms to have formal third-party risk assessments. Critical for our regulatory compliance." },
    @{ Title = "Remote Work Policy Update"; RequestedBy = "Lindiwe Nkosi"; Department = "HR"; Priority = "High"; Status = "Approved"; BusinessJustification = "Post-COVID hybrid work arrangements need formal policy update. Must address COIDA implications for work-from-home injuries and UIF contributions for flexible workers." },
    @{ Title = "Cryptocurrency and Digital Assets Policy"; RequestedBy = "Sipho Dlamini"; Department = "Finance"; Priority = "Medium"; Status = "Pending"; BusinessJustification = "With SARS (South African Revenue Service) issuing guidance on crypto taxation and FSCA licensing requirements, we need a policy on employee cryptocurrency activities and company digital asset management." }
)
$count = 0
foreach ($req in $requests) {
    try {
        Add-PnPListItem -List $listName -Values @{
            Title = $req.Title
            RequestedBy = $req.RequestedBy
            Department = $req.Department
            Priority = $req.Priority
            Status = $req.Status
            BusinessJustification = $req.BusinessJustification
            RequestedDate = $now.AddDays(-(Get-Random -Minimum 3 -Maximum 21)).ToString("o")
        } | Out-Null
        $count++
    } catch { }
}
Write-Host "    Seeded $count policy requests" -ForegroundColor Gray

# ============================================================================
# 10. PM_SLAConfigs — SLA targets
# ============================================================================
$listName = "PM_SLAConfigs"
Write-Host "[10] Seeding $listName..." -ForegroundColor Yellow
$slaConfigs = @(
    @{ Title = "Acknowledgement SLA"; SLAType = "Acknowledgement"; TargetDays = 14; WarningThresholdDays = 10; Category = "All"; IsActive = $true },
    @{ Title = "Approval SLA"; SLAType = "Approval"; TargetDays = 5; WarningThresholdDays = 3; Category = "All"; IsActive = $true },
    @{ Title = "Review SLA"; SLAType = "Review"; TargetDays = 30; WarningThresholdDays = 21; Category = "All"; IsActive = $true },
    @{ Title = "Critical Policy Ack SLA"; SLAType = "Acknowledgement"; TargetDays = 7; WarningThresholdDays = 5; Category = "Critical"; IsActive = $true },
    @{ Title = "Authoring SLA"; SLAType = "Authoring"; TargetDays = 21; WarningThresholdDays = 14; Category = "All"; IsActive = $true }
)
$count = 0
foreach ($sla in $slaConfigs) {
    try {
        Add-PnPListItem -List $listName -Values $sla | Out-Null
        $count++
    } catch { }
}
Write-Host "    Seeded $count SLA configurations" -ForegroundColor Gray

# ============================================================================
# 11. PM_NamingRules — Policy naming conventions
# ============================================================================
$listName = "PM_NamingRules"
Write-Host "[11] Seeding $listName..." -ForegroundColor Yellow
$namingRules = @(
    @{ Title = "Standard Policy"; Pattern = "POL-{CATEGORY}-{COUNTER}"; AppliesTo = "All Policies"; IsActive = $true; Example = "POL-HR-001" },
    @{ Title = "IT Policy"; Pattern = "POL-IT-{COUNTER}"; AppliesTo = "IT Policies"; IsActive = $true; Example = "POL-IT-005" },
    @{ Title = "Compliance Policy"; Pattern = "POL-COM-{COUNTER}"; AppliesTo = "Compliance Policies"; IsActive = $true; Example = "POL-COM-003" },
    @{ Title = "Health & Safety"; Pattern = "POL-HS-{COUNTER}"; AppliesTo = "All Policies"; IsActive = $true; Example = "POL-HS-001" },
    @{ Title = "Finance Policy"; Pattern = "POL-FI-{COUNTER}"; AppliesTo = "Finance Policies"; IsActive = $true; Example = "POL-FI-002" }
)
$count = 0
foreach ($rule in $namingRules) {
    try {
        Add-PnPListItem -List $listName -Values $rule | Out-Null
        $count++
    } catch { }
}
Write-Host "    Seeded $count naming rules" -ForegroundColor Gray

# ============================================================================
# 12. PM_PolicySubCategories — Sub-categories
# ============================================================================
$listName = "PM_PolicySubCategories"
Write-Host "[12] Seeding $listName..." -ForegroundColor Yellow
$subCats = @(
    @{ SubCategoryName = "Leave & Benefits"; ParentCategoryName = "HR Policies"; SortOrder = 1; IsActive = $true },
    @{ SubCategoryName = "Recruitment & Onboarding"; ParentCategoryName = "HR Policies"; SortOrder = 2; IsActive = $true },
    @{ SubCategoryName = "Employee Relations"; ParentCategoryName = "HR Policies"; SortOrder = 3; IsActive = $true },
    @{ SubCategoryName = "Network Security"; ParentCategoryName = "IT & Security"; SortOrder = 1; IsActive = $true },
    @{ SubCategoryName = "Data Protection"; ParentCategoryName = "IT & Security"; SortOrder = 2; IsActive = $true },
    @{ SubCategoryName = "Device Management"; ParentCategoryName = "IT & Security"; SortOrder = 3; IsActive = $true },
    @{ SubCategoryName = "Workplace Safety"; ParentCategoryName = "Health & Safety"; SortOrder = 1; IsActive = $true },
    @{ SubCategoryName = "Emergency Procedures"; ParentCategoryName = "Health & Safety"; SortOrder = 2; IsActive = $true },
    @{ SubCategoryName = "Regulatory Compliance"; ParentCategoryName = "Compliance"; SortOrder = 1; IsActive = $true },
    @{ SubCategoryName = "Data Privacy (POPIA)"; ParentCategoryName = "Compliance"; SortOrder = 2; IsActive = $true },
    @{ SubCategoryName = "Financial Controls"; ParentCategoryName = "Financial"; SortOrder = 1; IsActive = $true },
    @{ SubCategoryName = "Procurement"; ParentCategoryName = "Financial"; SortOrder = 2; IsActive = $true }
)
$count = 0
foreach ($sc in $subCats) {
    try {
        Add-PnPListItem -List $listName -Values @{
            Title = $sc.SubCategoryName
            SubCategoryName = $sc.SubCategoryName
            ParentCategoryName = $sc.ParentCategoryName
            SortOrder = $sc.SortOrder
            IsActive = $sc.IsActive
        } | Out-Null
        $count++
    } catch { }
}
Write-Host "    Seeded $count sub-categories" -ForegroundColor Gray

# ============================================================================
# 13. PM_ReportDefinitions — Report templates
# ============================================================================
$listName = "PM_ReportDefinitions"
Write-Host "[13] Seeding $listName..." -ForegroundColor Yellow
$reportDefs = @(
    @{ ReportName = "Department Compliance Report"; Description = "Full compliance status for all team members with acknowledgement breakdown by department"; Category = "Compliance"; Status = "Published"; IsPublic = $true; IsTemplate = $true },
    @{ ReportName = "Acknowledgement Status Report"; Description = "Detailed list of pending and overdue policy acknowledgements across the organisation"; Category = "Compliance"; Status = "Published"; IsPublic = $true; IsTemplate = $true },
    @{ ReportName = "SLA Performance Report"; Description = "Team SLA metrics for acknowledgement, review, and approval turnarounds"; Category = "Operational"; Status = "Published"; IsPublic = $true; IsTemplate = $true },
    @{ ReportName = "Audit Trail Export"; Description = "Complete log of all policy-related actions by team members"; Category = "Compliance"; Status = "Published"; IsPublic = $true; IsTemplate = $true },
    @{ ReportName = "Risk & Violations Report"; Description = "Identify non-compliant areas, policy violations, and risk exposure across departments"; Category = "Compliance"; Status = "Published"; IsPublic = $true; IsTemplate = $true },
    @{ ReportName = "Training Completion Report"; Description = "Track policy training modules completed by team members with pass rates"; Category = "HR"; Status = "Published"; IsPublic = $true; IsTemplate = $true }
)
$count = 0
foreach ($rd in $reportDefs) {
    try {
        Add-PnPListItem -List $listName -Values @{
            Title = $rd.ReportName
            ReportName = $rd.ReportName
            Description = $rd.Description
            Category = $rd.Category
            Status = $rd.Status
            IsPublic = $rd.IsPublic
            IsTemplate = $rd.IsTemplate
        } | Out-Null
        $count++
    } catch { }
}
Write-Host "    Seeded $count report definitions" -ForegroundColor Gray

# ============================================================================
# SUMMARY
# ============================================================================
Write-Host ""
Write-Host "================================================================" -ForegroundColor Green
Write-Host "  Demo Data Seeding Complete!" -ForegroundColor Green
Write-Host "  Lists seeded:" -ForegroundColor Green
Write-Host "    1.  PM_Configuration (app settings)" -ForegroundColor White
Write-Host "    2.  PM_UserProfiles (15 SA team members)" -ForegroundColor White
Write-Host "    3.  PM_PolicyAcknowledgements (realistic ack data)" -ForegroundColor White
Write-Host "    4.  PM_PolicyDistributions (5 campaigns)" -ForegroundColor White
Write-Host "    5.  PM_PolicyAuditLog (20 audit entries)" -ForegroundColor White
Write-Host "    6.  PM_ReportExecutions (10 report runs)" -ForegroundColor White
Write-Host "    7.  PM_ScheduledReports (5 scheduled reports)" -ForegroundColor White
Write-Host "    8.  PM_PolicyFeedback (8 feedback items)" -ForegroundColor White
Write-Host "    9.  PM_PolicyRequests (5 SA-specific requests)" -ForegroundColor White
Write-Host "    10. PM_SLAConfigs (5 SLA targets)" -ForegroundColor White
Write-Host "    11. PM_NamingRules (5 naming conventions)" -ForegroundColor White
Write-Host "    12. PM_PolicySubCategories (12 sub-categories)" -ForegroundColor White
Write-Host "    13. PM_ReportDefinitions (6 report templates)" -ForegroundColor White
Write-Host "" -ForegroundColor White
Write-Host "  NOTE: Run AFTER Deploy-SampleData.ps1 and" -ForegroundColor Yellow
Write-Host "        Seed-ApprovalAndNotificationData.ps1" -ForegroundColor Yellow
Write-Host "        (this script supplements those, not replaces)" -ForegroundColor Yellow
Write-Host "================================================================" -ForegroundColor Green
Write-Host ""
