# ============================================================================
# Script 23: Seed Fast Track Templates (PM_PolicyMetadataProfiles)
# Creates pre-configured templates for the Fast Track wizard mode
# ============================================================================

$listName = "PM_PolicyMetadataProfiles"
Write-Host "`n=== Seeding Fast Track Templates ===" -ForegroundColor Cyan

$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    Write-Host "  List $listName does not exist. Please run provisioning first." -ForegroundColor Red
    return
}

$templates = @(
    @{
        Title = "IT Security Policy"
        ProfileName = "IT Security Policy"
        PolicyCategory = "IT & Security"
        ComplianceRisk = "High"
        ReadTimeframe = "Week 1"
        RequiresAcknowledgement = $true
        RequiresQuiz = $true
        Description = "Standard IT security template — data classification, access controls, incident response. High risk, requires acknowledgement and quiz."
        IsActive = $true
    },
    @{
        Title = "HR Employee Policy"
        ProfileName = "HR Employee Policy"
        PolicyCategory = "HR Policies"
        ComplianceRisk = "Medium"
        ReadTimeframe = "Week 1"
        RequiresAcknowledgement = $true
        RequiresQuiz = $false
        Description = "Standard HR policy for employee handbook chapters. Medium risk, annual review, requires acknowledgement."
        IsActive = $true
    },
    @{
        Title = "Regulatory Compliance"
        ProfileName = "Regulatory Compliance"
        PolicyCategory = "Compliance"
        ComplianceRisk = "Critical"
        ReadTimeframe = "Day 3"
        RequiresAcknowledgement = $true
        RequiresQuiz = $true
        Description = "Critical compliance policy with regulatory framework references. Requires acknowledgement, quiz, and quarterly review."
        IsActive = $true
    },
    @{
        Title = "Health & Safety"
        ProfileName = "Health & Safety"
        PolicyCategory = "Health & Safety"
        ComplianceRisk = "High"
        ReadTimeframe = "Week 1"
        RequiresAcknowledgement = $true
        RequiresQuiz = $true
        Description = "Workplace health and safety policy — risk assessments, incident reporting, emergency procedures. High risk, biannual review."
        IsActive = $true
    },
    @{
        Title = "Financial Policy"
        ProfileName = "Financial Policy"
        PolicyCategory = "Financial"
        ComplianceRisk = "High"
        ReadTimeframe = "Week 2"
        RequiresAcknowledgement = $true
        RequiresQuiz = $false
        Description = "Financial controls and reporting policy. High risk, requires acknowledgement, annual review."
        IsActive = $true
    },
    @{
        Title = "Data Privacy (POPIA/GDPR)"
        ProfileName = "Data Privacy (POPIA/GDPR)"
        PolicyCategory = "Compliance"
        ComplianceRisk = "Critical"
        ReadTimeframe = "Day 3"
        RequiresAcknowledgement = $true
        RequiresQuiz = $true
        Description = "Data protection and privacy policy aligned to POPIA/GDPR. Critical risk, requires acknowledgement and quiz."
        IsActive = $true
    },
    @{
        Title = "Operational Procedure"
        ProfileName = "Operational Procedure"
        PolicyCategory = "Operational"
        ComplianceRisk = "Low"
        ReadTimeframe = "Week 2"
        RequiresAcknowledgement = $true
        RequiresQuiz = $false
        Description = "Standard operating procedure template. Low risk, requires acknowledgement only."
        IsActive = $true
    },
    @{
        Title = "Legal & Governance"
        ProfileName = "Legal & Governance"
        PolicyCategory = "Legal"
        ComplianceRisk = "High"
        ReadTimeframe = "Week 1"
        RequiresAcknowledgement = $true
        RequiresQuiz = $false
        Description = "Legal governance and corporate compliance policy. High risk, annual review."
        IsActive = $true
    },
    @{
        Title = "Environmental Policy"
        ProfileName = "Environmental Policy"
        PolicyCategory = "Environmental"
        ComplianceRisk = "Medium"
        ReadTimeframe = "Month 1"
        RequiresAcknowledgement = $true
        RequiresQuiz = $false
        Description = "Environmental sustainability and ESG policy. Medium risk, annual review cycle."
        IsActive = $true
    },
    @{
        Title = "Quick Internal Notice"
        ProfileName = "Quick Internal Notice"
        PolicyCategory = "Operational"
        ComplianceRisk = "Low"
        ReadTimeframe = "Day 1"
        RequiresAcknowledgement = $false
        RequiresQuiz = $false
        Description = "Lightweight internal notice or announcement. No acknowledgement required. Fastest path to publish."
        IsActive = $true
    }
)

foreach ($t in $templates) {
    try {
        $existing = Get-PnPListItem -List $listName -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($t.Title)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
        if ($null -eq $existing -or $existing.Count -eq 0) {
            Add-PnPListItem -List $listName -Values @{
                Title = $t.Title
                ProfileName = $t.ProfileName
                PolicyCategory = $t.PolicyCategory
                ComplianceRisk = $t.ComplianceRisk
                ReadTimeframe = $t.ReadTimeframe
                RequiresAcknowledgement = $t.RequiresAcknowledgement
                RequiresQuiz = $t.RequiresQuiz
                Description = $t.Description
                IsActive = $t.IsActive
            } | Out-Null
            Write-Host "  + $($t.Title)" -ForegroundColor Green
        } else {
            Write-Host "  ~ $($t.Title) (exists)" -ForegroundColor Gray
        }
    } catch {
        Write-Host "  ! $($t.Title): $_" -ForegroundColor Red
    }
}

Write-Host "`n=== Fast Track Templates Complete ===" -ForegroundColor Cyan
