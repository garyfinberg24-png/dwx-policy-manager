# ============================================================================
# Script 21: PM_EmailTemplates List
# Stores customizable email templates with merge tags for all notification events
# ============================================================================

$listName = "PM_EmailTemplates"
Write-Host "`n=== Creating $listName ===" -ForegroundColor Cyan

$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

# Title = Template Name (built-in)
Add-PnPField -List $listName -DisplayName "Event" -InternalName "Event" -Type Text -Required -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Category" -InternalName "Category" -Type Choice -Choices "Acknowledgement","Approval","Quiz","Review","Distribution","Compliance","Lifecycle","System" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Subject" -InternalName "Subject" -Type Text -Required -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Body" -InternalName "Body" -Type Note -Required -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Recipients" -InternalName "Recipients" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "IsActive" -Values @{DefaultValue="1"} -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Default" -InternalName "IsDefault" -Type Boolean -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "IsDefault" -Values @{DefaultValue="1"} -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Merge Tags" -InternalName "MergeTags" -Type Note -ErrorAction SilentlyContinue

Set-PnPField -List $listName -Identity "Event" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "IsActive" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  $listName configured" -ForegroundColor Green

# ============================================================================
# Seed Default Templates
# ============================================================================
Write-Host "`n  Seeding default email templates..." -ForegroundColor Yellow

$templates = @(
    @{ Title="New Policy Published"; Event="policy-published"; Category="Acknowledgement"; Subject="New Policy: {{PolicyTitle}}"; Body="<p>A new policy <strong>{{PolicyTitle}}</strong> has been published and requires your attention.</p><p>Please read and acknowledge by <strong>{{Deadline}}</strong>.</p><p><a href='{{PolicyUrl}}'>View Policy</a></p>"; Recipients="All Employees"; MergeTags="PolicyTitle,PolicyNumber,Deadline,PolicyUrl" },
    @{ Title="Acknowledgement Required"; Event="ack-required"; Category="Acknowledgement"; Subject="Action Required: Acknowledge {{PolicyTitle}}"; Body="<p>Hi {{RecipientName}},</p><p>You are required to read and acknowledge <strong>{{PolicyTitle}}</strong>.</p><p>Deadline: <strong>{{Deadline}}</strong></p><p><a href='{{PolicyUrl}}'>View & Acknowledge</a></p>"; Recipients="Assigned Users"; MergeTags="PolicyTitle,RecipientName,Deadline,PolicyUrl" },
    @{ Title="Ack Reminder (3-day)"; Event="ack-reminder-3day"; Category="Acknowledgement"; Subject="Reminder: {{PolicyTitle}} — 3 days remaining"; Body="<p>Hi {{RecipientName}},</p><p>This is a friendly reminder that you have <strong>3 days</strong> remaining to acknowledge <strong>{{PolicyTitle}}</strong>.</p><p><a href='{{PolicyUrl}}'>Acknowledge Now</a></p>"; Recipients="Assigned Users"; MergeTags="PolicyTitle,RecipientName,Deadline,PolicyUrl" },
    @{ Title="Ack Reminder (1-day)"; Event="ack-reminder-1day"; Category="Acknowledgement"; Subject="URGENT: {{PolicyTitle}} — due tomorrow"; Body="<p>Hi {{RecipientName}},</p><p><strong>Final reminder:</strong> Your acknowledgement of <strong>{{PolicyTitle}}</strong> is due <strong>tomorrow</strong>.</p><p><a href='{{PolicyUrl}}'>Acknowledge Now</a></p>"; Recipients="Assigned Users"; MergeTags="PolicyTitle,RecipientName,Deadline,PolicyUrl" },
    @{ Title="Acknowledgement Overdue"; Event="ack-overdue"; Category="Acknowledgement"; Subject="OVERDUE: {{PolicyTitle}} — acknowledgement required"; Body="<p>Hi {{RecipientName}},</p><p>Your acknowledgement of <strong>{{PolicyTitle}}</strong> is now <strong>overdue</strong>. Please complete this immediately.</p><p><a href='{{PolicyUrl}}'>Acknowledge Now</a></p>"; Recipients="Assigned Users"; MergeTags="PolicyTitle,RecipientName,DaysOverdue,PolicyUrl" },
    @{ Title="Ack Complete (Manager)"; Event="ack-complete"; Category="Acknowledgement"; Subject="{{EmployeeName}} acknowledged {{PolicyTitle}}"; Body="<p>{{EmployeeName}} has acknowledged <strong>{{PolicyTitle}}</strong>.</p><p>Team compliance: <strong>{{ComplianceRate}}%</strong></p>"; Recipients="Managers"; MergeTags="EmployeeName,PolicyTitle,ComplianceRate" },
    @{ Title="Approval Request"; Event="approval-request"; Category="Approval"; Subject="Approval Required: {{PolicyTitle}}"; Body="<p>A policy requires your approval:</p><p><strong>{{PolicyTitle}}</strong></p><p>Submitted by: {{AuthorName}}<br/>Category: {{Category}}<br/>Risk Level: {{RiskLevel}}</p><p><a href='{{PolicyUrl}}'>Review & Approve</a></p>"; Recipients="Approvers"; MergeTags="PolicyTitle,AuthorName,Category,RiskLevel,PolicyUrl" },
    @{ Title="Approval Approved"; Event="approval-approved"; Category="Approval"; Subject="Approved: {{PolicyTitle}}"; Body="<p>Great news! <strong>{{PolicyTitle}}</strong> has been approved by <strong>{{ApproverName}}</strong>.</p><p>{{Comments}}</p>"; Recipients="Policy Owners"; MergeTags="PolicyTitle,ApproverName,Comments" },
    @{ Title="Approval Rejected"; Event="approval-rejected"; Category="Approval"; Subject="Rejected: {{PolicyTitle}}"; Body="<p><strong>{{PolicyTitle}}</strong> has been rejected by <strong>{{ApproverName}}</strong>.</p><p><strong>Reason:</strong> {{Comments}}</p><p>Please review the feedback and resubmit.</p>"; Recipients="Policy Owners"; MergeTags="PolicyTitle,ApproverName,Comments" },
    @{ Title="Review Required"; Event="review-required"; Category="Review"; Subject="Review Required: {{PolicyTitle}}"; Body="<p>Hi {{RecipientName}},</p><p><strong>{{AuthorName}}</strong> has submitted <strong>{{PolicyTitle}}</strong> for your review.</p><p>Please log in to Policy Manager to review and provide feedback.</p><p><a href='{{PolicyUrl}}'>Review Policy</a></p>"; Recipients="Reviewers"; MergeTags="PolicyTitle,AuthorName,RecipientName,PolicyUrl" },
    @{ Title="Review Withdrawn"; Event="review-withdrawn"; Category="Review"; Subject="Review Cancelled: {{PolicyTitle}}"; Body="<p>The review for <strong>{{PolicyTitle}}</strong> has been withdrawn by the author.</p><p>No action is required.</p>"; Recipients="Reviewers"; MergeTags="PolicyTitle,AuthorName" },
    @{ Title="Policy Expiring"; Event="policy-expiring"; Category="Compliance"; Subject="Policy Expiring: {{PolicyTitle}}"; Body="<p><strong>{{PolicyTitle}}</strong> will expire on <strong>{{ExpiryDate}}</strong>.</p><p>Please review and either renew or retire this policy.</p><p><a href='{{PolicyUrl}}'>View Policy</a></p>"; Recipients="Policy Owners"; MergeTags="PolicyTitle,ExpiryDate,DaysUntilExpiry,PolicyUrl" },
    @{ Title="Policy Updated"; Event="policy-updated"; Category="Lifecycle"; Subject="Policy Updated: {{PolicyTitle}} v{{Version}}"; Body="<p><strong>{{PolicyTitle}}</strong> has been updated to version <strong>{{Version}}</strong>.</p><p>Changes: {{ChangeDescription}}</p><p><a href='{{PolicyUrl}}'>View Updated Policy</a></p>"; Recipients="All Employees"; MergeTags="PolicyTitle,Version,ChangeDescription,PolicyUrl" },
    @{ Title="Policy Retired"; Event="policy-retired"; Category="Lifecycle"; Subject="Policy Retired: {{PolicyTitle}}"; Body="<p><strong>{{PolicyTitle}}</strong> has been retired and is no longer in effect.</p><p>Replacement: {{ReplacementPolicy}}</p>"; Recipients="All Employees"; MergeTags="PolicyTitle,ReplacementPolicy,RetiredDate" },
    @{ Title="Welcome Email"; Event="user-welcome"; Category="System"; Subject="Welcome to Policy Manager"; Body="<p>Welcome, {{RecipientName}}!</p><p>Policy Manager is where you will find all company policies. Please review the policies assigned to you in <strong>My Policies</strong>.</p><p><a href='{{PolicyHubUrl}}'>Go to Policy Hub</a></p>"; Recipients="New Users"; MergeTags="RecipientName,PolicyHubUrl" }
)

foreach ($t in $templates) {
    try {
        $existing = Get-PnPListItem -List $listName -Query "<View><Query><Where><Eq><FieldRef Name='Event'/><Value Type='Text'>$($t.Event)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
        if ($null -eq $existing -or $existing.Count -eq 0) {
            Add-PnPListItem -List $listName -Values @{
                Title = $t.Title
                Event = $t.Event
                Category = $t.Category
                Subject = $t.Subject
                Body = $t.Body
                Recipients = $t.Recipients
                IsActive = $true
                IsDefault = $true
                MergeTags = $t.MergeTags
            } | Out-Null
            Write-Host "    + $($t.Title)" -ForegroundColor Green
        } else {
            Write-Host "    ~ $($t.Title) (exists)" -ForegroundColor Gray
        }
    } catch {
        Write-Host "    ! $($t.Title): $_" -ForegroundColor Red
    }
}

Write-Host "`n=== Email Templates Complete ===" -ForegroundColor Cyan
