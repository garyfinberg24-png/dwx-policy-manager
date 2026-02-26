# ============================================================================
# Policy Manager - SharePoint Pages Provisioning
# Creates all required SitePages for Policy Manager webparts
# Target: https://mf7m.sharepoint.com/sites/PolicyManager
# ============================================================================
#
# PREREQUISITE: You must already be connected to SharePoint via Connect-PnPOnline
#
# USAGE:
#   .\Provision-SharePointPages.ps1
#
# This script is idempotent — it checks for existing pages before creating.
# ============================================================================

$SiteUrl = "https://mf7m.sharepoint.com/sites/PolicyManager"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Policy Manager - SharePoint Pages Provisioning" -ForegroundColor Cyan
Write-Host "  Target: $SiteUrl" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Define all required pages with their webpart associations
$Pages = @(
    @{ Name = "PolicyHub";            Title = "Policy Hub";            Description = "Central policy discovery, browsing, and search" }
    @{ Name = "MyPolicies";           Title = "My Policies";           Description = "Personal policy dashboard for employees" }
    @{ Name = "PolicyAdmin";          Title = "Policy Administration"; Description = "Administrative settings and configuration" }
    @{ Name = "PolicyBuilder";        Title = "Policy Builder";        Description = "Create and edit policies" }
    @{ Name = "PolicyAuthor";         Title = "Policy Author";         Description = "Author dashboard — policies, approvals, delegations" }
    @{ Name = "PolicyDetails";        Title = "Policy Details";        Description = "View policy and acknowledge" }
    @{ Name = "PolicyPacks";          Title = "Policy Packs";          Description = "Manage policy bundles and assignments" }
    @{ Name = "QuizBuilder";          Title = "Quiz Builder";          Description = "Create and manage policy quizzes" }
    @{ Name = "PolicySearch";         Title = "Policy Search";         Description = "Dedicated search center" }
    @{ Name = "PolicyHelp";           Title = "Help Center";           Description = "Help articles, FAQs, shortcuts, and support" }
    @{ Name = "PolicyDistribution";   Title = "Policy Distribution";   Description = "Distribution campaign management and tracking" }
    @{ Name = "PolicyAnalytics";      Title = "Policy Analytics";      Description = "Executive analytics dashboard" }
    @{ Name = "PolicyManagerView";    Title = "Policy Manager View";   Description = "Manager compliance dashboard" }
)

$created = 0
$skipped = 0
$failed = 0

foreach ($page in $Pages) {
    $pageName = "$($page.Name).aspx"

    try {
        # Check if page already exists
        $existingPage = $null
        try {
            $existingPage = Get-PnPPage -Identity $page.Name -ErrorAction SilentlyContinue
        } catch {
            # Page doesn't exist — this is expected
        }

        if ($existingPage) {
            Write-Host "  [SKIP] $pageName already exists" -ForegroundColor Yellow
            $skipped++
        } else {
            # Create new blank page
            $newPage = Add-PnPPage -Name $page.Name -Title $page.Title -LayoutType Article -CommentsEnabled:$false

            if ($newPage) {
                Write-Host "  [CREATE] $pageName — $($page.Description)" -ForegroundColor Green
                $created++
            } else {
                Write-Host "  [WARN] $pageName — Add-PnPPage returned null" -ForegroundColor Yellow
                $skipped++
            }
        }
    } catch {
        Write-Host "  [ERROR] $pageName — $($_.Exception.Message)" -ForegroundColor Red
        $failed++
    }
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Summary" -ForegroundColor Cyan
Write-Host "  Created:  $created" -ForegroundColor Green
Write-Host "  Skipped:  $skipped (already existed)" -ForegroundColor Yellow
Write-Host "  Failed:   $failed" -ForegroundColor $(if ($failed -gt 0) { "Red" } else { "Green" })
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "NOTE: After creating pages, you must manually add the corresponding" -ForegroundColor Yellow
Write-Host "      webpart to each page via the SharePoint page editor." -ForegroundColor Yellow
Write-Host ""
Write-Host "  Page -> Webpart mapping:" -ForegroundColor Cyan
Write-Host "    PolicyHub.aspx         -> jmlPolicyHub" -ForegroundColor White
Write-Host "    MyPolicies.aspx        -> jmlMyPolicies" -ForegroundColor White
Write-Host "    PolicyAdmin.aspx       -> jmlPolicyAdmin" -ForegroundColor White
Write-Host "    PolicyBuilder.aspx     -> jmlPolicyAuthor" -ForegroundColor White
Write-Host "    PolicyAuthor.aspx      -> dwxPolicyAuthorView" -ForegroundColor White
Write-Host "    PolicyDetails.aspx     -> jmlPolicyDetails" -ForegroundColor White
Write-Host "    PolicyPacks.aspx       -> jmlPolicyPackManager" -ForegroundColor White
Write-Host "    QuizBuilder.aspx       -> dwxQuizBuilder" -ForegroundColor White
Write-Host "    PolicySearch.aspx      -> jmlPolicySearch" -ForegroundColor White
Write-Host "    PolicyHelp.aspx        -> jmlPolicyHelp" -ForegroundColor White
Write-Host "    PolicyDistribution.aspx-> jmlPolicyDistribution" -ForegroundColor White
Write-Host "    PolicyAnalytics.aspx   -> jmlPolicyAnalytics" -ForegroundColor White
Write-Host "    PolicyManagerView.aspx -> dwxPolicyManagerView" -ForegroundColor White
