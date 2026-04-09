# ============================================================================
# Verify-PublishPipeline.ps1
# Pre-deployment verification for the Draft → Publish → Distribution pipeline
#
# Assumes you are already connected to SharePoint via Connect-PnPOnline.
# Checks every dependency in the pipeline and reports pass/fail
# with remediation steps for each failure.
# ============================================================================

$SiteUrl = "https://mf7m.sharepoint.com/sites/PolicyManager"

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  DWx Policy Manager — Publish Pipeline Verification" -ForegroundColor Cyan
Write-Host "  Site: $SiteUrl" -ForegroundColor Cyan
Write-Host "  Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

$pass = 0
$fail = 0
$warn = 0

function Test-Check {
    param(
        [string]$Name,
        [string]$Category,
        [scriptblock]$Test,
        [string]$Remediation
    )

    Write-Host "  [$Category] $Name ... " -NoNewline
    try {
        $result = & $Test
        if ($result -eq $true) {
            Write-Host "PASS" -ForegroundColor Green
            $script:pass++
            return $true
        } else {
            Write-Host "FAIL" -ForegroundColor Red
            Write-Host "    Remediation: $Remediation" -ForegroundColor Yellow
            $script:fail++
            return $false
        }
    } catch {
        Write-Host "FAIL — $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "    Remediation: $Remediation" -ForegroundColor Yellow
        $script:fail++
        return $false
    }
}

function Test-Warning {
    param(
        [string]$Name,
        [string]$Category,
        [scriptblock]$Test,
        [string]$Advice
    )

    Write-Host "  [$Category] $Name ... " -NoNewline
    try {
        $result = & $Test
        if ($result -eq $true) {
            Write-Host "OK" -ForegroundColor Green
            $script:pass++
        } else {
            Write-Host "WARN" -ForegroundColor Yellow
            Write-Host "    Advice: $Advice" -ForegroundColor Yellow
            $script:warn++
        }
    } catch {
        Write-Host "WARN — $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "    Advice: $Advice" -ForegroundColor Yellow
        $script:warn++
    }
}

# ============================================================================
# SECTION 1: Core Lists (REQUIRED for Draft → Publish)
# ============================================================================

Write-Host "── SECTION 1: Core SharePoint Lists ──" -ForegroundColor White

$coreLists = @(
    @{ Name = "PM_Policies"; Desc = "Core policies list"; Script = "Create-PolicyManagementLists.ps1" },
    @{ Name = "PM_PolicyVersions"; Desc = "Version history"; Script = "Create-PolicyManagementLists.ps1" },
    @{ Name = "PM_PolicyReviewers"; Desc = "Reviewer assignments"; Script = "Create-PolicyTemplatesLibrary.ps1" },
    @{ Name = "PM_PolicyAuditLog"; Desc = "Audit trail"; Script = "Create-PolicyManagementLists.ps1" },
    @{ Name = "PM_PolicyAcknowledgements"; Desc = "User acknowledgements"; Script = "Create-PolicyManagementLists.ps1" },
    @{ Name = "PM_PolicyDistributions"; Desc = "Distribution tracking"; Script = "03-Exemption-Distribution-Lists.ps1" },
    @{ Name = "PM_NotificationQueue"; Desc = "Email delivery queue"; Script = "07-Notification-Lists.ps1" },
    @{ Name = "PM_Notifications"; Desc = "In-app notifications"; Script = "07-Notification-Lists.ps1" },
    @{ Name = "PM_Configuration"; Desc = "Admin configuration"; Script = "11-AdminConfig-Lists.ps1" }
)

foreach ($list in $coreLists) {
    Test-Check -Name "$($list.Name) ($($list.Desc))" -Category "LIST" -Test {
        $l = Get-PnPList -Identity $list.Name -ErrorAction Stop
        return ($null -ne $l)
    } -Remediation "Run: .\scripts\policy-management\$($list.Script)"
}

# ============================================================================
# SECTION 2: Distribution Infrastructure
# ============================================================================

Write-Host ""
Write-Host "── SECTION 2: Distribution Infrastructure ──" -ForegroundColor White

Test-Warning -Name "PM_DistributionQueue (server-side queue)" -Category "DIST" -Test {
    $l = Get-PnPList -Identity "PM_DistributionQueue" -ErrorAction SilentlyContinue
    return ($null -ne $l)
} -Advice "Without this list, distribution falls back to inline processing (slower, browser-dependent). Run provisioning script to create it."

Test-Warning -Name "PM_UserProfiles (target user directory)" -Category "DIST" -Test {
    $l = Get-PnPList -Identity "PM_UserProfiles" -ErrorAction SilentlyContinue
    if ($null -eq $l) { return $false }
    $count = $l.ItemCount
    Write-Host "$count users ... " -NoNewline
    return ($count -gt 0)
} -Advice "PM_UserProfiles has no data. Run EntraUserSyncService or seed manually. Without users, resolveTargetUsers() returns empty array and nobody gets notified."

Test-Warning -Name "PM_PolicyReadReceipts (acknowledgement receipts)" -Category "DIST" -Test {
    $l = Get-PnPList -Identity "PM_PolicyReadReceipts" -ErrorAction SilentlyContinue
    return ($null -ne $l)
} -Advice "Read receipts won't be saved without this list. Non-blocking but reduces compliance audit trail."

# ============================================================================
# SECTION 3: Critical Column Checks
# ============================================================================

Write-Host ""
Write-Host "── SECTION 3: Critical Columns on PM_Policies ──" -ForegroundColor White

$requiredColumns = @(
    "PolicyStatus", "PolicyName", "PolicyNumber", "PolicyCategory",
    "VersionNumber", "IsActive", "PolicyDescription",
    "ComplianceRisk", "ReadTimeframe", "ReadTimeframeDays",
    "RequiresAcknowledgement", "EffectiveDate", "PublishedDate",
    "DistributionScope", "TargetDepartments", "TargetRoles",
    "CreationMethod"
)

$fields = Get-PnPField -List "PM_Policies" -ErrorAction SilentlyContinue
$fieldNames = @()
if ($fields) { $fieldNames = $fields | ForEach-Object { $_.InternalName } }

foreach ($col in $requiredColumns) {
    Test-Check -Name "PM_Policies.$col" -Category "COL" -Test {
        return ($fieldNames -contains $col)
    } -Remediation "Column '$col' missing from PM_Policies. Run Deploy-AllPolicyLists.ps1 or add manually."
}

# ============================================================================
# SECTION 4: Critical Columns on PM_PolicyAcknowledgements
# ============================================================================

Write-Host ""
Write-Host "── SECTION 4: Acknowledgement Columns ──" -ForegroundColor White

$ackColumns = @(
    "PolicyId", "AckUserId", "UserEmail", "AckStatus",
    "AssignedDate", "DueDate", "AcknowledgedDate"
)

$ackFields = Get-PnPField -List "PM_PolicyAcknowledgements" -ErrorAction SilentlyContinue
$ackFieldNames = @()
if ($ackFields) { $ackFieldNames = $ackFields | ForEach-Object { $_.InternalName } }

foreach ($col in $ackColumns) {
    Test-Check -Name "PM_PolicyAcknowledgements.$col" -Category "COL" -Test {
        return ($ackFieldNames -contains $col)
    } -Remediation "Column '$col' missing from PM_PolicyAcknowledgements. Run Create-PolicyManagementLists.ps1."
}

# ============================================================================
# SECTION 5: Notification Queue Columns
# ============================================================================

Write-Host ""
Write-Host "── SECTION 5: Notification Queue Columns ──" -ForegroundColor White

$nqColumns = @("To", "Subject", "Message", "QueueStatus", "Priority", "NotificationType", "Channel")

$nqFields = Get-PnPField -List "PM_NotificationQueue" -ErrorAction SilentlyContinue
$nqFieldNames = @()
if ($nqFields) { $nqFieldNames = $nqFields | ForEach-Object { $_.InternalName } }

foreach ($col in $nqColumns) {
    Test-Check -Name "PM_NotificationQueue.$col" -Category "COL" -Test {
        return ($nqFieldNames -contains $col)
    } -Remediation "Column '$col' missing from PM_NotificationQueue. Run 07-Notification-Lists.ps1."
}

# ============================================================================
# SECTION 6: Notification Queue Health
# ============================================================================

Write-Host ""
Write-Host "── SECTION 6: Queue Health ──" -ForegroundColor White

Test-Warning -Name "Failed items in PM_NotificationQueue" -Category "QUEUE" -Test {
    $items = Get-PnPListItem -List "PM_NotificationQueue" -Query "<View><Query><Where><Eq><FieldRef Name='QueueStatus'/><Value Type='Text'>Failed</Value></Eq></Where></Query><RowLimit>10</RowLimit></View>" -ErrorAction SilentlyContinue
    $count = ($items | Measure-Object).Count
    if ($count -gt 0) {
        Write-Host "$count failed ... " -NoNewline
        return $false
    }
    return $true
} -Advice "There are failed items in PM_NotificationQueue. Check Logic App run history in Azure Portal."

Test-Warning -Name "Stuck Pending items (>1 hour old)" -Category "QUEUE" -Test {
    $cutoff = (Get-Date).AddHours(-1).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    $items = Get-PnPListItem -List "PM_NotificationQueue" -Query "<View><Query><Where><And><Eq><FieldRef Name='QueueStatus'/><Value Type='Text'>Pending</Value></Eq><Lt><FieldRef Name='Created'/><Value Type='DateTime'>$cutoff</Value></Lt></And></Where></Query><RowLimit>10</RowLimit></View>" -ErrorAction SilentlyContinue
    $count = ($items | Measure-Object).Count
    if ($count -gt 0) {
        Write-Host "$count stuck ... " -NoNewline
        return $false
    }
    return $true
} -Advice "Items stuck in Pending for >1 hour. Logic App may not be polling. Check Azure Portal > dwx-pm-email-sender-prod."

# ============================================================================
# SECTION 7: Email Template Check
# ============================================================================

Write-Host ""
Write-Host "── SECTION 7: Configuration Keys ──" -ForegroundColor White

$requiredKeys = @(
    @{ Key = "Admin.EventViewer.Enabled"; Required = $false },
    @{ Key = "Admin.Compliance.RequireAcknowledgement"; Required = $true },
    @{ Key = "Admin.Compliance.DefaultDeadlineDays"; Required = $true }
)

foreach ($keyDef in $requiredKeys) {
    $keyName = $keyDef.Key
    $isRequired = $keyDef.Required

    if ($isRequired) {
        Test-Check -Name "Config: $keyName" -Category "CONFIG" -Test {
            $items = Get-PnPListItem -List "PM_Configuration" -Query "<View><Query><Where><Eq><FieldRef Name='ConfigKey'/><Value Type='Text'>$keyName</Value></Eq></Where></Query><RowLimit>1</RowLimit></View>" -ErrorAction SilentlyContinue
            return (($items | Measure-Object).Count -gt 0)
        } -Remediation "Required config key '$keyName' not found in PM_Configuration. Set it in Admin Centre."
    } else {
        Test-Warning -Name "Config: $keyName" -Category "CONFIG" -Test {
            $items = Get-PnPListItem -List "PM_Configuration" -Query "<View><Query><Where><Eq><FieldRef Name='ConfigKey'/><Value Type='Text'>$keyName</Value></Eq></Where></Query><RowLimit>1</RowLimit></View>" -ErrorAction SilentlyContinue
            return (($items | Measure-Object).Count -gt 0)
        } -Advice "Optional config key '$keyName' not set. The app will use defaults."
    }
}

# ============================================================================
# SECTION 8: End-to-End Smoke Test (optional)
# ============================================================================

Write-Host ""
Write-Host "── SECTION 8: Write Test ──" -ForegroundColor White

Test-Check -Name "Can write to PM_PolicyAuditLog" -Category "WRITE" -Test {
    $item = Add-PnPListItem -List "PM_PolicyAuditLog" -Values @{
        Title = "[Pipeline Verify] Write test — safe to delete"
        AuditAction = "SystemCheck"
        EntityType = "Diagnostic"
        ActionDescription = "Publish pipeline verification script write test"
        PerformedByEmail = "verify-script@system"
    } -ErrorAction Stop
    # Clean up
    Remove-PnPListItem -List "PM_PolicyAuditLog" -Identity $item.Id -Force -ErrorAction SilentlyContinue
    return $true
} -Remediation "Cannot write to PM_PolicyAuditLog. Check list permissions — current user needs Contribute access."

Test-Check -Name "Can write to PM_NotificationQueue" -Category "WRITE" -Test {
    $item = Add-PnPListItem -List "PM_NotificationQueue" -Values @{
        Title = "[Pipeline Verify] Write test — safe to delete"
        QueueStatus = "Test"
        Channel = "Email"
        Priority = "Low"
    } -ErrorAction Stop
    Remove-PnPListItem -List "PM_NotificationQueue" -Identity $item.Id -Force -ErrorAction SilentlyContinue
    return $true
} -Remediation "Cannot write to PM_NotificationQueue. Check list permissions."

Test-Check -Name "Can write to PM_PolicyAcknowledgements" -Category "WRITE" -Test {
    $item = Add-PnPListItem -List "PM_PolicyAcknowledgements" -Values @{
        Title = "[Pipeline Verify] Write test — safe to delete"
        AckStatus = "Test"
        PolicyId = 0
    } -ErrorAction Stop
    Remove-PnPListItem -List "PM_PolicyAcknowledgements" -Identity $item.Id -Force -ErrorAction SilentlyContinue
    return $true
} -Remediation "Cannot write to PM_PolicyAcknowledgements. Check list permissions."

# ============================================================================
# SUMMARY
# ============================================================================

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  RESULTS" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  PASSED:   $pass" -ForegroundColor Green
Write-Host "  FAILED:   $fail" -ForegroundColor $(if ($fail -gt 0) { "Red" } else { "Green" })
Write-Host "  WARNINGS: $warn" -ForegroundColor $(if ($warn -gt 0) { "Yellow" } else { "Green" })
Write-Host ""

if ($fail -eq 0) {
    Write-Host "  ✓ Pipeline infrastructure is READY for Draft → Publish → Distribution" -ForegroundColor Green
} elseif ($fail -le 2) {
    Write-Host "  ⚠ Pipeline has $fail issue(s) that MUST be fixed before publishing" -ForegroundColor Yellow
} else {
    Write-Host "  ✗ Pipeline has $fail critical issues — DO NOT deploy until fixed" -ForegroundColor Red
}

if ($warn -gt 0) {
    Write-Host "  $warn warning(s) — non-blocking but may reduce functionality" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "  Next steps:" -ForegroundColor White
Write-Host "    1. Fix any FAIL items using the remediation instructions above" -ForegroundColor White
Write-Host "    2. Deploy policy-manager.sppkg to the App Catalog" -ForegroundColor White
Write-Host "    3. Hard refresh the browser (Ctrl+Shift+R)" -ForegroundColor White
Write-Host "    4. Test: Create draft → Submit → Approve → Publish" -ForegroundColor White
Write-Host "    5. Verify: PM_PolicyAcknowledgements has new Pending records" -ForegroundColor White
Write-Host "    6. Verify: PM_NotificationQueue has new Pending emails" -ForegroundColor White
Write-Host ""
