# ============================================================================
# DWx Policy Manager — Link PM_PolicyDocuments to PM_Policies
# ============================================================================
# Reads all files in PM_PolicyDocuments, then for each PM_Policies item
# attempts to find a matching document by PolicyNumber in the filename.
# Sets DocumentURL on PM_Policies and PolicyId on the document.
#
# Usage:
#   Connect-PnPOnline -Url "https://mf7m.sharepoint.com/sites/PolicyManager" -Interactive
#   .\link-documents-to-policies.ps1
#
# Matching logic:
#   1. Exact match: filename contains the PolicyNumber (e.g. POL-HR-001)
#   2. Fuzzy match: filename contains a normalised version of the policy Title
#   If no match is found, the policy is skipped with a warning.
# ============================================================================

$ErrorActionPreference = "Stop"

Write-Host "`n=== Linking PM_PolicyDocuments to PM_Policies ===" -ForegroundColor Cyan

# ── 1. Load all documents from PM_PolicyDocuments ────────────────────────────
Write-Host "`n  Loading documents from PM_PolicyDocuments..." -ForegroundColor Cyan
$documents = Get-PnPListItem -List "PM_PolicyDocuments" -PageSize 500 | ForEach-Object {
    [PSCustomObject]@{
        Id           = $_.Id
        FileName     = $_.FieldValues["FileLeafRef"]
        FileRef      = $_.FieldValues["FileRef"]
        PolicyId     = $_.FieldValues["PolicyId"]
        FileNameLower = ($_.FieldValues["FileLeafRef"]).ToLower()
    }
}

if (-not $documents -or $documents.Count -eq 0) {
    Write-Host "  No documents found in PM_PolicyDocuments. Upload files first." -ForegroundColor Yellow
    return
}

Write-Host "  Found $($documents.Count) document(s):" -ForegroundColor Green
foreach ($doc in $documents) {
    Write-Host "    - $($doc.FileName)" -ForegroundColor DarkGray
}

# ── 2. Load all policies from PM_Policies ────────────────────────────────────
Write-Host "`n  Loading policies from PM_Policies..." -ForegroundColor Cyan
$policies = Get-PnPListItem -List "PM_Policies" -PageSize 500 | ForEach-Object {
    [PSCustomObject]@{
        Id           = $_.Id
        Title        = $_.FieldValues["Title"]
        PolicyNumber = $_.FieldValues["PolicyNumber"]
        DocumentURL  = $_.FieldValues["DocumentURL"]
    }
}

Write-Host "  Found $($policies.Count) policy record(s)" -ForegroundColor Green

# ── 3. Match & Link ─────────────────────────────────────────────────────────
Write-Host "`n  Matching documents to policies..." -ForegroundColor Cyan

$linked   = 0
$skipped  = 0
$already  = 0

foreach ($policy in $policies) {
    $policyNumber = $policy.PolicyNumber
    $policyTitle  = $policy.Title

    # Skip if already linked
    if ($policy.DocumentURL) {
        Write-Host "    [LINKED]  $policyNumber — $($policy.Title) → $($policy.DocumentURL)" -ForegroundColor DarkGray
        $already++
        continue
    }

    # Try matching by PolicyNumber in filename (e.g. "POL-HR-001" in "POL-HR-001-Employee-Handbook.pdf")
    $match = $null
    if ($policyNumber) {
        $policyNumberLower = $policyNumber.ToLower()
        $match = $documents | Where-Object { $_.FileNameLower -like "*$policyNumberLower*" } | Select-Object -First 1
    }

    # Fallback: try matching by policy Title in filename
    if (-not $match -and $policyTitle) {
        # Normalise: "Employee Handbook" → "employee-handbook" or "employeehandbook"
        $titleNorm = $policyTitle.ToLower() -replace '[^a-z0-9]', ''
        $match = $documents | Where-Object {
            $fnNorm = $_.FileNameLower -replace '[^a-z0-9]', ''
            $fnNorm -like "*$titleNorm*"
        } | Select-Object -First 1
    }

    if ($match) {
        $docUrl = $match.FileRef
        # Update DocumentURL on PM_Policies
        Set-PnPListItem -List "PM_Policies" -Identity $policy.Id -Values @{
            DocumentURL = $docUrl
        } | Out-Null

        # Update PolicyId on PM_PolicyDocuments
        Set-PnPListItem -List "PM_PolicyDocuments" -Identity $match.Id -Values @{
            PolicyId = $policy.Id
        } | Out-Null

        Write-Host "    [MATCH]   $policyNumber — $policyTitle" -ForegroundColor Green
        Write-Host "              → $($match.FileName)" -ForegroundColor Green
        $linked++
    } else {
        Write-Host "    [NO MATCH] $policyNumber — $policyTitle" -ForegroundColor Yellow
        $skipped++
    }
}

# ── 4. Summary ───────────────────────────────────────────────────────────────
Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host "  Linking complete!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Results:" -ForegroundColor Yellow
Write-Host "    Newly linked:    $linked"
Write-Host "    Already linked:  $already"
Write-Host "    No match found:  $skipped"
Write-Host ""
if ($skipped -gt 0) {
    Write-Host "  Tip: For unmatched policies, rename the document to include" -ForegroundColor Yellow
    Write-Host "  the PolicyNumber (e.g. POL-HR-001-Employee-Handbook.pdf)" -ForegroundColor Yellow
    Write-Host "  then re-run this script." -ForegroundColor Yellow
    Write-Host ""
}
