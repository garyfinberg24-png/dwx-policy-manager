# ============================================================================
# DWx Policy Manager — Manual Link PM_PolicyDocuments to PM_Policies
# ============================================================================
# Hardcoded mapping of document filenames to PolicyNumber values.
# Sets DocumentURL on PM_Policies and PolicyId on PM_PolicyDocuments.
#
# Usage:
#   Connect-PnPOnline -Url "https://mf7m.sharepoint.com/sites/PolicyManager" -Interactive
#   .\manual-link-documents.ps1
# ============================================================================

$ErrorActionPreference = "Stop"

Write-Host "`n=== Manual Document-to-Policy Linking ===" -ForegroundColor Cyan

# ── Define the mapping: FileName → PolicyNumber ──────────────────────────────
# Adjust these mappings as needed before running.
$mappings = @(
    @{ FileName = "Code of Conduct Policy.docx";             PolicyNumber = "POL-HR-001" }   # Employee Code of Conduct
    @{ FileName = "Health and Safety Policy.docx";            PolicyNumber = "POL-HS-001" }   # Workplace Health and Safety
    @{ FileName = "Azure Marketplace Privacy policy.docx";    PolicyNumber = "POL-DP-001" }   # Data Protection and Privacy
    @{ FileName = "Cloud Seccurity Policy.docx";              PolicyNumber = "POL-IT-003" }   # Password and Authentication (closest IT security)
    @{ FileName = "Dress Code Policy.pdf";                    PolicyNumber = "POL-HR-002" }   # Anti-Harassment and Discrimination
    @{ FileName = "Sample Document 1.pdf";                    PolicyNumber = "POL-HR-003" }   # Remote Work Policy
    @{ FileName = "Sample Document 2.pdf";                    PolicyNumber = "POL-HR-004" }   # Leave and Time Off
    @{ FileName = "Sample Document 3.pdf";                    PolicyNumber = "POL-HR-005" }   # Performance Management
    @{ FileName = "Sample 8.pdf";                             PolicyNumber = "POL-IT-004" }   # Data Backup and Recovery
    @{ FileName = "Sample 13.pdf";                            PolicyNumber = "POL-IT-005" }   # BYOD Policy
    @{ FileName = "Sample 14.pdf";                            PolicyNumber = "POL-HS-002" }   # Emergency Evacuation Procedures
    @{ FileName = "Upload Sample Document.pdf";               PolicyNumber = "POL-CO-001" }   # Anti-Bribery and Corruption
)

Write-Host "  Defined $($mappings.Count) mapping(s)`n" -ForegroundColor Green

# ── Load documents ───────────────────────────────────────────────────────────
Write-Host "  Loading documents from PM_PolicyDocuments..." -ForegroundColor Cyan
$documents = Get-PnPListItem -List "PM_PolicyDocuments" -PageSize 500 | ForEach-Object {
    [PSCustomObject]@{
        Id       = $_.Id
        FileName = $_.FieldValues["FileLeafRef"]
        FileRef  = $_.FieldValues["FileRef"]
    }
}
Write-Host "  Found $($documents.Count) document(s)" -ForegroundColor Green

# ── Load policies ────────────────────────────────────────────────────────────
Write-Host "  Loading policies from PM_Policies..." -ForegroundColor Cyan
$policies = Get-PnPListItem -List "PM_Policies" -PageSize 500 | ForEach-Object {
    [PSCustomObject]@{
        Id           = $_.Id
        Title        = $_.FieldValues["Title"]
        PolicyNumber = $_.FieldValues["PolicyNumber"]
    }
}
Write-Host "  Found $($policies.Count) policy record(s)`n" -ForegroundColor Green

# ── Apply mappings ───────────────────────────────────────────────────────────
Write-Host "  Applying mappings..." -ForegroundColor Cyan

$linked  = 0
$failed  = 0

foreach ($map in $mappings) {
    $doc = $documents | Where-Object { $_.FileName -eq $map.FileName } | Select-Object -First 1
    $pol = $policies  | Where-Object { $_.PolicyNumber -eq $map.PolicyNumber } | Select-Object -First 1

    if (-not $doc) {
        Write-Host "    [SKIP] Document not found: $($map.FileName)" -ForegroundColor Yellow
        $failed++
        continue
    }
    if (-not $pol) {
        Write-Host "    [SKIP] Policy not found: $($map.PolicyNumber)" -ForegroundColor Yellow
        $failed++
        continue
    }

    # Update DocumentURL on PM_Policies
    Set-PnPListItem -List "PM_Policies" -Identity $pol.Id -Values @{
        DocumentURL = $doc.FileRef
    } | Out-Null

    # Update PolicyId on PM_PolicyDocuments
    Set-PnPListItem -List "PM_PolicyDocuments" -Identity $doc.Id -Values @{
        PolicyId = $pol.Id
    } | Out-Null

    Write-Host "    [LINKED] $($map.PolicyNumber) — $($pol.Title)" -ForegroundColor Green
    Write-Host "             → $($doc.FileName)" -ForegroundColor Green
    $linked++
}

# ── Summary ──────────────────────────────────────────────────────────────────
Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host "  Manual linking complete!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Results:" -ForegroundColor Yellow
Write-Host "    Linked:   $linked"
Write-Host "    Skipped:  $failed"
Write-Host ""
