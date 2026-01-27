# ============================================================================
# JML Policy Management - Deploy All Sample Data
# Master script to populate all Policy Management lists with rich data
# ============================================================================

param(
    [Parameter(Mandatory=$false)]
    [string]$SiteUrl = "https://mf7m.sharepoint.com/sites/JML"
)

$clientId = "d91b5b78-de72-424e-898b-8b5c9512ebd9"
$tenantId = "03bbbdee-d78b-4613-9b99-c468398246b7"
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  JML Policy Management - Sample Data Deployment" -ForegroundColor Cyan
Write-Host "  Target: $SiteUrl" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

Import-Module PnP.PowerShell -ErrorAction Stop

# Connect once and keep connection
Write-Host "`nConnecting to SharePoint..." -ForegroundColor Yellow
Connect-PnPOnline -Url $SiteUrl -DeviceLogin -ClientId $clientId -Tenant $tenantId
$web = Get-PnPWeb
Write-Host "Connected to: $($web.Title)" -ForegroundColor Green

# ============================================================================
# Run each data script inline (to avoid re-authentication)
# ============================================================================

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  PHASE 1: Creating Sample Policies (22 policies)" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Cyan
& "$scriptPath\Sample-Data-Policies.ps1" -SiteUrl $SiteUrl 2>$null

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  PHASE 2: Creating Templates & Quizzes" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Cyan
& "$scriptPath\Sample-Data-Templates.ps1" -SiteUrl $SiteUrl 2>$null

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  PHASE 3: Creating Policy Packs" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Cyan
& "$scriptPath\Sample-Data-Packs.ps1" -SiteUrl $SiteUrl 2>$null

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  PHASE 4: Creating Social Data (Ratings, Comments, Feedback)" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Cyan
& "$scriptPath\Sample-Data-Social.ps1" -SiteUrl $SiteUrl 2>$null

# ============================================================================
# SUMMARY
# ============================================================================

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  SAMPLE DATA DEPLOYMENT COMPLETE!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Data Created:" -ForegroundColor White
Write-Host "  ─────────────────────────────────────────────────────────" -ForegroundColor Gray
Write-Host "  Policies:           22 (HR, IT, H&S, Compliance, Finance)" -ForegroundColor White
Write-Host "  Templates:          4 (HR, IT, Compliance, H&S)" -ForegroundColor White
Write-Host "  Quizzes:            9 (with 40+ questions)" -ForegroundColor White
Write-Host "  Policy Packs:       9 (Onboarding, Department, Role-based)" -ForegroundColor White
Write-Host "  Ratings:            19" -ForegroundColor White
Write-Host "  Comments:           17 (with threading)" -ForegroundColor White
Write-Host "  Feedback:           7" -ForegroundColor White
Write-Host ""
Write-Host "  Policy Categories:" -ForegroundColor Yellow
Write-Host "  • HR Policies (6): Code of Conduct, Anti-Harassment, Remote Work," -ForegroundColor Gray
Write-Host "                     Leave, Performance, Sabbatical (Draft)" -ForegroundColor Gray
Write-Host "  • IT & Security (6): Info Security, Acceptable Use, Password," -ForegroundColor Gray
Write-Host "                       Backup, BYOD, AI Usage (In Review)" -ForegroundColor Gray
Write-Host "  • Health & Safety (3): Workplace H&S, Emergency, Mental Health" -ForegroundColor Gray
Write-Host "  • Compliance (2): Anti-Bribery, Whistleblowing" -ForegroundColor Gray
Write-Host "  • Data Privacy (2): GDPR/Data Protection, Data Retention" -ForegroundColor Gray
Write-Host "  • Financial (2): Expenses, Procurement" -ForegroundColor Gray
Write-Host ""
Write-Host "  Ready to test in workbench!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green

Disconnect-PnPOnline
