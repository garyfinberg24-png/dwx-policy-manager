# ============================================================================
# Create PolicyAdmin.aspx Page
# Creates a SharePoint modern page with the DWx Policy Admin SPA webpart
# ============================================================================

param(
    [Parameter(Mandatory=$false)]
    [string]$SiteUrl = "https://mf7m.sharepoint.com/sites/PolicyManager"
)

# Connect to SharePoint
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "DWx Policy Manager - Create PolicyAdmin Page" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

Write-Host "Connecting to SharePoint: $SiteUrl" -ForegroundColor Yellow
Connect-PnPOnline -Url $SiteUrl -Interactive
Write-Host "Connected successfully!`n" -ForegroundColor Green

# Check if page already exists
$pageName = "PolicyAdmin"
$existingPage = Get-PnPPage -Identity $pageName -ErrorAction SilentlyContinue

if ($existingPage) {
    Write-Host "Page '$pageName.aspx' already exists." -ForegroundColor Yellow
    $confirm = Read-Host "Do you want to recreate it? (y/n)"
    if ($confirm -ne 'y') {
        Write-Host "Aborted." -ForegroundColor Gray
        exit
    }
    Remove-PnPPage -Identity $pageName -Force
    Write-Host "Existing page removed." -ForegroundColor Gray
}

# Create the page as a single full-width column layout
Write-Host "Creating page: $pageName.aspx" -ForegroundColor Yellow
$page = Add-PnPPage -Name $pageName -LayoutType Article -HeaderLayoutType NoImage -CommentsEnabled:$false

# Find the SPFx webpart component by its ID from the manifest
$webPartComponentId = "e1c016f7-7b87-4531-8590-e19b8911cde2"

Write-Host "Looking up DWx Policy Admin component ($webPartComponentId)..." -ForegroundColor Yellow

# Use Add-PnPPageWebPart with the Component parameter
# First, get available components and find our webpart
$availableComponents = Get-PnPAvailableClientSideComponents -Page $pageName

$policyAdminComponent = $availableComponents | Where-Object { $_.Id -eq $webPartComponentId }

if (-not $policyAdminComponent) {
    Write-Host "`nERROR: DWx Policy Admin webpart not found!" -ForegroundColor Red
    Write-Host "Component ID: $webPartComponentId" -ForegroundColor Gray
    Write-Host "`nMake sure the .sppkg package has been deployed to the App Catalog" -ForegroundColor Yellow
    Write-Host "and the app has been added to this site." -ForegroundColor Yellow
    Write-Host "`nAvailable DWx components:" -ForegroundColor Yellow
    $availableComponents | Where-Object { $_.Name -like "*Dwx*" -or $_.Name -like "*Policy*" -or $_.Name -like "*Jml*" } | ForEach-Object {
        Write-Host "  - $($_.Name) ($($_.Id))" -ForegroundColor Gray
    }
    exit
}

Write-Host "  Found: $($policyAdminComponent.Name)" -ForegroundColor Green

# Add a full-width section and the webpart
Add-PnPPageSection -Page $pageName -SectionTemplate OneColumnFullWidth -Order 1

Add-PnPPageWebPart -Page $pageName -Component $policyAdminComponent -Section 1 -Column 1

# Set page title and publish
Set-PnPPage -Identity $pageName -Title "Policy Admin" -HeaderLayoutType NoImage
Set-PnPPage -Identity $pageName -Publish

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "PAGE CREATED SUCCESSFULLY!" -ForegroundColor Green
Write-Host "========================================`n" -ForegroundColor Cyan
Write-Host "Page URL: $SiteUrl/SitePages/$pageName.aspx" -ForegroundColor White
Write-Host "`nThe DWx Policy Admin webpart has been added to the page." -ForegroundColor Gray

Write-Host "Done!" -ForegroundColor Green
