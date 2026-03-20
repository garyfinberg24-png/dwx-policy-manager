# 16-TemplateType-Update.ps1
# Updates PM_PolicyTemplates.TemplateType choices to support Template Manager types
# Also adds DocumentTemplateURL field if missing
# Idempotent — safe to run multiple times

# Site: https://mf7m.sharepoint.com/sites/PolicyManager
# Assumes you are already connected via Connect-PnPOnline

Write-Host "Updating PM_PolicyTemplates for Template Manager..." -ForegroundColor Cyan

# Update TemplateType choices
$field = Get-PnPField -List "PM_PolicyTemplates" -Identity "TemplateType" -ErrorAction SilentlyContinue
if ($null -ne $field) {
    # Update the choices to include new template types
    $choices = [string[]]@("richtext","word","excel","powerpoint","corporate","regulatory","Standard Policy","Procedure","Guideline","Code of Conduct","Custom")
    Set-PnPField -List "PM_PolicyTemplates" -Identity "TemplateType" -Values @{Choices = $choices}
    Write-Host "  OK  TemplateType choices updated" -ForegroundColor Green
} else {
    Write-Host "  WARN  TemplateType field not found — run Create-PolicyTemplatesLibrary.ps1 first" -ForegroundColor Yellow
}

# Add DocumentTemplateURL field if missing
$docUrlField = Get-PnPField -List "PM_PolicyTemplates" -Identity "DocumentTemplateURL" -ErrorAction SilentlyContinue
if ($null -eq $docUrlField) {
    Add-PnPField -List "PM_PolicyTemplates" -DisplayName "DocumentTemplateURL" -InternalName "DocumentTemplateURL" -Type Note
    Write-Host "  OK  DocumentTemplateURL field added" -ForegroundColor Green
} else {
    Write-Host "  --  DocumentTemplateURL field already exists" -ForegroundColor DarkGray
}

# Add HTMLTemplate field if missing (some sites may not have it)
$htmlField = Get-PnPField -List "PM_PolicyTemplates" -Identity "HTMLTemplate" -ErrorAction SilentlyContinue
if ($null -eq $htmlField) {
    Add-PnPField -List "PM_PolicyTemplates" -DisplayName "HTMLTemplate" -InternalName "HTMLTemplate" -Type Note
    Write-Host "  OK  HTMLTemplate field added" -ForegroundColor Green
} else {
    Write-Host "  --  HTMLTemplate field already exists" -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "Template Manager schema update complete." -ForegroundColor Green
Write-Host "You can now create templates of type: richtext, word, excel, powerpoint, corporate, regulatory" -ForegroundColor Cyan
