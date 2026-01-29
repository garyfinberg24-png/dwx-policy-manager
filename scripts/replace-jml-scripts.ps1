# Script to replace JML_ prefix with PM_ in all policy management scripts
# Also updates site URL from JML to PolicyManager

$scriptsPath = "c:\Projects\SPFx\PolicyManager\policy-manager\scripts\policy-management"

$files = Get-ChildItem -Path $scriptsPath -Filter "*.ps1" -Recurse

$updatedCount = 0
foreach ($file in $files) {
    $content = Get-Content $file.FullName -Raw
    $modified = $false

    # Replace JML_ with PM_ for list names
    if ($content -match 'JML_') {
        $content = $content -replace 'JML_', 'PM_'
        $modified = $true
    }

    # Update site URLs
    if ($content -match 'sites/JML') {
        $content = $content -replace 'sites/JML', 'sites/PolicyManager'
        $modified = $true
    }

    if ($modified) {
        Set-Content -Path $file.FullName -Value $content -NoNewline
        Write-Host "Updated: $($file.Name)"
        $updatedCount++
    }
}

Write-Host "`nTotal script files updated: $updatedCount"
