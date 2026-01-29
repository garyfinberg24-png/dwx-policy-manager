# Script to replace JML_ prefix with PM_ in all source files (excluding SharePointListNames.ts legacy mapping)
$srcPath = "c:\Projects\SPFx\PolicyManager\policy-manager\src"

# Get all TypeScript files except SharePointListNames.ts and .bak files
$files = Get-ChildItem -Path $srcPath -Filter "*.ts" -Recurse |
         Where-Object { $_.Name -ne 'SharePointListNames.ts' -and $_.Extension -ne '.bak' }

$updatedCount = 0
foreach ($file in $files) {
    $content = Get-Content $file.FullName -Raw
    if ($content -match 'JML_') {
        $newContent = $content -replace 'JML_', 'PM_'
        Set-Content -Path $file.FullName -Value $newContent -NoNewline
        Write-Host "Updated: $($file.FullName.Replace($srcPath, ''))"
        $updatedCount++
    }
}

# Also update TSX files
$tsxFiles = Get-ChildItem -Path $srcPath -Filter "*.tsx" -Recurse
foreach ($file in $tsxFiles) {
    $content = Get-Content $file.FullName -Raw
    if ($content -match 'JML_') {
        $newContent = $content -replace 'JML_', 'PM_'
        Set-Content -Path $file.FullName -Value $newContent -NoNewline
        Write-Host "Updated: $($file.FullName.Replace($srcPath, ''))"
        $updatedCount++
    }
}

Write-Host "`nTotal files updated: $updatedCount"
