# Script to replace JML_ prefix with PM_ in all service files
$servicesPath = "c:\Projects\SPFx\PolicyManager\policy-manager\src\services"

$files = Get-ChildItem -Path $servicesPath -Filter "*.ts" -Recurse

$updatedCount = 0
foreach ($file in $files) {
    $content = Get-Content $file.FullName -Raw
    if ($content -match 'JML_') {
        $newContent = $content -replace 'JML_', 'PM_'
        Set-Content -Path $file.FullName -Value $newContent -NoNewline
        Write-Host "Updated: $($file.Name)"
        $updatedCount++
    }
}

Write-Host "`nTotal files updated: $updatedCount"
