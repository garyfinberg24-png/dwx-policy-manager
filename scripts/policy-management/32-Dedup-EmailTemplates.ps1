# ============================================================================
# 32-Dedup-EmailTemplates.ps1
# Removes duplicate email templates from PM_EmailTemplates list.
# Keeps the NEWEST item (highest Id) for each unique Title and deletes older dupes.
# ============================================================================
# PREREQUISITE: Already connected to SharePoint via Connect-PnPOnline
# ============================================================================

$listName = "PM_EmailTemplates"

Write-Host "`n══════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Email Template Deduplication — $listName" -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════════════`n" -ForegroundColor Cyan

# Load all templates
$items = Get-PnPListItem -List $listName -PageSize 500 | Select-Object Id, @{N='Title';E={$_.FieldValues.Title}}, @{N='Created';E={$_.FieldValues.Created}}

Write-Host "  Total templates found: $($items.Count)" -ForegroundColor White

# Group by Title
$groups = $items | Group-Object -Property Title

$dupeCount = 0
$deleteCount = 0
$dupeGroups = @()

foreach ($group in $groups) {
    if ($group.Count -gt 1) {
        $dupeCount++
        $sorted = $group.Group | Sort-Object -Property Id -Descending
        $keep = $sorted[0]
        $toDelete = $sorted | Select-Object -Skip 1

        Write-Host "`n  DUPLICATE: '$($group.Name)' — $($group.Count) copies found" -ForegroundColor Yellow
        Write-Host "    Keeping:  Id=$($keep.Id) (newest)" -ForegroundColor Green

        foreach ($dupe in $toDelete) {
            Write-Host "    Deleting: Id=$($dupe.Id)" -ForegroundColor Red
            try {
                Remove-PnPListItem -List $listName -Identity $dupe.Id -Force
                $deleteCount++
            } catch {
                Write-Host "    ✗ Failed to delete Id=$($dupe.Id): $_" -ForegroundColor Red
            }
        }
    }
}

if ($dupeCount -eq 0) {
    Write-Host "`n  ✓ No duplicates found — all templates are unique!" -ForegroundColor Green
} else {
    Write-Host "`n  ────────────────────────────────────────────" -ForegroundColor Cyan
    Write-Host "  $dupeCount duplicate groups found" -ForegroundColor Yellow
    Write-Host "  $deleteCount items deleted" -ForegroundColor Red
    Write-Host "  $(($items.Count - $deleteCount)) templates remaining" -ForegroundColor Green
}

# Show final template list
Write-Host "`n  Final template list:" -ForegroundColor Cyan
$remaining = Get-PnPListItem -List $listName -PageSize 500 | Select-Object Id, @{N='Title';E={$_.FieldValues.Title}}
$remaining | Sort-Object -Property Title | ForEach-Object {
    Write-Host "    [$($_.Id)] $($_.Title)" -ForegroundColor DarkGray
}

Write-Host "`n══════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "  Deduplication complete! $($remaining.Count) templates" -ForegroundColor Green
Write-Host "══════════════════════════════════════════════════`n" -ForegroundColor Green
