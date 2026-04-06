# Patch-PolicyReadReceipts.ps1
# Creates PM_PolicyReadReceipts list for compliance audit trail.
# Idempotent — safe to run multiple times.

$listName = "PM_PolicyReadReceipts"

$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

$columns = @(
    @{ Name = "UserId"; Type = "Number" },
    @{ Name = "UserEmail"; Type = "Text" },
    @{ Name = "UserDisplayName"; Type = "Text" },
    @{ Name = "PolicyId"; Type = "Number" },
    @{ Name = "PolicyNumber"; Type = "Text" },
    @{ Name = "PolicyName"; Type = "Text" },
    @{ Name = "PolicyVersion"; Type = "Text" },
    @{ Name = "ReadStartTime"; Type = "DateTime" },
    @{ Name = "ReadEndTime"; Type = "DateTime" },
    @{ Name = "ReadDurationSeconds"; Type = "Number" },
    @{ Name = "QuizRequired"; Type = "Boolean" },
    @{ Name = "QuizCompleted"; Type = "Boolean" },
    @{ Name = "QuizScore"; Type = "Number" },
    @{ Name = "QuizPassPercentage"; Type = "Number" },
    @{ Name = "QuizPassedDate"; Type = "DateTime" },
    @{ Name = "AcknowledgedDate"; Type = "DateTime" },
    @{ Name = "AcknowledgedTime"; Type = "Text" },
    @{ Name = "IPAddress"; Type = "Text" },
    @{ Name = "UserAgent"; Type = "Note" },
    @{ Name = "DeviceType"; Type = "Text" },
    @{ Name = "BrowserName"; Type = "Text" },
    @{ Name = "DigitalSignature"; Type = "Note" },
    @{ Name = "LegalConfirmationText"; Type = "Note" },
    @{ Name = "Notes"; Type = "Note" },
    @{ Name = "ReceiptNumber"; Type = "Text" }
)

foreach ($col in $columns) {
    $existing = Get-PnPField -List $listName -Identity $col.Name -ErrorAction SilentlyContinue
    if ($null -eq $existing) {
        Add-PnPField -List $listName -DisplayName $col.Name -InternalName $col.Name -Type $col.Type -ErrorAction SilentlyContinue
        Write-Host "  + $($col.Name)" -ForegroundColor Green
    }
}

Set-PnPField -List $listName -Identity "PolicyId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "UserEmail" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "ReceiptNumber" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "PM_PolicyReadReceipts ready." -ForegroundColor Green
