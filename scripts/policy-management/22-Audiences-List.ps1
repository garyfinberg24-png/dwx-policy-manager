# ============================================================================
# Script 22: PM_Audiences — Audience Rule Definitions
# Stores reusable audience targeting rules for policy distribution
# ============================================================================

$siteUrl = "https://mf7m.sharepoint.com/sites/PolicyManager"
Write-Host "`n=== Creating PM_Audiences ===" -ForegroundColor Cyan

# ── PM_Audiences List ──
$listName = "PM_Audiences"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

# Title = Audience Name (built-in)
Add-PnPField -List $listName -DisplayName "Description" -InternalName "Description" -Type Note -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Rules" -InternalName "Rules" -Type Note -Required -ErrorAction SilentlyContinue
# Rules JSON format: [{"field":"Department","operator":"equals","value":"Sales"},{"field":"Office","operator":"equals","value":"Cape Town"}]
Add-PnPField -List $listName -DisplayName "Combinator" -InternalName "Combinator" -Type Choice -Choices "AND","OR" -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "Combinator" -Values @{DefaultValue="AND"} -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Category" -InternalName "Category" -Type Choice -Choices "Department","Role","Location","Custom","Compliance","Onboarding" -AddToDefaultView -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "IsActive" -Values @{DefaultValue="1"} -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is System" -InternalName "IsSystem" -Type Boolean -ErrorAction SilentlyContinue
# System audiences can't be deleted (e.g., "All Employees", "All Managers")
Add-PnPField -List $listName -DisplayName "Estimated Count" -InternalName "EstimatedCount" -Type Number -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Last Evaluated" -InternalName "LastEvaluated" -Type DateTime -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Created By" -InternalName "CreatedByUser" -Type Text -ErrorAction SilentlyContinue

Set-PnPField -List $listName -Identity "IsActive" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "Category" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

Write-Host "  $listName configured" -ForegroundColor Green

# ── Ensure PM_UserProfiles has Office and Location fields ──
Write-Host "`n  Ensuring PM_UserProfiles has targeting fields..." -ForegroundColor Yellow
$upList = "PM_UserProfiles"
Add-PnPField -List $upList -DisplayName "Office" -InternalName "Office" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $upList -DisplayName "Location" -InternalName "Location" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $upList -DisplayName "Manager Email" -InternalName "ManagerEmail" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $upList -DisplayName "Company" -InternalName "Company" -Type Text -ErrorAction SilentlyContinue
Add-PnPField -List $upList -DisplayName "Employee Type" -InternalName "EmployeeType" -Type Choice -Choices "Employee","Contractor","Intern","Consultant" -ErrorAction SilentlyContinue
Add-PnPField -List $upList -DisplayName "Start Date" -InternalName "StartDate" -Type DateTime -ErrorAction SilentlyContinue
Write-Host "  PM_UserProfiles targeting fields added" -ForegroundColor Green

# ── Seed System Audiences ──
Write-Host "`n  Seeding system audiences..." -ForegroundColor Yellow

$audiences = @(
    @{ Title="All Employees"; Description="Every active employee in the organisation"; Rules='[{"field":"IsActive","operator":"equals","value":"true"}]'; Combinator="AND"; Category="Department"; IsSystem=$true; IsActive=$true },
    @{ Title="All Managers"; Description="Users with Manager role"; Rules='[{"field":"PMRole","operator":"contains","value":"Manager"}]'; Combinator="AND"; Category="Role"; IsSystem=$true; IsActive=$true },
    @{ Title="All Authors"; Description="Users with Author role"; Rules='[{"field":"PMRole","operator":"contains","value":"Author"}]'; Combinator="AND"; Category="Role"; IsSystem=$true; IsActive=$true },
    @{ Title="IT Department"; Description="All users in IT department"; Rules='[{"field":"Department","operator":"equals","value":"IT"}]'; Combinator="AND"; Category="Department"; IsSystem=$false; IsActive=$true },
    @{ Title="HR Department"; Description="All users in HR department"; Rules='[{"field":"Department","operator":"equals","value":"Human Resources"}]'; Combinator="AND"; Category="Department"; IsSystem=$false; IsActive=$true },
    @{ Title="Finance Department"; Description="All users in Finance department"; Rules='[{"field":"Department","operator":"equals","value":"Finance"}]'; Combinator="AND"; Category="Department"; IsSystem=$false; IsActive=$true },
    @{ Title="Compliance Team"; Description="All users in Compliance department"; Rules='[{"field":"Department","operator":"equals","value":"Compliance"}]'; Combinator="AND"; Category="Compliance"; IsSystem=$false; IsActive=$true },
    @{ Title="New Hires (90 days)"; Description="Employees who joined within the last 90 days"; Rules='[{"field":"StartDate","operator":"within_days","value":"90"}]'; Combinator="AND"; Category="Onboarding"; IsSystem=$true; IsActive=$true },
    @{ Title="Contractors"; Description="External contractors and consultants"; Rules='[{"field":"EmployeeType","operator":"equals","value":"Contractor"}]'; Combinator="AND"; Category="Custom"; IsSystem=$false; IsActive=$true }
)

foreach ($a in $audiences) {
    try {
        $existing = Get-PnPListItem -List $listName -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($a.Title)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
        if ($null -eq $existing -or $existing.Count -eq 0) {
            Add-PnPListItem -List $listName -Values @{
                Title = $a.Title
                Description = $a.Description
                Rules = $a.Rules
                Combinator = $a.Combinator
                Category = $a.Category
                IsSystem = $a.IsSystem
                IsActive = $a.IsActive
            } | Out-Null
            Write-Host "    + $($a.Title)" -ForegroundColor Green
        } else {
            Write-Host "    ~ $($a.Title) (exists)" -ForegroundColor Gray
        }
    } catch {
        Write-Host "    ! $($a.Title): $_" -ForegroundColor Red
    }
}

Write-Host "`n=== Audiences Complete ===" -ForegroundColor Cyan
