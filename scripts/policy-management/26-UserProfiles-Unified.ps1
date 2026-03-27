# ============================================================================
# Policy Manager — PM_UserProfiles (Unified User Directory)
# Single source of truth for all user data. Synced from Entra ID,
# read by AudienceRuleService, Distribution, Publish, and Admin Centre.
#
# Replaces the need for separate PM_Employees list.
# Assumes: Already connected via Connect-PnPOnline
# ============================================================================

$ErrorActionPreference = "Continue"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  PM_UserProfiles — Unified User Directory" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

$listName = "PM_UserProfiles"
$list = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $list) {
    New-PnPList -Title $listName -Template GenericList -EnableVersioning
    Write-Host "  Created list: $listName" -ForegroundColor Green
} else {
    Write-Host "  List already exists: $listName" -ForegroundColor Gray
}

Write-Host "  Adding core identity fields..." -ForegroundColor Yellow

# ── Core Identity (from Entra ID sync) ──
Add-PnPField -List $listName -DisplayName "First Name" -InternalName "FirstName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Last Name" -InternalName "LastName" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Email" -InternalName "Email" -Type Text -Required -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Employee Number" -InternalName "EmployeeNumber" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Entra Object ID" -InternalName "EntraObjectId" -Type Text -ErrorAction SilentlyContinue | Out-Null

Write-Host "  Adding organisational fields..." -ForegroundColor Yellow

# ── Organisational (from Entra ID sync + Admin manual edit) ──
Add-PnPField -List $listName -DisplayName "Job Title" -InternalName "JobTitle" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Department" -InternalName "Department" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Office" -InternalName "Office" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Location" -InternalName "Location" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Company" -InternalName "Company" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Cost Center" -InternalName "CostCenter" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Manager Email" -InternalName "ManagerEmail" -Type Text -ErrorAction SilentlyContinue | Out-Null

Write-Host "  Adding contact fields..." -ForegroundColor Yellow

# ── Contact ──
Add-PnPField -List $listName -DisplayName "Office Phone" -InternalName "OfficePhone" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Mobile Phone" -InternalName "MobilePhone" -Type Text -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Profile Photo" -InternalName "ProfilePhoto" -Type URL -ErrorAction SilentlyContinue | Out-Null

Write-Host "  Adding employment fields..." -ForegroundColor Yellow

# ── Employment Status ──
Add-PnPField -List $listName -DisplayName "Status" -InternalName "EmployeeStatus" -Type Choice -Choices "Active","Inactive","PreHire","OnLeave","Terminated","Retired" -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "EmployeeStatus" -Values @{DefaultValue="Active"} -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "IsActive" -Values @{DefaultValue="1"} -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "Employment Type" -InternalName "EmployeeType" -Type Choice -Choices "Employee","Full-Time","Part-Time","Contractor","Intern","Consultant","Temporary" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Start Date" -InternalName "StartDate" -Type DateTime -ErrorAction SilentlyContinue | Out-Null

Write-Host "  Adding Policy Manager role fields..." -ForegroundColor Yellow

# ── Policy Manager Roles ──
Add-PnPField -List $listName -DisplayName "PM Role" -InternalName "PMRole" -Type Choice -Choices "User","Author","Manager","Admin" -AddToDefaultView -ErrorAction SilentlyContinue | Out-Null
Set-PnPField -List $listName -Identity "PMRole" -Values @{DefaultValue="User"} -ErrorAction SilentlyContinue
Add-PnPField -List $listName -DisplayName "PM Roles" -InternalName "PMRoles" -Type Text -ErrorAction SilentlyContinue | Out-Null
# PMRoles: semicolon-separated multi-role string (e.g., "Author;Manager")
Add-PnPField -List $listName -DisplayName "Managed Departments" -InternalName "ManagedDepartments" -Type Note -ErrorAction SilentlyContinue | Out-Null
# ManagedDepartments: semicolon-separated department names for Manager role scoping

Write-Host "  Adding sync tracking fields..." -ForegroundColor Yellow

# ── Sync Tracking ──
Add-PnPField -List $listName -DisplayName "Last Synced At" -InternalName "LastSyncedAt" -Type DateTime -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Sync Source" -InternalName "SyncSource" -Type Choice -Choices "EntraID","Manual","Import" -ErrorAction SilentlyContinue | Out-Null
Add-PnPField -List $listName -DisplayName "Notes" -InternalName "Notes" -Type Note -ErrorAction SilentlyContinue | Out-Null

Write-Host "  Creating indexes..." -ForegroundColor Yellow

# ── Indexes for performance ──
Set-PnPField -List $listName -Identity "Email" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "Department" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "EntraObjectId" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "IsActive" -Values @{Indexed=$true} -ErrorAction SilentlyContinue
Set-PnPField -List $listName -Identity "PMRole" -Values @{Indexed=$true} -ErrorAction SilentlyContinue

# ── Migrate data from PM_Employees if it exists ──
Write-Host ""
Write-Host "  Checking for PM_Employees data to migrate..." -ForegroundColor Yellow
$empList = Get-PnPList -Identity "PM_Employees" -ErrorAction SilentlyContinue
if ($null -ne $empList) {
    $empItems = Get-PnPListItem -List "PM_Employees" -PageSize 500 -ErrorAction SilentlyContinue
    if ($null -ne $empItems -and $empItems.Count -gt 0) {
        Write-Host "  Found $($empItems.Count) employees to migrate" -ForegroundColor Yellow
        $migrated = 0
        foreach ($emp in $empItems) {
            $email = $emp.FieldValues.Email
            if (-not $email) { continue }

            # Check if already exists in PM_UserProfiles
            $existing = Get-PnPListItem -List $listName -Query "<View><Query><Where><Eq><FieldRef Name='Email'/><Value Type='Text'>$email</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
            if ($null -ne $existing -and $existing.Count -gt 0) {
                Write-Host "    ~ $email (already exists)" -ForegroundColor Gray
                continue
            }

            $values = @{
                Title = $emp.FieldValues.Title
                FirstName = $emp.FieldValues.FirstName
                LastName = $emp.FieldValues.LastName
                Email = $email
                EmployeeNumber = $emp.FieldValues.EmployeeNumber
                JobTitle = $emp.FieldValues.JobTitle
                Department = $emp.FieldValues.Department
                Location = $emp.FieldValues.Location
                OfficePhone = $emp.FieldValues.OfficePhone
                MobilePhone = $emp.FieldValues.MobilePhone
                ManagerEmail = $emp.FieldValues.ManagerEmail
                EmployeeType = $emp.FieldValues.EmploymentType
                CostCenter = $emp.FieldValues.CostCenter
                EntraObjectId = $emp.FieldValues.EntraObjectId
                PMRole = $emp.FieldValues.PMRole
                IsActive = if ($emp.FieldValues.Status -eq "Active") { $true } else { $false }
                EmployeeStatus = $emp.FieldValues.Status
                LastSyncedAt = $emp.FieldValues.LastSyncedAt
                SyncSource = "EntraID"
                Notes = $emp.FieldValues.Notes
            }
            # Remove null values
            $cleanValues = @{}
            foreach ($key in $values.Keys) {
                if ($null -ne $values[$key] -and $values[$key] -ne "") {
                    $cleanValues[$key] = $values[$key]
                }
            }
            if ($cleanValues.Count -gt 1) {
                Add-PnPListItem -List $listName -Values $cleanValues -ErrorAction SilentlyContinue | Out-Null
                $migrated++
                Write-Host "    + $email" -ForegroundColor Green
            }
        }
        Write-Host "  Migrated $migrated employees to PM_UserProfiles" -ForegroundColor Green
    } else {
        Write-Host "  PM_Employees is empty — nothing to migrate" -ForegroundColor Gray
    }
} else {
    Write-Host "  PM_Employees does not exist — no migration needed" -ForegroundColor Gray
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  PM_UserProfiles provisioned successfully" -ForegroundColor Green
Write-Host "  Columns: 25 | Indexes: 5" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
