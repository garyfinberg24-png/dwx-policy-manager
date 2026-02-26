# Seed-DwxAppRegistry.ps1
# Seeds the DWX_AppRegistry list on the DWx Hub site with app entries
# and ensures DWX_Notifications list exists for cross-app notifications.
#
# Prerequisites: Already connected to SharePoint via Connect-PnPOnline
# Target site: https://mf7m.sharepoint.com/sites/DWxHub

$HubUrl = "https://mf7m.sharepoint.com/sites/DWxHub"

# ============================================================================
# APP REGISTRY ENTRIES
# ============================================================================

$Apps = @(
    @{
        Title       = "Policy Manager"
        AppId       = "PolicyManager"
        SiteUrl     = "https://mf7m.sharepoint.com/sites/PolicyManager"
        AppVersion  = "1.2.2"
        IconName    = "Shield"
        ThemeColor  = "#0d9488"
        Description = "Policy Governance & Compliance"
        IsActive    = $true
    },
    @{
        Title       = "DWx Hub"
        AppId       = "DwxHub"
        SiteUrl     = "https://mf7m.sharepoint.com/sites/DWxHub"
        AppVersion  = "1.0.0"
        IconName    = "ViewAll"
        ThemeColor  = "#1a365d"
        Description = "DWx Suite Central Hub"
        IsActive    = $true
    }
    # Add more apps here as they come online:
    # @{
    #     Title       = "Contract Manager"
    #     AppId       = "ContractManager"
    #     SiteUrl     = "https://mf7m.sharepoint.com/sites/ContractManager"
    #     AppVersion  = "1.0.0"
    #     IconName    = "EntitlementPolicy"
    #     ThemeColor  = "#2563eb"
    #     Description = "Contract Lifecycle Management"
    #     IsActive    = $true
    # },
    # @{
    #     Title       = "Asset Manager"
    #     AppId       = "AssetManager"
    #     SiteUrl     = "https://mf7m.sharepoint.com/sites/AssetManager"
    #     AppVersion  = "1.0.0"
    #     IconName    = "ProductCatalog"
    #     ThemeColor  = "#7c3aed"
    #     Description = "IT Asset Management"
    #     IsActive    = $true
    # }
)

# ============================================================================
# ENSURE DWX_AppRegistry LIST EXISTS
# ============================================================================

$RegistryListName = "DWX_AppRegistry"

$existingList = Get-PnPList -Identity $RegistryListName -ErrorAction SilentlyContinue
if (-not $existingList) {
    Write-Host "Creating list: $RegistryListName" -ForegroundColor Yellow
    New-PnPList -Title $RegistryListName -Template GenericList -EnableVersioning

    # Add columns
    Add-PnPField -List $RegistryListName -DisplayName "AppId" -InternalName "AppId" -Type Text -Required
    Add-PnPField -List $RegistryListName -DisplayName "SiteUrl" -InternalName "SiteUrl" -Type URL
    Add-PnPField -List $RegistryListName -DisplayName "AppVersion" -InternalName "AppVersion" -Type Text
    Add-PnPField -List $RegistryListName -DisplayName "IconName" -InternalName "IconName" -Type Text
    Add-PnPField -List $RegistryListName -DisplayName "ThemeColor" -InternalName "ThemeColor" -Type Text
    Add-PnPField -List $RegistryListName -DisplayName "Description" -InternalName "Description" -Type Note
    Add-PnPField -List $RegistryListName -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean
    Add-PnPField -List $RegistryListName -DisplayName "LastHeartbeat" -InternalName "LastHeartbeat" -Type DateTime

    Write-Host "  Created $RegistryListName with all columns" -ForegroundColor Green
} else {
    Write-Host "List already exists: $RegistryListName" -ForegroundColor Cyan
}

# ============================================================================
# SEED APP ENTRIES (idempotent — skips if AppId already exists)
# ============================================================================

foreach ($app in $Apps) {
    $existing = Get-PnPListItem -List $RegistryListName -Query "<View><Query><Where><Eq><FieldRef Name='AppId'/><Value Type='Text'>$($app.AppId)</Value></Eq></Where></Query></View>"

    if ($existing) {
        Write-Host "  App already registered: $($app.Title) ($($app.AppId)) — updating" -ForegroundColor Cyan
        Set-PnPListItem -List $RegistryListName -Identity $existing.Id -Values @{
            Title         = $app.Title
            SiteUrl       = "$($app.SiteUrl), $($app.Title)"
            AppVersion    = $app.AppVersion
            IconName      = $app.IconName
            ThemeColor    = $app.ThemeColor
            Description   = $app.Description
            IsActive      = $app.IsActive
            LastHeartbeat = (Get-Date).ToUniversalTime().ToString("o")
        }
    } else {
        Write-Host "  Registering app: $($app.Title) ($($app.AppId))" -ForegroundColor Yellow
        Add-PnPListItem -List $RegistryListName -Values @{
            Title         = $app.Title
            AppId         = $app.AppId
            SiteUrl       = "$($app.SiteUrl), $($app.Title)"
            AppVersion    = $app.AppVersion
            IconName      = $app.IconName
            ThemeColor    = $app.ThemeColor
            Description   = $app.Description
            IsActive      = $app.IsActive
            LastHeartbeat = (Get-Date).ToUniversalTime().ToString("o")
        }
    }
}

Write-Host ""
Write-Host "App Registry seeded successfully." -ForegroundColor Green

# ============================================================================
# ENSURE DWX_Notifications LIST EXISTS
# ============================================================================

$NotifListName = "DWX_Notifications"

$existingNotifList = Get-PnPList -Identity $NotifListName -ErrorAction SilentlyContinue
if (-not $existingNotifList) {
    Write-Host ""
    Write-Host "Creating list: $NotifListName" -ForegroundColor Yellow
    New-PnPList -Title $NotifListName -Template GenericList -EnableVersioning

    # Core fields
    Add-PnPField -List $NotifListName -DisplayName "MessageBody" -InternalName "MessageBody" -Type Note
    Add-PnPField -List $NotifListName -DisplayName "NotificationType" -InternalName "NotificationType" -Type Choice -Choices "PolicyPublished","PolicyExpiring","PolicyAssigned","ContractExpiring","ContractCreated","ApprovalRequired","ApprovalCompleted","AcknowledgementDue","AssetAssigned","AssetRequested","ComplianceAlert","RecordLinked","SystemAlert"
    Add-PnPField -List $NotifListName -DisplayName "Priority" -InternalName "Priority" -Type Choice -Choices "Low","Medium","High","Critical"

    # Source fields
    Add-PnPField -List $NotifListName -DisplayName "SourceApp" -InternalName "SourceApp" -Type Text
    Add-PnPField -List $NotifListName -DisplayName "SourceItemId" -InternalName "SourceItemId" -Type Number
    Add-PnPField -List $NotifListName -DisplayName "SourceItemTitle" -InternalName "SourceItemTitle" -Type Text
    Add-PnPField -List $NotifListName -DisplayName "SourceItemUrl" -InternalName "SourceItemUrl" -Type URL

    # Recipient fields
    Add-PnPField -List $NotifListName -DisplayName "RecipientEmail" -InternalName "RecipientEmail" -Type Text

    # Status fields
    Add-PnPField -List $NotifListName -DisplayName "Status" -InternalName "Status" -Type Choice -Choices "Unread","Read","Dismissed","Actioned"
    Add-PnPField -List $NotifListName -DisplayName "IsRead" -InternalName "IsRead" -Type Boolean
    Add-PnPField -List $NotifListName -DisplayName "ReadDate" -InternalName "ReadDate" -Type DateTime

    # Metadata fields
    Add-PnPField -List $NotifListName -DisplayName "Category" -InternalName "Category" -Type Choice -Choices "Approval","Compliance","Alert","Info","Task"
    Add-PnPField -List $NotifListName -DisplayName "ActionUrl" -InternalName "ActionUrl" -Type URL
    Add-PnPField -List $NotifListName -DisplayName "ExpiryDate" -InternalName "ExpiryDate" -Type DateTime

    Write-Host "  Created $NotifListName with all columns" -ForegroundColor Green
} else {
    Write-Host "List already exists: $NotifListName" -ForegroundColor Cyan
}

# ============================================================================
# ENSURE DWX_LinkedRecords LIST EXISTS
# ============================================================================

$LinkedListName = "DWX_LinkedRecords"

$existingLinkedList = Get-PnPList -Identity $LinkedListName -ErrorAction SilentlyContinue
if (-not $existingLinkedList) {
    Write-Host ""
    Write-Host "Creating list: $LinkedListName" -ForegroundColor Yellow
    New-PnPList -Title $LinkedListName -Template GenericList -EnableVersioning

    Add-PnPField -List $LinkedListName -DisplayName "SourceApp" -InternalName "SourceApp" -Type Text
    Add-PnPField -List $LinkedListName -DisplayName "SourceListName" -InternalName "SourceListName" -Type Text
    Add-PnPField -List $LinkedListName -DisplayName "SourceItemId" -InternalName "SourceItemId" -Type Number
    Add-PnPField -List $LinkedListName -DisplayName "SourceItemTitle" -InternalName "SourceItemTitle" -Type Text
    Add-PnPField -List $LinkedListName -DisplayName "SourceSiteUrl" -InternalName "SourceSiteUrl" -Type URL

    Add-PnPField -List $LinkedListName -DisplayName "TargetApp" -InternalName "TargetApp" -Type Text
    Add-PnPField -List $LinkedListName -DisplayName "TargetListName" -InternalName "TargetListName" -Type Text
    Add-PnPField -List $LinkedListName -DisplayName "TargetItemId" -InternalName "TargetItemId" -Type Number
    Add-PnPField -List $LinkedListName -DisplayName "TargetItemTitle" -InternalName "TargetItemTitle" -Type Text
    Add-PnPField -List $LinkedListName -DisplayName "TargetSiteUrl" -InternalName "TargetSiteUrl" -Type URL

    Add-PnPField -List $LinkedListName -DisplayName "LinkType" -InternalName "LinkType" -Type Choice -Choices "References","DependsOn","RelatedTo","Supersedes","Governs","Requires"
    Add-PnPField -List $LinkedListName -DisplayName "Description" -InternalName "Description" -Type Note
    Add-PnPField -List $LinkedListName -DisplayName "CreatedByEmail" -InternalName "CreatedByEmail" -Type Text
    Add-PnPField -List $LinkedListName -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean

    Write-Host "  Created $LinkedListName with all columns" -ForegroundColor Green
} else {
    Write-Host "List already exists: $LinkedListName" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "  DWx Hub provisioning complete!" -ForegroundColor Green
Write-Host "  Lists: DWX_AppRegistry, DWX_Notifications, DWX_LinkedRecords" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
