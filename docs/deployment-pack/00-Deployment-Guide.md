# DWx Policy Manager - Deployment Guide

**Version**: 1.2.5
**Date**: 30 March 2026
**Company**: First Digital
**Suite**: DWx (Digital Workplace Excellence)

---

## Table of Contents

1. [Overview](#1-overview)
2. [Prerequisites](#2-prerequisites)
3. [Azure Infrastructure Setup](#3-azure-infrastructure-setup)
4. [SharePoint Site Setup](#4-sharepoint-site-setup)
5. [SharePoint List Provisioning](#5-sharepoint-list-provisioning)
6. [Configuration](#6-configuration)
7. [Sample Data Seeding](#7-sample-data-seeding)
8. [Entra ID User Sync](#8-entra-id-user-sync)
9. [Testing & Verification](#9-testing--verification)
10. [Troubleshooting](#10-troubleshooting)
11. [Appendices](#11-appendices)

---

## 1. Overview

### What is DWx Policy Manager?

DWx Policy Manager is an enterprise policy lifecycle management solution built on SharePoint Framework (SPFx). It provides end-to-end policy governance from authoring through approval, distribution, acknowledgement, and compliance tracking.

### Key Capabilities

- **Policy Authoring** --- Rich text editor with Word/Excel/PowerPoint/HTML templates, version control, and revision workflows
- **Approval Workflows** --- Multi-level approval chains with delegation and escalation
- **Distribution & Acknowledgement** --- Target audiences by department, role, or security group; track read receipts and acknowledgements
- **Quiz & Assessment** --- 11 question types with AI-powered question generation from policy documents
- **Analytics & Reporting** --- Executive dashboards, SLA compliance, audit trails
- **AI Assistant** --- GPT-4o-powered chat for policy Q&A, authoring help, and app guidance
- **Bulk Import** --- Drag-and-drop upload with AI classification and batch metadata assignment

### Architecture

```
+------------------------------------------------------------------+
|                    SharePoint Online                              |
|  +------------------------------------------------------------+  |
|  |  SPFx WebParts (15)          SharePoint Lists (30+)        |  |
|  |  - PolicyHub                 - PM_Policies                 |  |
|  |  - MyPolicies                - PM_PolicyVersions           |  |
|  |  - PolicyAdmin               - PM_PolicyAcknowledgements   |  |
|  |  - PolicyBuilder             - PM_Configuration            |  |
|  |  - PolicyDetails             - PM_Notifications            |  |
|  |  - PolicySearch              - PM_NotificationQueue        |  |
|  |  - PolicyAnalytics           - PM_Approvals                |  |
|  |  - PolicyDistribution        - PM_PolicyQuizzes            |  |
|  |  - PolicyAuthor              - PM_UserProfiles             |  |
|  |  - PolicyManagerView         - ... (30+ total)             |  |
|  |  - PolicyPacks                                             |  |
|  |  - QuizBuilder                                             |  |
|  |  - PolicyHelp                                              |  |
|  |  - PolicyAuthorReports                                     |  |
|  |  - PolicyBulkUpload                                        |  |
|  +------------------------------------------------------------+  |
+------------------------------------------------------------------+
         |              |              |              |
         v              v              v              v
+----------------+ +----------------+ +----------+ +----------------+
| Azure OpenAI   | | Azure Function | | Logic App| | Azure Function |
| GPT-4o         | | Quiz Generator | | Email    | | Doc Converter  |
| (swedencentral)| | (swedencentral)| | Sender   | | (swedencentral)|
+----------------+ +----------------+ | (au-east)| +----------------+
         |                             +----------+
         v                                    |
+----------------+                   +----------------+
| Key Vault      |                   | Office 365     |
| (swedencentral)|                   | (Email Send)   |
+----------------+                   +----------------+
```

### Technology Stack

| Component | Technology | Version |
|-----------|-----------|---------|
| Framework | SharePoint Framework (SPFx) | 1.20.0 |
| UI Library | React | 17.0.1 |
| Language | TypeScript | 4.7.4 |
| UI Components | Fluent UI v8 | 8.106.4 |
| Data Access | PnP/SP, PnP/Graph | 3.25.0 |
| AI Backend | Azure OpenAI (GPT-4o) | 2024-02-15 |
| Email | Azure Logic App + Office 365 | Consumption |
| Package Size | policy-manager.sppkg | ~9.0 MB |

---

## 2. Prerequisites

Refer to **01-Prerequisites-Checklist.md** for a printable checklist.

### Microsoft 365 Tenant

- Microsoft 365 E3/E5 or equivalent with SharePoint Online
- SharePoint App Catalog (tenant-level recommended)
- SharePoint Admin role for the deploying user
- Global Admin or Application Administrator for Entra ID app registration (if configuring AI features)

### Azure Subscription

Required only for AI and automation features. The app functions without Azure but the following features will be unavailable:

- AI Quiz Question Generation
- AI Chat Assistant
- Automated Email Delivery (Logic App)
- Document Format Conversion
- Bulk Distribution Processing
- Approval Escalation

**Required Azure roles**: Contributor on the subscription or target resource groups.

### Development Tools (for building from source)

- Node.js 18.17.1+, 20.x, or 22.x
- PnP PowerShell module (`PnP.PowerShell` 2.x)
- Azure CLI (`az`) 2.50+
- Git

### Admin Permissions Summary

| Action | Required Role |
|--------|--------------|
| Deploy .sppkg to App Catalog | SharePoint Admin |
| Create site collection | SharePoint Admin |
| Run provisioning scripts | Site Collection Admin |
| Deploy Azure resources | Azure Contributor |
| Authorize Logic App connections | Azure Contributor + SharePoint Admin |
| Register Entra ID app | Application Administrator |

---

## 3. Azure Infrastructure Setup

### 3.1 Shared Resources

These are deployed once and shared across multiple Azure Functions.

#### Azure OpenAI Service

| Property | Value |
|----------|-------|
| Resource Name | `dwx-pm-openai-prod` |
| Resource Group | `dwx-pm-quiz-rg-prod` |
| Region | swedencentral |
| Model | GPT-4o (deployment: `gpt-4o`) |

**Deploy**: Created as part of the Quiz Generator deployment (see 3.2).

#### Key Vault

| Property | Value |
|----------|-------|
| Resource Name | `dwx-pm-kv-*` (auto-generated suffix) |
| Resource Group | `dwx-pm-quiz-rg-prod` |
| Region | swedencentral |

**Deploy**: Created as part of the Quiz Generator deployment. Stores the Azure OpenAI API key. Other functions access it via cross-resource-group RBAC.

### 3.2 AI Quiz Generator

Generates quiz questions from policy document content using GPT-4o.

| Property | Value |
|----------|-------|
| Function App | `dwx-pm-quiz-func-prod` |
| Resource Group | `dwx-pm-quiz-rg-prod` |
| Region | swedencentral |
| Plan | Consumption (Y1) |
| Endpoint | `POST /api/generate-quiz-questions?code=<key>` |

**Deploy**:
```powershell
cd azure-functions/quiz-generator/infra
.\deploy.ps1 -Environment prod -Location swedencentral
```

### 3.3 AI Chat Assistant

Client-side RAG pattern: SPFx searches policies, builds context, sends to Azure Function which proxies to GPT-4o.

| Property | Value |
|----------|-------|
| Function App | `dwx-pm-chat-func-prod` |
| Resource Group | `dwx-pm-chat-rg-prod` |
| Region | swedencentral |
| Plan | Consumption (Y1) |
| Endpoint | `POST /api/policyChatCompletion?code=<key>` |

**Deploy**:
```powershell
cd azure-functions/policy-chat/infra
.\deploy.ps1 -Environment prod -Location swedencentral
```

**Note**: This deployment includes a `kvRbac.bicep` module that grants the Chat Function App access to the shared Key Vault in the Quiz resource group.

### 3.4 Email Sender (Logic App)

Polls SharePoint notification queue and sends emails via Office 365 connector.

| Property | Value |
|----------|-------|
| Logic App | `dwx-pm-email-sender-prod` |
| Resource Group | `dwx-pm-email-rg-prod` |
| Region | australiaeast |
| Trigger | Recurrence (every 5 minutes) |

**Deploy**:
```powershell
cd azure-functions/email-sender/infra
.\deploy.ps1 -Environment prod
# Optional: specify shared mailbox sender
.\deploy.ps1 -SenderEmail "noreply@company.com"
```

**Post-deployment (REQUIRED)**: The Logic App uses two API connections (Office 365 and SharePoint Online) that require manual OAuth authorization in the Azure Portal:
1. Navigate to the Resource Group > API Connection `office365-prod` > Edit API connection > Authorize > Save
2. Navigate to the Resource Group > API Connection `sharepointonline-prod` > Edit API connection > Authorize > Save

### 3.5 Distribution Processor

Server-side timer function that processes bulk distribution queues (survives browser close).

| Property | Value |
|----------|-------|
| Function App | `dwx-pm-dist-func-prod` |
| Resource Group | `dwx-pm-dist-rg-prod` |
| Region | australiaeast |
| Trigger | Timer (every 2 minutes) |

**Deploy**:
```powershell
cd azure-functions/distribution-processor/infra
.\deploy.ps1 -Environment prod -Location australiaeast
```

### 3.6 Document Converter

Converts .docx to styled HTML at publish time using mammoth.js.

| Property | Value |
|----------|-------|
| Function App | `dwx-pm-docconv-func-prod` |
| Resource Group | `dwx-pm-docconv-rg-prod` |
| Region | swedencentral |
| Endpoint | `POST /api/convert-document?code=<key>` |

**Deploy**:
```powershell
cd azure-functions/document-converter/infra
.\deploy.ps1 -Environment prod -Location swedencentral
```

### 3.7 Approval Escalation

Auto-escalates overdue approvals based on configured rules.

| Property | Value |
|----------|-------|
| Logic App | `dwx-pm-approval-escalation-prod` |
| Resource Group | `dwx-pm-escalation-rg-prod` |
| Region | australiaeast |

**Deploy**:
```powershell
cd azure-functions/approval-escalation/infra
.\deploy.ps1 -Environment prod
```

### Estimated Monthly Azure Cost

| Resource | SKU | Est. Monthly Cost (USD) |
|----------|-----|------------------------|
| Azure OpenAI (GPT-4o) | Pay-per-token | $10--40 (usage dependent) |
| Quiz Generator Function | Consumption (Y1) | $0--5 |
| Chat Function | Consumption (Y1) | $0--5 |
| Email Logic App | Consumption | $1--5 |
| Distribution Function | Consumption (Y1) | $0--5 |
| Document Converter Function | Consumption (Y1) | $0--5 |
| Approval Escalation Logic App | Consumption | $1--3 |
| Key Vault | Standard | $0.50 |
| Application Insights (x3) | Pay-per-GB | $2--10 |
| Storage Accounts (x4) | LRS | $1--4 |
| **Total** | | **$15--82/month** |

> Consumption plans are billed per execution. Low-usage tenants will be near the minimum; high-volume deployments (1000+ policies, 5000+ users) will be higher.

---

## 4. SharePoint Site Setup

### 4.1 Create the Site Collection

If the site does not already exist:

```powershell
# Using SharePoint Admin PowerShell
Connect-SPOService -Url "https://mf7m-admin.sharepoint.com"
New-SPOSite -Url "https://mf7m.sharepoint.com/sites/PolicyManager" `
            -Owner "admin@mf7m.onmicrosoft.com" `
            -StorageQuota 5120 `
            -Template "STS#3" `
            -Title "Policy Manager"
```

Or create via SharePoint Admin Center > Active Sites > Create > Communication Site.

### 4.2 Deploy the SPFx Package

1. **Build the package** (if building from source):
   ```bash
   npm install
   gulp clean && gulp bundle --ship && gulp package-solution --ship
   ```
   Output: `sharepoint/solution/policy-manager.sppkg` (~9.0 MB)

2. **Upload to App Catalog**:
   - Navigate to your tenant App Catalog (e.g., `https://mf7m.sharepoint.com/sites/AppCatalog`)
   - Go to Apps for SharePoint
   - Upload `policy-manager.sppkg`
   - When prompted, check **"Make this solution available to all sites in the organization"**
   - Click **Deploy**

3. **Trust API permissions** (if prompted):
   - SharePoint Admin Center > Advanced > API access
   - Approve any pending permission requests from "Policy Manager"

### 4.3 Add the App to the Site

```powershell
Connect-PnPOnline -Url "https://mf7m.sharepoint.com/sites/PolicyManager" -Interactive
Install-PnPApp -Identity "policy-manager-client-side-solution" -Scope Site
```

Or: Site Settings > Add an App > Policy Manager.

### 4.4 Create SharePoint Pages

Run the provisioning script to create all 15 pages:

```powershell
Connect-PnPOnline -Url "https://mf7m.sharepoint.com/sites/PolicyManager" -Interactive
.\scripts\policy-management\Provision-SharePointPages.ps1
```

This creates blank pages. After creation, add the corresponding webpart to each page:

| Page | Webpart to Add |
|------|---------------|
| PolicyHub.aspx | jmlPolicyHub |
| MyPolicies.aspx | jmlMyPolicies |
| PolicyAdmin.aspx | jmlPolicyAdmin |
| PolicyBuilder.aspx | jmlPolicyAuthor |
| PolicyAuthor.aspx | dwxPolicyAuthorView |
| PolicyDetails.aspx | jmlPolicyDetails |
| PolicyPacks.aspx | jmlPolicyPackManager |
| QuizBuilder.aspx | dwxQuizBuilder |
| PolicySearch.aspx | jmlPolicySearch |
| PolicyHelp.aspx | jmlPolicyHelp |
| PolicyDistribution.aspx | jmlPolicyDistribution |
| PolicyAnalytics.aspx | jmlPolicyAnalytics |
| PolicyManagerView.aspx | dwxPolicyManagerView |
| PolicyAuthorReports.aspx | dwxPolicyAuthorReports |
| PolicyBulkUpload.aspx | dwxPolicyBulkUpload |

**To add a webpart to a page**:
1. Navigate to the page (e.g., `/sites/PolicyManager/SitePages/PolicyHub.aspx`)
2. Click Edit (pencil icon)
3. Click **+** in the page section
4. Search for the webpart name (e.g., "jmlPolicyHub")
5. Click the webpart to add it
6. Click **Publish**

---

## 5. SharePoint List Provisioning

### 5.1 Master Provisioning Script

The fastest way to provision all lists is to run the master script:

```powershell
Connect-PnPOnline -Url "https://mf7m.sharepoint.com/sites/PolicyManager" -Interactive
.\scripts\policy-management\Run-AllListProvisioning.ps1
```

Alternatively, use the master deployment script (which includes connection):

```powershell
.\scripts\policy-management\Deploy-AllPolicyLists.ps1
```

All scripts are **idempotent** --- they check for existing lists and fields before creating.

### 5.2 Complete List Inventory

#### Core Policy Lists

| List Name | Purpose | Script |
|-----------|---------|--------|
| PM_Policies | Main policy records (80+ fields) | 01-Core-PolicyLists.ps1 |
| PM_PolicyVersions | Version history snapshots | 01-Core-PolicyLists.ps1 |
| PM_PolicyAcknowledgements | User read/acknowledgement tracking | 01-Core-PolicyLists.ps1 |
| PM_PolicyMetadataProfiles | Fast Track metadata presets | 01-Core-PolicyLists.ps1 |
| PM_PolicyReviewers | Reviewer assignments per policy | 01-Core-PolicyLists.ps1 |
| PM_PolicyCategories | Category definitions with sort order | 01-Core-PolicyLists.ps1 |
| PM_PolicyRequests | User-submitted policy creation requests | 14-SubCategory-And-Folders.ps1 |
| PM_PolicySubCategories | Nested subcategory tree | 14-SubCategory-And-Folders.ps1 |
| PM_PolicySourceDocuments | Document library with per-policy folders | 09-PolicySourceDocuments.ps1 |
| PM_PolicyTemplates | Reusable policy templates library | 10-CorporateTemplates.ps1 |
| PM_PolicyExemptions | Policy exemption records | 03-Exemption-Distribution-Lists.ps1 |
| PM_PolicyDistributions | Distribution campaign tracking | 03-Exemption-Distribution-Lists.ps1 |
| PM_DistributionQueue | Server-side bulk distribution queue | 16-DistributionQueue-List.ps1 |

#### Quiz Lists

| List Name | Purpose | Script |
|-----------|---------|--------|
| PM_PolicyQuizzes | Quiz definitions (settings, passing score) | 02-Quiz-Lists.ps1 |
| PM_PolicyQuizQuestions | Individual questions with type-specific data | 02-Quiz-Lists.ps1 |
| PM_PolicyQuizResults | User attempt results and scores | 02-Quiz-Lists.ps1 |

#### Approval Lists

| List Name | Purpose | Script |
|-----------|---------|--------|
| PM_Approvals | Individual approval records | 08-Approval-Lists.ps1 |
| PM_ApprovalHistory | Approval action audit trail | 08-Approval-Lists.ps1 |
| PM_ApprovalDelegations | Delegation assignments | 08-Approval-Lists.ps1 |

#### Notification Lists

| List Name | Purpose | Script |
|-----------|---------|--------|
| PM_Notifications | In-app notifications (notification bell) | 07-Notification-Lists.ps1 |
| PM_NotificationQueue | Email delivery queue (polled by Logic App) | 07-Notification-Lists.ps1 |
| PM_ReminderSchedule | Automated reminder schedule | 25-ReminderSchedule-List.ps1 |

#### Admin & Configuration Lists

| List Name | Purpose | Script |
|-----------|---------|--------|
| PM_Configuration | Key-value settings store | 11-AdminConfig-Lists.ps1 |
| PM_UserProfiles | User profiles synced from Entra ID | 12-UserManagement-Lists.ps1 / 26-UserProfiles-Unified.ps1 |
| PM_EmailTemplates | Email notification template definitions | 21-EmailTemplates-List.ps1 |

#### Policy Packs

| List Name | Purpose | Script |
|-----------|---------|--------|
| PM_PolicyPacks | Policy bundle definitions | 05-PolicyPack-Lists.ps1 |
| PM_PolicyPackAssignments | Pack assignments to users/groups | 05-PolicyPack-Lists.ps1 |

#### Analytics & Audit

| List Name | Purpose | Script |
|-----------|---------|--------|
| PM_PolicyAuditLog | Compliance audit trail | 06-Analytics-Audit-Lists.ps1 |
| PM_PolicyAnalytics | Usage analytics and metrics | 06-Analytics-Audit-Lists.ps1 |
| PM_PolicyFeedback | User feedback and support requests | 06-Analytics-Audit-Lists.ps1 |

#### Social Lists (V2 --- Planned)

| List Name | Purpose | Script |
|-----------|---------|--------|
| PM_PolicyRatings | 5-star policy ratings | 04-Social-Lists.ps1 / 27-Social-Lists.ps1 |
| PM_PolicyComments | Discussion comments on policies | 04-Social-Lists.ps1 / 27-Social-Lists.ps1 |
| PM_PolicyCommentLikes | Likes on comments | 04-Social-Lists.ps1 / 27-Social-Lists.ps1 |
| PM_PolicyShares | Share tracking | 04-Social-Lists.ps1 / 27-Social-Lists.ps1 |
| PM_PolicyFollowers | Policy followers | 04-Social-Lists.ps1 / 27-Social-Lists.ps1 |

#### Workflow Lists (V2 --- Planned)

| List Name | Purpose | Script |
|-----------|---------|--------|
| PM_WorkflowTemplates | Reusable workflow templates | 29-Workflow-Lists.ps1 |
| PM_WorkflowInstances | Active workflow instances | 29-Workflow-Lists.ps1 |
| PM_ApprovalChains | Approval chain progression | 29-Workflow-Lists.ps1 |
| PM_ApprovalTemplates | Reusable approval templates | 29-Workflow-Lists.ps1 |

#### Retention Lists (V2 --- Planned)

| List Name | Purpose | Script |
|-----------|---------|--------|
| PM_RetentionPolicies | Retention policy definitions | 28-Retention-Lists.ps1 |
| PM_LegalHolds | Legal hold records | 28-Retention-Lists.ps1 |
| PM_SLABreaches | SLA breach audit records | 26-SLABreaches-List.ps1 |

### 5.3 Individual Script Execution Order

If running scripts individually instead of the master script:

```
01-Core-PolicyLists.ps1
02-Quiz-Lists.ps1
03-Exemption-Distribution-Lists.ps1
04-Social-Lists.ps1
05-PolicyPack-Lists.ps1
06-Analytics-Audit-Lists.ps1
07-Notification-Lists.ps1
08-Approval-Lists.ps1
09-PolicySourceDocuments.ps1
10-CorporateTemplates.ps1
11-AdminConfig-Lists.ps1
12-UserManagement-Lists.ps1
13-Visibility-Columns.ps1
14-SubCategory-And-Folders.ps1
15-ManagedDepartments-Column.ps1
16-DistributionQueue-List.ps1
16-TemplateType-Update.ps1
17-ReportingLists.ps1
18-MissingColumns-Patch.ps1
19-PolicyRoleGroups.ps1
20-NotificationChoiceUpdate.ps1
21-EmailTemplates-List.ps1
22-Audiences-List.ps1
23-Seed-FastTrackTemplates.ps1
24-Distribution-Missing-Columns.ps1
25-ReminderSchedule-List.ps1
26-UserProfiles-Unified.ps1
26-SLABreaches-List.ps1
27-Social-Lists.ps1
27-Missing-Lists-Master.ps1
28-Retention-Lists.ps1
29-Workflow-Lists.ps1
```

---

## 6. Configuration

### 6.1 PM_Configuration Key-Value Pairs

After list provisioning, seed the following configuration entries in the `PM_Configuration` list:

| ConfigKey | ConfigValue | Category | Description |
|-----------|-------------|----------|-------------|
| `Integration.AI.Chat.Enabled` | `true` | AI | Enable/disable AI Chat Assistant |
| `Integration.AI.Chat.FunctionUrl` | `https://dwx-pm-chat-func-prod.azurewebsites.net/api/policyChatCompletion?code=<KEY>` | AI | Chat Function endpoint |
| `Integration.AI.Chat.MaxTokens` | `1000` | AI | Max response tokens (500/1000/1500/2000) |
| `Integration.AI.Quiz.FunctionUrl` | `https://dwx-pm-quiz-func-prod.azurewebsites.net/api/generate-quiz-questions?code=<KEY>` | AI | Quiz Generator endpoint |
| `Integration.DocConverter.FunctionUrl` | `https://dwx-pm-docconv-func-prod.azurewebsites.net/api/convert-document?code=<KEY>` | Integration | Document Converter endpoint |
| `Admin.SecureLibraries.Config` | `[]` | Admin | JSON array of secure library configurations |
| `Admin.Branding.CompanyName` | `First Digital` | Branding | Company name shown in headers/emails |
| `Admin.Branding.ProductName` | `Policy Manager` | Branding | Product name |
| `Admin.Upload.DocLimitMB` | `25` | Admin | Document upload size limit (MB) |
| `Admin.Upload.VideoLimitMB` | `100` | Admin | Video upload size limit (MB) |
| `Admin.Quiz.DefaultPassingScore` | `70` | Quiz | Default quiz passing score (%) |
| `Notifications.NewPolicy.Enabled` | `true` | Notifications | Notify on new policy published |
| `Notifications.PolicyUpdate.Enabled` | `true` | Notifications | Notify on policy updates |
| `Notifications.DailyDigest.Enabled` | `false` | Notifications | Enable daily digest emails |
| `Compliance.RequireAcknowledgement` | `true` | Compliance | Require ack for published policies |
| `Compliance.DefaultDeadlineDays` | `14` | Compliance | Default ack deadline (days) |
| `Compliance.ReviewFrequencyMonths` | `12` | Compliance | Default review cycle (months) |
| `Approval.RequireOnNew` | `true` | Approval | Require approval for new policies |
| `Approval.RequireOnUpdate` | `true` | Approval | Require approval on updates |
| `Approval.AllowSelfApproval` | `false` | Approval | Allow authors to approve own policies |

### 6.2 Role Setup

The application uses a 4-tier role hierarchy:

| Role | Target Users | Access Level |
|------|-------------|-------------|
| **User** | All employees | Browse, My Policies, Details, Search, Help |
| **Author** | Policy writers | + Create policies, Packs, Author View, Reports, Bulk Upload |
| **Manager** | Department managers | + Approvals, Delegations, Distribution, Analytics, Manager View |
| **Admin** | System administrators | + Quiz Builder, Admin Centre, all settings |

Roles are assigned via the **PM_UserProfiles** list (`PMRole` column) and managed in Admin Centre > Users & Roles.

### 6.3 Navigation Visibility

Navigation items can be toggled per role in Admin Centre > Navigation. Defaults are stored in `PM_Configuration` and cached in localStorage (`pm_nav_visibility`).

### 6.4 Email Templates

29 email templates are auto-created by the application on first use. They follow the premium design pattern with color-coded headers:

- **Teal**: Informational (new policy, acknowledgement complete)
- **Blue**: Updates (policy revised, version published)
- **Amber**: Warnings (acknowledgement reminder, SLA approaching)
- **Orange**: Urgent (final reminder, escalation)
- **Red**: Overdue/Rejected (overdue notice, approval rejected)
- **Green**: Success (approved, completed)
- **Slate**: Retired (policy retirement notice)

Templates can be customized in Admin Centre > Notifications.

---

## 7. Sample Data Seeding

Sample data scripts are provided for demonstration and testing purposes.

### Available Seed Scripts

| Script | What It Creates |
|--------|----------------|
| `Sample-Data-Policies.ps1` | 10+ sample policies across categories |
| `Sample-Data-Quizzes.ps1` | Sample quizzes with questions |
| `Sample-Data-Packs.ps1` | Policy pack bundles |
| `Sample-Data-Templates.ps1` | Policy templates (Corporate, Regulatory, etc.) |
| `Sample-Data-Social.ps1` | Ratings, comments, followers |
| `Sample-Data-Complete.ps1` | All sample data in one script |
| `Seed-PolicyPacks.ps1` | Policy pack definitions and assignments |
| `Seed-ComprehensiveDemoData.ps1` | Comprehensive demo data for all lists |
| `Seed-ApprovalAndNotificationData.ps1` | Sample approvals and notifications |
| `Seed-FastTrackTemplates.ps1` | Fast Track metadata templates |
| `Seed-CurrentUserData.ps1` | Creates profile for current user |

### Run All Sample Data

```powershell
Connect-PnPOnline -Url "https://mf7m.sharepoint.com/sites/PolicyManager" -Interactive
.\scripts\policy-management\Run-AllSampleData.ps1
```

Or individually:

```powershell
.\scripts\policy-management\Seed-ComprehensiveDemoData.ps1
.\scripts\policy-management\Seed-PolicyPacks.ps1
.\scripts\policy-management\Seed-FastTrackTemplates.ps1
```

---

## 8. Entra ID User Sync

The `PM_UserProfiles` list powers audience targeting, role-based access, and compliance tracking. It must be populated with user data.

### How It Works

1. **Admin Centre > Users & Roles > Sync from Entra ID** button triggers a sync
2. The sync reads users from Microsoft Graph API
3. User records are created/updated in `PM_UserProfiles`

### Fields Synced

| PM_UserProfiles Column | Source |
|-----------------------|--------|
| FirstName | `givenName` |
| LastName | `surname` |
| Email | `mail` / `userPrincipalName` |
| Department | `department` |
| JobTitle | `jobTitle` |
| Location | `officeLocation` |
| ManagerEmail | `manager.mail` |
| IsActive | `accountEnabled` |
| PMRole | Manually assigned (default: `User`) |
| ManagedDepartments | Manually assigned (semicolon-delimited) |

### Initial Population

For first deployment, you can:

1. Use the Admin Centre sync button (requires Graph API permissions)
2. Import from CSV using the template in `02-Client-Data-Templates.md`
3. Run `Seed-CurrentUserData.ps1` to create a profile for the deploying user

### Manual Import via CSV

Prepare a CSV with columns: `FirstName, LastName, Email, Department, JobTitle, Location, PMRole`

```powershell
$users = Import-Csv -Path "users.csv"
foreach ($user in $users) {
    Add-PnPListItem -List "PM_UserProfiles" -Values @{
        Title       = "$($user.FirstName) $($user.LastName)"
        FirstName   = $user.FirstName
        LastName    = $user.LastName
        Email       = $user.Email
        Department  = $user.Department
        JobTitle    = $user.JobTitle
        Location    = $user.Location
        PMRole      = $user.PMRole
        IsActive    = "TRUE"
    }
}
```

---

## 9. Testing & Verification

### 9.1 Automated Verification

Run the verification script after deployment:

```powershell
Connect-PnPOnline -Url "https://mf7m.sharepoint.com/sites/PolicyManager" -Interactive
.\scripts\deployment\07-Verify-Deployment.ps1
```

This checks:
- All required SharePoint lists exist
- PM_Configuration has required keys
- All 15 pages exist
- PM_UserProfiles is populated

### 9.2 Manual Testing Checklist

After running verification, walk through these scenarios:

#### Policy Lifecycle
- [ ] Navigate to PolicyHub.aspx --- confirm page loads with category facets
- [ ] Click "Create Policy" --- confirm wizard opens (requires Author role)
- [ ] Create a Draft policy with basic metadata
- [ ] Submit for Review --- confirm status changes, notification sent
- [ ] Approve the policy --- confirm status changes to Published
- [ ] Navigate to PolicyDetails.aspx?policyId=X --- confirm policy displays
- [ ] Acknowledge the policy --- confirm acknowledgement recorded

#### Search & Discovery
- [ ] Navigate to PolicySearch.aspx --- confirm hero banner and filters load
- [ ] Search for the test policy --- confirm it appears in results
- [ ] Navigate to MyPolicies.aspx --- confirm assigned policies appear

#### Quiz System
- [ ] Navigate to QuizBuilder.aspx --- create a quiz linked to the test policy
- [ ] Add questions (multiple choice, true/false)
- [ ] Take the quiz from PolicyDetails --- confirm scoring works

#### Admin Centre
- [ ] Navigate to PolicyAdmin.aspx --- confirm sidebar loads with all sections
- [ ] Check Users & Roles --- confirm PM_UserProfiles data shows
- [ ] Check Configuration --- confirm PM_Configuration entries display
- [ ] Toggle a navigation item --- confirm it hides/shows in the header

#### Notifications
- [ ] Create an item in PM_NotificationQueue with Status "Pending"
- [ ] Wait 5 minutes (Logic App poll interval)
- [ ] Verify email was received and queue item Status = "Sent"

#### AI Features (if Azure deployed)
- [ ] Open AI Chat panel (speech bubble icon in header)
- [ ] Ask "What policies do we have?" --- confirm GPT-4o response
- [ ] In QuizBuilder, click "AI Generate" --- confirm questions generated

---

## 10. Troubleshooting

### SPFx CDN Caching

**Symptom**: Deployed a new .sppkg but changes are not visible.
**Fix**: Increment the version in `package-solution.json`, redeploy to App Catalog, then hard refresh (`Ctrl+Shift+R`) in the browser. SPFx CDN caches aggressively.

### Logic App Email Not Sending

**Symptom**: PM_NotificationQueue items stay in "Pending" status.
**Fix**:
1. Check Azure Portal > Logic App > Run History for errors
2. Verify API connections are authorized (see Section 3.4 post-deployment steps)
3. Confirm the Logic App is enabled (not disabled)
4. Check that the SP list field name is `QueueStatus` (not `Status`)

### PeoplePicker Not Resolving Users

**Symptom**: User fields show "Unresolved" or fail to save.
**Fix**:
1. Ensure the user exists in the site's User Information List (`_api/web/ensureUser('user@domain.com')`)
2. Verify `webAbsoluteUrl` is correctly passed to PeoplePicker components

### Audience Targeting Shows 0 Members

**Symptom**: Distribution targets show "0 members" when department/role is selected.
**Fix**:
1. Verify `PM_UserProfiles` is populated (run Entra ID sync)
2. Check that user records have `IsActive = TRUE`
3. Confirm department names match exactly (case-sensitive)

### Webpart Not Appearing in Page Editor

**Symptom**: Cannot find the webpart when editing a SharePoint page.
**Fix**:
1. Verify the .sppkg is deployed and trusted in the App Catalog
2. Check that the app is installed on the site
3. Confirm the webpart manifest is registered in `config/config.json`
4. Try clearing browser cache and refreshing

### Azure Function Returns 401/403

**Symptom**: AI features return authentication errors.
**Fix**:
1. Verify the function key is included in the URL (`?code=<KEY>`)
2. Check Key Vault access policies --- the function's managed identity needs `Get` secret permission
3. For CORS issues: verify the SharePoint domain is in the function's CORS allowed origins

### Build Errors After Pull

**Symptom**: `gulp bundle --ship` fails after pulling new code.
**Fix**:
```bash
rm -rf node_modules
npm install
gulp clean
gulp bundle --ship
```

---

## 11. Appendices

### Appendix A: Complete PM_Configuration Keys

| ConfigKey | Default Value | Category |
|-----------|---------------|----------|
| Integration.AI.Chat.Enabled | true | AI |
| Integration.AI.Chat.FunctionUrl | (deployment-specific) | AI |
| Integration.AI.Chat.MaxTokens | 1000 | AI |
| Integration.AI.Quiz.FunctionUrl | (deployment-specific) | AI |
| Integration.DocConverter.FunctionUrl | (deployment-specific) | Integration |
| Admin.SecureLibraries.Config | [] | Admin |
| Admin.Branding.CompanyName | First Digital | Branding |
| Admin.Branding.ProductName | Policy Manager | Branding |
| Admin.Upload.DocLimitMB | 25 | Admin |
| Admin.Upload.VideoLimitMB | 100 | Admin |
| Admin.Quiz.DefaultPassingScore | 70 | Quiz |
| Notifications.NewPolicy.Enabled | true | Notifications |
| Notifications.PolicyUpdate.Enabled | true | Notifications |
| Notifications.DailyDigest.Enabled | false | Notifications |
| Compliance.RequireAcknowledgement | true | Compliance |
| Compliance.DefaultDeadlineDays | 14 | Compliance |
| Compliance.ReviewFrequencyMonths | 12 | Compliance |
| Compliance.ReminderEnabled | true | Compliance |
| Approval.RequireOnNew | true | Approval |
| Approval.RequireOnUpdate | true | Approval |
| Approval.AllowSelfApproval | false | Approval |
| Theme.CustomEnabled | false | Theme |
| Theme.PrimaryColor | #0d9488 | Theme |
| Theme.DarkColor | #0f766e | Theme |

### Appendix B: SharePoint Page to Webpart Mapping

| # | Page | Webpart ID | Webpart Display Name |
|---|------|-----------|---------------------|
| 1 | PolicyHub.aspx | jmlPolicyHub | Policy Hub |
| 2 | MyPolicies.aspx | jmlMyPolicies | My Policies |
| 3 | PolicyAdmin.aspx | jmlPolicyAdmin | Policy Admin |
| 4 | PolicyBuilder.aspx | jmlPolicyAuthor | Policy Author |
| 5 | PolicyAuthor.aspx | dwxPolicyAuthorView | Policy Author View |
| 6 | PolicyDetails.aspx | jmlPolicyDetails | Policy Details |
| 7 | PolicyPacks.aspx | jmlPolicyPackManager | Policy Pack Manager |
| 8 | QuizBuilder.aspx | dwxQuizBuilder | Quiz Builder |
| 9 | PolicySearch.aspx | jmlPolicySearch | Policy Search |
| 10 | PolicyHelp.aspx | jmlPolicyHelp | Policy Help |
| 11 | PolicyDistribution.aspx | jmlPolicyDistribution | Policy Distribution |
| 12 | PolicyAnalytics.aspx | jmlPolicyAnalytics | Policy Analytics |
| 13 | PolicyManagerView.aspx | dwxPolicyManagerView | Policy Manager View |
| 14 | PolicyAuthorReports.aspx | dwxPolicyAuthorReports | Policy Author Reports |
| 15 | PolicyBulkUpload.aspx | dwxPolicyBulkUpload | Policy Bulk Upload |

### Appendix C: Azure Resource Cost Estimate

See Section 3 "Estimated Monthly Azure Cost" table. Summary:

- **Minimum** (low usage, <100 users): ~$15/month
- **Typical** (500 users, moderate AI usage): ~$40/month
- **Maximum** (5000+ users, heavy AI usage): ~$80/month

All compute resources use Consumption (serverless) plans --- you only pay for what you use.

### Appendix D: Role Permission Matrix

| Feature | User | Author | Manager | Admin |
|---------|:----:|:------:|:-------:|:-----:|
| Policy Hub (browse) | Y | Y | Y | Y |
| My Policies | Y | Y | Y | Y |
| Policy Details | Y | Y | Y | Y |
| Policy Search | Y | Y | Y | Y |
| Help Centre | Y | Y | Y | Y |
| Create/Edit Policies | | Y | | Y |
| Author Dashboard | | Y | | Y |
| Author Reports | | Y | | Y |
| Bulk Upload | | Y | | Y |
| Policy Packs | | Y | | Y |
| Approvals & Delegations | | | Y | Y |
| Distribution Campaigns | | | Y | Y |
| Analytics Dashboard | | | Y | Y |
| Manager View | | | Y | Y |
| Quiz Builder | | | | Y |
| Admin Centre | | | | Y |
| Users & Roles | | | | Y |
| System Configuration | | | | Y |

> Note: Manager does NOT inherit Author permissions. Each role has explicitly defined access. A user who needs both authoring and management capabilities should be assigned the Admin role.

---

*Document prepared for DWx Policy Manager v1.2.5 deployment. For support, contact the First Digital DWx team.*
