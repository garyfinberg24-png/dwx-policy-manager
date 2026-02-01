# DWx Decoupling Strategy
## Spinning Off JML Monolith Apps into Standalone DWx Applications

**Version**: 1.0.0
**Date**: 1 February 2026
**Author**: First Digital — Solution Architecture
**Status**: DRAFT — Pending Approval

---

## 1. Executive Summary

The JML_Enterprise monolith currently contains **53+ webparts, 130+ services, and 56+ models** serving the entire Digital Workplace Excellence (DWx) platform. Two apps have already been successfully decoupled — **Policy Manager** and **Contract Manager** — each into standalone SPFx projects with their own Git repos, .sppkg packages, and SharePoint sites.

This document defines a **repeatable, consistent process** for decoupling the remaining **13 apps** from the JML monolith into standalone DWx applications, each in its own project folder under `C:\Projects\SPFx\`.

---

## 2. Applications to Decouple

| # | App Name | Tagline | JML Webpart(s) | Estimated Complexity |
|---|----------|---------|-----------------|---------------------|
| 1 | Asset Manager | IT Asset Tracking & Management | jmlAssetManager | Medium |
| 2 | Signing Service | Digital Signing Service | jmlSigningService | High |
| 3 | CV Management | Candidate Resume Repository | jmlCVManagement | Low-Medium |
| 4 | External Sharing Hub | Secure External Collaboration | jmlExternalSharingHub | Low-Medium |
| 5 | Gamification | Rewards & Recognition Platform | jmlGamification, jmlGamificationAdmin | Medium |
| 6 | License Management | Software License Tracking | jmlLicenseManagement | Low-Medium |
| 7 | Procurement Manager | Purchase Order Workflows | jmlProcurementManager | Medium-High |
| 8 | Recruitment Manager | Talent Acquisition Platform | jmlTalentDashboard | High |
| 9 | Training & Skills | Learning Management System | jmlTrainingSkillsBuilder | Medium-High |
| 10 | Survey Management | Employee Feedback Platform | jmlSurveyManagement, jmlMySurveys | Medium |
| 11 | Reports Builder | Dynamic Report Generation | jmlReportsBuilder | Medium |
| 12 | Document Hub | Enterprise Document Management | jmlDocumentHub, jmlDocumentBuilder, jmlDocumentGeneration | High |
| 13 | Integration Hub | Connect Your Enterprise Systems | jmlIntegrationHub | Medium |

### Complexity Drivers
- **High**: Multiple webparts, 5+ dedicated services, Azure Functions, complex workflows
- **Medium**: 1-2 webparts, 2-4 services, standard CRUD + admin
- **Low**: Single webpart, 1-2 services, minimal cross-dependencies

---

## 3. Lessons Learned from Policy Manager & Contract Manager

Having already decoupled two apps, these are the proven patterns and pitfalls:

### What Worked Well
1. **Separate Git repo per app** — clean history, independent versioning, isolated CI/CD
2. **Own SharePoint site per app** — `https://mf7m.sharepoint.com/sites/{AppName}`
3. **List prefix convention** — `PM_` for Policy Manager prevents collisions
4. **CLAUDE.md per project** — comprehensive context file for AI-assisted development
5. **Copying shared components** (JmlAppLayout, header, styles) rather than abstracting into a shared library — simpler, no cross-project dependency management
6. **Idempotent PowerShell provisioning** — safe to re-run, checks before creating
7. **Role-based nav filtering** — consistent pattern across apps

### Pitfalls to Avoid
1. **Don't rename `jml` prefixes in webpart internal names during initial extraction** — causes manifest ID mismatches; rename in a separate follow-up pass
2. **Don't forget `config/config.json`** — every webpart must be registered in both `bundles` and `localizedResources`
3. **SPFx CDN caching** — version bump + app catalog re-upload + hard refresh required to see updates
4. **`@ts-nocheck` debt** — carry it forward during extraction, clean up later
5. **Service constructor pattern** — services expect `new ServiceName(sp)`, don't change the pattern during extraction

---

## 4. Standard DWx App Architecture

Every decoupled app follows this standard structure:

### 4.1 Project Folder Structure

```
C:\Projects\SPFx\{AppName}\{app-slug}\
├── .claude/                    # Claude Code settings
├── config/
│   ├── config.json             # WebPart bundle registration
│   ├── package-solution.json   # Solution package config (unique GUID)
│   ├── deploy-azure-storage.json
│   ├── serve.json
│   └── write-manifests.json
├── src/
│   ├── webparts/               # App-specific webparts
│   │   ├── dwx{AppName}/       # Main webpart (primary entry point)
│   │   ├── dwx{AppName}Admin/  # Admin/settings webpart (if applicable)
│   │   └── dwx{AppName}{Sub}/  # Additional webparts as needed
│   ├── components/             # Shared layout components (copied from template)
│   │   ├── DwxAppLayout/       # Full-page wrapper with role filtering
│   │   ├── DwxAppHeader/       # Branded header with nav
│   │   ├── DwxAppFooter/       # Footer
│   │   ├── DwxSplashScreen/    # Branded splash screen
│   │   └── PageSubheader/      # Page subheader
│   ├── services/               # App-specific services
│   ├── models/                 # App-specific TypeScript interfaces
│   ├── constants/
│   │   └── SharePointListNames.ts  # {PREFIX}_ list names
│   ├── hooks/                  # Custom React hooks
│   ├── styles/
│   │   └── fluent-mixins.scss  # Shared SCSS mixins (copied from template)
│   └── utils/
│       ├── pnpConfig.ts        # PnP/SP initialization
│       └── injectPortalStyles.ts
├── scripts/                    # PowerShell provisioning scripts
│   ├── 01-Core-Lists.ps1       # Core entity lists
│   ├── 02-Supporting-Lists.ps1 # Supporting lists
│   ├── Deploy-AllLists.ps1     # Master provisioning
│   └── Deploy-SampleData.ps1   # Sample/seed data
├── docs/                       # Documentation + mockups
├── e2e/                        # Playwright e2e tests (if applicable)
├── CLAUDE.md                   # AI development context
├── README.md                   # Project readme
├── package.json
├── tsconfig.json
├── gulpfile.js
└── .gitignore
```

### 4.2 Naming Conventions

| Element | Convention | Example (Asset Manager) |
|---------|-----------|--------------------------|
| Project Folder | PascalCase | `C:\Projects\SPFx\AssetManager\asset-dashboard\` |
| Git Repo | kebab-case | `dwx-asset-manager` |
| Package Name | kebab-case | `dwx-asset-manager` |
| Solution ID | Unique GUID | `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx` |
| SP List Prefix | 2-3 uppercase letters + `_` | `AM_` |
| WebPart Internal | camelCase with dwx prefix | `dwxAssetManager` |
| Component Files | PascalCase | `AssetManager.tsx` |
| Service Files | PascalCase + Service | `AssetService.ts` |
| Model Files | I + PascalCase | `IAsset.ts` |
| SCSS Modules | PascalCase.module.scss | `AssetManager.module.scss` |
| SharePoint Site | PascalCase (no spaces) | `/sites/AssetManager` |

### 4.3 SharePoint List Prefix Registry

Each app gets a unique 2-3 character prefix to prevent list name collisions when apps share a tenant:

| App | Prefix | Example List |
|-----|--------|-------------|
| Policy Manager | `PM_` | PM_Policies |
| Contract Manager | `CM_` | CM_Contracts |
| Asset Manager | `AM_` | AM_Assets |
| CV Management | `CV_` | CV_Resumes |
| Document Hub | `DH_` | DH_Documents |
| External Sharing Hub | `ES_` | ES_GuestAccess |
| Gamification | `GM_` | GM_Points |
| Integration Hub | `IH_` | IH_Integrations |
| License Management | `LM_` | LM_Licenses |
| Procurement Manager | `PR_` | PR_Requisitions |
| Reports Builder | `RB_` | RB_ReportDefinitions |
| Survey Management | `SM_` | SM_Surveys |
| Recruitment Manager | `RM_` | RM_JobRequisitions |
| Signing Service | `SS_` | SS_SigningRequests |
| Training & Skills | `TS_` | TS_Programs |

### 4.4 DWx Brand Color Theme per App

Each app uses the **DWx Brand Identity** (Primary Blue #1a5a8a, Inter font) with an app-specific **accent color** for headers and gradients. The accent color comes from the splash screen designs:

| App | Theme | Primary Accent | Gradient |
|-----|-------|---------------|----------|
| Asset Manager | Slate Blue | #475569 | `135deg, #475569, #334155` |
| CV Management | Teal | #0d9488 | `135deg, #0d9488, #0f766e` |
| Document Hub | Corporate Blue | #1a5a8a | `135deg, #1a5a8a, #0d3a5c` |
| External Sharing Hub | Indigo | #4f46e5 | `135deg, #4f46e5, #3730a3` |
| Gamification | Orange/Amber | #f59e0b | `135deg, #f59e0b, #d97706` |
| Integration Hub | Emerald | #059669 | `135deg, #059669, #047857` |
| License Management | Sky Blue | #0284c7 | `135deg, #0284c7, #0369a1` |
| Procurement Manager | Purple | #7c3aed | `135deg, #7c3aed, #6d28d9` |
| Reports Builder | Rose | #e11d48 | `135deg, #e11d48, #be123c` |
| Survey Management | Cyan | #06b6d4 | `135deg, #06b6d4, #0891b2` |
| Recruitment Manager | Fuchsia | #c026d3 | `135deg, #c026d3, #a21caf` |
| Signing Service | Warm Gray | #78716c | `135deg, #78716c, #57534e` |
| Training & Skills | Lime | #65a30d | `135deg, #65a30d, #4d7c0f` |

---

## 5. The Decoupling Process — Step by Step

This is the **standard operating procedure** for extracting each app. Follow this process for every app in the list.

### Phase 1: Preparation (Pre-Extraction)

#### Step 1.1 — Inventory the App in JML_Enterprise
Before writing any code, catalog everything the app touches in the monolith:

- [ ] List all webpart folders in `JML_Enterprise/src/webparts/` for this app
- [ ] List all services in `JML_Enterprise/src/services/` used by this app
- [ ] List all models in `JML_Enterprise/src/models/` used by this app
- [ ] List all components in `JML_Enterprise/src/components/` used by this app
- [ ] List all SharePoint lists referenced (search constants files)
- [ ] List all provisioning scripts related to this app
- [ ] Identify dashboard widgets in `jmlEmployeeDashboard/components/widgets/`
- [ ] Identify workflow handlers in `workflow/handlers/`
- [ ] Identify any Azure Functions
- [ ] Document cross-dependencies with other apps

**Output**: A checklist of every file to extract, saved to `docs/extraction-inventory.md`.

#### Step 1.2 — Create the Project Scaffold
```bash
# Create project folder
mkdir C:\Projects\SPFx\{AppName}\{app-slug}
cd C:\Projects\SPFx\{AppName}\{app-slug}

# Initialize SPFx project
yo @microsoft/sharepoint
# Select: SharePoint Online only, Y for tenant-wide, WebPart
# Framework: React, Name: dwx{AppName}

# Initialize Git
git init
git remote add origin https://github.com/garyfinberg24-png/dwx-{app-slug}.git
```

#### Step 1.3 — Generate a Unique Solution ID
Every app needs a unique GUID in `config/package-solution.json`:
```json
{
  "solution": {
    "name": "dwx-{app-slug}",
    "id": "{NEW-UNIQUE-GUID}",
    "version": "1.0.0.0"
  }
}
```

### Phase 2: Foundation Setup

#### Step 2.1 — Copy Shared Infrastructure
Copy these files from the **DWx App Template** (Policy Manager as reference):

| Source (Policy Manager) | Destination | Modify? |
|-------------------------|-------------|---------|
| `src/styles/fluent-mixins.scss` | Same path | No — use as-is |
| `src/utils/pnpConfig.ts` | Same path | No — use as-is |
| `src/utils/injectPortalStyles.ts` | Same path | No — use as-is |
| `src/components/JmlAppLayout/` | `src/components/DwxAppLayout/` | Yes — rename, update role enum |
| `src/components/PolicyManagerHeader/` | `src/components/DwxAppHeader/` | Yes — rename, update colors/nav |
| `src/components/PolicyManagerSplashScreen/` | `src/components/DwxSplashScreen/` | Yes — update branding |
| `src/components/PageSubheader/` | Same path | Minimal |

#### Step 2.2 — Configure the App Theme
Update the copied header component with the app's accent color:

```scss
// DwxAppHeader.module.scss
.headerGradient {
  background: linear-gradient(135deg, {APP_PRIMARY} 0%, {APP_DARK} 100%);
}
```

#### Step 2.3 — Create the Constants File
```typescript
// src/constants/SharePointListNames.ts

export const {PREFIX}_LISTS = {
  // Core entity lists
  {ENTITY_PLURAL}: '{PREFIX}_{EntityPlural}',
  // ... all lists for this app
} as const;

// Legacy mapping (if migrating from JML_ lists)
export const LEGACY_LIST_MAPPING: Record<string, string> = {
  'JML_{OldName}': '{PREFIX}_{NewName}',
};
```

#### Step 2.4 — Create the Role Service
Each app gets its own role service following the PolicyRoleService pattern:

```typescript
// src/services/{AppName}RoleService.ts

export enum {AppName}Role {
  User = 0,      // Can view/use the app
  Editor = 1,    // Can create/edit items
  Manager = 2,   // Can approve, manage team
  Admin = 3      // Can configure, full access
}

// Nav item visibility per role
const NAV_KEY_MIN_ROLE: Record<string, {AppName}Role> = {
  'dashboard': {AppName}Role.User,
  'my-items': {AppName}Role.User,
  'create': {AppName}Role.Editor,
  'admin': {AppName}Role.Admin,
};
```

### Phase 3: Service & Model Extraction

#### Step 3.1 — Extract Models
For each model file identified in Step 1.1:

1. Copy the interface file from `JML_Enterprise/src/models/` to the new project
2. Remove any fields/types that reference other apps
3. Replace `JML_` list name references with `{PREFIX}_`
4. Add any shared types (IBaseListItem, IUser, ICommon) to a local `ICommon.ts`

#### Step 3.2 — Extract Services
For each service file:

1. Copy from `JML_Enterprise/src/services/`
2. Update imports to use local models
3. Replace hardcoded list names with constants from `SharePointListNames.ts`
4. Remove dependencies on other apps' services
5. Preserve the `constructor(private sp: SPFI)` pattern
6. Keep `@ts-nocheck` if present — clean up later

#### Step 3.3 — Extract Shared Services
Copy these common services that most apps need:

| Service | Purpose | Always Needed? |
|---------|---------|----------------|
| `LoggingService.ts` | Audit logging | Yes |
| `CacheService.ts` | Data caching | Yes |
| `GraphService.ts` | MS Graph calls | If using Graph |
| `NotificationService.ts` | In-app notifications | If notifications exist |
| `RoleDetectionService.ts` | SP group → role mapping | Yes |
| `SearchService.ts` | Search infrastructure | If search exists |

### Phase 4: WebPart & Component Extraction

#### Step 4.1 — Extract WebPart Shells
For each webpart:

1. Copy the webpart folder from `JML_Enterprise/src/webparts/`
2. Update the manifest.json:
   - Generate a **new unique GUID** for the `id` field
   - Update `alias`, `title`, `description`
   - Set `group` to `"DWx {App Name}"`
   - Update `officeFabricIconFontName`
3. Update the WebPart.ts file:
   - Update class name
   - Ensure `onInit()` calls `getSP(this.context)`
   - Pass `sp` instance to root component
4. Register in `config/config.json` (bundles + localizedResources)

#### Step 4.2 — Extract Components
1. Copy component TSX and SCSS files
2. Update imports to use local services, models, constants
3. Replace JML-branded text/colors with DWx + app-specific branding
4. Update navigation items to match the app's pages

#### Step 4.3 — Create SharePoint Pages
Define the app's page structure:

```
{AppSite}/SitePages/
├── Dashboard.aspx          → Main dashboard webpart
├── Admin.aspx              → Admin/settings webpart (if applicable)
├── {EntityDetail}.aspx     → Detail view webpart
├── Search.aspx             → Search center (if applicable)
└── Help.aspx               → Help center (if applicable)
```

### Phase 5: SharePoint Provisioning

#### Step 5.1 — Create Provisioning Scripts
Follow the numbered script pattern from Policy Manager:

```powershell
# scripts/01-Core-{Entity}Lists.ps1

Import-Module PnP.PowerShell -ErrorAction Stop

$siteUrl = "https://mf7m.sharepoint.com/sites/{AppName}"

# Idempotent list creation
$listName = "{PREFIX}_{EntityPlural}"
$existingList = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($null -eq $existingList) {
    Write-Host "Creating list: $listName" -ForegroundColor Yellow
    New-PnPList -Title $listName -Template GenericList

    # Add fields
    Add-PnPField -List $listName -DisplayName "FieldName" -InternalName "FieldName" -Type Text
    # ... more fields

    Write-Host "✓ Created $listName" -ForegroundColor Green
} else {
    Write-Host "List $listName already exists — skipping" -ForegroundColor Cyan
}
```

#### Step 5.2 — Create Master Deploy Script
```powershell
# scripts/Deploy-AllLists.ps1

Write-Host "=== Deploying {App Name} Lists ===" -ForegroundColor Cyan
& "$PSScriptRoot\01-Core-Lists.ps1"
& "$PSScriptRoot\02-Supporting-Lists.ps1"
Write-Host "=== Deployment Complete ===" -ForegroundColor Green
```

#### Step 5.3 — Create Sample Data Script
```powershell
# scripts/Deploy-SampleData.ps1
# Populate lists with realistic sample data for development/demo
```

### Phase 6: Build, Test & Deploy

#### Step 6.1 — Build Verification
```bash
# Clean build
npm run clean

# Development build (verify no errors)
npm run build

# Ship build (verify all manifests)
gulp clean && gulp bundle --ship
# Confirm: "Found {N} web part manifest(s)"
```

#### Step 6.2 — Create SharePoint Site
```powershell
# Create dedicated site (if not already existing)
New-PnPSite -Type TeamSite -Title "{App Name}" -Alias "{AppName}" -IsPublic
```

#### Step 6.3 — Provision Lists
```powershell
Connect-PnPOnline -Url "https://mf7m.sharepoint.com/sites/{AppName}" -Interactive
& ".\scripts\Deploy-AllLists.ps1"
& ".\scripts\Deploy-SampleData.ps1"
```

#### Step 6.4 — Deploy the Package
```bash
# Package (only with explicit approval)
gulp package-solution --ship

# Upload to tenant app catalog
# sharepoint/solution/dwx-{app-slug}.sppkg
```

#### Step 6.5 — Create Pages & Add WebParts
In the SharePoint site:
1. Create each page defined in Step 4.3
2. Add the corresponding webpart to each page
3. Configure webpart properties
4. Publish each page

### Phase 7: Documentation & Handoff

#### Step 7.1 — Create CLAUDE.md
Every app gets a comprehensive `CLAUDE.md` following the Policy Manager template:

```markdown
# {App Name} - Claude Code Context

## Project Overview
- App Name, Suite (DWx), Company (First Digital)
- Tagline, Version, Package ID, SharePoint Site URL

## Technology Stack
(Standard SPFx stack — copy from PM)

## Architecture Overview
- WebParts (list all)
- SharePoint Pages (page → webpart mapping)
- Directory Structure

## Design System
- App-specific color palette (from Section 4.4 above)
- Typography (Inter, standard DWx)
- Spacing Scale

## SharePoint Lists
- All lists with {PREFIX}_ naming
- Provisioning script mapping

## Key Models
- Core entity interface
- Status lifecycle

## Build Commands
(Standard — copy from PM)

## Development Guidelines
(Standard patterns — copy from PM, update app-specific details)

## Session State
(Initialize empty for first session)
```

#### Step 7.2 — Update README.md
Standard readme with quick start, build commands, architecture overview.

#### Step 7.3 — Git Tag & Release
```bash
git add -A
git commit -m "Initial standalone DWx {App Name} extraction from JML monolith"
git tag v1.0.0
git push origin main --tags
```

---

## 6. Cross-Cutting Concerns

### 6.1 JML Integration Bridge
Some apps still need to communicate with the original JML process engine (Joiner/Mover/Leaver). For these apps:

1. Create a `{App}JMLIntegrationService.ts` that bridges the standalone app to JML lists
2. Use the SharePoint REST API to read/write to JML lists on the JML site
3. This service is **optional** and can be removed once JML is fully deprecated

**Apps that need JML integration**:
- Asset Manager (asset assignment during onboarding)
- Recruitment Manager (new hire pipeline → JML Joiner)
- Training & Skills (onboarding training tasks)

### 6.2 Shared Approval Workflow
Several apps use approval workflows. Rather than duplicating the full workflow engine:

1. Each app copies the core `ApprovalService.ts` pattern
2. Approval lists use the app's own prefix: `{PREFIX}_Approvals`, `{PREFIX}_ApprovalChains`
3. The workflow pattern is consistent but the data is independent per app

**Apps with approval workflows**:
- Signing Service (signing chain workflows — has its own SigningWorkflowEngine)
- Document Hub (document approvals)
- Procurement Manager (PO approvals)
- Recruitment Manager (offer approvals)

### 6.3 Dashboard Widget Pattern
Each app should expose a summary widget that can be embedded in a future "DWx Home" dashboard:

```typescript
// src/components/widgets/{AppName}Widget.tsx
export class {AppName}Widget extends React.Component<IWidgetProps, IWidgetState> {
  // Compact card showing 2-3 KPIs + "View All" link
}
```

### 6.4 Notification Pattern
Apps that need notifications should create their own notification list:
- `{PREFIX}_Notifications` — In-app notifications
- `{PREFIX}_NotificationQueue` — Email queue (if applicable)

### 6.5 Help Center Pattern
Every app should include a Help page with:
- Getting Started articles
- FAQs
- Keyboard shortcuts (if applicable)
- Support contact form

This is a direct copy of the `jmlPolicyHelp` pattern with app-specific content.

---

## 7. Verification Checklist

Use this checklist after extracting each app to ensure completeness:

### Build & Package
- [ ] `npm run build` succeeds with zero errors
- [ ] `gulp bundle --ship` succeeds with correct manifest count
- [ ] All webparts appear in SharePoint toolbox after deployment
- [ ] Splash screen displays correct branding (logo, colors, app name)

### Functionality
- [ ] App layout renders correctly (header, nav, content, footer)
- [ ] Role-based nav filtering works (User vs Editor vs Manager vs Admin)
- [ ] All nav items navigate to correct pages
- [ ] Search functionality works (if applicable)
- [ ] CRUD operations work against SharePoint lists
- [ ] Notifications display correctly (if applicable)

### Data
- [ ] All SharePoint lists created with correct `{PREFIX}_` names
- [ ] Sample data loads correctly
- [ ] No references to `JML_` prefixed lists remain in code
- [ ] No hardcoded URLs to JML SharePoint site

### Code Quality
- [ ] No imports from JML_Enterprise paths
- [ ] No references to other apps' services or models
- [ ] CLAUDE.md is complete and accurate
- [ ] config/config.json has all webparts registered
- [ ] package-solution.json has unique GUID

### Branding
- [ ] Header uses correct app accent color
- [ ] Splash screen matches DWx Brand Guide
- [ ] App name and tagline are correct throughout
- [ ] Inter font family is configured

---

## 8. App-Specific Extraction Notes

### 8.1 Asset Manager (Extraction #1)

**JML Source**: `jmlAssetManager/`
**Services**: AssetService, AssetTrackingService, JMLAssetIntegrationService
**Models**: IAsset.ts
**Components**: AssetCheckout, AssetForm, AssetRegistry, AssetReports
**Lists**: AM_Assets, AM_AssetAudit, AM_AssetCategories, AM_AssetAssignments
**Widgets**: ITAssetOverviewWidget, MyAssetsWidget
**Notes**: Has JML integration for onboarding asset assignment — extract as optional service

### 8.2 Signing Service (Extraction #2)

**JML Source**: `jmlSigningService/`
**Services**: SigningService, SigningWorkflowEngine, SigningPowerAutomateService, SigningNotificationService
**Models**: ISigning.ts (1415 lines — 80+ interfaces, 14 enums)
**Components**: SigningDashboard, RequestDetailPanel, KPICards, SigningExperience, SignaturePad, CreateSigningRequest, TemplateLibrary, TemplateEditor, SigningAnalytics, AuditLogViewer
**Lists**: SS_SigningRequests, SS_SigningChains, SS_Signers, SS_SigningTemplates, SS_SigningAuditLog, SS_SignatureConfig, SS_SigningDocuments, SS_SigningWebhooks, SS_SigningWebhookLog
**Provisioning**: Deploy-SigningService-Lists.ps1 (existing, needs prefix rename JML_ → SS_)
**Workflow Types**: Sequential, Parallel, Hybrid, FirstSigner, ApprovalThenSign, Custom
**Provider Integrations**: Internal, DocuSign, AdobeSign, SigningHub, HelloSign, PandaDoc
**Notes**: **Complex extraction** — 4 dedicated services, workflow engine with scheduled tasks (reminders, escalations, expirations), Power Automate webhook integration, multi-provider signature capture. The SigningWorkflowEngine handles level advancement, auto-approval, and certificate generation. Extract Power Automate integration as optional.

### 8.3 CV Management (Extraction #3)

**JML Source**: `jmlCVManagement/`
**Services**: Embedded in components (no dedicated service — create one)
**Models**: ICVManagement.ts
**Components**: CVManagement, CVBulkOperationsPanel, CVDetailsPanel, CVUploadPanel, CVScoringCharts
**Lists**: CV_Resumes, CV_Versions, CV_Skills
**Widgets**: CVManagementWidget
**Notes**: Relatively self-contained. Service layer needs to be extracted from components.

### 8.4 External Sharing Hub (Extraction #4)

**JML Source**: `jmlExternalSharingHub/`
**Services**: ExternalSharingService, ExternalSharingAuditService
**Models**: IExternalSharing.ts
**Components**: JmlExternalSharingHub, ExternalSharingDashboard
**Hooks**: useExternalSharing.ts
**Lists**: ES_GuestAccess, ES_SharingAudit, ES_SharingKPIs
**Notes**: Compact extraction. Uses MS Graph heavily for guest user management.

### 8.5 Gamification (Extraction #5)

**JML Source**: `jmlGamification/`, `jmlGamificationAdmin/`
**Services**: GamificationService, GamificationAdminService, GamificationBridgeService
**Components**: JmlGamification, GamificationAdmin
**Lists**: GM_Rules, GM_Points, GM_Leaderboards, GM_Badges, GM_Achievements
**Provisioning**: Create-QuizAndGamificationLists.ps1 (extract gamification portion)
**Notes**: Two webparts (user-facing + admin). Bridge service connects to other apps for point triggers — may need stub/mock.

### 8.6 License Management (Extraction #6)

**JML Source**: `jmlLicenseManagement/`
**Services**: LicenseService, M365LicenseService
**Models**: ILicense.ts
**Components**: LicenseManager, M365LicenseManager, LicenseManagementDashboard
**Hooks**: useLicense.ts
**Lists**: LM_Licenses, LM_Assignments, LM_Compliance
**Widgets**: SoftwareLicensesWidget
**Notes**: M365 license sync requires MS Graph permissions (Directory.Read.All or similar).

### 8.7 Procurement Manager (Extraction #7)

**JML Source**: `jmlProcurementManager/`
**Services**: ProcurementService, PurchaseOrderService, BudgetService
**Models**: IProcurement.ts
**Components**: JmlProcurementManager, ProcurementHeader, ProcurementNav
**Lists**: PR_Requisitions, PR_PurchaseOrders, PR_Vendors, PR_BudgetTracking, PR_Approvals
**Widgets**: ProcurementPipelineWidget
**Notes**: Has approval workflow for purchase orders. Budget tracking may reference other apps' spend data.

### 8.8 Recruitment Manager (Extraction #8)

**JML Source**: `jmlTalentDashboard/`
**Services**: RecruitmentService, CandidateService, OfferService, InterviewService, TalentJMLIntegrationService
**Models**: ITalentManagement.ts
**Components**: TalentDashboard, RecruitmentMetricsWidget, RecruitmentOverviewWidget, RecruiterPipelineProgressBar
**Lists**: RM_JobRequisitions, RM_Candidates, RM_Interviews, RM_Offers, RM_Pipeline
**Provisioning**: Provision-RecruitmentLists.ps1
**Notes**: **Second largest extraction** — 5 services, JML integration for new-hire handoff. Has approval workflow for offers.

### 8.9 Training & Skills (Extraction #9)

**JML Source**: `jmlTrainingSkillsBuilder/`
**Services**: ExternalTrainingService, SkillsCompetenciesService
**Models**: ITraining.ts
**Components**: JmlTrainingSkillsBuilder, TeamTrainingWidget, TrainingProgramsWidget, TrainingVideosView
**Lists**: TS_Programs, TS_Enrollment, TS_Certifications, TS_SkillsFramework, TS_UserSkills
**Notes**: Has JML integration for onboarding training assignments. Video content may reference external LMS.

### 8.10 Survey Management (Extraction #10)

**JML Source**: `jmlSurveyManagement/`, `jmlMySurveys/`
**Extensions**: surveyInstanceForm/ (form customizer)
**Services**: Embedded in components
**Models**: ISurvey.ts
**Components**: SurveyManagement, SurveyManagementForm, SurveyForm, MySurveys, SurveyInstanceForm
**Lists**: SM_Surveys, SM_Questions, SM_Responses, SM_Distribution
**Widgets**: MySurveysWidget
**Notes**: Two webparts + a form customizer extension. Service layer needs to be extracted from components.

### 8.11 Reports Builder (Extraction #11)

**JML Source**: `jmlReportsBuilder/`
**Services**: ReportDefinitionService, ReportNarrativeService, ScheduledReportService
**Models**: IReportBuilder.ts
**Components**: JmlReportsBuilder, ReportBuilderCanvas, UsageReports
**Lists**: RB_Definitions, RB_Schedules, RB_Cache, RB_Executions
**Provisioning**: Create-AnalyticsAndReportingLists.ps1
**Notes**: Report canvas is a drag-and-drop widget builder. May reference data from other apps' lists for cross-app reporting — extract with data-source abstraction.

### 8.12 Document Hub (Extraction #12)

**JML Source**: `jmlDocumentHub/`, `jmlDocumentBuilder/`, `jmlDocumentGeneration/`
**Services**: DocumentHubService, DocumentService, DocumentTemplateService, DocumentWorkflowService, DocumentRegistryService, DocumentApprovalService, BulkDocumentService
**Models**: IDocumentHub.ts, IJmlDocument.ts
**Components**: DocumentBuilderHub, DocumentBuilderWizard, DocumentGenerationHub, DocumentGenerationWizard, DocumentUploader, DocumentVersionHistory, DocumentBrowserView
**Lists**: DH_Documents, DH_Versions, DH_Metadata, DH_Workflows, DH_LegalHolds, DH_Templates
**Widgets**: MyDocumentsWidget
**Notes**: **Largest extraction** — 3 webparts, 7 services, many components. Consider phasing: Hub first, then Builder, then Generation.

### 8.13 Integration Hub (Extraction #13)

**JML Source**: `jmlIntegrationHub/`
**Services**: IntegrationService
**Models**: IIntegration.ts
**Components**: IntegrationHub, DocumentIntegrationsPanel, CalendarIntegrationWidget
**Lists**: IH_Integrations, IH_Configs, IH_SyncLogs
**Notes**: Meta-app that connects other apps. Extract last, after all other apps are standalone.

---

## 9. Confirmed Extraction Order

The following order has been approved. Each app will be fully extracted, verified, and deployed before moving to the next.

| Order | App | Complexity | Rationale |
|-------|-----|-----------|-----------|
| 1 | Asset Manager | Medium | Standalone, validates the decoupling process |
| 2 | Signing Service | High | High-value standalone module with workflow engine |
| 3 | CV Management | Low-Medium | Self-contained, quick extraction |
| 4 | External Sharing Hub | Low-Medium | Compact, Graph-heavy but isolated |
| 5 | Gamification | Medium | Two webparts, bridge service needs stubbing |
| 6 | License Management | Low-Medium | Compact, M365 Graph integration |
| 7 | Procurement Manager | Medium-High | Approval workflows, budget tracking |
| 8 | Recruitment Manager | High | 5 services, JML integration for new-hire handoff |
| 9 | Training & Skills | Medium-High | JML integration for onboarding training |
| 10 | Survey Management | Medium | Two webparts + form customizer extension |
| 11 | Reports Builder | Medium | Cross-app data source abstraction needed |
| 12 | Document Hub | High | Largest extraction — 3 webparts, 7 services |
| 13 | Integration Hub | Medium | Meta-app — extract last after all targets are standalone |

---

## 10. Process Efficiency Tips

### Template Automation
After the first 2-3 extractions, create a scaffold script:

```powershell
# scripts/New-DwxApp.ps1
param(
    [string]$AppName,        # "AssetManager"
    [string]$AppSlug,        # "asset-dashboard"
    [string]$ListPrefix,     # "AD"
    [string]$PrimaryColor,   # "#475569"
    [string]$DarkColor       # "#334155"
)

# 1. Create project folder
# 2. Initialize SPFx project
# 3. Copy shared infrastructure
# 4. Generate constants file
# 5. Generate role service
# 6. Generate CLAUDE.md template
# 7. Initialize Git repo
```

### Batch Provisioning
Create a master provisioning script that provisions all app sites:

```powershell
# scripts/Deploy-AllDwxApps.ps1
$apps = @(
    @{ Name="AssetManager"; Lists=".\AssetManager\scripts\Deploy-AllLists.ps1" },
    @{ Name="CVManagement"; Lists=".\CVManagement\scripts\Deploy-AllLists.ps1" },
    # ...
)

foreach ($app in $apps) {
    Connect-PnPOnline -Url "https://mf7m.sharepoint.com/sites/$($app.Name)" -Interactive
    & $app.Lists
}
```

### Parallel Development
Once the process is proven, multiple apps can be extracted in parallel by different developers/sessions, since each app is fully independent.

---

## 11. Post-Decoupling: JML Monolith Retirement Plan

As apps are extracted, the JML_Enterprise monolith shrinks. Track progress:

| App | Extracted | Verified | JML Code Removed |
|-----|-----------|----------|------------------|
| Policy Manager | ✅ | ✅ | Pending |
| Contract Manager | ✅ | ✅ | Pending |
| 1. Asset Manager | ⬜ | ⬜ | ⬜ |
| 2. Signing Service | ⬜ | ⬜ | ⬜ |
| 3. CV Management | ⬜ | ⬜ | ⬜ |
| 4. External Sharing Hub | ⬜ | ⬜ | ⬜ |
| 5. Gamification | ⬜ | ⬜ | ⬜ |
| 6. License Management | ⬜ | ⬜ | ⬜ |
| 7. Procurement Manager | ⬜ | ⬜ | ⬜ |
| 8. Recruitment Manager | ⬜ | ⬜ | ⬜ |
| 9. Training & Skills | ⬜ | ⬜ | ⬜ |
| 10. Survey Management | ⬜ | ⬜ | ⬜ |
| 11. Reports Builder | ⬜ | ⬜ | ⬜ |
| 12. Document Hub | ⬜ | ⬜ | ⬜ |
| 13. Integration Hub | ⬜ | ⬜ | ⬜ |

Once all 15 apps are extracted and verified:
1. Archive JML_Enterprise repository
2. Remove JML webparts from tenant app catalog
3. Decommission JML SharePoint site (after data migration verification)

---

## Appendix A: DWx Brand Standards Quick Reference

| Element | Value |
|---------|-------|
| Logo | Digital Blocks "DWx" with "FIRST DIGITAL" above |
| Primary Blue | #1a5a8a |
| Dark Blue | #0d3a5c |
| Light Blue | #2d7ab8 |
| Font Family | Inter, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif |
| Hero Size | 42px / 800 weight |
| Title Size | 28px / 700 weight |
| Subtitle Size | 20px / 600 weight |
| Body Size | 16px / 400 weight |
| Caption Size | 12px / 500 weight |

## Appendix B: Git Repository Naming

| App | Repo Name | GitHub URL |
|-----|-----------|-----------|
| Asset Manager | dwx-asset-manager | github.com/garyfinberg24-png/dwx-asset-manager |
| CV Management | dwx-cv-management | github.com/garyfinberg24-png/dwx-cv-management |
| Document Hub | dwx-document-hub | github.com/garyfinberg24-png/dwx-document-hub |
| External Sharing Hub | dwx-external-sharing-hub | github.com/garyfinberg24-png/dwx-external-sharing-hub |
| Gamification | dwx-gamification | github.com/garyfinberg24-png/dwx-gamification |
| Integration Hub | dwx-integration-hub | github.com/garyfinberg24-png/dwx-integration-hub |
| License Management | dwx-license-management | github.com/garyfinberg24-png/dwx-license-management |
| Procurement Manager | dwx-procurement-manager | github.com/garyfinberg24-png/dwx-procurement-manager |
| Reports Builder | dwx-reports-builder | github.com/garyfinberg24-png/dwx-reports-builder |
| Survey Management | dwx-survey-management | github.com/garyfinberg24-png/dwx-survey-management |
| Recruitment Manager | dwx-recruitment-manager | github.com/garyfinberg24-png/dwx-recruitment-manager |
| Signing Service | dwx-signing-service | github.com/garyfinberg24-png/dwx-signing-service |
| Training & Skills | dwx-training-skills | github.com/garyfinberg24-png/dwx-training-skills |
