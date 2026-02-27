# Policy Manager - Claude Code Context

## Instructions for Claude

1. **Always read CLAUDE.md before you do anything**
2. **Always ask questions if you are unsure of the task or requirement**
3. **Be systematic in your planning, and execution**
4. **After you complete a task, always validate the result**
5. **We are working in https://mf7m.sharepoint.com/sites/PolicyManager**

### Task Execution Rules (MANDATORY)

- **One task at a time, verify before moving on.** When given multiple items, complete the first one, then re-read the changed file and describe exactly what changed to yourself before starting the next item. Do NOT batch all items into one large edit.
- **A successful build does NOT mean the task is done.** After every edit, re-read the relevant section of the file and verify it matches the user's requirements point by point. Check: Does the output match what was asked? Did I skip anything? Is the logic correct?
- **NEVER package or ship without explicit user approval.** Do not run `gulp package-solution --ship` unless the user explicitly says "ship it", "package it", or "deploy it". A build check (`gulp bundle --ship`) is fine, but packaging for production requires permission.
- **Track every sub-item in the todo list.** If the user asks for 3 things, create 3 separate todo items. Mark each as completed only AFTER verifying the implementation against the requirement — not after the build passes.
- **When delegating to sub-agents, include the EXACT requirements verbatim.** Do not paraphrase or summarize the user's request when passing to a Task agent. Copy their words directly so nothing is lost in translation.
- **If a task was done wrong before, flag it.** When re-implementing something that failed previously, add a note in the todo: "REDO — previously incorrect" so it gets extra attention during verification.

### Task Communication Process (MANDATORY)

Follow this loop for every task:

1. **User requests a task** — Read and understand the requirement fully.
2. **You explain it back** — Describe what you will do in your own words so the user can confirm you understood correctly. For visual/UI changes, create an HTML mockup first.
3. **User approves or adjusts** — Wait for the user to say "go for it", "approved", or similar. If they adjust, update your understanding and repeat step 2.
4. **You implement** — Only after explicit approval. Follow the Task Execution Rules above.
5. **Iterate** — Continue this loop until the user is satisfied.

Do NOT skip step 2 and jump straight to implementation. The user must confirm understanding before code is written.

---

## Project Overview

**Policy Manager** is an enterprise Policy Management Solution being spun off from the integrated JML (Joiner, Mover, Leaver) solution to become a completely standalone application within the **DWx (Digital Workplace Excellence)** suite by First Digital.

### Application Identity
- **App Name**: Policy Manager
- **Suite**: DWx (Digital Workplace Excellence)
- **Company**: First Digital
- **Tagline**: Policy Governance & Compliance
- **Current Version**: 1.2.3
- **Package ID**: `12538121-8a6b-4e41-8bc7-17f252d5c36e`
- **SharePoint Site**: https://mf7m.sharepoint.com/sites/PolicyManager

---

## Technology Stack

| Category | Technology | Version |
|----------|------------|---------|
| Framework | SharePoint Framework (SPFx) | 1.20.0 |
| UI Library | React | 17.0.1 |
| Language | TypeScript | 4.7.4 |
| UI Components | Fluent UI v8 | 8.106.4 |
| Data Access | PnP/SP, PnP/Graph | 3.25.0 |
| Build System | Gulp | 4.0.2 |
| Node | Node.js | 18.17.1+, 20.x, or 22.x |

---

## Architecture Overview

### WebParts (14 total)
1. **jmlMyPolicies** - Personal policy dashboard for employees
2. **jmlPolicyHub** - Central policy discovery, browsing, and search
3. **jmlPolicyAdmin** - Administrative interface (sidebar + content layout)
4. **jmlPolicyAuthor** - Policy creation and editing (rich text editor)
5. **jmlPolicyDetails** - Detailed policy view with acknowledgement
6. **jmlPolicyPackManager** - Policy package bundling and assignment
7. **dwxQuizBuilder** - Quiz creation and management
8. **jmlPolicySearch** - Dedicated search center (hero + filters + results)
9. **jmlPolicyHelp** - Help center with articles, FAQs, shortcuts, videos, support
10. **jmlPolicyDistribution** - Distribution campaign management and tracking
11. **jmlPolicyAnalytics** - Executive analytics dashboard (6 tabs: Executive, Policy Metrics, Acknowledgements, SLA, Compliance, Audit)
12. **dwxPolicyAuthorView** - Author dashboard with policies, approvals, delegations, activity tabs
13. **dwxPolicyManagerView** - Manager dashboard for team compliance, approvals, delegations, reviews, reports

### SharePoint Pages
| Page | WebPart | Purpose |
|------|---------|---------|
| PolicyHub.aspx | jmlPolicyHub | Main dashboard, browsing, recently viewed |
| MyPolicies.aspx | jmlMyPolicies | User's assigned policies |
| PolicyAdmin.aspx | jmlPolicyAdmin | Admin settings and configuration |
| PolicyBuilder.aspx | jmlPolicyAuthor | Create/edit policies |
| PolicyAuthor.aspx | dwxPolicyAuthorView | Author dashboard (policies, approvals, delegations) |
| PolicyDetails.aspx | jmlPolicyDetails | View policy + acknowledge |
| PolicyPacks.aspx | jmlPolicyPackManager | Manage policy packs |
| QuizBuilder.aspx | dwxQuizBuilder | Create quizzes |
| PolicySearch.aspx | jmlPolicySearch | Search center |
| PolicyHelp.aspx | jmlPolicyHelp | Help center |
| PolicyDistribution.aspx | jmlPolicyDistribution | Distribution campaigns |
| PolicyAnalytics.aspx | jmlPolicyAnalytics | Executive analytics dashboard |
| PolicyManagerView.aspx | dwxPolicyManagerView | Manager compliance dashboard |

### Directory Structure
```
policy-manager/
├── src/
│   ├── webparts/          # 14 SPFx webparts
│   │   ├── jmlMyPolicies/
│   │   ├── jmlPolicyHub/
│   │   ├── jmlPolicyAdmin/
│   │   ├── jmlPolicyAuthor/
│   │   ├── jmlPolicyDetails/
│   │   ├── jmlPolicyPackManager/
│   │   ├── dwxQuizBuilder/
│   │   ├── jmlPolicySearch/
│   │   ├── jmlPolicyHelp/
│   │   ├── jmlPolicyDistribution/
│   │   ├── jmlPolicyAnalytics/
│   │   ├── dwxPolicyAuthorView/
│   │   └── dwxPolicyManagerView/
│   ├── components/        # Shared components
│   │   ├── JmlAppLayout/       # Full-page layout wrapper (with role filtering)
│   │   ├── JmlAppHeader/       # App header with navigation
│   │   ├── PageSubheader/      # Page subheader component
│   │   ├── PolicyManagerHeader/ # Policy Manager branded header with role-based nav
│   │   ├── PolicyManagerSplashScreen/
│   │   ├── QuizBuilder/        # Quiz creation, AI generation, question management
│   │   └── QuizTaker/          # Quiz-taking component (11 question types)
│   ├── services/          # 145+ business logic services + PolicyRoleService
│   ├── models/            # 56+ TypeScript interfaces
│   ├── hooks/             # Custom React hooks (useDialog, etc.)
│   ├── constants/         # SharePointListNames.ts, etc.
│   ├── styles/            # Centralized styling (fluent-mixins.scss)
│   └── utils/             # pnpConfig, injectPortalStyles, etc.
├── azure-functions/
│   └── quiz-generator/    # Azure Function — AI Quiz Question Generator
│       ├── src/functions/  # generateQuizQuestions.ts (HTTP trigger)
│       ├── infra/          # Bicep IaC + deploy.ps1 deployment script
│       ├── host.json
│       ├── package.json
│       └── tsconfig.json
├── scripts/
│   └── policy-management/ # PnP PowerShell provisioning scripts
├── docs/                  # Documentation + HTML mockups
├── config/                # SPFx build configurations
├── e2e/                   # Playwright e2e tests
├── testsprite_tests/      # TestSprite test cases (14 test files)
└── CLAUDE.md              # This file
```

---

## Design System

### Color Palette — Forest Teal Theme
The application uses a **Forest Teal** color scheme throughout, distinct from the DWx Blue brand.

| Name | Hex | Usage |
|------|-----|-------|
| Primary Teal | #0d9488 | Headers, active states, accents |
| Dark Teal | #0f766e | Gradient endpoints, hover states |
| Light Teal BG | #ccfbf1 | Active nav items, badges |
| Pale Teal BG | #f0fdfa | Summary bars, info panels |
| Text Primary | #0f172a | Headlines, primary text |
| Text Secondary | #334155 | Body text |
| Text Muted | #605e5c | Labels, descriptions |
| Border | #e2e8f0 | Card borders |
| Border Light | #edebe9 | Section separators |
| Sidebar BG | #f1f5f9 | Admin sidebar background |
| Warning | #d97706 | SLA warnings, thresholds |
| Error/Danger | #dc2626 | Auto-delete indicators |
| Success | #059669 | Active status indicators |

### Gradients
```css
/* Header gradient (teal) */
background: linear-gradient(135deg, #0d9488 0%, #0f766e 100%);

/* Summary bar gradient */
background: linear-gradient(135deg, #f0fdfa 0%, #ecfdf5 100%);
```

### Typography
- **Font Family**: Segoe UI (SharePoint default), system fallbacks
- **Font Weights**: 400 (body), 500 (labels), 600 (headings), 700 (KPI numbers)
- Components use Fluent UI `Text` component with `variant` props

### Spacing Scale
- XS: 4px | SM: 8px | MD: 12px | LG: 16px | XL: 20px | XXL: 24px

---

## Policy Admin Structure

The Policy Admin page uses a **sidebar + content area** layout pattern:

### Navigation Sections
```
CONFIGURATION
├── Templates           — Manage reusable policy templates
├── Metadata Profiles   — Configure metadata presets
├── Approval Workflows  — Approval chains and routing
├── Compliance Settings — Risk levels, acknowledgement, review settings
├── Notifications       — Email templates and alerts
├── Naming Rules        — Naming convention builder (NEW)
├── SLA Targets         — Service level agreements (NEW)
├── Data Lifecycle      — Data retention and archival (NEW)
└── Navigation          — Toggle app navigation items (NEW)

MANAGEMENT
├── Reviewers & Approvers — SharePoint group management
├── Audit Log             — Policy change history
└── Data Export           — CSV/report exports
```

### Admin Component Details

| Component | Description | State Fields |
|-----------|-------------|-------------|
| Naming Rules | Card-based naming convention builder with segment chips (prefix, counter, date, category) | `namingRules: INamingRule[]` |
| SLA Targets | Grid of SLA cards with target days, warning thresholds, progress bars | `slaConfigs: ISLAConfig[]` |
| Data Lifecycle | Retention policies per entity type with auto-delete and archive toggles | `lifecyclePolicies: IDataLifecyclePolicy[]` |
| Navigation | Toggle switches per nav item, Enable/Disable All, protected items | `navToggles: INavToggleItem[]` |

---

## Critical Development Notes

### JML Coupling (LEGACY)
The codebase retains JML-prefixed names for components and webparts. SharePoint list names have been migrated to `PM_` prefix via `src/constants/SharePointListNames.ts`.

#### Key Constants File
```typescript
import { PM_LISTS } from '../constants/SharePointListNames';
// Use: .getByTitle(PM_LISTS.POLICIES)
```

### Component Patterns
- **Class components** are used throughout (React 17 pattern, consistent with SPFx)
- **@ts-nocheck** is used in several large components to suppress strict warnings
- Services are instantiated with `new ServiceName(props.sp)` pattern
- Dialog management uses `createDialogManager()` from `src/hooks/useDialog`

### Role-Based Access Control
The application uses a 4-tier role hierarchy defined in `src/services/PolicyRoleService.ts`:

| Role | Who | Nav Access |
|------|-----|------------|
| **User** | All employees | Browse, My Policies, Details |
| **Author** | Policy writers | + Create, Packs, Author View |
| **Manager** | Department managers | + Approvals, Delegations, Distribution, Analytics, Manager View, Settings cog |
| **Admin** | System admins | + Quiz Builder, Admin panel |

Role detection flows: `WebPart.onInit()` → `RoleDetectionService` → `PolicyRoleService.mapToRole()` → passed via `JmlAppLayout` → `PolicyManagerHeader` → nav items filtered by `filterNavForRole()`.

### Build Configuration
All 14 webparts must be registered in `config/config.json`:
- `bundles` section: entry point + manifest for each webpart
- `localizedResources` section: locale file path for each webpart
- Missing entries will cause webparts to not appear in SharePoint

---

## SharePoint Lists

**All lists use the `PM_` prefix (Policy Manager).**

### Core Lists
| List Name | Purpose |
|-----------|---------|
| PM_Policies | Core policy records |
| PM_PolicyVersions | Version history |
| PM_PolicyAcknowledgements | User acknowledgements |
| PM_PolicyMetadataProfiles | Metadata presets |

### Quiz Lists
| List Name | Purpose |
|-----------|---------|
| PM_PolicyQuizzes | Quiz definitions |
| PM_PolicyQuizQuestions | Quiz questions |
| PM_PolicyQuizResults | Quiz results |

### Approval Lists
| List Name | Purpose | Script |
|-----------|---------|--------|
| PM_Approvals | Individual approval records | 08-Approval-Lists.ps1 |
| PM_ApprovalChains | Approval chain instances | 08-Approval-Lists.ps1 |
| PM_ApprovalHistory | Action audit trail | 08-Approval-Lists.ps1 |
| PM_ApprovalDelegations | Delegation assignments | 08-Approval-Lists.ps1 |
| PM_ApprovalTemplates | Reusable workflow templates | 08-Approval-Lists.ps1 |

### Notification Lists
| List Name | Purpose | Script |
|-----------|---------|--------|
| PM_Notifications | In-app notifications | 07-Notification-Lists.ps1 |
| PM_NotificationQueue | Email notification queue | 07-Notification-Lists.ps1 |

### Workflow Lists
| List Name | Purpose |
|-----------|---------|
| PM_PolicyExemptions | Exemption management |
| PM_PolicyDistributions | Distribution tracking |
| PM_PolicyTemplates | Policy templates library |

### Social/Engagement Lists
| List Name | Purpose |
|-----------|---------|
| PM_PolicyRatings | User ratings |
| PM_PolicyComments | Discussion comments |
| PM_PolicyCommentLikes | Comment likes |
| PM_PolicyShares | Share tracking |
| PM_PolicyFollowers | Policy followers |

### Policy Packs
| List Name | Purpose |
|-----------|---------|
| PM_PolicyPacks | Policy bundle definitions |
| PM_PolicyPackAssignments | Pack assignments |

### Analytics & Audit
| List Name | Purpose |
|-----------|---------|
| PM_PolicyAuditLog | Audit trail |
| PM_PolicyAnalytics | Usage analytics |
| PM_PolicyFeedback | User feedback |
| PM_PolicyDocuments | Supporting documents |

### Configuration & Admin Lists

| List Name | Purpose |
|-----------|---------|
| PM_Configuration | Key-value configuration store (ConfigKey, ConfigValue, Category, IsActive, IsSystemConfig) |
| PM_PolicySubCategories | Subcategory definitions (SubCategoryName, ParentCategoryId, ParentCategoryName, IconName, Description, SortOrder, IsActive) |
| PM_PolicyRequests | Policy creation requests with status tracking |
| PM_PolicySourceDocuments | Document library with per-policy folders (auto-created by PolicyService) |

### User Management Lists

| List Name        | Purpose                |
|------------------|------------------------|
| PM_UserProfiles  | User profile data      |
| PM_UserGroups    | User group definitions |

### Provisioning Scripts
Located in `scripts/policy-management/`:
| Script | Lists |
|--------|-------|
| `Create-PolicyManagementLists.ps1` | Core lists (Policies, Versions, Acknowledgements, etc.) |
| `02-Quiz-Lists.ps1` | Quiz system lists |
| `Create-PolicySocialLists.ps1` | Social engagement lists |
| `Create-PolicyTemplatesLibrary.ps1` | Templates document library |
| `07-Notification-Lists.ps1` | PM_Notifications, PM_NotificationQueue |
| `08-Approval-Lists.ps1` | 5 approval-related lists |
| `11-AdminConfig-Lists.ps1` | Admin configuration lists |
| `12-UserManagement-Lists.ps1` | User management lists |
| `13-Visibility-Columns.ps1` | Visibility + TargetSecurityGroups columns on PM_Policies |
| `14-SubCategory-And-Folders.ps1` | PM_PolicySubCategories list + SubCategory, LinkedQuizId, SourceRequestId on PM_Policies |
| `Deploy-AllPolicyLists.ps1` | Master deployment script (runs all numbered scripts) |
| `Seed-ApprovalAndNotificationData.ps1` | Sample data for approvals + notifications |
| `Deploy-SampleData.ps1` | Master sample data deployment |

Located in `scripts/` (root):

- `upgrade-quiz-questions-list.ps1` — Adds Quiz Builder columns to PM_PolicyQuizQuestions + creates PM_Configuration (idempotent)
- `Create-PM_PolicyRequests.ps1` — Creates PM_PolicyRequests list with status tracking fields

---

## Key Models

### IPolicy (src/models/IPolicy.ts)
- 80+ fields covering all policy aspects
- Supports versioning, acknowledgement, quizzes
- `PolicyVisibility` enum (AllEmployees, Department, Role, SecurityGroup, Custom)
- `Visibility`, `TargetSecurityGroups`, `SubCategory`, `LinkedQuizId`, `SourceRequestId` fields
- Data classification and retention
- Regulatory compliance mapping

### IAdminConfig (src/models/IAdminConfig.ts)

- 15+ interfaces: IPolicySubCategory, INamingRule, ISLAConfig, IDataLifecyclePolicy, INavToggleItem, IComplianceSetting, IMetadataProfile, etc.

### IJmlApproval (src/models/IJmlApproval.ts)
- IJmlApproval, IJmlApprovalChain, IJmlApprovalLevel
- IJmlApprovalHistory, IJmlApprovalDelegation, IJmlApprovalTemplate
- Enums: ApprovalStatus, ApprovalType, EscalationAction

### Policy Status Lifecycle
```
Draft → In Review → Pending Approval → Approved → Published → Archived/Retired
                  ↓
               Rejected
```

### Read Timeframes
- Immediate, Day 1, Day 3, Week 1, Week 2, Month 1, Month 3, Month 6, Custom

### Compliance Risk Levels
- Critical, High, Medium, Low, Informational

---

## Build Commands

```bash
# Install dependencies
npm install

# Development build
npm run build

# Production build (ship)
gulp clean && gulp bundle --ship && gulp package-solution --ship

# Package location
sharepoint/solution/policy-manager.sppkg

# Clean build artifacts
npm run clean
```

---

## Development Guidelines

### Styling
1. Use Forest Teal color scheme (#0d9488 primary)
2. Follow the teal gradient for headers: `linear-gradient(135deg, #0d9488, #0f766e)`
3. Use SCSS modules per component (ComponentName.module.scss)
4. Shared mixins in `src/styles/fluent-mixins.scss`

### Components
1. Use class components (consistent with existing codebase)
2. Use `JmlAppLayout` for full-page webparts
3. Use `createDialogManager()` for modal dialogs
4. Follow Fluent UI v8 patterns (Stack, Text, Icon, Toggle, etc.)

### Services
1. Use the singleton SPFI instance via `getSP()`
2. Use `PM_LISTS` constants for all list names
3. Add comprehensive audit logging for compliance
4. Handle errors gracefully with user-friendly messages

### PowerShell / Provisioning Scripts
1. **Always assume the user is already connected to SharePoint** — never include `Connect-PnPOnline` or `Disconnect-PnPOnline` in scripts
2. When a SharePoint site URL is needed, use: `https://mf7m.sharepoint.com/sites/PolicyManager`
3. Scripts should be idempotent — check for existing lists/fields before creating
4. `Add-PnPField` does NOT support `-DefaultValue` — use `Set-PnPField -Values @{DefaultValue="..."}` separately after creating the field

### Adding a New Webpart
1. Create webpart folder under `src/webparts/`
2. Add manifest.json, WebPart.ts, components/, loc/
3. Register in `config/config.json` → `bundles` and `localizedResources`
4. Build and verify manifest count in output

---

## DWx Standard Components

All DWx suite applications (Policy Manager, Asset Manager, Contract Manager, and future apps) implement these **4 standard components** for consistency across the suite. Each app adapts them to its own theme color and domain.

### 1. Breadcrumb Navigation (DWx Standard)
- **Interface**: `IBreadcrumb { text: string; href?: string }` exported from the app header
- **Location**: Rendered as a full-width strip below the navigation bar inside the header component
- **Pattern**: `AppHeader` → breadcrumb bar (light bg) → `/` separated items → last item is current page (bold, no link)
- **Policy Manager**: `PolicyManagerHeader` renders breadcrumbs, passed via `JmlAppLayout`

### 2. Search Center (DWx Standard)
- **Component**: `{App}Search.tsx` + `.module.scss`
- **Pattern**: Hero section (gradient + search input + quick-filter chips) → sidebar filters + result cards
- **Policy Manager**: `jmlPolicySearch` webpart with dedicated SharePoint page
- **Asset Manager**: Embedded view within single webpart, searches across assets/vendors/contracts

### 3. Admin Center (DWx Standard)
- **Component**: `{App}AdminCenter.tsx` + `.module.scss`
- **Pattern**: 280px sidebar (collapsible category groups → nav items) + content area (switch/case on `activeSection`)
- **Standard sections**: Navigation toggles, General Settings, Notifications, Data Export, Audit Log, System Info
- **Navigation persistence**: `{prefix}_nav_visibility` localStorage key (e.g., `am_nav_visibility`, `pm_nav_visibility`)
- **Policy Manager**: `jmlPolicyAdmin` webpart — 21 admin sections with full CRUD

### 4. Help Center (DWx Standard)
- **Component**: `{App}HelpCenter.tsx` + `.module.scss`
- **Pattern**: Hero section + tab bar (Home, Articles, FAQs, Shortcuts, Support)
- **Content**: Hardcoded initially, can be wired to SharePoint lists (`PM_HelpArticles`, `PM_Cheatsheets`, `PM_HelpTickets`)
- **Policy Manager**: `jmlPolicyHelp` webpart with `HelpCenterService` for full SP list integration

### DWx Standard Content Layout
All DWx apps use the same content area layout:

```css
max-width: 1400px;
width: 100%;
margin: 0 auto;
padding: 24px;
box-sizing: border-box;
```

### DWx App Color Themes

| App | Primary | Dark | Theme Name |
|-----|---------|------|------------|
| Policy Manager | #0d9488 | #0f766e | Forest Teal |
| Asset Manager | #475569 | #334155 | Slate Blue |
| Contract Manager | #475569 | #334155 | Slate Blue |

### 5. DWx Hub Integration (Cross-App Services)

Each DWx app optionally connects to a shared DWx Hub site for cross-app features. Integration uses graceful degradation — apps work fully standalone when Hub is unavailable.

**Architecture:**

```
WebPart.onInit() → DwxHubService(context, sp)
  → isHubAvailable() — checks if Hub site exists
  → DwxAppRegistryService.heartbeat('AppName', version)
  → Pass dwxHub to component props (undefined if unavailable)

Component constructor → PolicyService.setDwxServices(notificationSvc, activitySvc)
  → Cross-app notifications on policy publish
  → Activity feed logging on approve/publish
```

**Key classes from `@dwx/core`:**

| Class | Purpose |
|-------|---------|
| `DwxHubService` | Hub site detection and connection |
| `DwxAppRegistryService` | App heartbeat/registration |
| `DwxNotificationService` | Cross-app notification delivery |
| `DwxActivityService` | Cross-app activity feed logging |
| `DwxLinkedRecordService` | Cross-app record linking (PolicyDetails) |
| `DwxNotificationBell` | Cross-app notification bell UI component |
| `DwxAppSwitcher` | Suite app switching UI component |

**Props pattern:** `dwxHub?: DwxHubService` — optional on all component props interfaces. WebPart sets to `undefined` if Hub unavailable.

**Notification integration:** `PolicyNotificationService` accepts optional `DwxNotificationService` in constructor. When Hub is available, cross-app notifications fire alongside local SP list storage on policy events.

---

## Azure Functions — AI Quiz Generator

### Architecture
The Quiz Builder integrates with Azure OpenAI GPT-4o via an Azure Function to generate quiz questions from policy documents.

```
QuizBuilder (SPFx) → Azure Function (Node.js 18) → Azure OpenAI (GPT-4o)
                                                  ↗ Key Vault (API key)
```

### Deployed Resources (Resource Group: `dwx-pm-quiz-rg-prod`)
| Resource | Name | Region |
|----------|------|--------|
| Azure OpenAI | `dwx-pm-openai-prod` | swedencentral |
| Function App | `dwx-pm-quiz-func-prod` | swedencentral |
| Key Vault | `dwx-pm-kv-ziqv6cfh2ck3o` | swedencentral |
| Storage Account | `dwxpmstziqv6cfh2ck3o` | swedencentral |
| App Insights | `dwx-pm-quiz-insights-prod` | swedencentral |
| Log Analytics | `dwx-pm-quiz-logs-prod` | swedencentral |
| App Service Plan | `dwx-pm-quiz-plan-prod` (Y1 Consumption) | swedencentral |

### Function Endpoint
```
POST https://dwx-pm-quiz-func-prod.azurewebsites.net/api/generate-quiz-questions?code=<function-key>
```

### Infrastructure as Code
- **Bicep template**: `azure-functions/quiz-generator/infra/main.bicep`
- **Parameters**: `azure-functions/quiz-generator/infra/main.parameters.json`
- **Deployment script**: `azure-functions/quiz-generator/infra/deploy.ps1`

### Redeployment
```powershell
cd azure-functions/quiz-generator/infra
.\deploy.ps1 -Environment prod -Location swedencentral
```

---

## Quiz System

### Question Types (11 total)
1. Multiple Choice — 4 options (A-D), single correct answer
2. True/False — Binary choice
3. Multiple Select — Multiple correct answers (semicolon-separated)
4. Short Answer — Free text with expected answer
5. Fill in the Blank — Blank positions with accepted answers (JSON)
6. Matching — Left-right pair matching (JSON array)
7. Ordering — Sequence ordering (JSON array with correctOrder)
8. Rating Scale — Numeric scale with tolerance
9. Essay — Long-form with word count limits
10. Image Choice — Image-based multiple choice
11. Hotspot — Click-on-image coordinate selection

### Quiz Lists (SharePoint)
| List | Purpose |
|------|---------|
| PM_PolicyQuizzes | Quiz definitions (settings, passing score, attempts) |
| PM_PolicyQuizQuestions | Individual questions with type-specific fields |
| PM_PolicyQuizResults | User attempt results and scores |
| PM_PolicyQuizAttempts | Individual attempt tracking |
| PM_PolicyQuizAnswers | Per-question answer records |
| PM_PolicyQuizFeedback | User feedback on quizzes |

### AI Question Generation
The QuizBuilder's "AI Generate" panel calls the Azure Function with:
- Policy document text (extracted from SharePoint)
- Question count, difficulty level, question types
- Returns structured JSON questions ready for import into SharePoint lists

---

## Session State (Last Updated: 27 Feb 2026 — Session 9)

### Recently Completed (Session 9 — 27 Feb 2026)

#### Phase 1: Policy Versioning (Compliance)

- **PolicyService.ts** — `createVersion()` made public; added `createEditableVersion()` (snapshots content → bumps minor version → sets Draft); `publishPolicy()` bumps to next major version (1.1 → 2.0)
- **PolicyDetails.tsx** — Version History panel (PanelType.medium) with version cards, type badges (Major green / Minor blue), date, description, "Current" badge, "Compare with Current" button; Version Comparison panel (PanelType.extraLarge) with `dangerouslySetInnerHTML` rendering LCS diff HTML via `PolicyDocumentComparisonService`
- **PolicyAuthorEnhanced.tsx** — Wired version history/comparison buttons (replaced alert stubs); Edit Published Policy flow shows confirmation dialog → calls `createEditableVersion()` → loads new draft into wizard

#### Phase 2: Policy Security & Visibility

- **IPolicy.ts** — Added `PolicyVisibility` enum (AllEmployees, Department, Role, SecurityGroup, Custom), `Visibility` and `TargetSecurityGroups` fields
- **PolicyHubService.ts** — Added `IUserVisibilityContext` interface and `filterByVisibility()` method; Admin/Manager bypass; Author sees own; 5 visibility modes with graceful defaults
- **PolicyHub.tsx** — Enhanced `initializeUserContext()` to resolve department, jobTitle, groupNames from SPFx page context + SP groups; applied visibility filter after `searchPolicyHub()` results
- **13-Visibility-Columns.ps1** — New provisioning script; Visibility (Choice) + TargetSecurityGroups (Note) columns on PM_Policies; uses `Set-PnPField` for default value (not `-DefaultValue` param)

#### Phase 3: Virtual Folders (Subcategories) + Per-Policy Documents

- **IAdminConfig.ts** — Added `IPolicySubCategory` interface (Id, SubCategoryName, ParentCategoryId, ParentCategoryName, IconName, Description, SortOrder, IsActive)
- **SharePointListNames.ts** — Added `POLICY_SUB_CATEGORIES: 'PM_PolicySubCategories'`
- **AdminConfigService.ts** — New service (818 lines); full admin configuration CRUD for templates, metadata, compliance, naming rules, SLA targets, subcategories
- **PolicyAdmin.tsx** — Major expansion (+2,288 lines); Sub-Categories management section in sidebar with DetailsList, create/edit Panel, parent category dropdown filter; all 21+ admin sections wired to AdminConfigService
- **PolicyHub.tsx** — Category tree navigation in facets panel with expand/collapse, policy counts per node, click-to-filter
- **PolicyService.ts** — `ensurePolicyFolder(policyNumber)` and `getPolicyDocuments(policyNumber)` for per-policy document folders in PM_PolicySourceDocuments
- **PolicyDetails.tsx** — Collapsible Documents section with lazy-loaded file list
- **14-SubCategory-And-Folders.ps1** — New provisioning script; PM_PolicySubCategories list (7 columns) + SubCategory, LinkedQuizId, SourceRequestId on PM_Policies

#### Phase 4: Quiz Sequencing + Request-to-Policy Improvements

- **PolicyAuthorEnhanced.tsx** — Disabled "Create New Quiz" in wizard Step 3 → "Create Quiz After Publish" (disabled + tooltip); post-publish quiz reminder dialog when RequiresQuiz && !LinkedQuizId; "Quiz Missing" amber badge in Author View; expanded `onCreatePolicyFromRequest` field mapping (7 new fields: BusinessJustification→policySummary, Priority→complianceRisk, DesiredEffectiveDate→effectiveDate, TargetAudience→departments, AdditionalNotes, AttachmentUrls, Id→sourceRequestId)
- **PolicyRequestsTab.tsx** — "Accept & Start Drafting" now also calls `onCreatePolicyFromRequest()` to open wizard with pre-filled data
- **PolicyService.ts** — Auto-complete source request on publish (updates PM_PolicyRequests status to 'Completed' with ResultingPolicyId)
- **IPolicy.ts** — Added `LinkedQuizId` and `SourceRequestId` fields, persisted in `handleSaveDraft()`

#### Component Decomposition & Refactoring

- Extracted 6 tab components from PolicyAuthorEnhanced.tsx: AnalyticsTab, DelegationsTab, PolicyPacksTab, PolicyRequestsTab, QuizBuilderTab + shared types.ts and index.ts barrel export
- Extracted PolicyWizard.tsx wizard component
- Net reduction: ~1,500 lines removed from PolicyAuthorEnhanced.tsx

#### New Services & Infrastructure

- **AdminConfigService.ts** — Full admin configuration CRUD (818 lines)
- **UserManagementService.ts** — User/group management against PM_UserProfiles, PM_UserGroups (336 lines)
- **AudienceService.ts** — Audience targeting and resolution (291 lines)
- **retryUtils.ts** — Exponential backoff retry utility with configurable attempts
- **pnp-sp-augmentations.d.ts** — PnP/SP type augmentations

#### Test Infrastructure

- **jest.config.js** — Jest configuration for SPFx project
- 5 mock files: dwxCoreMock.js, pnpMock.js, setupGlobals.js, spfxMock.js, styleMock.js
- 6 unit test suites: AudienceService, LoggingService, PolicyDistributionService, PolicyRoleService, RecentlyViewedService, UserManagementService

#### Build & Push

- **Ship build** — Zero errors, all 14 webpart manifests, v1.2.4
- **Commit** — `3a383b1` — pushed to both ADO and GitHub remotes
- **Provisioning** — 13-Visibility-Columns.ps1 and 14-SubCategory-And-Folders.ps1 run successfully on live site

### Previously Completed (Session 8 — 26 Feb 2026)

- Live data wiring: PolicyAnalytics (5 SP lists), PolicyDistribution (3 SP lists), PolicyHub + PolicyDetails recently viewed tracking
- New services: PolicyDistributionService.ts, RecentlyViewedService.ts
- Application Insights telemetry via Beacon API (LoggingService rewrite)
- Admin nav toggles wired end-to-end (PolicyAdmin ↔ PolicyManagerHeader via localStorage)
- DWx Hub integration expansion (NotificationBell, AppSwitcher, LinkedRecords)
- Provisioning scripts: Seed-DwxAppRegistry.ps1, Provision-SharePointPages.ps1

### Previously Completed (Sessions 1-7)

- Enterprise hardening — 9/10 critical security + performance fixes (Session 7)
- Image templates, quiz selection, fullscreen viewer, DWx Hub integration (Session 6)
- QuizBuilder UX overhaul, AI pipeline hardening, Recently Viewed dropdown (Session 5)
- QuizTaker rewrite (11 question types), Azure Function AI quiz generation (Session 4)
- PolicyRoleService, role-based nav, Author/Manager Views, Distribution, Analytics, Search, Help webparts, MyPolicies rewrite, Admin restructure, approval lists (Sessions 1-3)

### Enterprise Readiness Assessment

| Area | Score | Notes |
| ------ | ------- | ------- |
| Security | 8/10 | XSS fixed, OData sanitized, auth checks, visibility filtering. Remaining: expand sanitization to all 100+ filter sites |
| Performance | 7.5/10 | Parallelized loads, reduced query limits, live data. Remaining: React.lazy, virtualization |
| Reliability | 8/10 | Error boundaries, Application Insights ready, LoggingService type-checked. Remaining: @ts-nocheck removal from ~198 files |
| Code Quality | 7/10 | Types extracted, tab decomposition started (6 tabs extracted from god component), AdminConfigService. Remaining: complete decomposition |
| Testing | 3.5/10 | Jest config + mocks + 6 unit test suites. Remaining: integration tests, component tests, E2E |
| Accessibility | 3/10 | Basic Fluent UI a11y only. No ARIA roles, keyboard nav, screen reader testing |
| Overall | ~74/100 | Up from 72/100. Versioning, visibility, subcategories, test infrastructure, component decomposition |

### Known Issues

- PowerShell scripts starting with numbers need `.\` prefix to execute
- SPFx CDN caching may require version bump + app catalog re-upload + hard refresh to see updates
- `az` CLI not in PATH in VSCode terminal — use full path: `C:\Program Files (x86)\Microsoft SDKs\Azure\CLI2\wbin\az.cmd`
- Image upload metadata (`DocumentType`, `FileStatus`, etc.) fails with 400 if PM_PolicySourceDocuments custom columns not provisioned — non-blocking, caught by try/catch
- DWx Hub integration is experimental — `@dwx/core` services are wired but Hub site may not exist yet; all calls degrade gracefully
- OData sanitization covers 9 HIGH RISK sites; ~90+ MEDIUM/LOW risk sites remain (numeric IDs, enum constants — lower priority)
- `@ts-nocheck` remains in ~198 files — PolicyService.ts and LoggingService.ts are fully type-checked
- PnP PowerShell `Add-PnPField` does NOT support `-DefaultValue` — use `Set-PnPField -Values @{DefaultValue="..."}` separately
- Disk space on dev machine was critically low (162MB free / 879GB) during Session 8 — may need cleanup

### Next Steps

- User testing of versioning: create draft → publish → edit published → version history → comparison
- User testing of visibility: set policy visibility to Department/SecurityGroup, verify non-matching users don't see it
- User testing of subcategories: create subcategories in admin, assign to policies, verify tree navigation in PolicyHub
- Run `Provision-SharePointPages.ps1` to create all 13 SharePoint pages
- Run `10-CorporateTemplates.ps1` to provision PM_PolicySourceDocuments custom columns (fixes metadata 400 error)
- Initialize Application Insights in production — set connection string in PM_Configuration or Admin Settings
- Wire remaining webparts to live SharePoint data (MyPolicies featured/stats, PolicyHub featured section)
- Continue @ts-nocheck removal from critical services (ApprovalService, QuizService, PolicyHubService)
- Expand unit test coverage (currently 6 suites — need component tests, integration tests)
- Add accessibility improvements (ARIA roles, keyboard navigation, screen reader support)
- Continue component decomposition — extract remaining sections from PolicyAuthorEnhanced.tsx
