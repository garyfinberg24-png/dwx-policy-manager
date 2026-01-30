# Policy Manager - Claude Code Context

## Instructions for Claude

1. **Always read CLAUDE.md before you do anything**
2. **Always ask questions if you are unsure of the task or requirement**
3. **Be systematic in your planning, and execution**
4. **After you complete a task, always validate the result**
5. **We are working in https://mf7m.sharepoint.com/sites/PolicyManager**

---

## Project Overview

**Policy Manager** is an enterprise Policy Management Solution being spun off from the integrated JML (Joiner, Mover, Leaver) solution to become a completely standalone application within the **DWx (Digital Workplace Excellence)** suite by First Digital.

### Application Identity
- **App Name**: Policy Manager
- **Suite**: DWx (Digital Workplace Excellence)
- **Company**: First Digital
- **Tagline**: Policy Governance & Compliance
- **Current Version**: 1.1.0
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
│   ├── services/          # 141+ business logic services + PolicyRoleService
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
| `Deploy-AllPolicyLists.ps1` | Master deployment script |
| `Seed-ApprovalAndNotificationData.ps1` | Sample data for approvals + notifications |
| `Deploy-SampleData.ps1` | Master sample data deployment |

---

## Key Models

### IPolicy (src/models/IPolicy.ts)
- 80+ fields covering all policy aspects
- Supports versioning, acknowledgement, quizzes
- Data classification and retention
- Regulatory compliance mapping

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

### Adding a New Webpart
1. Create webpart folder under `src/webparts/`
2. Add manifest.json, WebPart.ts, components/, loc/
3. Register in `config/config.json` → `bundles` and `localizedResources`
4. Build and verify manifest count in output

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

## Session State (Last Updated: 30 Jan 2026 — Session 4)

### Recently Completed (Session 4 — 30 Jan 2026)

#### Quiz System Overhaul
- **QuizTaker rewrite** — Complete rewrite with proper TypeScript (removed `@ts-nocheck`), all 11 question type renderers, timer leak fix, retake fix, remaining attempts fix
- **QuizService type safety** — Removed `@ts-nocheck`, fixed unused members, fixed `delete` operator on non-optional properties
- **QuizBuilder enhancements** — Removed `@ts-nocheck`, added AI Generate panel, move up/down/duplicate buttons per question
- **Quiz provisioning** — Updated `02-Quiz-Lists.ps1` with 6 quiz-related lists and all question type fields

#### AI Quiz Generator (Azure Function)
- **Azure Function** — `generateQuizQuestions.ts` HTTP trigger with PDF extraction, GPT-4o prompt engineering, structured JSON output
- **Azure Infrastructure** — Bicep template provisioning OpenAI, Functions, Key Vault, Storage, App Insights, Log Analytics, RBAC
- **Deployed to production** — `dwx-pm-quiz-func-prod` in swedencentral, tested and working
- **QuizBuilder integration** — Function URL hardcoded as default in AI Generate panel

#### Policy Details Integration
- **QuizTaker wired to live data** — PolicyDetails looks up active published quiz via `QuizService.getQuizzesByPolicy()`, renders `<QuizTaker>` with fallback to mock quiz

#### Version & Packaging
- **Version bump** — Solution 1.0.0.0 → 1.1.0.0, package.json 1.0.0 → 1.1.0 (CDN cache busting)
- **Ship build** — Zero errors, all 14 webpart manifests

### Previously Completed (Sessions 1-3)
- PolicyRoleService — 4-tier role hierarchy with nav filtering
- Role-based nav filtering threaded through JmlAppLayout → PolicyManagerHeader
- DWx Policy Author View webpart — 4 tabs (My Policies, Approvals, Delegations, Activity)
- DWx Policy Manager View webpart — 6 tabs (Dashboard, Team Compliance, Approvals, Delegations, Reviews, Reports)
- Policy Distribution webpart — Campaign management with 4 tabs
- Policy Analytics webpart — Executive dashboard with 6 tabs
- Search Center, Help Center webparts
- MyPolicies rewrite, Policy Admin restructure with 12 nav sections
- Approval lists provisioning, seed data
- FOUC prevention, splash screen, layout enhancements

### Known Issues
- PowerShell scripts starting with numbers need `.\` prefix to execute
- Featured Policies and Recently Viewed sections hidden by default until Admin Navigation toggle is wired
- SPFx CDN caching may require version bump + app catalog re-upload + hard refresh to see updates
- `az` CLI not in PATH in VSCode terminal — use full path: `C:\Program Files (x86)\Microsoft SDKs\Azure\CLI2\wbin\az.cmd`

### Next Steps
- User testing of Quiz Builder AI generation with real policy documents
- Wire remaining webparts to live SharePoint data (Analytics, Distribution)
- Wire Admin Navigation toggles to control nav item visibility
- Create remaining SharePoint pages if not already created
- Connect Distribution webpart to live data from PolicyDistributionService
