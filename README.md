# DWx Policy Manager

Enterprise-grade Policy Lifecycle Management system built on SharePoint Framework (SPFx). Part of the DWx (Digital Workplace Excellence) suite by First Digital.

## SharePoint Framework Version

![version](https://img.shields.io/badge/SPFx-1.20.0-green.svg)
![version](https://img.shields.io/badge/React-17.0.1-blue.svg)
![version](https://img.shields.io/badge/Fluent_UI-8.x-purple.svg)
![version](https://img.shields.io/badge/PnPjs-3.x-orange.svg)

## Solution

| Solution | Author(s) |
| --- | --- |
| DWx Policy Manager | First Digital / Gary Finberg |

## Web Parts (14)

| Web Part | Description | Page |
| --- | --- | --- |
| **JmlPolicyHub** | Main policy browsing interface with KPI dashboard, category tree sidebar, visibility filtering, advanced filtering, table/card views | `PolicyHub.aspx` |
| **JmlMyPolicies** | Personal dashboard showing assigned policies, due dates, completion status | `MyPolicies.aspx` |
| **JmlPolicyAdmin** | Admin panel with sidebar navigation — 21+ sections including templates, metadata, workflows, compliance, SLA, naming rules, lifecycle, navigation toggles, sub-categories | `PolicyAdmin.aspx` |
| **JmlPolicyAuthor** | Policy creation wizard with rich text editor, metadata, workflow submission, version history, edit-published-policy flow, request-to-policy mapping | `PolicyBuilder.aspx` |
| **DwxPolicyAuthorView** | Author dashboard — policies, approvals, delegations, activity tabs | `PolicyAuthor.aspx` |
| **JmlPolicyDetails** | Full policy viewer with version history + comparison panels, per-policy documents, acknowledgement, quiz, recently viewed tracking, cross-app record linking | `PolicyDetails.aspx` |
| **JmlPolicyPackManager** | Bundle policies into packs, assign to users/groups, track completion | `PolicyPacks.aspx` |
| **DwxQuizBuilder** | Quiz creation and management with AI-powered question generation (Azure OpenAI GPT-4o) | `QuizBuilder.aspx` |
| **JmlPolicySearch** | Dedicated search center with hero section, sidebar filters, result cards | `PolicySearch.aspx` |
| **JmlPolicyHelp** | Help center with articles, FAQs, shortcuts, videos, support tabs | `PolicyHelp.aspx` |
| **JmlPolicyDistribution** | Distribution campaign management — create, track, and manage policy distribution with live SharePoint data | `PolicyDistribution.aspx` |
| **JmlPolicyAnalytics** | Executive analytics dashboard (6 tabs: Executive, Policy Metrics, Acknowledgements, SLA, Compliance, Audit) with live SP data | `PolicyAnalytics.aspx` |
| **DwxPolicyManagerView** | Manager compliance dashboard — team compliance, approvals, delegations, reviews, reports | `PolicyManagerView.aspx` |

## Architecture

```text
src/
  webparts/                    # 14 SPFx web parts
    jmlPolicyHub/              # Main hub — browsing & dashboard
    jmlMyPolicies/             # My assigned policies
    jmlPolicyAdmin/            # Admin panel (sidebar layout, 21+ sections, sub-categories)
    jmlPolicyAuthor/           # Policy authoring wizard + version history
      components/tabs/         # Extracted tab components (6 tabs)
      components/wizard/       # Extracted PolicyWizard component
    dwxPolicyAuthorView/       # Author dashboard (4 tabs)
    jmlPolicyDetails/          # Policy detail viewer + acknowledgement
    jmlPolicyPackManager/      # Policy pack management
    dwxQuizBuilder/            # Quiz builder + AI generation
    jmlPolicySearch/           # Search center
    jmlPolicyHelp/             # Help center
    jmlPolicyDistribution/     # Distribution campaigns (live SP data)
    jmlPolicyAnalytics/        # Executive analytics (6 tabs, live SP data)
    dwxPolicyManagerView/      # Manager compliance dashboard
  components/                  # Shared components
    JmlAppLayout/              # Full-page layout wrapper with role filtering
    JmlAppHeader/              # App header with DWx Hub integration
    PolicyManagerHeader/       # Global header with nav icons, admin toggle filtering
    PageSubheader/             # Page subheader component
    QuizBuilder/               # Quiz creation, AI generation, question management
    QuizTaker/                 # Quiz-taking component (11 question types)
    ErrorBoundary/             # React error boundary with retry
  models/                      # 58+ TypeScript interfaces
    IPolicy.ts                 # Core policy models (80+ fields, versioning, visibility)
    IAdminConfig.ts            # Admin config interfaces (15+ types)
    IJmlApproval.ts            # Approval workflow models
  services/                    # 150+ service layer
    PolicyService.ts           # Core CRUD, versioning, document folders, authorization (type-checked)
    PolicyHubService.ts        # Hub search, visibility filtering (IUserVisibilityContext)
    PolicyDistributionService.ts  # Distribution campaigns — live SP CRUD
    PolicyNotificationService.ts  # In-app + email + DWx cross-app notifications
    AdminConfigService.ts      # Admin config CRUD (templates, SLA, naming, subcategories)
    RecentlyViewedService.ts   # localStorage recently viewed tracking
    LoggingService.ts          # Dual-mode telemetry (console + Application Insights)
    PolicyRoleService.ts       # 4-tier RBAC (User, Author, Manager, Admin)
    QuizService.ts             # Quiz CRUD, results, AI generation
    __tests__/                 # 6 unit test suites
  constants/                   # Configuration
    SharePointListNames.ts     # All PM_ list name constants
  styles/                      # Shared styles
    fluent-mixins.scss         # SCSS mixins for full-bleed layouts
  types/                       # TypeScript type augmentations
  utils/                       # pnpConfig, retryUtils, SharePointOverrides, injectPortalStyles
azure-functions/
  quiz-generator/              # Azure Function — AI Quiz Question Generator (GPT-4o)
    infra/                     # Bicep IaC + deployment script
scripts/
  policy-management/           # PnP PowerShell provisioning
    Deploy-AllPolicyLists.ps1  # Master list deployment
    Provision-SharePointPages.ps1  # Create all 13 SharePoint pages
    Deploy-SampleData.ps1      # Sample data seeding
  Seed-DwxAppRegistry.ps1      # DWx Hub app registry seeding
docs/                          # Architecture docs, proposals, mockups
```

## Key Services

| Service | Purpose |
| --- | --- |
| `PolicyService` | Core CRUD, versioning (createEditableVersion, version bumps), per-policy document folders, authorization checks (fully type-checked) |
| `PolicyHubService` | Hub search, visibility filtering (filterByVisibility — 5 modes with Admin/Manager bypass) |
| `AdminConfigService` | Admin configuration CRUD — templates, metadata, compliance, naming rules, SLA, subcategories |
| `PolicyDistributionService` | Distribution campaign CRUD against PM_PolicyDistributions |
| `PolicyNotificationService` | In-app + email notifications, optional DWx cross-app delivery |
| `RecentlyViewedService` | localStorage-based recently viewed policies (max 10 items) |
| `LoggingService` | Dual-mode: console-only or Azure Application Insights (Beacon API, no npm dep) |
| `PolicyRoleService` | 4-tier RBAC: User → Author → Manager → Admin |
| `QuizService` | Quiz CRUD, AI question generation pipeline |

## Role-Based Access Control

| Role | Access |
| --- | --- |
| **User** | Browse, My Policies, Policy Details |
| **Author** | + Create, Policy Packs, Author Dashboard |
| **Manager** | + Approvals, Distribution, Analytics, Manager Dashboard |
| **Admin** | + Quiz Builder, Admin panel, all settings |

## DWx Hub Integration

Optional cross-app integration via `@dwx/core`. All apps work fully standalone when Hub is unavailable (graceful degradation).

- **DwxNotificationBell** — Cross-app notification bell in header
- **DwxAppSwitcher** — Switch between DWx suite apps
- **DwxLinkedRecordService** — Cross-app record linking (PolicyDetails)
- **DwxNotificationService** — Cross-app notification delivery on policy events

## SharePoint Lists

All lists use the `PM_` prefix. Full definitions in `src/constants/SharePointListNames.ts`.

- **Policy Core**: PM_Policies, PM_PolicyVersions, PM_PolicyAcknowledgements, PM_PolicyExemptions, PM_PolicyDistributions, PM_PolicyTemplates, PM_PolicySourceDocuments (library)
- **Quiz**: PM_PolicyQuizzes, PM_PolicyQuizQuestions, PM_PolicyQuizResults, PM_PolicyQuizAttempts, PM_PolicyQuizAnswers
- **Approval**: PM_Approvals, PM_ApprovalChains, PM_ApprovalHistory, PM_ApprovalDelegations, PM_ApprovalTemplates
- **Analytics**: PM_PolicyAnalytics, PM_PolicyAuditLog, PM_PolicyFeedback
- **Notifications**: PM_Notifications, PM_NotificationQueue
- **Social**: PM_PolicyRatings, PM_PolicyComments, PM_PolicyCommentLikes, PM_PolicyShares, PM_PolicyFollowers
- **Policy Packs**: PM_PolicyPacks, PM_PolicyPackAssignments
- **Admin/Config**: PM_Configuration, PM_PolicySubCategories, PM_PolicyRequests
- **User Management**: PM_UserProfiles, PM_UserGroups

## Design System

**Forest Teal** color palette throughout:

| Token | Hex | Usage |
| --- | --- | --- |
| Primary | `#0d9488` | Buttons, links, active states, sidebar accents |
| Primary Dark | `#0f766e` | Gradient endpoints, hover states |
| Primary Light | `#ccfbf1` | Active backgrounds, highlights |
| Sidebar BG | `#f1f5f9` | Left sidebar background |

Gradient: `linear-gradient(135deg, #0d9488 0%, #0f766e 100%)`

Font: Segoe UI (system fallbacks)

## Enterprise Features (v1.2.4)

- **Policy Versioning** — Full compliance versioning with version history panels, side-by-side LCS diff comparison, minor version bumps on edit (1.0 → 1.1), major version bumps on publish (1.1 → 2.0), and "Edit Published Policy" flow that creates a new draft version.
- **Policy Visibility & Security** — Client-side visibility filtering with 5 modes: All Employees, Department, Role, Security Group, Custom. Admin/Manager bypass all filters. Authors always see their own policies. Built on `IUserVisibilityContext` resolved from SPFx page context + SP groups.
- **Category Tree Navigation** — Hierarchical browsing via Category > SubCategory in PolicyHub facets panel. Managed through `PM_PolicySubCategories` list with full CRUD in Admin panel.
- **Per-Policy Document Folders** — Auto-created folders in `PM_PolicySourceDocuments` library per policy number. Documents listed in collapsible section on PolicyDetails.
- **Quiz Sequencing** — Quiz creation disabled during policy drafting (Step 3) — only available after publish. Post-publish reminder dialog when quiz required but not linked. "Quiz Missing" badge in Author View.
- **Request-to-Policy Flow** — Expanded field mapping (7 fields) from policy requests to wizard state. "Accept & Start Drafting" opens wizard with pre-filled data. Auto-complete source request on publish.

## Admin Navigation Toggles

Admin panel controls which nav items are visible across the app. Settings persist to `pm_nav_visibility` in localStorage and are read by `PolicyManagerHeader` to filter the navigation bar. Protected items (Policy Hub, Administration) cannot be disabled.

## Application Insights Telemetry

`LoggingService` supports dual-mode telemetry:

- **Console-only** (default) — all telemetry goes to `console.log/warn/error`
- **Application Insights** — when initialized with a connection string, sends telemetry via Beacon API (zero npm dependencies)

Methods: `trackPageView`, `trackEvent`, `trackException`, `trackMetric`, `trackDependency`

## Provisioning

```powershell
# Assumes you are already connected to SharePoint via Connect-PnPOnline
cd scripts/policy-management

# Deploy all lists
.\Deploy-AllPolicyLists.ps1

# Create all 13 SharePoint pages
.\Provision-SharePointPages.ps1

# Seed sample data
.\Deploy-SampleData.ps1

# Seed DWx Hub registry (if Hub site exists)
cd ../
.\Seed-DwxAppRegistry.ps1
```

## Prerequisites

- Node.js 18.x, 20.x, or 22.x
- SharePoint Online tenant with App Catalog
- PnP PowerShell module (for provisioning)
- SharePoint Site: `https://mf7m.sharepoint.com/sites/PolicyManager`

## Build & Deploy

```bash
# Install dependencies
npm install

# Development build
npm run build

# Production build
gulp clean && gulp bundle --ship && gulp package-solution --ship

# Output: sharepoint/solution/policy-manager.sppkg
```

Upload the `.sppkg` file to the SharePoint App Catalog, then add web parts to their respective pages.

## Repository

- **Azure DevOps**: `https://dev.azure.com/gfinberg/DWx/_git/dwx-policy-manager`
- **GitHub (mirror)**: `https://github.com/garyfinberg24-png/dwx-policy-manager`

## Version History

| Version | Date | Comments |
| --- | --- | --- |
| 1.2.4 | February 2026 | Enterprise features: policy versioning + comparison, visibility/security filtering, subcategory tree navigation, per-policy document folders, quiz sequencing fix, request-to-policy flow, admin config CRUD, component decomposition (6 extracted tabs), 6 unit test suites, test infrastructure |
| 1.2.3 | February 2026 | Live data wiring (Analytics, Distribution), Application Insights telemetry, admin nav toggles, DWx Hub expansion, RecentlyViewedService, provisioning scripts |
| 1.2.2 | February 2026 | Image templates, quiz selection, fullscreen viewer, DWx Hub integration, enterprise hardening (9/10 security + performance fixes) |
| 1.2.1 | January 2026 | Quiz Builder UX overhaul, AI pipeline hardening, Recently Viewed dropdown |
| 1.1.0 | January 2026 | QuizTaker rewrite (11 question types), Azure Function AI quiz generation, deployed to production |
| 1.0.0 | January 2026 | 14 webparts, Forest Teal theme, admin sidebar, search/help centers, approval workflows, role-based access |

## References

- [SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview)
- [Fluent UI React v8](https://developer.microsoft.com/en-us/fluentui)
- [PnP/PnPjs v3](https://pnp.github.io/pnpjs/)
- [PnP PowerShell](https://pnp.github.io/powershell/)

## License

Proprietary — First Digital. All rights reserved.
