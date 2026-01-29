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

## Web Parts (9)

| Web Part | Description | Page |
| --- | --- | --- |
| **JmlPolicyHub** | Main policy browsing interface with KPI dashboard, advanced filtering, table/card views | `PolicyHub.aspx` |
| **JmlPolicyDetails** | Full policy viewer with version history, acknowledgement, quiz, feedback | `PolicyDetails.aspx` |
| **JmlPolicyAuthor** | Policy authoring and editing with rich text editor, metadata, workflow submission | `PolicyAuthor.aspx` |
| **JmlPolicyAdmin** | Admin panel with sidebar navigation — templates, metadata, workflows, compliance, SLA, naming rules, lifecycle, navigation | `PolicyAdmin.aspx` |
| **JmlPolicyPackManager** | Bundle policies into packs, assign to users/groups, track completion | `PolicyPacks.aspx` |
| **JmlMyPolicies** | Personal dashboard showing assigned policies, due dates, completion status | `MyPolicies.aspx` |
| **DwxQuizBuilder** | Quiz creation and management for policy comprehension testing | `QuizBuilder.aspx` |
| **JmlPolicySearch** | Dedicated search center with filters, category chips, result cards | `PolicySearch.aspx` |
| **JmlPolicyHelp** | Help center with articles, FAQs, shortcuts, videos, support tabs | `PolicyHelp.aspx` |

## Architecture

```
src/
  webparts/                    # 9 SPFx web parts
    jmlPolicyHub/              # Main hub - browsing & dashboard
    jmlPolicyDetails/          # Policy detail viewer
    jmlPolicyAuthor/           # Policy authoring
    jmlPolicyAdmin/            # Admin panel (sidebar layout)
    jmlPolicyPackManager/      # Policy pack management
    jmlMyPolicies/             # My assigned policies
    dwxQuizBuilder/            # Quiz builder
    jmlPolicySearch/           # Search center
    jmlPolicyHelp/             # Help center
  components/                  # Shared components
    JmlAppLayout/              # Shared header/footer layout wrapper
    PolicyManagerHeader/       # Global header with nav icons
    PolicyManagerFooter/       # Global footer
  models/                      # TypeScript interfaces
    IPolicy.ts                 # Core policy models (80+ fields)
    IJmlApproval.ts            # Approval workflow models
    ICommon.ts                 # Shared types (IUser, etc.)
  services/                    # Service layer
    PolicyService.ts           # SharePoint CRUD operations
  constants/                   # Configuration
    SharePointListNames.ts     # All PM_ list name constants
  styles/                      # Shared styles
    fluent-mixins.scss         # SCSS mixins for full-bleed layouts
```

## SharePoint Lists

All lists use the `PM_` prefix. Key list groups:

- **Policy Core**: PM_Policies, PM_PolicyVersions, PM_PolicyAcknowledgements, PM_PolicyExemptions, PM_PolicyDistributions, PM_PolicyTemplates, PM_PolicyCategories
- **Quiz**: PM_PolicyQuizzes, PM_PolicyQuizQuestions, PM_PolicyQuizResults, PM_QuizAttempts, PM_QuizCertificates
- **Workflow**: PM_WorkflowTemplates, PM_WorkflowInstances, PM_ApprovalDecisions, PM_Delegations, PM_EscalationRules
- **Analytics**: PM_PolicyAnalytics, PM_UserActivityLog, PM_ComplianceViolations, PM_AuditTrail, PM_AuditReports
- **Social**: PM_PolicyRatings, PM_PolicyComments, PM_PolicyFollowers
- **Policy Packs**: PM_PolicyPacks, PM_PolicyPackAssignments
- **Notifications**: PM_PolicyNotifications, PM_NotificationQueue, PM_ReminderSchedule
- **Retention**: PM_RetentionPolicies, PM_LegalHolds, PM_RetentionArchive

Full list definitions in `src/constants/SharePointListNames.ts`.

## Design System

**Forest Teal** color palette throughout:

| Token | Hex | Usage |
| --- | --- | --- |
| Primary | `#0d9488` | Buttons, links, active states, sidebar accents |
| Primary Dark | `#0f766e` | Gradient endpoints, hover states |
| Primary Light | `#ccfbf1` | Active backgrounds, highlights |
| Sidebar BG | `#f1f5f9` | Left sidebar background |
| Content BG | `#ffffff` | Main content areas |

Gradient: `linear-gradient(135deg, #0d9488 0%, #0f766e 100%)`

Font: Inter (fallback: Segoe UI, system fonts)

Full style guide: `docs/DWx-Style-Guide-Preview.html`

## Policy Admin Sections

The Admin panel uses a left sidebar (280px) + right content area layout with 12 sections:

**CONFIGURATION**: Templates, Metadata Profiles, Approval Workflows, Compliance Settings, Notifications, Naming Rules, SLA Targets, Data Lifecycle, Navigation

**MANAGEMENT**: Reviewers & Approvers, Audit Log, Data Export

## Provisioning

PowerShell scripts for SharePoint list creation:

| Script | Purpose |
| --- | --- |
| `scripts/provision-approval-lists.ps1` | Approval workflow lists (PnP PowerShell) |
| `scripts/provision-notification-lists.ps1` | Notification lists (PnP PowerShell) |

## Prerequisites

- Node.js 18.x LTS
- SharePoint Online tenant with App Catalog
- PnP PowerShell module (for provisioning)
- Site URL: configured via web part properties

## Build & Deploy

```bash
# Install dependencies
npm install

# Development (local workbench)
gulp serve

# Production build
gulp clean && gulp bundle --ship && gulp package-solution --ship

# Output: sharepoint/solution/policy-manager.sppkg
```

Upload the `.sppkg` file to the SharePoint App Catalog, then add web parts to their respective pages.

## Version History

| Version | Date | Comments |
| --- | --- | --- |
| 2.0 | January 29, 2026 | 9 webparts, Forest Teal theme, admin sidebar, search/help centers, approval workflows |
| 1.0 | January 2026 | Initial implementation — Policy Hub, Details, Author, Admin, Pack Manager, MyPolicies, QuizBuilder |

## References

- [SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview)
- [Fluent UI React v8](https://developer.microsoft.com/en-us/fluentui)
- [PnP/PnPjs v3](https://pnp.github.io/pnpjs/)
- [PnP PowerShell](https://pnp.github.io/powershell/)
