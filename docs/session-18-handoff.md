# Session 18 Handoff Document — Policy Manager
**Date**: 25-27 March 2026
**Agent**: Claude Opus 4.6 (1M context)
**Status**: Production-ready with demo flow working end-to-end

---

## What Was Accomplished (Session 18)

### 1. Email Notification Pipeline (LIVE)
- **Architecture**: SPFx writes to `PM_NotificationQueue` → Azure Logic App polls every 5 min → Office 365 sends → Outlook
- **Critical fix**: SharePoint reserved field name `Status` was silently dropping values via REST API. Renamed to `QueueStatus` across all 5 files + Logic App definition.
- **Logic App**: `dwx-pm-email-sender-prod` in resource group `dwx-pm-email-rg-prod` (subscription `7784acd0-42ce-4073-b0e8-b85841c6c7e3` — different from main subscription)
- **Logic App JSON**: Full definition was provided and deployed. Uses `QueueStatus eq 'Pending'` filter, reads `RecipientEmail`, `Title`, `Message` fields.
- **Email Template Engine**: `PM_EmailTemplates` list with 15 seeded templates, merge tags (`{{PolicyTitle}}`, `{{AuthorName}}`, etc.), 5-min cache, fallback to hardcoded templates. `PolicyNotificationService.queueTemplatedEmail()` is the main entry point.
- **Branded HTML emails**: All emails use Forest Teal header gradient, policy title card, CTA button, First Digital footer via `buildEmailShell()`.

### 2. Review Mode (mode=review)
- **URL**: `PolicyDetails.aspx?policyId=X&mode=review`
- **UI**: Teal banner "Review Mode", policy content left panel, decision panel right (Approve/Request Changes/Reject), review checklist, review chain, previous comments
- **Actions**: Updates PM_PolicyReviewers status, updates PM_Policies PolicyStatus, writes audit log, sends branded email to author
- **After reviewer approves**: If final approvers exist → policy moves to "Pending Approval". If no approvers → moves to "Approved"
- **After reject/changes**: Policy → "Draft", ALL reviewer statuses reset to Pending

### 3. Approval Mode (mode=approve)
- **URL**: `PolicyDetails.aspx?policyId=X&mode=approve`
- **UI**: Green/emerald banner "Approval Mode", same structure as Review Mode but different branding
- **Decisions**: "Approve for Publication" / "Return to Author" / "Reject"
- **After approver approves**: Policy → "Approved" (ready to Publish)
- **After return/reject**: Policy → "Draft", ALL statuses reset

### 4. Fast Track Wizard Mode
- **Mode Selection**: First screen shows two cards: Fast Track (4 steps) vs Standard Wizard (8 steps)
- **Fast Track Steps**: Template Selection → Policy Details (pre-filled with override) → Content → Review & Submit
- **Templates loaded from**: `PM_PolicyMetadataProfiles` list (10 seeded via script 23)
- **Validation**: Separate `validateFastTrackStep()` method for 4-step indices
- **FAST_TRACK_STEPS** defined in `src/models/IPolicyAuthor.ts`

### 5. Audience System
- **AudienceRuleService**: `src/services/AudienceRuleService.ts` — rule evaluation engine
- **PM_Audiences list**: 9 seeded system audiences (All Employees, All Managers, departments, etc.)
- **Rule format**: JSON array `[{"field":"Department","operator":"equals","value":"Sales"}]` with AND/OR combinator
- **Wizard Step 4**: Loads audience cards from PM_Audiences, click to select, shows estimated user count
- **Admin Centre UI**: NOT YET BUILT (backlog item)

### 6. Field Name Audit (20+ fixes)
Major field mismatches fixed across the entire codebase:
- `Action` → `AuditAction` (PM_PolicyAuditLog) — was causing ALL audit logs to have null action data
- `NotificationType` → `Type` (PM_Notifications)
- `LinkUrl` → `ActionUrl` (PM_Notifications)
- `Description` → `PolicyDescription` (PM_Policies)
- `ReviewerIds` removed from PM_Policies (doesn't exist)
- `Status` → `ExemptionStatus` (PM_PolicyExemptions)
- `UserId` → `AckUserId` (PM_PolicyAcknowledgements)
- All `/sites/JML/` URLs → `/sites/PolicyManager/`
- `PM_EmailQueue` → `PM_NotificationQueue` in NotificationRouter
- `Body` → `Message` in NotificationRouter

### 7. Lifecycle Audit Fixes
- **CRITICAL-5**: Added Publish button to pipeline for Approved policies
- **CRITICAL-4**: `updateApprovalStatus` now writes to SharePoint (was client-only setState)
- **CRITICAL-2+3**: Fixed `handleSubmitForReviewFromKanban` — wrong method, field, status
- **HIGH-2**: Audience validation allows `targetAllEmployees` OR `selectedAudienceId`
- **HIGH-6**: Reviewer statuses reset on rejection
- **HIGH-8**: Fast Track validation with separate method
- **MEDIUM-3**: `handleSaveDraft` no longer overwrites non-Draft PolicyStatus

### 8. Pipeline Enhancements
- KPI cards with arrow progression (workflow visual)
- Inline workflow status indicator (dot-line per row)
- Default view = Drafts, "All" moved to last
- Action icons: Edit, View, Publish, Submit, Duplicate, Withdraw, Delete, Revision, Retire
- "New Policy" as first item in Author dropdown

### 9. Global Form Styling
- **3-layer approach**: SCSS `:global` + MutationObserver + fallback CSS `<style>` tag
- All Fluent UI controls: `border-radius: 6px`, `#d1d5db` borders, teal focus ring
- Dropdown callouts: rounded, subtle shadow, teal selected state
- Focus outlines: blue → teal everywhere
- **File**: `src/styles/global-dropdown-fix.scss` + `src/utils/injectPortalStyles.ts`

### 10. Distribution Page
- Action buttons moved inline with progress bar
- Stat cards with left accent borders (colour per stat type)
- Pause/Resume handler (toggles campaign status)
- Send Reminder handler (queues branded emails to pending recipients)

---

## Current Architecture

### Wizard Step Order (Standard — 8 steps)
```
0: Creation Method (Word/Excel/PPT/HTML/Infographic/Upload)
1: Basic Information (name, category, department, summary, owner)
2: Compliance & Risk (metadata profile selection)
3: Audience (PM_Audiences card grid)
4: Effective Dates (dates, review frequency, supersedes)
5: Review Workflow (reviewers + approvers via PeoplePicker)
6: Policy Content (rich text or linked document)
7: Review & Submit (accordion summary)
```

### Wizard Step Order (Fast Track — 4 steps)
```
0: Fast Track Template (PM_PolicyMetadataProfiles card grid)
1: Policy Details (name, summary, owner + pre-filled overrides)
2: Policy Content (same as Standard step 6)
3: Review & Submit (same as Standard step 7)
```

### Policy Status Lifecycle
```
Draft → In Review → Pending Approval → Approved → Published → Archived/Retired
                  ↓                   ↓
             Rejected            Returned
              (→ Draft)           (→ Draft)
```

### Notification Flow
```
SPFx Action → PM_NotificationQueue (QueueStatus: 'Pending') → Logic App → Office 365 → Outlook
                                                              ↓
                                                         QueueStatus: 'Processing' → 'Sent'/'Failed'
```

### Key Lists
| List | Purpose |
|------|---------|
| PM_Policies | Core policy records |
| PM_PolicyReviewers | Reviewer/approver assignments per policy |
| PM_PolicyAuditLog | Audit trail (AuditAction field) |
| PM_NotificationQueue | Email queue (QueueStatus field — NOT Status!) |
| PM_Notifications | In-app notifications (Type field — NOT NotificationType!) |
| PM_EmailTemplates | Customisable email templates with merge tags |
| PM_Audiences | Audience rule definitions (JSON rules) |
| PM_UserProfiles | Synced user data for audience evaluation |
| PM_PolicyMetadataProfiles | Fast Track templates / metadata presets |
| PM_ApprovalDelegations | Delegation assignments (DelegatedById/DelegatedToId) |

---

## Key Files Modified

| File | What Changed |
|------|-------------|
| `PolicyAuthorEnhanced.tsx` | Wizard: Fast Track mode, step reorder, creation method strip, audience cards, validation |
| `PolicyAuthorView.tsx` | Pipeline: KPI arrows, action buttons, Publish/Retire/Revision, approvals tab fix, delegation panel |
| `PolicyDetails.tsx` | Review Mode + Approval Mode (two separate render methods) |
| `PolicyService.ts` | `logAudit()` fix (AuditAction), `submitForReview` fix, `rejectPolicy` fix |
| `PolicyNotificationService.ts` | Email template engine, `queueTemplatedEmail()`, `loadEmailTemplate()` |
| `ApprovalService.ts` | Field fixes (Type, ActionUrl), JML URL removal |
| `NotificationRouter.ts` | PM_EmailQueue → PM_NotificationQueue, field fixes |
| `PolicyManagerHeader.tsx` | Nav reorder, "New Policy" in Author dropdown, duplicate removal |
| `PolicyDistribution.tsx` | Action buttons inline, stat card accents, Pause/Resume/Send Reminder |
| `AudienceRuleService.ts` | NEW — rule evaluation engine |
| `global-dropdown-fix.scss` | Global Forest Teal form control styling |
| `injectPortalStyles.ts` | Portal styles updated to Forest Teal |
| `IPolicyAuthor.ts` | FAST_TRACK_STEPS, WizardStep type union, step title rename |

---

## Provisioning Scripts Run This Session

| Script | What It Does |
|--------|-------------|
| `20-NotificationChoiceUpdate.ps1` | Added approval/review types to PM_Notifications and PM_NotificationQueue choice fields |
| `21-EmailTemplates-List.ps1` | Created PM_EmailTemplates + seeded 15 default templates |
| `22-Audiences-List.ps1` | Created PM_Audiences + enhanced PM_UserProfiles + seeded 9 system audiences |
| `23-Seed-FastTrackTemplates.ps1` | Seeded 10 Fast Track templates in PM_PolicyMetadataProfiles |

### Manual Column Fixes Run
```powershell
# PM_PolicyReviewers — PolicyId column was missing
Add-PnPField -List "PM_PolicyReviewers" -DisplayName "PolicyId" -InternalName "PolicyId" -Type Number -AddToDefaultView

# PM_NotificationQueue — Status renamed to QueueStatus (reserved name conflict)
Remove-PnPField -List "PM_NotificationQueue" -Identity "Status" -Force
Add-PnPField -List "PM_NotificationQueue" -DisplayName "Queue Status" -InternalName "QueueStatus" -Type Choice -Choices "Pending","Processing","Sent","Failed","Retry" -Required -AddToDefaultView
Set-PnPField -List "PM_NotificationQueue" -Identity "QueueStatus" -Values @{DefaultValue="Pending"; Indexed=$true}

# PM_UserProfiles — Email column
Add-PnPField -List "PM_UserProfiles" -DisplayName "Email" -InternalName "Email" -Type Text -AddToDefaultView
```

---

## Logic App Configuration

The Logic App `dwx-pm-email-sender-prod` must use this definition (already deployed):
- **List**: `PM_NotificationQueue`
- **Filter**: `QueueStatus eq 'Pending'`
- **Order**: `Priority desc, Created asc`
- **Fields**: `RecipientEmail` → To, `Title` → Subject, `Message` → Body
- **Status updates**: `QueueStatus: 'Processing'` → `QueueStatus: 'Sent'` or `QueueStatus: 'Failed'`

Full JSON definition was provided to the user and deployed. See Bicep template at `azure-functions/email-sender/infra/main.bicep` for the canonical definition.

---

## Backlog (Next Session)

1. **Audience Rule Builder UI** — Admin Centre CRUD for creating/editing audience rules
2. **Wire audience resolution into publish/distribution flow** — resolve audience → create PM_PolicyPackAssignments
3. **Save as Fast Track Template** — option on standard wizard Review step
4. **StyledSelect component** — custom dropdown replacing Fluent UI v8 dropdowns
5. **Submit for Review consistency check** — verify works from all paths
6. **PM_ReminderSchedule provisioning** — automated revision reminders
7. **Delete dead code** — PolicyWizard.tsx (~600 lines, never used)
8. **Archive view** — retired policies visible in pipeline

---

## Known Issues

1. **PM_PolicyDistributions.CampaignName** column doesn't exist — Distribution page shows error (pre-existing, not from this session)
2. **PolicyAuditService** has ~20 fields not provisioned on PM_PolicyAuditLog — enhanced audit writes silently drop data (basic audit via PolicyService.logAudit works fine)
3. **PM_ReminderSchedule** list not provisioned — automated reminders won't fire
4. **Retired policies** disappear from pipeline entirely — no archive view
5. **PeoplePicker in Distribution** shows "No results found" — may need to search PM_UserProfiles instead of Entra ID

---

## Commits (Session 18)

| Hash | Tag | Summary |
|------|-----|---------|
| `bcbaaca` | [feat][fix][refactor] | Email pipeline, 20+ field mismatches, wizard redesign |
| `f1d5091` | [docs] | CLAUDE.md checkpoint |
| `fe61fed` | [feat][fix] | Review Mode, Email Templates, Audiences, Fast Track mockup |
| `f7e5f49` | [feat][fix] | Lifecycle audit fixes, Fast Track, pipeline, PeoplePicker audit |
| `b3ff6be` | [feat][fix][refactor] | Global form styling, Distribution actions, dropdown fixes |
| `683a7ac` | [docs] | CLAUDE.md final, border-radius 6px |
| `487bb5d` | [fix][feat] | QueueStatus rename, Pending Approval flow, email URL fixes |
| `3835115` | [feat] | Approval Mode — differentiated UI for final approvers |

---

## Demo Flow (End-to-End)

1. **Author** → New Policy → Fast Track → Select "IT Security Policy" template → Enter name/summary → Create Word doc → Review & Submit → Submit for Review
2. **Reviewer** receives styled email → clicks "Review Policy" → sees **Review Mode** (teal) → Approves with comments
3. Policy moves to **Pending Approval** → **Approver** receives email → clicks "Review & Approve Policy" → sees **Approval Mode** (green/emerald) → Grants approval
4. Policy moves to **Approved** → Author sees Publish icon in pipeline → clicks Publish → Status → **Published**
5. Published email sent → Policy appears in Policy Hub for all users

---

## Critical Patterns to Remember

- **QueueStatus** not Status for PM_NotificationQueue (SharePoint reserved name)
- **AuditAction** not Action for PM_PolicyAuditLog
- **Type** not NotificationType for PM_Notifications
- **escapeHtml()** mandatory for all user content in email templates
- **mode=review** for reviewer emails, **mode=approve** for approver emails
- **saveReviewers** deletes all existing then re-adds — if re-add fails, reviewers are wiped
- **PolicyStatus: 'Draft'** only set on NEW policies in handleSaveDraft (not on updates)
- **3-layer CSS**: SCSS :global + MutationObserver + fallback `<style>` tag for portal elements
