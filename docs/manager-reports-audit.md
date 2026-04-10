# Manager > Reports Tab — Deep Audit & Assessment

**Date**: 10 April 2026  
**Scope**: PolicyManagerView.tsx Reports tab (3 sub-tabs), supporting services, SharePoint lists, provisioning scripts  
**Verdict**: **Partially Implemented** — UI shell is polished; backend plumbing has critical gaps

---

## 1. Executive Summary

The Manager Reports tab provides a visually polished, 3-sub-tab reporting interface (Report Hub, Report Builder, Reports Analytics). Scheduling CRUD against SharePoint is fully wired. However, several foundational capabilities are incomplete:

| Area | Status | Score |
|------|--------|-------|
| **UI / Layout** | Complete | 9/10 |
| **Report Hub (browse & filter)** | Functional | 8/10 |
| **Report Scheduling** | Fully Functional | 9/10 |
| **Report Generation (CSV download)** | Working (CSV only) | 6/10 |
| **Report Builder Parameters** | Decorative only | 2/10 |
| **Report Preview** | Hardcoded mock data | 1/10 |
| **PDF / Excel Export** | Not implemented | 0/10 |
| **Email Distribution** | Not implemented | 0/10 |
| **Scheduled Execution (backend)** | Not implemented | 0/10 |
| **Service Utilisation** | Severely underutilised | 2/10 |

**Overall: 4/10 for functional completeness**

---

## 2. Architecture Overview

```
PolicyManagerView.tsx (2,823 lines)
├── Reports Tab (renderReportsTab)
│   ├── Report Hub (renderReportHub)
│   ├── Report Builder (renderReportBuilder)
│   └── Reports Analytics (renderExecDashboard)
│
├── Report Flyout Panel (renderReportFlyout)
├── Schedule Panel (renderSchedulePanel)
│
├── Data Loading
│   ├── loadReportExecutions() → PM_ReportExecutions (real SP)
│   └── loadScheduledReports() → PM_ScheduledReports (real SP)
│
├── Actions
│   ├── handleGenerateReport() → PolicyReportExportService → CSV download + log to SP
│   ├── handleSaveSchedule() → PM_ScheduledReports CRUD (real SP)
│   ├── handleDeleteSchedule() → PM_ScheduledReports (real SP)
│   └── handleToggleSchedule() → PM_ScheduledReports (real SP)
│
└── Services Used
    └── PolicyReportExportService (CSV export only)

Services Available but NOT Used:
├── ReportDefinitionService (full CRUD for PM_ReportDefinitions)
├── ReportNarrativeService (AI-powered report narratives)
├── ScheduledReportService (scheduling logic)
├── AnalyticsService (generic analytics)
├── PolicyAnalyticsService (KPI calculations)
├── ProcessAnalyticsService (process metrics)
├── PredictiveAnalyticsService (forecasting)
├── ROIAnalyticsService (ROI calculations)
├── ROIExportService (ROI export)
├── WorkflowAnalyticsService (workflow metrics)
└── ReportHtmlGenerator (imported but never called)
```

---

## 3. What Works (Fully Functional)

### 3.1 Report Hub — Browse & Filter
- **8 report cards** displayed in a responsive grid
- **Search box** filters by title/description (real-time)
- **Category pill filters**: All, Compliance, Acknowledgement, SLA, Audit, Delegation, Training
- **Each card** shows: icon, title, description, format badge, last generated date
- **"Generate" button** triggers CSV download via PolicyReportExportService
- **"Schedule" button** opens scheduling panel
- **"View Details" button** opens flyout panel

### 3.2 Report Scheduling (CRUD)
- **Create**: Frequency (Daily/Weekly/Monthly/Quarterly), Format (PDF/Excel/CSV), Recipients, Active toggle
- **Edit**: Opens panel pre-populated with existing values
- **Delete**: Confirmation dialog → SP delete
- **Toggle**: Enable/disable schedule
- **NextRun calculation**: Correctly computed from frequency
- **Persistence**: Full CRUD against `PM_ScheduledReports` SharePoint list
- **PeoplePicker**: Recipients selection (though stored as text, not resolved user IDs)

### 3.3 Report Execution Logging
- Every "Generate" click logs to `PM_ReportExecutions` with:
  - ReportName, ReportType, GeneratedByName, GeneratedByEmail
  - Format, RecordCount, FileSize, ExecutionTime, ExecutionStatus, ExecutedAt
- **Recent executions** displayed in Report Builder and Reports Analytics sub-tabs
- **Timeline view** in Reports Analytics shows execution history

### 3.4 CSV Export (via PolicyReportExportService)
The following export methods work and download real CSV files:

| Report Key | Service Method | Data Source | Status |
|------------|---------------|-------------|--------|
| `dept-compliance` | `exportComplianceSummary({ groupBy: 'department' })` | PM_Policies + PM_PolicyAcknowledgements | **Working** |
| `ack-status` | `exportAcknowledgementStatus({})` | PM_PolicyAcknowledgements | **Working** |
| `sla-performance` | `exportExecutiveSummary()` | PM_Policies + PM_PolicyAcknowledgements | **Working** |
| `risk-violations` | `exportOverdueReport()` | PM_PolicyAcknowledgements | **Working** |
| `training-completion` | `exportQuizResults()` | PM_PolicyQuizResults | **Working** |
| `audit-trail` | `exportPolicyInventory({})` | PM_Policies | **WRONG** — should export from PM_PolicyAuditLog |
| `delegation-summary` | `exportPolicyInventory({})` | PM_Policies | **WRONG** — should export from PM_ApprovalDelegations |
| `review-schedule` | `exportPolicyInventory({})` | PM_Policies | **WRONG** — should export from PM_Policies filtered by NextReviewDate |

---

## 4. What Is Broken or Incorrect

### 4.1 CRITICAL: 3 Report Types Map to Wrong Export Method
Lines 770-779 in PolicyManagerView.tsx:
```typescript
case 'audit-trail':
  result = await this.reportExportService.exportPolicyInventory({}); // WRONG
case 'delegation-summary':
  result = await this.reportExportService.exportPolicyInventory({}); // WRONG
case 'review-schedule':
  result = await this.reportExportService.exportPolicyInventory({}); // WRONG
```

- **`audit-trail`** should export from `PM_PolicyAuditLog` — the service has no `exportAuditTrail()` method
- **`delegation-summary`** should export from `PM_ApprovalDelegations` — the service has no `exportDelegations()` method
- **`review-schedule`** should export from `PM_Policies` filtered by `NextReviewDate` — the service has no `exportReviewSchedule()` method

All three fall through to `exportPolicyInventory({})`, which exports the generic policy list. **Users get the wrong data.**

### 4.2 CRITICAL: Report Builder Parameters Are Decorative
The Report Builder sub-tab has parameter inputs that are **never passed** to the export service:

| Parameter | State Variable | Passed to Export? |
|-----------|---------------|-------------------|
| Date Range Start | `builderDateStart` | **NO** |
| Date Range End | `builderDateEnd` | **NO** |
| Departments | `builderDepartments` | **NO** |
| Output Format | `builderFormat` | **YES** (format only) |
| Include summary charts | Uncontrolled checkbox | **NO** (not tracked in state) |
| Include individual breakdown | Uncontrolled checkbox | **NO** (not tracked in state) |
| Include historical comparison | Uncontrolled checkbox | **NO** (not tracked in state) |
| Include risk assessment | Uncontrolled checkbox | **NO** (not tracked in state) |

The checkboxes use `defaultChecked` with no `onChange` handler — their values are never read.

Despite `PolicyReportExportService` accepting `dateRangeStart`, `dateRangeEnd`, and `departments` in its option interfaces, the Manager View never passes them.

### 4.3 CRITICAL: PDF and Excel Export Not Implemented
- Report cards advertise PDF and Excel formats via format badges
- The format dropdown in Report Builder offers PDF/Excel/CSV options
- **Only CSV export is implemented** in `PolicyReportExportService`
- There is no `downloadPDF()` or `downloadExcel()` method
- `ReportHtmlGenerator` is imported but **never called** — it generates premium HTML suitable for PDF conversion but isn't wired
- No library (jsPDF, xlsx, etc.) is installed for PDF/Excel generation

### 4.4 HIGH: "Email Report" Button Opens Schedule Panel Instead
Line 2118-2119:
```typescript
<DefaultButton text="Email Report" iconProps={{ iconName: 'Mail' }}
  onClick={() => this.openSchedulePanel(selectedReport.key, selectedReport.title)} />
```
The "Email Report" button opens the scheduling panel — it should open a panel to email the report immediately to selected recipients.

### 4.5 HIGH: Report Flyout Shows Hardcoded Mock Data
The flyout panel (renderReportFlyout, lines 2452-2551) displays:
- 3 hardcoded stat cards: Compliance Rate 87.3%, Team Members 8, Pending Items 12
- 5 hardcoded sample rows (Thabo Mokoena, Lindiwe Nkosi, etc.)

These never change regardless of which report is selected.

### 4.6 MEDIUM: Report Preview Is Entirely Mock
The Report Builder preview section (lines 2131-2177) shows:
- 4 hardcoded KPI values (87.3%, 8, 24, 12)
- 5 hardcoded department rows with static numbers

Preview should run a limited query against real data using the selected parameters.

### 4.7 MEDIUM: lastGenerated Dates Are Hardcoded
All 8 report cards have hardcoded `lastGenerated` values (23-30 Jan 2026). These should be derived from `PM_ReportExecutions` data.

### 4.8 LOW: Department List Is Hardcoded
The Report Builder department multi-select dropdown (lines 2055-2064) has 8 hardcoded departments. Should be loaded from PM_Policies distinct `DepartmentOwner` values or a configuration list.

---

## 5. What Is Not Implemented

### 5.1 Scheduled Report Execution (Backend)
The scheduling UI creates/edits/deletes schedule records in `PM_ScheduledReports`, but there is **no backend process** that:
- Polls `PM_ScheduledReports` for items where `NextRun <= now() AND Enabled = true`
- Executes the report generation
- Emails the result to recipients
- Updates `LastRun` and recalculates `NextRun`

This requires an **Azure Function or Logic App** timer trigger — similar to the existing `dwx-pm-email-sender-prod` pattern.

### 5.2 Report Email Distribution
No mechanism exists to:
- Generate a report and attach it to an email
- Send a one-time email with a report to selected recipients
- Process scheduled report deliveries

### 5.3 Custom Report Builder Persistence
The Report Builder lets users configure parameters but:
- Cannot save a custom report configuration for reuse
- Cannot share a saved report with other managers
- `PM_ReportDefinitions` list exists and `ReportDefinitionService` has full CRUD, but neither is used

### 5.4 Real-Time Report Preview
The "Preview" button sets `showReportPreview: true` but renders hardcoded data. Should run a limited SP query using selected parameters and display actual results.

### 5.5 Report Download from History
The execution timeline shows past report runs, but:
- No generated file is stored anywhere (no document library, no blob)
- "Re-generate" button re-runs the export (new download, not the original)
- No download link for previously generated reports

---

## 6. Unused Services Inventory

These services exist in `src/services/` and are **never imported by any webpart**:

| Service | Lines | Purpose | Could Power |
|---------|-------|---------|-------------|
| `ReportDefinitionService` | ~310 | CRUD for PM_ReportDefinitions | Custom report builder, saved reports, report templates |
| `ReportNarrativeService` | ~100 | AI-generated report narratives via Azure OpenAI | Executive summary text, natural language insights |
| `ScheduledReportService` | ~150 | Scheduling logic, execution tracking | Backend scheduled execution |
| `AnalyticsService` | ~200 | Generic analytics queries | Dashboard KPIs with real data |
| `PolicyAnalyticsService` | ~300 | Policy-specific KPI calculations | Real compliance rates, trends |
| `ProcessAnalyticsService` | ~200 | Process metrics (cycle times, bottlenecks) | SLA performance report |
| `PredictiveAnalyticsService` | ~250 | Forecasting, trend prediction | Compliance trend forecasting |
| `ROIAnalyticsService` | ~350 | ROI calculations | ROI report (not in current 8 reports) |
| `ROIExportService` | ~150 | ROI export | ROI CSV/PDF export |
| `WorkflowAnalyticsService` | ~400 | Workflow metrics | Approval turnaround report |

**Total unused code: ~2,410 lines of report/analytics services.**

Additionally, `ReportHtmlGenerator` (841 lines) is **imported** in PolicyManagerView.tsx but **never called**. It generates premium branded HTML suitable for PDF conversion.

---

## 7. Overlap & Duplication Analysis

### 7.1 Manager Reports vs. Analytics Webpart (jmlPolicyAnalytics)

| Capability | Manager Reports | Analytics Webpart | Duplication? |
|-----------|----------------|-------------------|-------------|
| Compliance overview | Hardcoded mock | 7 real tabs (Executive, Metrics, Acks, SLA, Compliance, Audit, Quiz) | Analytics is superior |
| SLA tracking | One report card (CSV) | Full SLA tab with breach table | Analytics is superior |
| Audit export | Falls through to policy inventory | Full audit tab with SP queries | Analytics is superior |
| Acknowledgement tracking | One report card (CSV) | Full ack tab with department breakdown | Analytics is superior |
| Quiz analytics | One report card (CSV) | Full quiz tab | Analytics is superior |
| Scheduling | Full CRUD | None | Manager only |
| Export/download | CSV (6 types) | None | Manager only |

**Finding**: The Analytics webpart has **far richer dashboards** for the same data domains. The Manager Reports tab adds value primarily through **export/download** and **scheduling** — capabilities the Analytics webpart lacks.

### 7.2 Manager Reports vs. Author Reports (dwxPolicyAuthorReports)

| Capability | Manager Reports | Author Reports | Overlap? |
|-----------|----------------|----------------|----------|
| Scope | All policies (team/org) | Author's own policies only | Different scope |
| Ack tracking | CSV export | Per-policy ack table with progress bars | Author has richer UI |
| Policy lifecycle | Not implemented | Visual pipeline (Draft→Published→Retired) | Author only |
| Review schedule | Falls through to policy inventory | Grouped by urgency (Overdue/Due Soon/Upcoming) | Author has richer UI |
| Activity history | Not implemented | Timeline with colour-coded action badges | Author only |
| Data source | PolicyReportExportService | Direct SP queries | Both real |
| Export | CSV download | No export | Manager only |

**Finding**: Author Reports has **richer per-policy analytics** with real data. Manager Reports has **broader scope** but less depth. No direct code duplication.

---

## 8. SharePoint List Assessment

### Lists Used

| List | Provisioned? | Script | Used Correctly? |
|------|-------------|--------|----------------|
| `PM_ScheduledReports` | Yes | 17-ReportingLists.ps1 | Yes — full CRUD |
| `PM_ReportExecutions` | Yes | 17-ReportingLists.ps1 | Yes — logging works |
| `PM_ReportDefinitions` | Yes | 17-ReportingLists.ps1 | **No — never read by any webpart** |
| `PM_Policies` | Yes | Core provisioning | Yes |
| `PM_PolicyAcknowledgements` | Yes | Core provisioning | Yes |
| `PM_PolicyQuizResults` | Yes | Quiz provisioning | Yes |

### Lists That Should Be Used But Aren't

| List | Should Power | Currently |
|------|-------------|-----------|
| `PM_PolicyAuditLog` | Audit Trail Export report | Falls through to policy inventory |
| `PM_ApprovalDelegations` | Delegation Summary report | Falls through to policy inventory |
| `PM_ReportDefinitions` | Custom report builder persistence | Never queried |

---

## 9. Code Quality Assessment

### Positives
- Clean error handling with `_isMounted` guards
- LoggingService used (no console.log in PolicyManagerView)
- Proper try/catch around all SP operations
- Non-blocking log failure for execution recording (line 811)
- Responsive grid layout
- Consistent Forest Teal theming
- Dialog confirmation before schedule deletion

### Issues
- **ReportHtmlGenerator imported but unused** — dead import (line 44)
- **ReportDefinitionService has 14 console.log statements** — development debugging left in
- **Uncontrolled checkboxes** in Report Builder — `defaultChecked` with no state tracking
- **`builderDateStart`/`builderDateEnd` typed as `string` in state** (line 172-173) but assigned `Date` objects from DatePicker — type mismatch
- **No OData sanitisation** on department filter values passed to export service (though currently hardcoded, this becomes a risk if made dynamic)
- **Recipients stored as raw text** in PM_ScheduledReports — should use resolved user IDs for reliability

---

## 10. Security & Compliance Notes

- **Role gate**: Manager role required at component entry — correct
- **No PII exposure**: Report execution logs user name/email (should use ID instead for GDPR)
- **CSV injection risk**: `downloadCSV()` wraps values containing commas/quotes in escaped quotes, but does not prefix `=`, `+`, `-`, `@` characters that can trigger formula injection in Excel — **medium risk**
- **No audit trail for report downloads**: Report generation is logged, but there's no compliance audit of who downloaded what data (important for policy governance applications)

---

## 11. Opportunities for Improvement

### Quick Wins (Low Effort, High Value)

1. **Wire date range & department parameters** to `handleGenerateReport` → pass to export service options
2. **Track checkbox state** — add state variables for the 4 "Include in Report" options
3. **Fix 3 wrong report mappings** — create `exportAuditTrail()`, `exportDelegationSummary()`, `exportReviewSchedule()` methods in PolicyReportExportService
4. **Derive lastGenerated from PM_ReportExecutions** — match report type to latest execution
5. **Remove dead import** of ReportHtmlGenerator (or wire it up for PDF generation)
6. **Remove console.log statements** from ReportDefinitionService (14 occurrences)

### Medium Effort

7. **Wire ReportDefinitionService** to Report Builder for saved report configurations
8. **Real-time preview** — run limited SP query (top 10) using selected parameters
9. **Load departments dynamically** from PM_Policies distinct DepartmentOwner values
10. **PDF generation** — use ReportHtmlGenerator + Blob URL approach (similar to existing print pattern)
11. **Excel generation** — evaluate lightweight xlsx library or CSV-to-Excel conversion
12. **Fix "Email Report" button** — create a one-time email panel with PeoplePicker + format selector

### High Effort (Strategic)

13. **Backend scheduled execution** — Azure Function timer trigger that polls PM_ScheduledReports, generates reports, emails via PM_NotificationQueue
14. **Report storage** — Save generated reports to PM_PolicySourceDocuments or a dedicated document library for later download
15. **Wire analytics services** — Connect PolicyAnalyticsService, ProcessAnalyticsService to provide real KPIs in preview/dashboard
16. **AI narrative generation** — Use ReportNarrativeService to add natural language summaries to exported reports
17. **Consolidate with Analytics webpart** — Add export/schedule capability to Analytics tabs rather than duplicating dashboard views in Reports

---

## 12. Consolidation Recommendations

### Option A: Enhance Manager Reports (Current Path)
Keep the 3-sub-tab structure but fix all gaps. This preserves the current architecture but requires significant work to reach parity with the Analytics webpart's dashboard quality.

### Option B: Merge Export into Analytics (Recommended)
Add "Export" and "Schedule" buttons to each Analytics tab. Move Report Builder into Analytics as a 7th tab. This:
- Eliminates dashboard duplication
- Leverages Analytics' existing real-data queries
- Keeps the Manager Reports tab focused on scheduling and history
- Reduces the 2,823-line PolicyManagerView.tsx

### Option C: Standalone Reports Webpart
Extract Reports into its own webpart (`dwxPolicyReports`) with:
- ReportDefinitionService for custom reports
- Full parameter → export pipeline
- Backend execution via Azure Function
- Shared between Manager and Admin views

---

## 13. Summary Matrix

| Feature | Current State | What Exists | What's Missing |
|---------|--------------|-------------|----------------|
| Report Hub | Polished UI, 8 cards | Search, filter, generate, schedule | Real lastGenerated dates |
| Report Builder | UI complete | Date/dept/format inputs, checkboxes | Parameters not passed to service, preview is mock |
| Reports Analytics | KPI cards + tables | Real execution data | Quick reports are mock, no drill-down |
| CSV Export | 5/8 correct | downloadCSV + BOM | 3 reports map to wrong method |
| PDF Export | Not implemented | ReportHtmlGenerator exists (841 lines) | Not wired |
| Excel Export | Not implemented | Nothing | Library needed |
| Scheduling UI | Fully functional | Full CRUD on PM_ScheduledReports | N/A |
| Scheduled Execution | Not implemented | ScheduledReportService exists (unused) | Azure Function timer |
| Email Distribution | Not implemented | PM_NotificationQueue + Logic App exist | Wiring needed |
| Custom Reports | Not implemented | ReportDefinitionService + PM_ReportDefinitions exist | Wiring needed |
| Real Preview | Not implemented | Export service can query data | Wiring needed |

---

## 14. Risk Assessment

| Risk | Severity | Impact |
|------|----------|--------|
| Users click "Generate Audit Trail" and get policy inventory CSV | **HIGH** | Incorrect compliance data exported; potential regulatory issue |
| Users configure schedules that never execute | **HIGH** | False sense of automated reporting; no emails ever sent |
| Report Builder parameters ignored | **MEDIUM** | Users think they're filtering by date/department but get unfiltered data |
| PDF/Excel format badges imply capability that doesn't exist | **MEDIUM** | User expectation mismatch; all exports are CSV regardless of selection |
| Mock preview data could be mistaken for real metrics | **MEDIUM** | 87.3% compliance rate is hardcoded, not actual |
| CSV formula injection | **LOW** | Cells starting with `=`, `+`, `-`, `@` could execute in Excel |
| GeneratedByEmail in PM_ReportExecutions | **LOW** | PII in logs; should use user ID |

---

*End of audit.*
