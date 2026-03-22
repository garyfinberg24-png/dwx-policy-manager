# Policy Manager — Production Readiness Checklist

Read CLAUDE.md and all memory files first.

This is DWx Policy Manager — a 17-session SPFx build that is now functionally complete but needs production hardening. The app is deployed at:
- SharePoint: https://mf7m.sharepoint.com/sites/PolicyManager
- Azure Functions: https://dwx-pm-quiz-func-prod.azurewebsites.net, https://dwx-pm-chat-func-prod.azurewebsites.net, https://dwx-pm-email-sender-prod (Logic App)
- Repo: https://dev.azure.com/gfinberg/DWx/_git/dwx-policy-manager + GitHub mirror

DO NOT build new features. DO NOT do broad sweeps.

Instead, work through this checklist ONE ITEM AT A TIME, depth-first.
For each item: test it, fix it, verify the fix, then move to the next.
Package and deploy after every batch of fixes so I can test live.

## MANDATORY RULES

### Rule 1: Border Radius Standards
All controls such as dropdowns, filters, search fields, text input boxes, buttons etc. ALWAYS have border-radius of 4px. KPI cards are generally 8px. Enforce this consistently across all pages. If you find any control that deviates, fix it.

### Rule 2: No Stubs — Full Implementation Required
If you identify something that is a stub, is marked for later implementation, or is not fully implemented yet, YOU MUST IMPLEMENT IT FULLY. DO NOT MARK "FOR LATER". The goal is to go deep and focused and get the app ready for production. LEAVE NOTHING UNIMPLEMENTED. Every button must work. Every form must save. Every feature must function end-to-end.

### Rule 3: Do Not Touch Fluent Panels
DO NOT MAKE ANY CHANGES TO ANY FLUENT PANEL. If there are any issues with panels (StyledPanel, Panel, slide-in panels), note them and log them for user attention. We will review panel issues manually later. This includes panel headers, panel content, panel footers, panel styling.

## PRODUCTION READINESS CHECKLIST

### Phase 1: Core Policy CRUD Flows (test each end-to-end)
1. Create a new policy via Policy Builder wizard → verify it appears in PM_Policies list
2. Save as Draft → verify PolicyStatus = 'Draft' and all fields persist (including audience, dates, review frequency, key points, owner)
3. Edit existing Draft policy → verify all fields load correctly (audience, dates, key points, quiz link, reviewers)
4. Submit for Review → verify status changes to 'In Review', notification created in PM_Notifications, audit log entry in PM_PolicyAuditLog
5. View policy in Policy Hub card view → verify all badges, category strip, version chip display
6. View policy in Policy Hub list view → verify all columns, row click opens StyledPanel
7. Open policy in Simple Reader (browse mode) → verify content renders (HTML and PDF), toolbar works (Download, Print, Fullscreen), "Back to Policy Hub" button works
8. Open policy from My Policies (assigned read mode) → verify full read flow: Read → Quiz → Acknowledge → Complete with certificate
9. Complete acknowledgement → verify PM_PolicyAcknowledgements record created, "Return to My Policies" button goes to MyPolicies.aspx
10. Create a Policy Pack → verify it appears in PM_PolicyPacks list
11. Assign a Policy Pack → verify PM_PolicyPackAssignments record created
12. Create a Distribution Campaign → verify PM_PolicyDistributions record created
13. Version a policy → verify PM_PolicyVersions record, version number bumps
14. Create a quiz via Quiz Builder → verify PM_PolicyQuizzes and PM_PolicyQuizQuestions records
15. AI Generate quiz questions → verify Azure Function call returns structured JSON
16. Take a quiz → verify PM_PolicyQuizResults record with score

### Phase 2: Azure Functions (test each endpoint)
17. POST /api/generate-quiz-questions with policy text → verify JSON response with questions
18. POST /api/policyChatCompletion with policy-qa mode → verify AI response with citations
19. POST /api/policyChatCompletion with author-assist mode → verify writing coach response
20. POST /api/policyChatCompletion with general-help mode → verify app navigation response
21. Verify Logic App email pipeline → queue email to PM_EmailQueue → verify Logic App processes → verify status changes to Sent
22. POST /api/convertDocument with .docx file → verify HTML conversion returned

### Phase 3: Admin Centre (test each section saves/loads)
23. General Settings → change company name → save → reload → verify persisted in PM_Configuration
24. Templates → create template → save → verify in PM_PolicyTemplates
25. Metadata Profiles → create profile → save → verify in PM_PolicyMetadataProfiles
26. Approval Workflows → edit settings → save → verify
27. Compliance Settings → edit defaults → save → verify in PM_Configuration
28. Notifications → toggle settings → save → verify in PM_Configuration
29. Naming Rules → create rule → save → verify in PM_NamingRules
30. SLA Targets → edit values → save → verify in PM_SLAConfigs
31. Data Lifecycle → edit retention → save → verify
32. Navigation → toggle nav items → save → verify items hide/show in header
33. Reviewers & Approvers → verify group management works
34. Audit Log → verify entries load with filters, search works
35. Data Export → generate CSV export → verify file downloads
36. Provisioning → verify all lists show status and item counts
37. AI Assistant → enable/disable → save → verify chat panel responds

### Phase 4: Manager Features (test each view)
38. Manager Dashboard → verify KPIs populate from live data, compliance ring works
39. Team Compliance → verify team members load from PM_PolicyAcknowledgements, grouped by user
40. Approvals → verify approval cards load from PM_Approvals, Approve/Return buttons work
41. Delegations → verify delegation cards load from PM_ApprovalDelegations, create/edit/delete work
42. Policy Reviews → verify review schedule loads from PM_Policies (NextReviewDate), status badges correct
43. Reports → Report Hub → click Generate → verify CSV downloads with live data
44. Reports → Report Builder → select parameters → Generate → verify export works
45. Reports → Schedule → create schedule → verify PM_ScheduledReports record, edit/delete/toggle work
46. Reports → Analytics → verify KPIs from PM_ReportExecutions, quick reports work, timeline shows history
47. Distribution → create campaign → verify cards render, progress bars update

### Phase 5: Navigation & UI Polish
48. Every nav item links to correct SharePoint page
49. Every breadcrumb links correctly
50. Every StyledPanel opens/closes correctly with X and light dismiss (DO NOT modify panels — log issues only per Rule 3)
51. Every dropdown in header (Author, Manager, Secure Policies) opens/closes correctly
52. Help icon → navigates to PolicyHelp.aspx (not panel)
53. Search in header → opens PolicySearch.aspx
54. Settings cog → opens PolicyAdmin.aspx (respects role — Manager+ only)
55. "+ New Policy" button → opens PolicyBuilder.aspx
56. Policy Hub Featured Policy → accordion expand/collapse works
57. Policy Hub facet sidebar → checkbox filtering works (Category, Risk, Department)
58. My Policies hero → compliance ring, greeting, KPI cards display correctly
59. All hero banners consistent (Policy Hub, Help, My Policies)
60. Footer → teal gradient matching header, links work
61. Analytics → pill tabs switch correctly, no visual clash with KPI cards
62. Border radius audit → all controls (dropdowns, search fields, buttons, inputs) use 4px border-radius per Rule 1

### Phase 6: Error Handling & Edge Cases
63. What happens when a SP list is empty? (each page should show empty state, not crash)
64. What happens when PM_Policies has no Published policies? (Policy Hub should show empty state)
65. What happens when user has no acknowledgements? (My Policies should show "All caught up")
66. What happens when Azure Function is down? (Quiz Builder and AI Chat should show graceful error)
67. What happens with special characters in policy name? (quotes, ampersands, angle brackets)
68. What happens when PolicyContent AND HTMLContent are both empty? (Simple Reader should show "No content" message)
69. What happens when a policy has no DocumentURL? (Reader should not crash)
70. What happens when the user doesn't have Manager role? (Manager pages should show access denied)
71. What happens when PM_Configuration list doesn't exist? (Admin should degrade gracefully)

### Phase 7: Security & Data Integrity
72. HTML content sanitised via sanitizeHtml() before rendering (no XSS)
73. Email templates use escapeHtml() for all user-controlled content
74. OData filters use sanitizeForOData() — no injection
75. Admin-only pages check userRole === 'Admin' in componentDidMount
76. Upload validation: file extension + MIME cross-validation via ALLOWED_MIME_MAP
77. Document size limits enforced (25MB doc, 100MB video)
78. JSON.parse of external data wrapped in try/catch
79. localStorage/sessionStorage URLs validated (protocol check)

For each item, report: PASS / FAIL / PARTIAL
If FAIL: fix it immediately, don't continue until fixed.
If a stub or incomplete implementation is found: implement it fully per Rule 2.
If a panel issue is found: log it for manual review per Rule 3.
