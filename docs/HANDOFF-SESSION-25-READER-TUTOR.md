# PolicyIQ — Session 26 Handoff: Reader Tutor Feature

> **For:** New agent picking up this work
> **From:** Session 25 agent (16-18 Apr 2026)
> **Date:** 18 April 2026
> **Feature:** Reader Tutor Panel — AI-powered document comprehension for PolicyIQ
> **Priority:** This is the next major feature to build.

---

## 1. MANDATORY: Read These First

Before writing any code, read these files in this order:

1. **`CLAUDE.md`** — Project rules, architecture, conventions, session history. READ ALL OF IT.
2. **`docs/reader-tutor-brief-for-policyiq-contractiq.md`** — YOUR PRIMARY SPEC. This is the implementation brief scoped specifically for PolicyIQ. Follow it closely.
3. **`docs/dwx-document-tutor-pattern.md`** — The full DWx Document Tutor pattern specification. The brief references specific sections. Read §§1-3 at minimum.
4. **`docs/interactive-pdf-research-and-plan.md`** — LearnIQ's fuller research plan. Use for Azure architecture reference, NOT as a spec (PolicyIQ's scope is simpler).
5. **`docs/interactive-pdf-ux-mockup.html`** — THE definitive UX mockup. Open in a browser. The Reader Tutor panel must behave identically to this.
6. **`docs/dwx-interactive-ebook-concept.md`** — Context only. LearnIQ is building the full ebook version. You are building the simpler "agentic Q&A sidebar on existing documents" version.

---

## 2. What You Are Building (90-Second Summary)

A **right-docked panel (420px)** that sits alongside any PDF or Word document in the PolicyIQ reader. The panel has three tabs:

- **Chat** — Grounded Q&A with citations that jump to the exact page/section in the document
- **Study** — Auto-generated artefacts (summary, glossary, key sections, attestation questionnaire)
- **Progress** — Engagement score, reading time, attestation gate

Plus a **proactive tutor modal** that interrupts skim-reading with a comprehension check.

The key difference from the existing AI Chat in PolicyIQ: the existing chat searches across ALL published policies. The Reader Tutor is grounded on exactly ONE document — the one the user is currently reading — with page-anchored citations.

### Ship Criteria (MVP)

- User opens a PDF policy → Reader Tutor panel available within 30 seconds
- "What does section 3 mean?" → cited answer with page highlight
- Question the document doesn't answer → graceful "I don't have a reliable source for that" refusal
- Skimming triggers proactive tutor within 60 seconds
- "I have read and understood" button locked until `engagement_score >= 0.65`
- Full audit trail exportable to compliance team
- PII never leaves the tenant

---

## 3. Current State of PolicyIQ (What Already Exists)

### Application

| Attribute | Value |
|-----------|-------|
| Framework | SPFx 1.20, React 17, TypeScript 4.7, Fluent UI v8 |
| Site | https://mf7m.sharepoint.com/sites/PolicyManager |
| Version | 1.2.4 |
| Webparts | 17 |
| Services | 150+ |
| Theme | Forest Teal (#0d9488) |

### Existing Reader (PolicyDetails.tsx)

The current policy reader is in `src/webparts/jmlPolicyDetails/components/PolicyDetails.tsx`. It already has:

- **HTML rendering** via `dangerouslySetInnerHTML` for converted documents
- **PDF rendering** via native browser `<object>` embed
- **Office Online iframe** for Word/Excel/PPT
- **4-step wizard flow**: Read → Quiz → Acknowledge → Complete
- **Bottom bar** with "I Have Read This Policy" button
- **Scroll tracking** (95% scroll to enable the button)
- **Read time tracking**

The Reader Tutor panel will sit alongside this existing content area, NOT replace it.

### Existing AI Infrastructure

| Resource | Name | Purpose |
|----------|------|---------|
| Azure OpenAI | `dwx-pm-openai-prod` | GPT-4o, deployed in Sweden Central |
| Chat Function | `dwx-pm-chat-func-prod` | Existing chat proxy (3 modes) |
| Quiz Function | `dwx-pm-quiz-func-prod` | AI quiz generation |
| Key Vault | `dwx-pm-kv-ziqv6cfh2ck3o` | API key storage |

**New Azure resources needed for Reader Tutor:**
- Azure Document Intelligence (for PDF/DOCX → markdown + bounding polygons)
- Azure AI Search index (for per-document hybrid retrieval)
- New Azure Function endpoints (doc-ingest, doc-copilot, doc-telemetry, etc.)

### Existing Email Pipeline

All notifications go through `PM_NotificationQueue → Azure Logic App → Office 365`. Use this existing pipeline for any Reader Tutor notifications.

### Role System

4 roles: User, Author, Manager, Admin. Roles are stored in `PM_UserProfiles.PMRole` and detected via `RoleDetectionService`. The Reader Tutor panel should be visible to ALL roles (everyone reads policies).

### Key Conventions (from CLAUDE.md — MANDATORY)

- **Class components** throughout (React 17 SPFx pattern)
- **`@ts-nocheck`** on large files (don't fight it)
- **Forest Teal** colour scheme (#0d9488 primary, #0f766e dark)
- **`PM_LISTS`** constants for all SharePoint list names
- **`sanitizeHtml()` / `escapeHtml()`** for all user content
- **`_isMounted` guard** on all async setState calls
- **`ErrorBoundary`** wrapper on all webpart renders
- **No `Promise.allSettled()`** — SPFx 1.20 targets ES2017
- **No mock data fallbacks** — show real errors (learned the hard way in Session 25)
- **Single code path per operation** — never duplicate logic across multiple files
- **One task at a time, verify before moving on**
- **Never package without explicit user approval**
- **Explain your plan before coding — get approval first**

---

## 4. Implementation Order (from the brief §8)

Follow this exact sequence. Each step unblocks the next.

| Day | Task | Spec Reference |
|-----|------|----------------|
| 1-2 | Infrastructure confirmation — Azure OpenAI, AI Search, Document Intelligence, Blob access | Brief §8.1 |
| 3-4 | PDF renderer wrapper — PDF.js v5 legacy build as React component, 7 required events | Brief §3.2, Pattern §3.1 |
| 5-6 | Ingestion function — `/api/doc-ingest`: Doc Intelligence → chunks → embed → AI Search index → sidecar JSON | Brief §3.3, Pattern §3.3 |
| 7-8 | Copilot function — `/api/doc-copilot`: hybrid search + L2 rerank + GPT-4o + SSE streaming + citations | Brief §3.4, Pattern §3.5 |
| 9-10 | Panel UI — 3-tab panel (Chat, Study, Progress), citation chips, wire to copilot function | Brief §3.1, Mockup HTML |
| 11-12 | Telemetry + engagement score — `/api/doc-telemetry`, batch events, compute score server-side | Brief §3.6, Pattern §3.2 |
| 13 | Proactive tutor — rules engine + gpt-4o-mini escalation + interrupt modal | Brief §3.5, Pattern §3.6 |
| 14 | Artefact pipeline — summary, glossary, key sections, attestation questionnaire + author review | Brief §3.7, Pattern §3.4 |
| 15-16 | Attestation gate + audit export — gate the "I have read and understood" button, export JSON package | Brief §3.6, Pattern §3.7 |
| 17-18 | PII scrubber, privilege flag, Guardian tier gate | Brief §5, Pattern §3.8 |
| 19 | WCAG 2.2 AA pass | Brief §8.11 |
| 20 | Playwright tests + audit-package load test + documentation | Brief §8.12 |

---

## 5. Architecture Decisions (Already Made)

These are locked. Don't reconsider them.

| Decision | Choice | Why |
|----------|--------|-----|
| PDF renderer | PDF.js v5 legacy build (ES5) | SPFx compatibility, Apache 2.0, full event access |
| PDF extract | Azure Document Intelligence `prebuilt-layout` | Bounding polygons for citation jump, markdown output, OCR for scans |
| Chunking | 512 tokens / 128 overlap, section-heading-aware | Microsoft-published sweet spot |
| Embeddings | Azure OpenAI `text-embedding-3-large` | Already in our AOAI deployment |
| Vector store | Azure AI Search with hybrid + L2 reranker | Already deployed, BM25 + vector critical for legal terms |
| Agent framework | Semantic Kernel Process Framework (.NET 8) | Matches existing Function app runtime |
| Chat LLM | GPT-4o (primary) + GPT-4o-mini (proactive tutor) | Split by task importance |
| Streaming | Server-Sent Events (SSE) | Simpler than WebSockets, works through SP |
| Word handling | Convert to PDF on ingestion | Cleanest path for MVP |
| Telemetry | xAPI into PM_xAPIStatements (new list) | Standard vocabulary |
| Naming | "DWx Copilot: Document Tutor" | Suite-consistent branding |

---

## 6. What NOT to Build

Explicitly out of scope for MVP (from brief §11):

- Multimedia widgets (video pause-points, scenarios, drag-drops) — that's LearnIQ's track
- Ebook authoring UI
- Branching content
- SCORM/xAPI export
- Voice mode / audio overview (Phase 2)
- Flashcards as in-document widgets (Phase 2)
- Peer annotations / collaborative reading (Phase 3)
- Policy-change diff view + auto-reattestation (Phase 2)

---

## 7. Integration Points with Existing PolicyIQ Code

### Where the Reader Tutor Panel Attaches

The panel goes into `PolicyDetails.tsx` — the same component that renders the policy reader. It should appear as a right-docked panel alongside the document content (HTML/PDF/Office viewer).

**Entry points (from brief §6.1):**
- From My Policies → "Read with Copilot" (default for policies with Reader Tutor enabled)
- From email notification links
- From Policy Hub browse mode (optional — browse mode could show read-only panel)

### Engagement Score Replaces Scroll-Based Gating

Currently the "I Have Read This Policy" button enables at 95% scroll. The Reader Tutor replaces this with `engagement_score >= 0.65` — a richer signal that includes dwell time, scroll coverage, interaction rate, and comprehension score.

### PolicyIQ-Specific Metadata

The panel reads these existing policy fields:

```
policyId, version (VersionNumber), owner (PolicyOwner),
effectiveDate, nextReviewDate,
RequiresAcknowledgement, RequiresQuiz,
ComplianceRisk, ReadTimeframe
```

New field needed on PM_Policies: `ReaderTutorEnabled` (boolean, defaults true for new policies).

### PolicyIQ-Specific Mechanics

1. **Role-scoped applicability** — on open, panel checks the user's role/department and collapses sections that don't apply (not hidden — collapsed with "Not applicable to your role" marker)
2. **Attestation cadence** — if a policy version changes, all prior attesters are flagged for re-attestation
3. **Applicability checklist** — auto-generated questions like "does this section apply to my role?" that must be answered before attestation

---

## 8. SharePoint Lists to Create

| List | Purpose | Key Fields |
|------|---------|------------|
| PM_xAPIStatements | Telemetry events | StatementId, Verb, Actor, Object, Result, Context, Timestamp |
| PM_DocumentIndex | Document ingestion status | DocumentId, PolicyId, IndexStatus, ChunkCount, LastIndexed |
| PM_AIArtefacts | Generated study artefacts | ArtefactId, PolicyId, Type, Content, Status (Draft/Reviewed), ReviewedBy |
| PM_EngagementScores | Per-user per-policy engagement | PolicyId, UserId, Score, Components (JSON), LastUpdated |
| PM_AIInteractionLog | Immutable audit trail | InteractionId, PolicyId, UserRef, Action, PromptHash, TokensIn, TokensOut, Citations |

---

## 9. Azure Resources to Deploy

| Resource | Name (suggested) | Region | Purpose |
|----------|-------------------|--------|---------|
| Document Intelligence | dwx-pm-docint-prod | swedencentral | PDF/DOCX → markdown + polygons |
| AI Search | dwx-pm-search-prod | swedencentral | Per-document hybrid retrieval |
| Function App | dwx-pm-tutor-func-prod | swedencentral | 7 Reader Tutor endpoints |
| Storage Account | (reuse existing) | swedencentral | Sidecar JSON blobs |

Reuse existing: Azure OpenAI (`dwx-pm-openai-prod`), Key Vault (`dwx-pm-kv-*`), App Insights.

---

## 10. Session 25 Fixes (Context for the New Agent)

These are fixes that were made in Session 25 that the new agent should be aware of:

1. **Role detection** — PM_UserProfiles is now the source of truth for roles (JmlAppLayout.tsx, RoleDetectionService.ts). IsSiteAdmin is fallback only when no profile record exists.
2. **Single code path for submit-for-review** — PolicyAuthorView now delegates to PolicyService.submitForReview(). Don't create inline notification code.
3. **No mock data fallbacks** — PolicyDetails.tsx had a `loadMockPolicyDetails()` that silently showed fake data when SP queries failed. It was deleted. Never add fallback mock data.
4. **select('*') for SP queries** — Don't use explicit column lists in `.select()` calls. Columns may not be provisioned. Use `select('*')`.
5. **Nav permissions** — getDefaultPermissions() keys must exactly match nav item keys in PolicyManagerHeader.tsx.
6. **Settings cog** — Admin-only (not Manager+).
7. **AI bulk classification** — Content extraction sends whatever it has (even <100 chars). Strip markdown code fences from AI responses before JSON parse.
8. **PDF viewer height** — Use `minHeight: calc(100vh - 200px)` for PDF/iframe embeds in flex containers. `height: 100%` collapses to near-zero.

---

## 11. Known Issues & Technical Debt

- PolicyDetails review/approval decisions still have ~220 lines of inline code (direct SP writes). Should be consolidated to PolicyService in a future session — but it works, don't break it while building the Reader Tutor.
- `@ts-nocheck` on ~220 files. Don't try to fix this.
- The existing AI Chat (PolicyChatPanel) searches across ALL policies. The Reader Tutor is per-document. They are separate features. Don't merge them.
- TinyMCE is bundled (not CDN) — same approach needed for PDF.js.

---

## 12. User Communication Rules (from CLAUDE.md — CRITICAL)

The user (Gary) has specific expectations for how you work:

1. **Always explain your plan before coding.** Describe what you will do and wait for approval.
2. **One task at a time.** Complete and verify before starting the next.
3. **Never package without explicit permission.** `gulp bundle --ship` is fine. `gulp package-solution --ship` requires "ship it" / "package it" from the user.
4. **Never add mock data.** Show real errors.
5. **If a task was done wrong before, flag it.** Note "REDO" in your todo list.
6. **A successful build does NOT mean the task is done.** Verify the actual output.
7. **Create HTML mockups for UI changes** before implementing.

---

## 13. File References (Quick Access)

| File | Purpose |
|------|---------|
| `CLAUDE.md` | Full project context |
| `src/webparts/jmlPolicyDetails/components/PolicyDetails.tsx` | Current reader — Reader Tutor panel integrates here |
| `src/services/PolicyService.ts` | Core policy CRUD + lifecycle |
| `src/services/PolicyChatService.ts` | Existing AI chat (different scope — searches all policies) |
| `src/services/PolicyRoleService.ts` | Role detection + nav filtering |
| `src/components/PolicyManagerHeader/PolicyManagerHeader.tsx` | App header + nav |
| `src/components/JmlAppLayout/JmlAppLayout.tsx` | Full-page layout wrapper |
| `src/utils/EmailTemplateBuilder.ts` | Email template builder |
| `src/utils/sanitizeHtml.ts` | HTML sanitization (sanitizeHtml + escapeHtml) |
| `src/constants/SharePointListNames.ts` | All PM_ list name constants |
| `src/constants/BuildInfo.ts` | Build version number |
| `azure-functions/policy-chat/` | Existing chat Azure Function |
| `azure-functions/quiz-generator/` | Existing quiz Azure Function |
| `config/config.json` | SPFx webpart registration |
| `docs/reader-tutor-brief-for-policyiq-contractiq.md` | **YOUR PRIMARY SPEC** |
| `docs/dwx-document-tutor-pattern.md` | Full pattern specification |
| `docs/interactive-pdf-research-and-plan.md` | LearnIQ research (architecture reference) |
| `docs/interactive-pdf-ux-mockup.html` | **Definitive UX mockup** |

---

## 14. Git & Deployment

- **Repos:** ADO `dev.azure.com/gfinberg/DWx/_git/dwx-policy-manager` + GitHub mirror
- **Branch:** `master` (direct commits, no PR flow currently)
- **Push:** `git push origin master` pushes to both remotes
- **Build:** `npx gulp clean && npx gulp bundle --ship && npx gulp package-solution --ship`
- **Package:** `sharepoint/solution/policy-manager.sppkg` — upload to App Catalog
- **Azure:** `az` CLI available, Bicep for IaC. See existing `azure-functions/*/infra/` for patterns.

---

## 15. Done Criteria (from brief §13)

You can claim MVP complete when:

- [ ] Upload a real PDF → ingestion completes in 30-60 seconds
- [ ] "What does section 3 mean?" → cited answer that scrolls and highlights the passage
- [ ] Ungrounded question → graceful refusal
- [ ] Skimming triggers proactive tutor within 60 seconds
- [ ] Attestation button locked below 0.65, unlocks above
- [ ] Audit export JSON validates against pattern §3.7 schema
- [ ] PII detection redacts a test query with a client name
- [ ] WCAG 2.2 AA keyboard-only navigation works end-to-end
- [ ] Pattern compliance checklist (pattern spec §6) is 100% ticked

---

*End of handoff document. Good luck — this is a great feature to build.*
