# Reader Tutor MVP — Implementation Brief for PolicyIQ and ContractIQ

> **Audience:** An AI coding agent working in the PolicyIQ or ContractIQ codebase (not LearnIQ)
> **Scope:** Build a **Reader Tutor Panel** — a simpler, PDF/Word-focused subset of the [DWx Document Tutor pattern](dwx-document-tutor-pattern.md) that works over existing documents with zero re-authoring required
> **Context:** LearnIQ is building the full agentic ebook version of this pattern. You are building the leaner **"agentic Q&A sidebar on any existing document"** version. Same architectural DNA; much smaller surface area.
> **Source of truth for the full pattern:** [dwx-document-tutor-pattern.md](dwx-document-tutor-pattern.md) — read §§1-3 before starting. This brief is the simplified subset.
> **Companion:** [interactive-pdf-research-and-plan.md](interactive-pdf-research-and-plan.md) is LearnIQ's fuller plan. Use it as a reference for rendering patterns and Azure plumbing, not as a spec.
> **Version:** 1.0 · 2026-04-18

---

## 0. TL;DR (90 seconds)

Build a **right-docked panel** that sits alongside any PDF or Word document in the app. The panel has three tabs — **Chat**, **Study**, **Progress** — plus a proactive-tutor modal that interrupts skim-reading with a comprehension check.

Every answer is **grounded in the document** (cites a specific page and section) and **refuses gracefully** when the document doesn't contain the answer. Every interaction is logged for audit. Learner/employee PII never leaves the tenant — the AI sees only `{userRef, docId, anonymised question context}`.

**Ship criteria (MVP):**
- Can point the panel at any existing PDF in the app's document store and get useful, cited Q&A within 30 seconds.
- Same for Word (.docx) via conversion on ingestion.
- Proactive tutor fires on skim signal; user can dismiss.
- Attestation gate: user cannot click "I have read and understood" (PolicyIQ) / cannot advance signoff workflow (ContractIQ) until `engagement_score >= 0.65`.
- Full audit trail exportable to compliance team.

**Target build effort:** 15-20 dev-days. Tight scope on purpose. See §9 for the breakdown.

---

## 1. Why this is different from LearnIQ's track

LearnIQ is building an **interactive multimedia ebook** with authored widgets (videos, scenarios, drag-drops). That requires new authoring tooling, a widget library, a content JSON model, and ~40+ dev-days.

PolicyIQ and ContractIQ **do not need any of that**. Their users already have a library of existing PDFs and Word documents. The win is to layer an agentic reader on top of those documents without asking anyone to re-author anything.

This means you build:

- A **PDF and Word renderer** (not an ebook renderer, not a widget library)
- A **tutor chat panel** with grounded Q&A and citations
- A **proactive-tutor modal** for skim detection
- A **progress/attestation tab** with engagement score
- The **ingestion pipeline** that indexes a document once when it's first opened (or uploaded)

You do NOT build:

- Any widget library (no MCQ, no scenarios, no drag-drop)
- Any authoring UI
- Any page-layout presets
- Any multimedia content model
- Any SCORM/xAPI export authoring tooling

**What stays the same as the LearnIQ track:** the event contract, the engagement score formula, the audit trail schema, the PII-scrubbing discipline, the Azure AI Search + Document Intelligence plumbing, the Semantic Kernel agent, the refusal-when-ungrounded rule. You inherit all of that directly from the pattern spec.

---

## 2. Answers to the "can we use existing docs?" questions

Before you start, these are the answers to the questions the product owner is almost certainly going to ask.

| Question | Answer |
|---|---|
| Can we use existing PDFs with no modification? | **Yes.** Upload to the document store → ingest pipeline runs once → tutor works. |
| Can we use existing Word (.docx) docs? | **Yes.** Azure Document Intelligence handles DOCX natively and returns markdown + bounding polygons. Alternatively, pre-convert to PDF on ingestion and treat as PDF. For ContractIQ, Word is the primary format. |
| Do documents need to be re-authored for the tutor to work? | **No.** The tutor grounds itself on the document content as-is. No markers, no embedded metadata, no widgets. |
| What about scanned/image PDFs? | **Yes.** Document Intelligence's `prebuilt-layout` model handles OCR natively. Quality is better for digitally-generated PDFs but scans work. |
| What about password-protected or DRM'd PDFs? | **No.** Must be decrypted before ingestion. Flag this explicitly at upload — it's a common footgun. |
| What about very large documents (>200 pages)? | **Yes but ingest is slower and costs more.** Budget ~$0.02/page in Document Intelligence + embedding. Index once and reuse forever. |
| Do we need to chunk by clause or page number? | **Chunk by section heading** (markdown heading from Doc Intelligence output) with a 512-token / 128-overlap fallback for long sections. Store page number as metadata for citation. |
| Does ContractIQ need playbook comparison? | **Out of MVP scope.** That's the ContractIQ-specific extension in pattern spec §4.3. Ship base Reader Tutor first; layer playbook comparison as a Phase 2 extension. |

---

## 3. Feature scope — what you are building

### 3.1 The Reader Tutor Panel

A right-docked panel, 420px wide, that sits alongside the document viewer. Based exactly on the Copilot Panel in [interactive-pdf-ux-mockup.html](interactive-pdf-ux-mockup.html) (the LearnIQ PDF mockup). Study that file. Reproduce the interaction rhythm; don't invent a new UX.

**Panel structure:**

```
╔═══════════════════════════════════╗
║  DWx Copilot                      ║
║  [DOCUMENT TUTOR MODE]            ║ ← header (A3 gradient, 64px)
╠═══════════════════════════════════╣
║  [ Chat | Study | Progress ]      ║ ← tabs
╠═══════════════════════════════════╣
║                                   ║
║  (tab body — chat messages,       ║
║   or study artefacts,             ║
║   or engagement score)            ║
║                                   ║
╠═══════════════════════════════════╣
║  [ input field ][send]            ║ ← chat input (chat tab only)
║  AI cites sources · No PII...     ║ ← disclaimer
╚═══════════════════════════════════╝
```

### 3.2 The document renderer

- **PDF**: use PDF.js v5 legacy build (ES5-compatible). See §5.1 of [interactive-pdf-research-and-plan.md](interactive-pdf-research-and-plan.md) for the full wrapper pattern. Bundle, don't CDN.
- **Word**: convert to PDF on ingestion (cleanest path). Serve PDF to the renderer. Alternative: inline HTML via Mammoth.js if live-edit is needed later — skip this for MVP.

The renderer must emit the event stream defined in [dwx-document-tutor-pattern.md §3.1](dwx-document-tutor-pattern.md#31-component-1--document-renderer): `docOpened`, `pageChanged`, `dwellTick`, `textSelected`, `annotationCreated`, `citationJumped`, `sessionEnded`. This is non-negotiable — the rest of the system assumes these events exist.

### 3.3 Ingestion pipeline (one-off per document)

Exactly as specified in [pattern spec §3.3](dwx-document-tutor-pattern.md#33-component-3--document-index):

1. Upload → document store (your app's document library)
2. Azure Document Intelligence `prebuilt-layout` in markdown mode → returns markdown + polygons per paragraph
3. Chunk: 512 tokens / 128 overlap, section-heading-aware split
4. Embed: Azure OpenAI `text-embedding-3-large` (trimmable to 1536 dim)
5. Index into Azure AI Search with fields: `documentId, productCode, corpusKey, pageNumber, boundingPolygon, sectionPath, content, contentVector, authorReviewed, dataClassification, tenantId`
6. Sidecar JSON (paragraph polygons) → Azure Blob (so citation-to-coordinate jump works without a round-trip at query time)

**`productCode` values to use:**
- PolicyIQ → `"PolicyIQ"`
- ContractIQ → `"ContractIQ"`

**`corpusKey` patterns:**
- PolicyIQ: `policy:{PolicyId}`
- ContractIQ: `contract:{MatterRef}/{DocId}`

### 3.4 The tutor agent

Semantic Kernel Process Framework in .NET 8, running as an Azure Function. Endpoints from [pattern spec §5.3](dwx-document-tutor-pattern.md#53-shared-endpoint-contract):

```
POST /api/doc-ingest                    # kick off ingestion pipeline
POST /api/doc-copilot                   # SSE-streamed Q&A turn
POST /api/doc-generate-artefact         # summary, glossary, attestation Qs
POST /api/doc-telemetry                 # batch events from renderer
POST /api/doc-proactive-tutor           # skim detection response
GET  /api/doc-sidecar/{docId}           # sidecar JSON for citation jumping
POST /api/doc-audit-export              # audit package per pattern §3.7
```

The agent's **system prompt** must include:

- The document title and docId it's grounded on
- The corpus description (what the fallback knowledge corpus is — e.g. "the wider firm policy library" or "the firm's contract playbook")
- Persona parameter (`default` / `basic` / `regulator` / `client`)
- The four refusal triggers (pattern spec §3.5): ungrounded, privileged, high-stakes, PII-contaminated

Use exactly this template as the starting point:

```
You are the DWx Document Tutor for {productName}.
You are grounded in exactly ONE document: "{documentTitle}" ({docId}).
You may also reference the {corpusDescription} when the user explicitly asks.

Rules:
1. Every factual claim MUST cite at least one passage from the document using the {CITE:page:anchor} marker.
2. If no retrieved passage scores above the relevance threshold, respond:
   "I don't have a reliable source for that in this document."
   Then offer to search {fallbackCorpusName} if it exists.
3. Never output legal, financial, or compliance advice — offer interpretation of the source document only.
4. Never echo or solicit PII: client names, employee names, matter references, email addresses.
   If the user's question contains PII, rewrite it without the identifiers before answering.
5. Persona: {persona} — adjust tone and depth, never change facts.
```

### 3.5 Proactive tutor

Rules-first, LLM-escalated. Exactly as specified in [pattern §3.6](dwx-document-tutor-pattern.md#36-component-6--proactive-tutor). Rules fire when:

- WPM > 400 on a page with > 300 words
- 3+ consecutive pages with < 1s dwell
- zero selections + zero mouse movement over 3 pages
- scroll-past rate > 40% across session
- user opened doc, visited only last page, then tried to unlock attestation

Only on a rule flag: call `gpt-4o-mini` to compose a section-specific comprehension check. Never block; always dismissible. Logs to audit trail.

### 3.6 Engagement score + attestation gate

Canonical formula from [pattern §3.7](dwx-document-tutor-pattern.md#37-component-7--engagement-score--comprehension-attestation):

```
engagement_score =
    0.4 * normalisedDwell
  + 0.2 * scrollCoverage
  + 0.2 * interactionRate
  + 0.2 * comprehensionScore
```

**Gate behaviour per app:**

- **PolicyIQ** — the "I have read and understood" / attestation button is disabled until `engagement_score >= 0.65` AND every applicability-checklist item is answered (if the policy has an applicability checklist). On click, produce a signed attestation record with tamper-evident signature; store against the user + policy version.
- **ContractIQ** — the partner-signoff workflow cannot advance until `engagement_score >= 0.65` AND every flagged risky clause has been acknowledged. On gate pass, the engagement-score package + Q&A transcript attach to the matter file.

### 3.7 Artefact pipeline (MVP subset)

Generate the following at ingestion, save as `Draft` until an owner reviews:

- **Executive summary** (1 page, gpt-4o)
- **Glossary** (auto-extracted terms, gpt-4o-mini)
- **Key-sections index** (deterministic — table of contents)

**PolicyIQ also generates:**
- **Applicability checklist** ("does this section apply to my role?")
- **Attestation questionnaire** (3-5 questions the attester must get right)

**ContractIQ also generates:**
- **Clause index with risk score** (gpt-4o grounded on firm playbook if available)

Skip for MVP (Phase 2): audio overview, flashcards, regulator's-perspective critique, scenario quiz, playbook-alignment report, deviation heatmap.

**Author review gate is mandatory.** Generated artefacts carry `authorReviewed: false` until a human has reviewed and accepted. Only reviewed artefacts are served to end users.

---

## 4. Audit trail — the most important part for your product

PolicyIQ and ContractIQ users care about attestation evidence more than tutoring quality. Get this right.

Every AI interaction writes one immutable row:

```
{
  interactionId, ts, userRef, docId, corpusKey, productCode,
  action, promptHash, modelId, tokensIn, tokensOut, latencyMs,
  citations, redactionsApplied, persona, toolsUsed
}
```

`userRef` is a reference code — `EMP-XXXX` (PolicyIQ) or `LAW-XXXX` (ContractIQ). **Never** store full names or emails in audit rows that the AI can see.

Audit export package per document per user follows [pattern §3.7](dwx-document-tutor-pattern.md#37-component-7--engagement-score--comprehension-attestation). This is what your compliance team exports and hands to auditors. Build it as one endpoint — don't make users assemble it.

---

## 5. PII and privilege handling

Before any user query reaches OpenAI:

1. Run **Azure AI Language PII detection** on the query
2. Redact or block per policy:
   - **PolicyIQ**: redact client names and employee names; warn the user
   - **ContractIQ**: privilege-default is on — every ContractIQ document carries `privilegeWarning: true` unless explicitly disabled. The tutor prepends a privilege disclaimer and disables the "regulator's perspective" / "opposing counsel" tools.
3. Log redactions to the audit trail (not the redacted content, just the fact that it happened)
4. Never persist PII in Q&A transcripts stored against the document

Document-side: every indexed chunk carries `dataClassification` (`public` / `internal` / `confidential` / `privileged`). Retrieval filters by user clearance level.

---

## 6. Per-app UI integration

### 6.1 PolicyIQ integration

The Reader Tutor sits inside the **Policy Reader** surface. Entry points:

- From the policy library ("Read with Copilot")
- From an email link delivered to attesters
- From an HR onboarding flow (mandatory new-starter reading)
- Embedded in the internal portal's policy-of-the-month widget

Policy metadata the panel reads:

```
policyId, version, owner, effectiveDate, nextReviewDate,
attestationRequired: bool,
attestationRoles: string[],
attestationCadence: 'annual' | 'on-change' | 'once',
privilegeWarning: false
```

The panel's Progress tab shows: last attestation date (if any), current policy version, whether a re-attestation is pending (e.g. because the policy version changed since last attestation).

Unique PolicyIQ mechanic: **role-scoped applicability**. On open, panel asks the user their role (or reads from AD), then surfaces only sections that apply to them. Irrelevant sections show a collapsed "Not applicable to your role" marker. Don't hide them — collapse them. Users need to see the full doc exists.

### 6.2 ContractIQ integration

The Reader Tutor sits inside the **Contract Reader** surface. Entry points:

- From a matter's document list
- From a contract upload flow
- From DMS link-in (iManage / NetDocuments) in Phase 2

Contract metadata the panel reads:

```
matterId, clientRef, dealType, practiceArea, role (firm-side | counterparty),
playbookId, playbookVersion (optional — absent in Phase 1),
privilegeWarning: true (default)
```

Unique ContractIQ mechanic: **deviation flag acknowledgement**. For contracts where a playbook is set (Phase 2), the panel surfaces clauses that deviate from standard firm language. Each flag requires learner acknowledgement before signoff. In Phase 1 without a playbook, skip this — it becomes a clause-index view only.

Progress tab additionally shows: who else has reviewed this contract, their attestations, any partner overrides logged.

---

## 7. Shared infrastructure — what to build vs what to reuse

### 7.1 Reuse (from suite-wide shared code)

These belong in shared suite packages per [pattern spec §5](dwx-document-tutor-pattern.md#5-reusable-building-blocks). If they don't exist yet in your monorepo, **coordinate with the DWx Deployment Manager team** — do not fork them into your product.

- `@dwx/document-renderer-core` — event contract interface
- `@dwx/document-renderer-pdf` — PDF.js wrapper
- `@dwx/document-tutor-client` — copilot panel UI, chat, citations
- `@dwx/document-tutor-types` — shared TypeScript types
- `@dwx/engagement-scorer` — the canonical formula
- `DWx.DocumentTutor.Ingestion` (.NET)
- `DWx.DocumentTutor.Retrieval` (.NET)
- `DWx.DocumentTutor.Agent` (.NET)
- `DWx.DocumentTutor.ProactiveTutor` (.NET)
- `DWx.DocumentTutor.Telemetry` (.NET)
- `DWx.DocumentTutor.Shared` (.NET — DTOs, PII scrubber, privilege-flag handler)

### 7.2 Build (PolicyIQ/ContractIQ-specific)

- The surface page/web-part (Policy Reader or Contract Reader)
- System prompt specialisations per product (see §3.4)
- Tool registrations per product (see [pattern §2 table](dwx-document-tutor-pattern.md#2-why-its-reusable-and-why-the-same-architecture-fits-all-three))
- Artefact type specialisations (PolicyIQ: applicability checklist, attestation questionnaire; ContractIQ: clause index, deviation flags)
- Attestation / signoff workflow integration with the rest of your product
- Role-based applicability (PolicyIQ) / clause-flag acknowledgement (ContractIQ)

### 7.3 Do not build

- Any widget library
- Any ebook authoring UI
- Any SCORM/xAPI export tooling
- A separate branding or naming for the copilot — **reuse "DWx Copilot"** across the suite

---

## 8. Concrete implementation order

Work through this sequence. Don't jump ahead — each step unblocks the next.

1. **Day 1-2 — Infrastructure confirmation.** Confirm you have access to Azure OpenAI, Azure AI Search, Azure Document Intelligence, Azure Blob. If the shared packages in §7.1 don't exist, coordinate before you write any code.

2. **Day 3-4 — PDF renderer wrapper.** Build the PDF.js v5 legacy wrapper as a React component. Must emit the 7 required events. Test with a real 50-page PDF.

3. **Day 5-6 — Ingestion function.** Implement `/api/doc-ingest`. Test by ingesting a PDF end-to-end: Document Intelligence → chunks → embed → Azure AI Search index → sidecar JSON.

4. **Day 7-8 — Copilot function.** Implement `/api/doc-copilot` with hybrid search + L2 rerank + Azure OpenAI GPT-4o + SSE streaming. System prompt from §3.4. Refusal behaviour enforced. Every response includes citation markers.

5. **Day 9-10 — Panel UI.** Build the three-tab panel. Wire it to the copilot function. Citation chips that click-jump to the PDF. Test against the mockup [interactive-pdf-ux-mockup.html](interactive-pdf-ux-mockup.html) — it should behave identically to the chat interactions in that mockup.

6. **Day 11-12 — Telemetry + engagement score.** Implement `/api/doc-telemetry`. Batch client events every 30s. Compute engagement score server-side. Show in Progress tab.

7. **Day 13 — Proactive tutor.** Rules engine + gpt-4o-mini escalation + interrupt modal.

8. **Day 14 — Artefact pipeline.** Summary + glossary + key-sections + (PolicyIQ: applicability checklist / attestation Qs) / (ContractIQ: clause index). Author review UI as simple accept-all-or-edit.

9. **Day 15-16 — Attestation gate + audit export.** Gate the product's attestation button / signoff workflow. Implement `/api/doc-audit-export` producing the JSON package in [pattern §3.7](dwx-document-tutor-pattern.md#37-component-7--engagement-score--comprehension-attestation).

10. **Day 17-18 — PII scrubber, privilege flag, Guardian tier gate.** Wire the safety rails per §5 + pattern §3.8.

11. **Day 19 — WCAG 2.2 AA pass.** Keyboard nav, focus-trap, aria-live, screen-reader announcements. Referenced in pattern §3.5.

12. **Day 20 — Playwright tests + audit-package load test + documentation.**

---

## 9. Compliance with the pattern spec

Before you claim the work is complete, walk through the [Pattern Compliance Checklist in §6 of the spec](dwx-document-tutor-pattern.md#6-pattern-compliance-checklist). Tick every item. If you can't tick one, that's an open ticket, not a shipped feature.

Declare compliance in your product manifest:

```
"documentTutor": {
  "specVersion": "1.0",
  "profile": "PolicyIQ"   // or "ContractIQ"
}
```

---

## 10. What to do if you hit ambiguity

1. **Read the pattern spec.** 95% of questions are answered there.
2. **Read LearnIQ's plan.** The [interactive-pdf-research-and-plan.md](interactive-pdf-research-and-plan.md) has the full Azure plumbing decisions with rationale.
3. **Don't invent new architecture.** If your instinct is to introduce a new service, a new storage format, or a new event type — stop and ask. The pattern is deliberately opinionated; deviations compound.
4. **Don't skip the author review gate.** CLAUDE.md (LearnIQ) §18 is the origin, but this applies suite-wide: AI-generated artefacts never auto-publish.
5. **Don't send PII to OpenAI.** Every time.

---

## 11. Out of scope for MVP (explicitly)

- Multimedia widgets (video pause-points, scenarios, drag-drops)
- Ebook authoring UI
- Branching content
- SCORM / xAPI export of ebooks
- Voice mode / audio overview (Phase 2)
- Flashcards as in-document widgets (Phase 2 — they can exist as study artefacts but not inline)
- Peer annotations / collaborative reading (Phase 3)
- DMS integration (ContractIQ Phase 2)
- Playbook alignment and deviation heatmap (ContractIQ Phase 2)
- Policy-change diff view + auto-reattestation (PolicyIQ Phase 2)
- Regulatory-update auto-generated courses (Phase 3, LearnIQ-led)

If in doubt whether something is in scope, default to out. Ship MVP first.

---

## 12. Open questions for the product owner

Before you start, get explicit answers to these from your product owner:

1. **Document source of truth** — where do PolicyIQ / ContractIQ store their documents today? SharePoint document library? Blob? A proprietary doc store?
2. **Who is the "author" / "reviewer"** for the artefact review gate in each product? Policy owner for PolicyIQ. Supervising partner for ContractIQ?
3. **Guardian tier positioning** — is Reader Tutor a base-tier feature for PolicyIQ/ContractIQ, or a Professional/Enterprise add-on?
4. **Identity system** — does the product already issue ref codes (`EMP-XXXX`, `LAW-XXXX`), or do you need to create a mapping?
5. **Deployment Manager manifest** — has this product been added to the DWx Deployment Manager yet? If yes, extend its manifest with the `documentTutor` section. If not, coordinate.
6. **Language / multi-tenant localisation** — MVP is English-only. Is that acceptable?

---

## 13. Done criteria

You can claim MVP complete when:

- [ ] A product owner can upload a real PDF or Word doc, wait 30-60 seconds for ingestion, and open the Reader Tutor panel
- [ ] Asking "what does section 3 mean?" returns a cited answer that scrolls and highlights the right passage
- [ ] Asking a question the document doesn't answer produces a graceful "I don't have a reliable source for that" refusal
- [ ] Skipping through the doc triggers the proactive tutor within 60 seconds
- [ ] The attestation / signoff gate is locked below 0.65 and unlocks above it
- [ ] The audit export JSON validates against the [pattern §3.7 schema](dwx-document-tutor-pattern.md#37-component-7--engagement-score--comprehension-attestation)
- [ ] PII-detection redacts a test query that contains a client name
- [ ] The privilege flag (ContractIQ) disables the "regulator's perspective" tool
- [ ] WCAG 2.2 AA keyboard-only navigation works end-to-end
- [ ] Pattern compliance checklist is 100% ticked

Only then is this shippable.

---

## 14. Further reading

- [DWx Document Tutor Pattern Specification](dwx-document-tutor-pattern.md) — the authoritative reference
- [LearnIQ Interactive PDF Plan](interactive-pdf-research-and-plan.md) — fuller Azure architecture detail and market context
- [LearnIQ Interactive Ebook Concept](dwx-interactive-ebook-concept.md) — what LearnIQ is building beyond this brief (you don't need to build any of it, but it's useful context for where the shared packages are headed)
- [Interactive PDF UX mockup](interactive-pdf-ux-mockup.html) — **the definitive UX reference for your Reader Tutor panel**
- [Interactive Ebook mockup](interactive-ebook-mockup.html) — LearnIQ-specific; the Copilot panel portion still applies

---

*End of brief. Deliver this to the PolicyIQ and ContractIQ agents as the starting point for their implementation.*
