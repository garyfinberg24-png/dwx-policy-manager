# DWx Document Tutor Pattern — Reusable Specification

> **Pattern name:** DWx Document Tutor
> **Intended reusers:** PolicyIQ, ContractIQ, LearnIQ, and any future DWx app that renders long-form documents to professional users
> **Reference implementation:** LearnIQ `InteractivePDF` lesson type (this repository)
> **Status:** Specification — MVP pending in LearnIQ; pattern is stable enough to document for reuse
> **Version:** 1.0 · 2026-04-17
> **Owner:** First Digital — DWx Suite

---

## 1. What this pattern is

The **DWx Document Tutor** is a reusable interaction pattern that turns any long-form professional document into a **structured, auditable, agentic learning surface** without changing the document itself.

The user reads the document as they always have. Alongside it, an agentic AI tutor is always available to:

- **Answer grounded questions** about the document with page-anchored citations
- **Proactively intervene** when the reader skims, asking a comprehension check tied to the section they just sped through
- **Generate reusable artefacts** from the document once (summary, glossary, flashcards, scenarios, "regulator's perspective" critique, audio overview) which the user can consume instead of, or alongside, the full document
- **Capture engagement telemetry** (dwell + scroll + selection + Q&A + comprehension-check accuracy) into an immutable audit trail that can later prove *understanding*, not just attendance

Everything the tutor says is grounded in the document (and any adjacent corpus — firm policies, playbooks, case law, regulatory guidance). Everything the user does is logged. PII never leaves the tenant.

### 1.1 One-line description for each app

- **LearnIQ** — a training lesson type: the document becomes a course, the tutor becomes a teacher, comprehension becomes CPD evidence.
- **PolicyIQ** — a policy-owner companion: any firm policy becomes a living Q&A surface; the tutor answers "does this apply to X?" for every employee, every policy, every day; engagement logs prove attestation.
- **ContractIQ** — a deal-lawyer copilot: any contract becomes an interactive negotiating surface; the tutor explains clauses, flags deviations from firm playbook, and rehearses counter-arguments with juniors; every interaction is a learning event for the firm's institutional knowledge.

The **pattern is the same**. The **surface, prompts, and tool-set differ per app.**

---

## 2. Why it's reusable (and why the same architecture fits all three)

The underlying machinery — a renderer that emits selection/scroll events, a RAG grounded on one document, a streaming chat with citations back to page coordinates, a skim-detector that fires LLM-cheap comprehension questions, a pipeline that pre-generates artefacts at ingestion, an engagement scorer that gates the next step — is **identical** across the three use cases.

The only things that change per app are:

| Dimension | LearnIQ | PolicyIQ | ContractIQ |
|---|---|---|---|
| **Document type** | Training material (any format) | Firm policies, regulatory texts | Contracts, term sheets, NDAs, deal docs |
| **Typical length** | 5-100 pages | 2-50 pages | 5-500+ pages |
| **Typical audience** | Associate learners, CPD cohorts | All firm staff, periodically | Deal teams, partners, clients |
| **Primary outcome** | Comprehension attestation → CPD cert | Policy acknowledgement + Q&A audit trail | Clause explanations + deviation flags + playbook alignment |
| **Engagement gate** | Unlocks assessment | Unlocks "I have read and understood" | Unlocks partner signoff / deal progression |
| **Tool set available to the agent** | `search_doc`, `generate_quiz_question`, `jump_to_page`, `log_misconception` | `search_doc`, `search_firm_policy_corpus`, `explain_in_context`, `flag_for_policy_owner` | `search_doc`, `search_playbook`, `compare_to_precedent`, `draft_markup`, `flag_deviation` |
| **Artefact mix (auto-generated)** | Summary, glossary, flashcards, scenario quiz, regulator critique | Summary, plain-English explainer, attestation questionnaire, applicability checklist | Clause index, playbook alignment report, deviation heatmap, opposing-counsel critique, redline suggestions |
| **Proactive tutor triggers** | Skim detection → comprehension check | Section not visited → prompt "this section applies to your role" | Known risky clause → "let me explain why this matters" |
| **Audit context** | LPC/SRA CPD | POPIA §18 attestation, internal compliance | Deal file, DMS metadata, partner review trail |

Everything else is the same pattern.

---

## 3. Pattern anatomy

A DWx Document Tutor implementation has exactly **seven components**. Any conforming implementation must address all seven, and should not add a new one without updating this specification.

```
                                  ┌──────────────────────────────┐
                                  │  1. Document Renderer        │
                                  │     - paints the document    │
                                  │     - emits events           │
                                  └──────────────┬───────────────┘
                                                 │
                                                 ▼
                                  ┌──────────────────────────────┐
                                  │  2. Engagement Telemetry     │
                                  │     - dwell, scroll, select  │
                                  │     - xAPI out               │
                                  └──────────────┬───────────────┘
                                                 │
    ┌──────────────────────────────┐             │             ┌──────────────────────────────┐
    │  5. Tutor Agent              │◀────────────┼────────────▶│  6. Proactive Tutor          │
    │     - streaming Q&A          │             │             │     - rules-first skim det   │
    │     - tool use               │             │             │     - LLM-escalated prompts  │
    │     - citations              │             │             │     - polite interrupt modal │
    └──────────────┬───────────────┘             │             └──────────────────────────────┘
                   │                             │                           │
                   └─────────────┬───────────────┴───────────────┬───────────┘
                                 │                               │
                                 ▼                               ▼
                  ┌──────────────────────────────┐ ┌──────────────────────────────┐
                  │  3. Document Index           │ │  7. Engagement Score +       │
                  │     - Doc Intelligence       │ │     Comprehension Attestation│
                  │     - hybrid search          │ │     - gate next step         │
                  │     - per-doc filter         │ │     - audit export           │
                  └──────────────────────────────┘ └──────────────────────────────┘
                                 ▲
                                 │
                  ┌──────────────────────────────┐
                  │  4. Artefact Pipeline        │
                  │     - summary, glossary,     │
                  │       flashcards, scenarios, │
                  │       regulator critique,    │
                  │       audio overview         │
                  │     - author review gate     │
                  └──────────────────────────────┘
```

Component contracts below.

### 3.1 Component 1 — Document Renderer

**Contract:** Paints the document and emits a stable event stream.

**Required events:**

| Event | Payload |
|-------|---------|
| `docOpened` | `{docId, learnerRef, ts}` |
| `pageChanged` | `{docId, page, ts, viaCitation?: boolean}` |
| `dwellTick` | `{docId, page, dwellMs, wpmEstimate, scrollPct}` (every 5s on visible tab) |
| `textSelected` | `{docId, page, polygon, text, section}` |
| `annotationCreated` | `{docId, page, polygon, kind: highlight\|note, colour}` |
| `citationJumped` | `{docId, fromMsgId, toPage, anchorId}` |
| `sessionEnded` | `{docId, learnerRef, totalDwellMs, pagesVisited}` |

**Required methods:**

| Method | Purpose |
|--------|---------|
| `jumpToPage(page, anchorId?)` | scrolls + flashes highlight |
| `highlightPolygon(page, polygon, style)` | paints overlay from server-supplied coordinates |
| `showInterrupt(modalConfig)` | renders a proactive-tutor modal anchored to current position |

**Context payload for renderer-aware tutor turns:**

When the tutor is invoked, the renderer supplies a structured context object so the agent can narrate the learner's exact position. The minimum shape:

```jsonc
{
  "docId":     "...",
  "page":      2,            // 1-indexed
  "sectionId": "s-2a",       // optional — present for ebook renderer
  "layout":    "2x2",        // optional — one of full | 2-col | 2-col-1-2 | 2-col-2-1 | 3-col | 2x2
  "slot":      3,            // optional — 0-indexed within the section
  "widgetId":  "w-7",        // optional — present when a widget is focused or in interaction
  "widgetState": { }         // optional — widget-type-specific state (e.g. MCQ selection, scenario path)
}
```

Renderers that don't have sections (PDF, DOCX-as-PDF) simply omit the section/layout/slot fields. The tutor tolerates their absence.

**Implementations:**

- **PDFs** — PDF.js v5 legacy build + custom overlay DOM (SPFx, React 17, ES5). Context payload: `{docId, page}` only. See §5.1 of [interactive-pdf-research-and-plan.md](interactive-pdf-research-and-plan.md).
- **DOCX** — Convert to PDF on ingest and reuse PDF renderer *(recommended for ContractIQ MVP)*, OR use Mammoth.js for inline HTML rendering if live-edit is needed later.
- **HTML / multimedia ebook** — Custom React renderer (NOT EPUB at rest — see [dwx-interactive-ebook-concept.md §5](dwx-interactive-ebook-concept.md#5-technical-architecture--what-changes-vs-the-pdf-plan) for the JSON-source-of-truth decision). Context payload: full `{docId, page, sectionId, layout, slot, widgetId, widgetState}`.
- **SharePoint Modern Pages** — iframe with postMessage bridge emitting the same events.

**Rule:** whatever you render with, it must emit the event contract. The other components must work unchanged regardless of renderer.

### 3.2 Component 2 — Engagement Telemetry

**Contract:** Captures the event stream, batches to server, writes xAPI statements, computes the engagement score.

**xAPI verb mapping:**

| Event | xAPI verb | Extension fields |
|-------|-----------|------------------|
| `docOpened` | `experienced` | `docId` |
| `pageChanged` | `experienced` | `page` |
| `dwellTick` | `progressed` | `dwellMs`, `wpm`, `scrollPct` |
| `textSelected` | `interacted` | `selectedText.length`, `section` |
| `annotationCreated` | `interacted` | `annotationId`, `kind`, `colour` |
| AI Q&A asked | `asked` (custom URI `http://dwx.firstdigital.com/xapi/verb/asked`) | `query`, `chunksUsed`, `responseLatencyMs` |
| AI Q&A cited | `interacted` | `citations[]` |
| Proactive check answered | `answered` | `correct`, `sectionPath` |
| `sessionEnded` | `terminated` | `totalDwellMs`, `pagesVisited` |

**Engagement score formula (per doc, per learner):**

```
engagement_score =
    0.4 * normalisedDwell        // clamped actual/expected per page
  + 0.2 * scrollCoverage         // % of text lines intersected > 400ms
  + 0.2 * interactionRate        // (selections + highlights + Q&A) per 10min, normalised
  + 0.2 * comprehensionScore     // rolling correct rate on proactive checks
```

**Default thresholds (can be overridden per-doc):**

- `< 0.35` → low engagement; tutor may proactively nudge
- `0.35 - 0.65` → active; no intervention
- `>= 0.65` AND `all_mandatory_sections_have_dwell > 0` → **gate unlocks**

**Storage:** xAPI statements into the app's xAPI list (LearnIQ: `LMS_xAPIStatements`); engagement score into the app's insights list (LearnIQ: `LMS_AIInsights`).

**Privacy rule:** learner identifier is always a reference code (`LRN-XXXX`, `EMP-XXXX`, `LAW-XXXX`), never a name or email. PII scrubbing happens on the Function boundary before OpenAI sees anything.

### 3.3 Component 3 — Document Index

**Contract:** One-off server-side ingestion that makes a single document retrievable with page-anchored citations.

**Pipeline:**

1. Source upload → object storage (SharePoint library + Azure Blob shadow copy)
2. **Azure Document Intelligence `prebuilt-layout`** in markdown mode — returns markdown with headings *and* bounding polygons per paragraph, per line, per word
3. `MarkdownHeaderTextSplitter → RecursiveCharacterTextSplitter` at **512 tokens / 128 overlap**, preserving `sectionPath` metadata
4. Embed with **Azure OpenAI `text-embedding-3-large`** (trimmable to 1536 dim if index growth is a concern)
5. Index fields:

```
  documentId           string     // stable doc identifier
  productCode          string     // 'LearnIQ' | 'PolicyIQ' | 'ContractIQ' — enables per-app filter
  corpusKey            string     // 'course:POPIA-2026-001' | 'policy:FP-4.2' | 'matter:WW/AcmeCo/M&A'
  pageNumber           int
  boundingPolygon      json       // for citation → coordinate jump
  sectionPath          string     // 'Part III > Section 12 > (3)'
  content              string     // the chunk text (BM25)
  contentVector        vector     // embedding (kNN)
  authorReviewed       bool       // only served if true
  dataClassification   string     // 'public' | 'internal' | 'confidential' | 'privileged'
  tenantId             string     // for multi-tenant isolation
```

6. **Sidecar JSON** in Blob per doc: `{pages: [{num, w, h}], paragraphs: [{id, page, polygon, text, sectionPath}]}` — the renderer downloads this once at open and resolves citations to screen coordinates without a round trip.

**Retrieval at query time:**

```
Hybrid (BM25 + vector kNN=50) + L2 semantic reranker
Filter: documentId eq '{docId}' and authorReviewed eq true
top=8, select="content,pageNumber,boundingPolygon,sectionPath"
```

**Escalation to agentic retrieval** (Azure AI Search preview feature) on compound queries — gate behind a complexity heuristic: query length > 25 tokens *or* compare/contrast/multi-entity cues.

**Rule:** one Azure AI Search index per `productCode+corpusKey` grouping, never a global index. Tenant isolation is achieved through `tenantId` filters *plus* index partitioning. See §3.7 tenancy rules.

### 3.4 Component 4 — Artefact Pipeline

**Contract:** At ingestion time (and on doc update), generate a fixed set of reusable learning artefacts and save them as *Draft* (never auto-published).

**Base artefact types (all three apps):**

| Artefact | Typical use | Model | Author review |
|----------|-------------|-------|---------------|
| Executive summary | Quick-read substitute for full doc | gpt-4o | Required |
| Plain-English explainer | For non-specialist audiences | gpt-4o-mini | Required |
| Glossary (auto-extracted) | Terminology reference | gpt-4o-mini | Required |
| Key sections index | Click-through navigation | deterministic | Not required |
| Audio overview (2-host podcast) | Commute-friendly | Azure Speech + gpt-4o | Required |

**App-specific artefacts:**

LearnIQ adds:
- Flashcard deck (12-20 cards, spaced-rep ready) — gpt-4o-mini
- Scenario quiz (5-10 scenarios, AI-graded) — gpt-4o
- "What would the Regulator ask?" critique (14 audit questions + model answers) — gpt-4o

PolicyIQ adds:
- Applicability checklist ("does this apply to my role?") — gpt-4o-mini
- Attestation questionnaire (3-5 questions signed attester must get right) — gpt-4o
- "Top 10 mistakes I'd catch in practice" — gpt-4o

ContractIQ adds:
- Clause index with risk score — gpt-4o
- Playbook alignment report — gpt-4o with RAG on firm playbook
- Deviation heatmap (vs firm standard language) — gpt-4o
- Opposing-counsel critique ("what would the other side push back on?") — gpt-4o
- Redline suggestions (auto-drafted markup) — gpt-4o with playbook grounding

**Author review gate (MANDATORY):**

- Every artefact is saved as **`Status: Draft`** and carries `authorReviewed: false`
- Artefacts are **never served to end users** until an author has explicitly reviewed them
- Review UI provides: accept / accept-with-edits / regenerate / reject, per artefact
- Every accept/reject decision logged to audit trail

**Cost optimisation:**

- **Prompt caching**: the document markdown is the common prefix for all artefacts — structure prompts with document first, instruction last. One ingest, five+ artefacts reuses cached prefix.
- **Batch API**: 50% discount on non-interactive artefacts (run overnight at ingest).
- **Model routing**: gpt-4o-mini for deterministic/mechanical artefacts, gpt-4o for generative/scenario ones.
- **De-dup by content hash**: if document hasn't changed, skip regeneration.

Typical cost per 50-page legal document across all LearnIQ artefacts: **~$0.30 one-off**.

### 3.5 Component 5 — Tutor Agent

**Contract:** A streaming, grounded chat agent with tool use, accessible from the copilot panel alongside the document.

**Core behaviours:**

1. **Grounded-or-refuse**: if no chunk scores above a relevance threshold, the agent says "I don't have a reliable source for that in this document — would you like me to check the firm's wider policy corpus?" instead of hallucinating.
2. **Citations are non-optional**: every factual claim references at least one chunk's `{page, polygon, sectionPath}`.
3. **Streaming via Server-Sent Events** from the Function App to the client. Tokens appear as they generate.
4. **Tool use**: the agent has a defined toolbox per app (see §2 table). Tool calls are logged verbatim in the audit trail.
5. **Persona awareness**: a `persona` parameter — `default` / `basic` / `regulator` / `client` — changes tone and depth but never changes factual content.

**Framework recommendation:** Semantic Kernel Process Framework (.NET 8) — matches Azure Functions runtime, native Azure OpenAI + Azure AI Search connectors, stateful and streaming, human-in-the-loop ready. Alternative: LangGraph (Python) in a separate Function app if team is Python-first.

**Agent node graph (applies to all three apps):**

```
             ┌──────────────┐
  user turn ─▶│  Router      │─▶ intent = qa | explain | compare | draft | navigate
             └──────┬───────┘
                    │
       ┌────────────┼─────────────┬──────────────┬────────────┐
       ▼            ▼             ▼              ▼            ▼
   ┌───────┐  ┌──────────┐  ┌──────────┐  ┌──────────┐  ┌─────────┐
   │Retrieve│  │CompareTool│  │DraftTool│  │Navigate │  │Escalate │
   └───┬───┘  └─────┬────┘  └─────┬────┘  └────┬─────┘  └────┬────┘
       │            │             │            │             │
       └────────────┴──────┬──────┴────────────┴─────────────┘
                           ▼
             ┌──────────────────────────┐
             │ Grounded Answer Composer │
             │  - cites {page, polygon} │
             │  - refuses if ungrounded │
             └────────────┬─────────────┘
                          ▼
                   Stream to client
                          ▼
                 Telemetry + Audit
```

**System prompt template (base — per-app overrides in product prompt files):**

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

Tools available: {toolList}
```

**Refusal and escalation:**

The tutor must refuse in three specific cases:

1. **Ungrounded** — no chunk above relevance threshold. Offer fallback corpus or escalation to human owner.
2. **Privileged-document flag** — if the document carries `privilegeWarning: true`, prepend a privilege disclaimer and disable `compare` / `regulator-critique` tools.
3. **High-stakes question** — if the agent classifies the question as regulatory, disciplinary, or litigation-exposure, the response is drafted but gated: "Please review with the supervising partner before relying on this."

### 3.6 Component 6 — Proactive Tutor

**Contract:** A separate, cheap server-side loop that watches the telemetry stream and fires a polite interrupt when engagement signals indicate the reader isn't actually reading.

**Rules-first design (to keep cost low):**

```
Every 30 seconds the client POSTs a batch of events.
Server evaluates deterministic rules first (free).
Only if a rule fires does the server call gpt-4o-mini to compose an interrupt.

Rules:
- wpm > 400 on a page with > 300 words                          → "speed"
- 3+ consecutive pages with < 1s dwell                          → "skim"
- zero selections + zero mouse movement over 3 pages            → "absent"
- scroll-past rate > 40% across session                         → "avoidance"
- user opened doc, visited only last page, then clicked unlock  → "shortcut"
```

**Interrupt UX contract:**

- Never steals focus abruptly; shows as a badge + soft chime + optional modal
- Always dismissible ("Skip this check"); dismissals are logged
- Always polite ("Quick check — you're moving fast through this"); never scolding
- Always section-specific (references the actual section the reader just sped through)
- Answer recorded in `LMS_AIInsights` (or app-specific equivalent) and feeds back into the `comprehensionScore` component of engagement
- **Never blocks** the user from continuing — they can always dismiss and read on

**Per-app triggers (override defaults):**

- LearnIQ: default triggers
- PolicyIQ: add "opened but didn't read applicability section for their role" → "This section applies to you specifically — worth a read?"
- ContractIQ: add "known risky clause in this template hasn't been dwelt on" → "Just flagging — this indemnity clause has a known deviation. Want me to walk through it?"

### 3.7 Component 7 — Engagement Score + Comprehension Attestation

**Contract:** The computed engagement score gates the app's next step, and produces an audit-grade evidence package.

**Gate semantics:**

- LearnIQ: `engagement_score >= 0.65 AND all_mandatory_pages_have_dwell > 0` → assessment unlocks
- PolicyIQ: same threshold + all applicability-checklist items answered → "I have read and understood" button unlocks
- ContractIQ: same threshold + all risky-clause flags acknowledged → partner signoff workflow advances

**Audit export package (all three apps generate the same structure):**

```json
{
  "documentId": "FP-4.2-v3.1",
  "learnerRef": "LRN-0042",
  "sessionStart": "2026-04-17T08:32:14Z",
  "sessionEnd":   "2026-04-17T09:14:58Z",
  "totalDwellMs": 2521000,
  "pagesVisited": [1, 2, 3, 4],
  "engagementScore": 0.78,
  "components": {
    "normalisedDwell":     0.82,
    "scrollCoverage":      0.91,
    "interactionRate":     0.64,
    "comprehensionScore":  0.75
  },
  "interactions": {
    "aiQuestionsAsked": 5,
    "proactiveChecksAnswered": {"correct": 2, "incorrect": 1, "skipped": 0},
    "highlightsCreated": 4,
    "notesCreated": 2
  },
  "citationsFollowed": [
    {"fromMsgId": "m7", "toPage": 2, "toSection": "3.2"}
  ],
  "attestation": {
    "result": "PASS",
    "comprehensionCertificateId": "CC-POPIA-2026-001-0042",
    "signedByPipelineHash": "sha256:a7b4...",
    "timestamp": "2026-04-17T09:14:58Z"
  },
  "piiScrubbing": {
    "userQueriesScanned": 7,
    "userQueriesRedacted": 1,
    "redactionsByCategory": {"clientName": 1}
  }
}
```

This package is what an auditor asks for. All three apps produce it — only the `attestation.result` semantics differ (LearnIQ: CPD cert; PolicyIQ: attestation; ContractIQ: clause acknowledgement).

### 3.8 Multi-tenancy, privacy, and compliance (cross-cutting)

These rules apply to every implementation of this pattern.

1. **One Azure AI Search index per `productCode + corpusKey`** (LearnIQ: per-course; PolicyIQ: per-policy-family; ContractIQ: per-matter or per-playbook). Never a global index.
2. **`tenantId` filter on every query** — enforced server-side, never trusted from client payload.
3. **PII scrubbing on the Function boundary** — Azure AI Language PII detection skill + regex before the OpenAI call. If PII detected, user is warned and either allowed to rewrite or the query is blocked per app policy.
4. **LRN-XXXX codes never names** in any prompt or log that touches OpenAI.
5. **Data residency**: Azure SA North for storage, Azure OpenAI Sweden Central (or equivalent approved region). Per tenant override available.
6. **Audit log row per AI interaction** — immutable, exportable via the app's audit export endpoint.
7. **Privilege flag** on documents — when set, tutor prepends disclaimer and disables comparative / regulator tools.
8. **POPIA / GDPR deletion workflow** — when a learner deletion is requested, all Q&A transcripts, engagement scores, and audit rows for that learner-ref are purged within 30 days.
9. **Guardian tier gating** — all apps use the existing DWx Guardian service to gate tiers (Free / Professional / Enterprise). Base Q&A is Professional; Proactive Tutor and Enterprise-grade artefacts (Regulator Critique, Opposing-Counsel Critique, Deviation Heatmap) are Enterprise.

---

## 4. Per-app profiles

### 4.1 LearnIQ profile

**Surface:** Lesson type `InteractivePDF` (and later `InteractiveEbook`) inside [LmsCoursePlayer](../src/webparts/lmsCoursePlayer).
**Corpus key pattern:** `course:{CourseCode}`
**Audit target:** LPC / SRA CPD
**Gate:** engagement_score ≥ 0.65 → assessment unlocks → on pass, certificate generated server-side via `/api/generate-certificate`
**App-specific tools:** `generate_quiz_question`, `log_misconception`, `recommend_next_lesson`
**Comprehension Certificate format:** tamper-evident PDF signed with pipeline hash, attached to LMS_Certificates.

Reference implementation: this repository. See [interactive-pdf-research-and-plan.md](interactive-pdf-research-and-plan.md) for full build plan.

### 4.2 PolicyIQ profile

**Surface:** New `PolicyTutor` page/web-part. Accessible from a policy library, from an email link, or embedded in an HR/onboarding flow.
**Corpus key pattern:** `policy:{PolicyId}` with optional `fallbackCorpus: firm-policies`
**Audit target:** POPIA §18 (subject access), internal compliance committee, ISO 27001 / SOC 2 evidence
**Gate:** engagement_score ≥ 0.65 AND applicability checklist complete → "I have read and understood" button unlocks → signed attestation stored with tamper-evident signature
**App-specific tools:** `search_firm_policy_corpus`, `check_applicability_to_role`, `flag_for_policy_owner`
**Per-policy metadata:**
```
  policyId, version, owner, effectiveDate, nextReviewDate
  attestationRequired: boolean
  attestationRoles: string[]        // which roles must attest
  attestationCadence: 'annual' | 'on-change' | 'once'
  privilegeWarning: boolean
```
**Unique mechanics:**
- **Attestation cadence tracking** — if a policy version changes, all prior attesters are flagged for re-attestation automatically.
- **Role-scoped applicability** — tutor opens by asking the user their role (or reads it from AD), then surfaces only the sections that apply to them. Irrelevant sections are collapsed with "Not applicable to your role" markers.
- **Policy-change impact agent** — when a new policy version is uploaded, the pipeline diffs against the previous version and generates a "What changed" executive summary that the tutor proactively surfaces on first open.

### 4.3 ContractIQ profile

**Surface:** New `ContractTutor` page/web-part. Opens from a matter record, from a contract upload, or from a DMS link.
**Corpus key pattern:** `contract:{MatterRef}/{DocId}` with `fallbackCorpus: firm-playbook:{PracticeArea}`
**Audit target:** Matter file, partner review trail, E&O insurance defensibility
**Gate:** engagement_score ≥ 0.65 AND all `risky-clause` flags acknowledged → partner signoff workflow advances; deviation report attached to matter file
**App-specific tools:** `search_playbook`, `compare_to_precedent`, `draft_markup`, `flag_deviation`, `run_negotiation_simulation`
**Per-contract metadata:**
```
  matterId, clientRef, dealType, practiceArea, role (firm-side | counterparty)
  playbookId, playbookVersion
  comparandaCorpusKey: string       // past deals of same type
  privilegeWarning: true (default)  // contracts are usually privileged
```
**Unique mechanics:**
- **Playbook alignment scoring** — each clause is scored against firm playbook standard language. Deviations get a risk score (green/amber/red).
- **Counterparty-clause detection** — tutor identifies clauses the counterparty added vs firm-standard, and surfaces them first.
- **Negotiation rehearsal mode** — multi-turn agent role-plays as opposing counsel; junior lawyer practises counter-arguments; session is saved and partner can review.
- **Redline drafting tool** — on deviation flag, tutor can draft the redline inline; partner approves/edits before it's written back.
- **Institutional knowledge feed** — every partner-approved redline contributes (anonymised, firm-internal only) to the firm's playbook for next time.

---

## 5. Reusable building blocks

These should be built as shared libraries, not re-implemented per app.

### 5.1 Shared npm packages (suggested)

```
@dwx/document-renderer-core         — event contract, renderer interface, telemetry batcher
@dwx/document-renderer-pdf          — PDF.js v5 legacy wrapper implementing the interface
@dwx/document-renderer-ebook        — interactive-ebook wrapper (see ebook concept doc)
@dwx/document-tutor-client          — copilot panel, chat UI, citation chips, artefacts tabs
@dwx/document-tutor-types           — shared TypeScript types (ES5-safe)
@dwx/engagement-scorer              — scoring formula, threshold helpers
```

These packages all target React 17 + ES5 (for SPFx compatibility). A sibling `@dwx/*-react18` variant can exist for marketplace or standalone SPAs.

### 5.2 Shared Azure Function projects (suggested)

```
DWx.DocumentTutor.Ingestion         — Doc Intelligence + chunking + embedding + indexing
DWx.DocumentTutor.Retrieval         — hybrid search wrappers, reranking, fallback corpus logic
DWx.DocumentTutor.Agent             — Semantic Kernel process graph, tool registry
DWx.DocumentTutor.ProactiveTutor    — rules engine + LLM-escalated prompts
DWx.DocumentTutor.ArtefactPipeline  — summary, glossary, flashcards etc. generation
DWx.DocumentTutor.Telemetry         — xAPI writer, engagement scorer, audit logger
DWx.DocumentTutor.Shared            — DTOs, PII scrubber, privilege-flag handler
```

Each product app composes these with product-specific:

- Surface web parts / pages
- System prompts (per app, per persona)
- Tool registrations
- Artefact type set
- Audit export formatters

### 5.3 Shared endpoint contract

Every DWx Document Tutor implementation exposes the same core endpoints (prefixed per app):

```
POST /api/doc-ingest                   { source, productCode, corpusKey, metadata }
POST /api/doc-copilot                  { docId, messages, persona, tools }    → SSE stream
POST /api/doc-generate-artefact        { docId, artefactType }
POST /api/doc-telemetry                { docId, events[] }
POST /api/doc-proactive-tutor          { docId, currentPage, recentEvents }
GET  /api/doc-sidecar/{docId}          → sidecar JSON
POST /api/doc-audit-export             { docId, learnerRef, format }
```

LearnIQ prefix: `/api/` (existing).
PolicyIQ prefix: `/api/policy/` (recommended).
ContractIQ prefix: `/api/contract/` (recommended).

Each prefix resolves to its own Function app — but the *schema* of every endpoint is identical, and the underlying shared Function projects are common.

---

## 6. Pattern compliance checklist

To claim "DWx Document Tutor" compliance, an implementation must tick every item:

**Renderer:**
- [ ] Emits all 7 required events with the defined payload contract
- [ ] Implements `jumpToPage`, `highlightPolygon`, `showInterrupt` methods
- [ ] Respects the renderer/tutor separation — tutor must work unchanged if renderer is swapped

**Telemetry:**
- [ ] All events map to xAPI verbs per §3.2
- [ ] Engagement score uses the canonical formula
- [ ] Statements written to app's xAPI store with LRN-XXXX identifiers only

**Index:**
- [ ] One Azure AI Search index per `productCode+corpusKey`
- [ ] Hybrid search + L2 semantic reranker
- [ ] Per-paragraph bounding polygons stored; sidecar JSON served via endpoint
- [ ] `authorReviewed` and `dataClassification` fields present and filtered

**Artefacts:**
- [ ] Base set (summary, glossary, audio overview) generated at ingestion
- [ ] All artefacts saved as `Draft`, none auto-published
- [ ] Author review UI present with accept/edit/regenerate/reject
- [ ] Every review decision logged to audit trail

**Tutor agent:**
- [ ] Grounded-or-refuses on low-relevance retrieval
- [ ] Every factual claim carries a citation to `{page, polygon, sectionPath}`
- [ ] Streaming (SSE) from server to client
- [ ] Persona parameter implemented
- [ ] Privilege-flag behaviour implemented
- [ ] Tool use logged verbatim

**Proactive tutor:**
- [ ] Rules-first; LLM-escalated only on flag
- [ ] Polite, dismissible, section-specific
- [ ] Answers feed back into `comprehensionScore`
- [ ] Never blocks the user

**Gate + attestation:**
- [ ] Engagement score gates the next step
- [ ] Audit export package matches §3.7 JSON schema
- [ ] Tamper-evident signature on attestation outcome

**Compliance:**
- [ ] PII scrubber on Function boundary
- [ ] Tenant isolation via `tenantId` filter on every query
- [ ] Guardian tier gating wired
- [ ] POPIA/GDPR deletion workflow functional
- [ ] Data residency honoured per tenant config

---

## 7. Versioning and evolution

This specification is versioned. Breaking changes require a new major version and a migration note.

- **v1.0 (2026-04-17)** — initial specification, based on LearnIQ reference design
- Planned **v1.1** — multimedia ebook renderer profile (see [dwx-interactive-ebook-concept.md](dwx-interactive-ebook-concept.md))
- Planned **v1.2** — workflow-embedded just-in-time triggers (matter open, email thread mention, etc.)
- Planned **v1.3** — cross-document synthesis agent (answer spans multiple documents in the same corpus)

Compliant apps should declare the spec version they target, e.g. `DocumentTutor-Spec: 1.0` in their manifest.

---

## 8. Open questions for DWx architecture

Before PolicyIQ and ContractIQ start implementing:

1. **Shared package hosting** — Azure Artifacts feed under `dev.azure.com/gfinberg/DWx/_artifacts`? Or npm private scope?
2. **Shared Function project hosting** — single repo with multi-project solution? Or git submodules per app?
3. **Cross-app telemetry aggregation** — does the DWx Admin Portal aggregate engagement and audit data across LearnIQ/PolicyIQ/ContractIQ for the same tenant? If so, schema alignment matters from day one.
4. **Deployment Manager manifest schema** — add a `documentTutor` section to `IProductManifest.ts` so apps can declare their profile (renderer, corpus key pattern, artefact mix, tool registrations)?
5. **Guardian tier semantics** — are "Professional" and "Enterprise" tiers named the same across PolicyIQ/ContractIQ/LearnIQ, or do they have distinct names per product? Implementation-neutral is easier.
6. **Shared branding for the tutor** — "DWx Copilot" works for LearnIQ. Should it be the same name across the suite (consistent brand), or have per-product variants (e.g. "DWx Copilot for Policies")?

---

*End of specification.*
