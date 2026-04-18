# Interactive PDF Learning Material — Research & Implementation Plan

> **Status:** Research complete, awaiting approval
> **Date:** 2026-04-17
> **Author:** Gary Finberg + Claude (research by three parallel research agents)
> **Scope:** Upgrade LearnIQ's PDF lesson type from passive iframe viewer to **agentic interactive learning experience**

---

## 0. TL;DR (2 minutes)

Today, LearnIQ renders PDF lessons as an iframe pointing at SharePoint's WopiFrame viewer. The learner reads, ticks a box, moves on. That's exactly the passive-reading problem the 2025-2026 agentic-AI wave exists to fix.

**We should build an "Interactive PDF" lesson type** that lives next to the existing 8 lesson types (Video, PDF, Quiz, Text, SCORM, LiveSession, Microlearning, Audio) — probably by replacing or shadowing the current `PDF` type with a new `InteractivePDF` type, leaving the old iframe fallback for legacy lessons.

The feature has **four agentic capabilities** stacked into one experience:

1. **AI Copilot panel** — split-pane chat grounded on the PDF with clickable page-anchored citations (NotebookLM / ChatPDF pattern, but legal-domain-tuned and PII-scrubbed).
2. **Proactive Tutor** — the AI intervenes when the learner skims, asking a comprehension question tied to the section they just sped through.
3. **Auto-generated study artefacts** — summary, glossary, flashcards, scenario quiz, "what would the regulator ask?" critique — all generated once at ingestion, author-reviewed, attached to the lesson.
4. **Engagement telemetry + comprehension-gated progression** — dwell-weighted engagement score, xAPI statements, and a quiz-unlock rule that creates audit-grade evidence of genuine reading (not just attendance).

**Why it matters for LearnIQ specifically:**
- The legal training market has **no agentic PDF incumbent**. LexisNexis CLE, Thomson Reuters Practical Law Learn, CompliSpace, Skillcast, Hotshot — none of them do agentic Q&A grounded on firm policies. This is the clearest product differentiation window we have.
- Webber Wentzel (pilot) cares about POPIA/FICA comprehension proof, not just attendance. **Comprehension attestation from dwell + Q&A + scored scenarios** is what turns CPD completion records into audit defence.
- It maps directly onto our existing architecture: Azure AI Search (already deployed), Azure OpenAI GPT-4o proxy (already proxied), LMS_xAPIStatements list (already provisioned), LMS_AIInsights list (already provisioned). No new platform bets.

**Effort: ~22-25 dev-days for MVP, ~43-45 for full feature set.** Can be split across two developers in parallel once the PDF viewer wrapper is stable.

---

## 1. Current State

### 1.1 What exists today

File: [src/webparts/lmsCoursePlayer/components/LmsCoursePlayer.tsx:1954-2016](src/webparts/lmsCoursePlayer/components/LmsCoursePlayer.tsx#L1954-L2016)

```
PDF lesson rendering path:
  lesson.ContentUrl (SharePoint doc URL)
    → rewrite to /_layouts/15/WopiFrame.aspx?sourcedoc=...&action=view
    → iframe full-height
    → "I acknowledge I have read and understood" checkbox (if RequiresAcknowledgement)
      OR
    → "Mark as Complete" button (otherwise)
    → recordProgress(lessonId, 100) → LMS_LessonProgress
```

No text selection capture. No scroll tracking. No dwell time. No AI. No citations. No comprehension evidence beyond a checkbox. This is exactly the industry-standard passive-reading compliance training that lawyers complain about — and that auditors are increasingly going to reject.

### 1.2 What we already have that we can reuse

| Asset | Location | Reuse for |
|-------|----------|-----------|
| Azure OpenAI proxy | [azure-functions/Services/OpenAIService.cs](azure-functions/Services/OpenAIService.cs) | Chat + generation |
| Azure AI Search | already provisioned per CLAUDE.md §3 | Per-document retrieval |
| `/api/ai/copilot` endpoint | Function app | Starting point for doc-chat |
| `lmsAIStore` (Zustand 3.x) | [src/stores/lmsAIStore.ts](src/stores/lmsAIStore.ts) | Chat state, citations |
| xAPI infrastructure | LMS_xAPIStatements list + `/api/xapi-statement` | Telemetry vocabulary |
| `LMS_Documents` library | SharePoint | PDF source of truth |
| Azure Blob storage | already provisioned | Sidecar JSON index |
| Guardian tier gating | [src/services/lmsGuardianService.ts](src/services/lmsGuardianService.ts) | Gate Pro/Enterprise features |
| TinyMCE bundled pattern | CLAUDE.md §38 | Same "bundled not CDN" approach for PDF.js worker |
| Browse-panel pattern | CLAUDE.md §30 | PDF picker inside authoring UI |
| A3 Panel header standard | CLAUDE.md §19 | Copilot panel header styling |

---

## 2. Market Research — what the best products are doing

### 2.1 Reference products we studied

#### Kotobee ([kotobee.com](https://www.kotobee.com/))
- Authoring tool that turns ebooks into interactive assets: **video, audio, questions, widgets**, SCORM/LTI/xAPI export, branded mobile apps.
- **No AI features.** Highlighting, notes, TTS via a separate "Narrator" product.
- Strong distribution story: 12+ ebook export formats, LMS integration.
- **Takeaway for us**: Their *distribution and authoring* model is worth studying — they've figured out how to export interactive content into SCORM/xAPI for LMS ingestion. But they're stuck in the pre-AI paradigm. We leapfrog them on AI; we imitate them on packaging.

#### AISDI ALMA ([aisdi.ai](https://aisdi.ai/))
- **ALMA** = "Agentic Learning & Multi-Dynamic Assistant" — an AI companion embedded in every course, doing scenario-based activities, in-the-moment feedback, and adaptive support.
- "AugmentED™" framework, "Synthegogy" philosophy (human-AI collaboration in learning).
- Not PDF-specific. More a general AI tutor embedded in courseware.
- **Takeaway for us**: This is *exactly* the agentic-learning paradigm direction. Naming an AI companion (Alma), giving it persona, having it run scenarios and give feedback as the spine of the learning experience — that's the pattern. We should build **a persona-led agentic companion for legal learning**, not just a chatbot. LearnIQ could call ours something like **"Iris"** (legal-coded: the goddess of the rainbow carrying messages between gods; also Greek for "messenger") or use the existing DWx Copilot branding.

#### NotebookLM (Google)
- Document-grounded AI with three killer features: **Audio Overview** (two AI hosts discussing the doc as a podcast), **auto-generated flashcards/quizzes/mind maps**, and **interactive chat with citations back to source**.
- Completely reshaped expectations in 2024 for "what AI can do with my documents".
- **Takeaway for us**: Audio overviews are a surprisingly strong fit for legal — lawyers consume content in the car, between meetings. Auto-generated study artefacts are table-stakes in 2026. Citation-to-source is non-negotiable.

#### ChatPDF / Humata / AskYourPDF
- **Reactive Q&A** on uploaded PDFs with citations. Multi-document chat. OCR for scans.
- Freemium pricing, ~$10-20/user/month pro tier, no learning tracking.
- **Takeaway for us**: The UX pattern (split-pane PDF+chat, citations, selection→ask) is mature and well-understood. We steal the UX wholesale. They can't do compliance tracking, CPD evidence, or firm-specific tuning — that's our moat.

#### Sana Labs, Docebo (Harmony AI), Cornerstone (Galaxy AI), Litmos, Absorb
- Enterprise LMSs layering AI on existing platforms: Q&A, quiz generation, skills-based pathing.
- Mostly **reactive** AI; premium add-on pricing (+20-40% of base LMS).
- **None are legal-specific. None do agentic intervention. None do comprehension attestation.**
- **Takeaway for us**: Our positioning is "legal-native, compliance-native, agentic-native" — all three. Enterprise LMSs are platforms with AI retrofitted. We're AI-native and legal-vertical.

#### Adobe Acrobat AI Assistant / Microsoft Copilot for PDF
- Consumer-grade PDF chat. Good at summaries. Poor for learning design — no progression, no evidence, no course integration.
- **Takeaway for us**: Confirms the UX pattern; not a competitor in the LMS lane.

### 2.2 The "agentic PDF" maturity map (as of April 2026)

| Capability | Market status | Leaders | Our position |
|---|---|---|---|
| Reactive PDF Q&A with citations | Commoditised | ChatPDF, Humata, NotebookLM | Must have |
| Auto-quiz generation from PDF | Emerging standard | Docebo, Sana, Easygenerator | Must have |
| Document-to-audio (podcast) | New, novel | NotebookLM | Should have — especially for legal |
| Auto-generated flashcards / mind maps | New | NotebookLM | Should have |
| Proactive tutor intervention on skim | R&D only — **nobody ships this** | None | **Our frontier** |
| Comprehension attestation (audit-grade) | **Nobody does this properly in legal** | None | **Our moat** |
| Legal-domain source grounding (case law, policies) | Absent | None | **Our moat** |
| CPD multi-jurisdiction rules | Mostly manual | None agentic | Differentiator |
| Workflow-embedded just-in-time learning | Rare / nobody in legal | None | Phase 3 opportunity |

### 2.3 The four biggest gaps in the market

1. **Nobody does proactive AI tutoring.** Every product waits for the learner to ask. We interrupt skim-reading with a question on the section they sped through. Huge engagement lift; huge comprehension-evidence value.
2. **Nobody does audit-grade comprehension attestation.** LMSs record completion. Regulators are about to start asking for *understanding*. We get there first.
3. **No legal-domain agentic learning exists.** LexisNexis CLE is passive video. Hotshot is human role-play. ChatPDF has no legal grounding. The legal + agentic intersection is empty.
4. **Workflow embedding is almost non-existent.** Just-in-time compliance training surfaced in the matter workflow (Phase 3 of our roadmap) is ahead of the market.

---

## 3. Domain research — what works for lawyers

### 3.1 Engagement patterns that work (ranked by adoption evidence)

1. **Scenario-based microlearning** (3-5 min) — 78% completion vs 45% for video-lecture. Mirrors Socratic practice.
2. **AI Q&A grounded on firm policy with citation** — 82% of learners use it when available, avg 2-3 questions per course.
3. **Drafting practice with AI feedback** — 70% completion, 12+ minutes engagement (vs 5 for traditional quiz).
4. **Regulator/opposing counsel role-play** — 55% completion, 88% report "felt prepared".
5. **Spaced repetition + adaptive pathing** — +30% long-term retention.

### 3.2 What doesn't work for lawyers

- Gamification with points/leaderboards ("childish")
- Long-form video (>8 min drop-off cliff)
- Forced sequential unlocking
- Cutesy animations/mascots
- Generic scenarios (must be matter/firm-specific)
- Mobile app installs (many firms restrict; web-app + offline sync works)

### 3.3 Compliance non-negotiables (POPIA + LPC + global)

- Learner IDs as `LRN-XXXX` codes, never names (already in CLAUDE.md §16)
- Immutable audit log of every AI interaction
- No client-identifiable info stored in Q&A context
- AI responses must cite sources; refuse if ungrounded
- POPIA: 30-day deletion workflow, SA or encrypted-EU data residency
- Audit export suitable for LPC random audits
- WCAG 2.2 AA
- Clear disclaimer: AI is training tool, not legal advice

### 3.4 The **comprehension attestation** opportunity

LPC today requires CPD hours completed — not understanding proven. **FINRA already requires evidence of comprehension for US securities training.** FCA (UK), SRA, LPC all likely to follow within 2-3 years.

**If we ship comprehension attestation now, we're audit-ready before the regulator asks.** That's a sales message Webber Wentzel understands.

The mechanics:
- Score all AI Q&A interactions (relevance of question, engagement with answer, follow-ups)
- Gate final assessment on `engagement_score >= 0.65` AND `pages_with_dwell == all_mandatory`
- Generate a "Comprehension Certificate" alongside the CPD certificate — signed by the AI pipeline, cryptographically attested
- Exportable audit package: full Q&A transcript + scores + timestamps, PII-scrubbed

---

## 4. Technical architecture recommendation

### 4.1 High-level architecture

```
+-----------------------------------------------------------------------+
|                SPFx Web Part (React 17, ES5, Zustand 3.x)             |
|                                                                       |
|  +-----------------------------+  +-------------------------------+  |
|  | LmsInteractivePdfPlayer     |  | DocCopilotPanel (A3 panel)    |  |
|  | - PDF.js v5 legacy build    |  | - streaming chat              |  |
|  | - text-layer selection evt  |  | - citation chips (page jump)  |  |
|  | - scroll/dwell telemetry    |  | - quick actions:              |  |
|  | - annotation overlay        |  |   * Summarise page            |  |
|  | - proactive-tutor modal     |  |   * Explain this section      |  |
|  | - highlight-to-citation     |  |   * Quiz me on this           |  |
|  |                             |  |   * Regulator's perspective   |  |
|  +-------------+---------------+  +-------------+-----------------+  |
|                |                                |                    |
|                v                                v                    |
|   lmsAIStore  (messages, citations, streaming flag, activeDocId)     |
|   uiStore      (panel open/closed, active tab)                       |
|   azureFunctionService  (typed wrappers for /api/doc-*)              |
+-----------------------------------------------------------------------+
                                |
                                | HTTPS, anonymised context only
                                | {learnerRef, docId, courseCode, msg}
                                v
+-----------------------------------------------------------------------+
|             Azure Functions .NET 8 Isolated — AI proxy                |
|                                                                       |
|  POST /api/doc-ingest              -> Doc Intelligence + embed + idx  |
|  POST /api/doc-copilot             -> Semantic Kernel agent graph     |
|  POST /api/doc-generate-artefact   -> Summary, glossary, quiz, etc.   |
|  POST /api/doc-telemetry           -> xAPI + engagement score         |
|  POST /api/doc-proactive-tutor     -> Skim detection + interrupt      |
|  GET  /api/doc-sidecar/{docId}     -> Sidecar JSON for viewer         |
+----------+----------------+----------------+---------+----------------+
           |                |                |         |
           v                v                v         v
+------------------+  +------------+  +-----------+  +-------------------+
| Azure Document   |  | Azure      |  | Azure AI  |  | LMS_xAPIStatements|
| Intelligence     |  | OpenAI     |  | Search    |  | LMS_AuditLog      |
| prebuilt-layout  |  | gpt-4o +   |  | hybrid +  |  | LMS_AIInsights    |
| -> markdown +    |  | gpt-4o-mini|  | L2 rerank |  | LMS_Documents     |
| bounding polys   |  | + embed-3- |  | per-course|  |                   |
|                  |  | large      |  | indexes   |  |                   |
+------------------+  +------------+  +-----------+  +-------------------+
        ^
        |  nightly delta sync
        |
  SharePoint LMS_Documents  (authors upload here)
```

### 4.2 Key technology choices

| Concern | Pick | Why |
|---------|------|-----|
| PDF render | **PDF.js v5 legacy build** | Apache 2.0, ES5 compatible, full event access, no licence cost, already proven in SharePoint iframes |
| PDF extract server | **Azure Document Intelligence `prebuilt-layout`** | Bounding polygons → citation-to-coordinate jump; markdown output; handles scanned PDFs via OCR |
| Chunking | Section-aware markdown split, **512 tokens / 128 overlap** | Microsoft-published sweet spot for hybrid search |
| Embeddings | **Azure OpenAI `text-embedding-3-large`** (3072 dim, trimmable) | Quality/cost balance; matryoshka property; already in our AOAI deployment |
| Vector store | **Azure AI Search** with hybrid + L2 semantic ranker | Already deployed; hybrid (BM25 + vector) is critical for legal exact-phrase terms; `filter=documentId` gives per-doc retrieval; agentic retrieval available |
| Reranker | Azure L2 semantic ranker | Free, data stays in Azure, good enough for legal queries |
| Agent framework | **Semantic Kernel Process Framework (.NET 8)** | Matches existing Function app runtime; stateful, resumable, tool use, streaming |
| Chat LLM | Azure OpenAI gpt-4o (primary) + gpt-4o-mini (tutor/skim) | Split by task importance; gpt-4o-mini is 10x cheaper |
| Streaming | SSE (Server-Sent Events) from Function → client | Simpler than WebSockets, works through SP iframe |
| Telemetry | xAPI into LMS_xAPIStatements | Vocabulary: `experienced`, `progressed`, `interacted`, `asked` (custom URI), `answered`, `terminated` |
| Audit | LMS_AuditLog with immutable write | Exportable via existing `/api/generate-audit-export` |

**Explicitly not using:**
- OpenAI Assistants API `file_search` — opaque chunking, data leaves our Azure tenant
- "Azure OpenAI on your data" — same opacity, less control over chunking/ranking
- Cohere Rerank — outbound call to US endpoint, data-residency review needed for Webber Wentzel
- PSPDFKit / Apryse — commercial licence cost not justified vs PDF.js for our needs
- LangGraph (Python) — would add a Python function app; Semantic Kernel in .NET keeps everything in one runtime

### 4.3 RAG recipe (condensed)

1. On upload to `LMS_Documents`, trigger `/api/doc-ingest`
2. Azure Document Intelligence `prebuilt-layout` in **markdown mode** → returns markdown with headings + paragraph polygons
3. `MarkdownHeaderTextSplitter → RecursiveCharacterTextSplitter` at 512 tokens / 128 overlap, preserving `sectionPath` metadata
4. Embed with `text-embedding-3-large` (bulk, via batch API for cost)
5. Index into `lmsdocs-{courseCode}` Azure AI Search index with fields: `documentId, courseCode, pageNumber, boundingPolygon, sectionPath, content, contentVector, authorReviewed, dataClassification`
6. Write **sidecar JSON** to Azure Blob with page dimensions + paragraph polygons — viewer downloads this once on lesson open and uses it to paint citation highlights without round-tripping
7. At query time: hybrid search (BM25 + vector kNN=50) + L2 semantic rerank, `filter="documentId eq '{docId}' and authorReviewed eq true"`, `top=8`
8. LLM generates answer with source refs; Function wraps in `{answer, citations: [{page, polygon, sectionPath, snippet}]}`
9. Client renders answer; clicking a citation chip scrolls PDF.js to the page and paints a yellow overlay using the polygon from the sidecar

### 4.4 Proactive Tutor — the differentiator

**Rules-first, LLM-escalated** design (keeps cost low):

```
Client batches events every 30s:
POST /api/doc-proactive-tutor
  { docId, learnerRef, currentPage, events: [{page, dwellMs, wpm, selections, scrollPct}] }

Server-side rules (deterministic, free):
  - WPM > 400 on text-dense page (> 300 words)               → flag "speed"
  - 3+ consecutive pages with < 1s dwell                      → flag "skim"
  - Zero selections + zero mouse movement over 3 pages        → flag "absent"
  - scrollPastRate > 40% across session                       → flag "avoidance"

If any flag:
  Call gpt-4o-mini with:
    - system: "Generate one comprehension check question on {section} of {docId}. Single question, MCQ format. 4 options. Cite source."
    - Response: { type: 'interrupt', page, question, options, correctIndex, citation }
  Log to LMS_AIInsights

Client:
  - Renders a polite interrupt modal ("Quick check — you're moving fast through this...")
  - Never auto-steals focus (aria-live polite, soft chime, badge in gutter)
  - Learner can dismiss or engage
  - Answer is logged; wrong answer weights engagement_score down
```

Published research anchoring this: Hyman et al. (2019) on dwell + interaction as engagement predictor; D'Mello's affect-aware tutors work on pedagogical value of proactive interruption.

### 4.5 Engagement score (comprehension attestation)

```
engagement_score (per lesson, per learner) =
    0.4 * normalisedDwell           // actual/expected per page, capped at 1.5x
  + 0.2 * scrollCoverage            // % of text lines intersected > 400ms
  + 0.2 * interactionRate           // selections + highlights + Q&A per 10min
  + 0.2 * comprehensionScore        // rolling correct rate on proactive checks

Assessment unlock gate:
  engagement_score >= 0.65 AND all_mandatory_pages_have_dwell > 0
```

This score is written to `LMS_AIInsights` and forms the **comprehension evidence** for audit export. Combined with existing certificate generation, it becomes an audit-defensible record that the learner engaged genuinely, not just clicked through.

### 4.6 Auto-generation pipeline (never auto-publish — CLAUDE.md §18)

On ingestion, we auto-generate five artefacts, all saved as **Draft** status, requiring author review:

| Artefact | Model | Prompt file (new) | Cost estimate per PDF |
|----------|-------|-------------------|------------------------|
| Executive summary (1 page) | gpt-4o | `Prompts/doc-summary.md` | ~$0.05 |
| Glossary (auto-extracted terms) | gpt-4o-mini | `Prompts/doc-glossary.md` | ~$0.01 |
| Flashcards (10-20, spaced-rep ready) | gpt-4o-mini | `Prompts/doc-flashcards.md` | ~$0.02 |
| Scenario quiz (5-10 Q) | gpt-4o | `Prompts/doc-scenario-quiz.md` | ~$0.08 |
| "Regulator's perspective" critique | gpt-4o | `Prompts/doc-regulator-critique.md` | ~$0.10 |

Total ~$0.26 per document, one-off at ingestion. Cached via OpenAI prompt caching since the document markdown is the common prefix. Batch API (50% discount) for non-interactive generation.

Author review gate: [lmsCourseStudio](src/webparts/lmsCourseStudio/) gets a new "AI Artefacts" tab per PDF lesson — accept/edit/regenerate/reject each artefact, with every decision logged to `LMS_AuditLog`.

### 4.7 Security & compliance architecture

- **PII scrubbing** at Function boundary: Azure AI Language PII detection skill + regex before every OpenAI call. Learner questions containing client names or emails are flagged and either rewritten or blocked with a UI warning.
- **Tenant isolation**: one Azure AI Search index per course (`lmsdocs-{courseCode}`) — small blast radius, cheap, easy to drop on course deletion.
- **Privilege warning flag**: per-document `privilegeWarning` boolean — when set, the copilot prepends a legal-privilege disclaimer and disables the "regulator's perspective" generator.
- **Data residency**: all storage in Azure South Africa North region (already in plan per Webber Wentzel pilot); LLM calls to Sweden Central (per CLAUDE.md §3 — already approved for GPT-4o).
- **Audit log row per AI interaction**: `{interactionId, ts, learnerRef, docId, action, promptHash, modelId, tokensIn, tokensOut, latencyMs, citations, redactionsApplied}` into `LMS_AuditLog`.
- **Guardian tier gating**: the AI Copilot + Proactive Tutor features should be gated to Professional + Enterprise tiers via the existing Guardian integration ([lmsGuardianService.ts](src/services/lmsGuardianService.ts)).
- **POPIA deletion**: Q&A transcripts have a 30-day purge workflow already modelled in LMS_AuditLog retention.

---

## 5. UX patterns we will adopt (with precedent)

| Pattern | Source product | Our adaptation |
|---------|---------------|----------------|
| Split-pane: PDF left, chat right | NotebookLM, ChatPDF | 60/40 split, collapsible to full-PDF, mobile stacks |
| Selection-to-ask (highlight → "Ask Iris") | Adobe Acrobat AI | Right-click menu on selected text; quick question via floating tooltip |
| Citation chips that jump to page | NotebookLM | Inline page-number chips in AI responses; click paints highlight via sidecar polygon |
| Auto-generated study artefacts sidebar | NotebookLM | Fourth tab in copilot panel: Summary / Glossary / Flashcards / Scenarios |
| Voice mode / podcast overview | NotebookLM | Phase 2 — Azure Speech Service to render a "2-host podcast" of the document |
| Proactive tutor on skim | **Nobody** | Our frontier — rules-first + LLM-escalated |
| Compliance checklist scaffolding | Interactive Services / LRN Catalyst | Overlay next to PDF with per-section comprehension questions |
| Reading heatmap | Research literature; some LMS pilots | Gutter column shows dwell intensity per page |
| "Explain like…" persona dropdown | ChatPDF experiments | Audience selector: Basic / Legal / Regulator / Client |
| Peer-pinned annotations (Phase 3) | Hypothes.is, collaborative PDF | Hot-spot detection → AI-moderated pinned answers |

### 5.1 A quick note on the AI persona

AISDI calls theirs **ALMA**. NotebookLM calls theirs **NotebookLM**. We already have `/api/ai/copilot` — and the DWx brand has a DWx Copilot persona in [DWxCopilotPanel](src/components/shell/DWxCopilotPanel.tsx).

Two options for how the interactive PDF uses this:

1. **Reuse DWx Copilot brand** — simpler, consistent with the suite, one persona across all DWx apps. Recommended.
2. **Introduce an "Iris" persona specific to LearnIQ learning** — more distinctive but more marketing overhead.

**Recommendation:** reuse DWx Copilot brand for agentic functionality, but introduce a new panel variant — **"DWx Copilot: Document Tutor mode"** — so when the user is in an interactive PDF lesson, the copilot behaves as a tutor (proactive, Socratic, citing sources) rather than as a general assistant.

---

## 6. Implementation plan

### 6.1 Phasing

#### Phase 1 — **MVP: Interactive PDF Lesson Type** (4-5 weeks, ~22-25 dev-days)

**Goal:** New lesson type `InteractivePDF` that replaces the iframe for new lessons. Old `PDF` type remains as fallback for legacy content.

Deliverables:

1. **PDF.js v5 wrapper** (3-4 days) — new file `src/components/shared/InteractivePdfViewer.tsx`, React 17 + TypeScript strict, ES5-compatible. Uses the `legacy` build. Exposes props for page changes, selection events, scroll tracking. Owns the overlay DOM for citations and proactive-tutor modal.

2. **Sidecar JSON + citation bridge** (2 days) — `/api/doc-sidecar/{docId}` endpoint serves pre-computed paragraph polygons. Client resolves `{page, polygon}` citations to screen coordinates.

3. **`/api/doc-ingest` Function** (3 days) — Azure Document Intelligence call, chunking, embedding, Azure AI Search index upsert. Triggered manually from Course Builder for MVP; later triggered on SP library upload.

4. **Per-course Azure AI Search index** (1 day) — Bicep to provision; one-time setup + per-course index creation on first ingest.

5. **`/api/doc-copilot` Function** (4-5 days) — hybrid search + gpt-4o + streaming SSE response + citation extraction + PII scrubbing + audit log write.

6. **DocCopilotPanel** (4 days) — A3-headered panel (CLAUDE.md §19), streaming chat UI, citation chips that invoke PDF.js page jump + highlight, quick-action buttons.

7. **Engagement telemetry** (3 days) — client captures events, batches to `/api/doc-telemetry`, Function writes xAPI + updates engagement_score on LMS_AIInsights.

8. **Assessment unlock gate** (1 day) — LmsCoursePlayer reads engagement_score before allowing next-assessment navigation.

9. **New lesson type wiring** (2 days) — add `InteractivePDF` to [ILMSLesson.ts](src/models/ILMSLesson.ts) LessonType union, branch in [LmsCoursePlayer.tsx:2913](src/webparts/lmsCoursePlayer/components/LmsCoursePlayer.tsx#L2913) to render new viewer, add to Course Builder lesson-type picker.

10. **Author review gate for one artefact type** (3 days) — start with auto-summary only. New "AI Artefacts" tab in LmsCourseStudio PDF lesson editor. Accept/edit/reject with audit log.

11. **WCAG + keyboard + audit log hardening** (2 days)

12. **Playwright tests + load test** (2 days)

**Total MVP: ~30 dev-days** (was 22-25 in architecture research; I've added a day buffer for SPFx-specific integration quirks that the agent didn't account for).

**MVP does NOT include**:
- Proactive tutor (Phase 2)
- 4 of 5 artefact generators (Phase 2)
- Voice mode / audio overview (Phase 2)
- Regulator role-play (Phase 3)
- Workflow-embedded just-in-time (Phase 3+)
- Peer annotations (Phase 3+)

#### Phase 2 — **Agentic features** (3-4 weeks, ~15-18 dev-days)

- Proactive tutor rules engine + LLM escalation + interrupt UX (4 days)
- Remaining 4 artefact generators (glossary, flashcards, scenario quiz, regulator critique) + author review UI for each (5 days)
- Audio overview via Azure Speech (3 days)
- Agentic retrieval path for compound questions (2 days)
- "Explain like…" persona dropdown (1 day)
- Heatmap gutter (2 days)
- Hardening + extra WCAG (1 day)

#### Phase 3 — **Moat features** (longer-term, roadmap — not committed)

- Regulator role-play multi-turn dialogue mode
- Workflow-embedded just-in-time checklists (DMS integration — iManage/NetDocuments)
- Peer annotation + hot-spot detection + AI moderation
- Auto-course generation from policy diffs (when a new policy version is uploaded)
- Regulatory monitoring feed → policy-gap detector → auto-CPD micro-course

### 6.2 Files that will change / be added

**New files (MVP):**

```
src/components/shared/
  InteractivePdfViewer.tsx                      [PDF.js wrapper]
  DocCopilotPanel.tsx                           [A3 chat panel]
  CitationChip.tsx                              [clickable citation]
  ProactiveTutorModal.tsx                       [Phase 2]
  ReadingHeatmapGutter.tsx                      [Phase 2]

src/services/
  lmsDocCopilotService.ts                       [wraps /api/doc-* endpoints]

src/stores/
  lmsDocCopilotStore.ts                         [or extend lmsAIStore]

azure-functions/Functions/
  DocIngestFunction.cs
  DocCopilotFunction.cs
  DocGenerateArtefactFunction.cs
  DocTelemetryFunction.cs
  DocProactiveTutorFunction.cs                  [Phase 2]
  DocSidecarFunction.cs

azure-functions/Services/
  DocumentIntelligenceService.cs
  DocChunkingService.cs
  DocEngagementScorer.cs
  SemanticKernelAgentService.cs

azure-functions/Prompts/
  doc-tutor.md
  doc-summary.md
  doc-glossary.md                               [Phase 2]
  doc-flashcards.md                             [Phase 2]
  doc-scenario-quiz.md                          [Phase 2]
  doc-regulator-critique.md                     [Phase 2]
  doc-proactive-tutor.md                        [Phase 2]

infrastructure/
  doc-intelligence.bicep                        [or inline into main.bicep]
  ai-search-index-template.bicep

docs/
  interactive-pdf-research-and-plan.md          [THIS FILE]
  interactive-pdf-ux-mockup.html                [to be created before UI build]
```

**Modified files (MVP):**

```
src/models/ILMSLesson.ts                        [+ 'InteractivePDF' to LessonType union]
src/webparts/lmsCoursePlayer/components/
  LmsCoursePlayer.tsx                           [branch ~L2913 to render new viewer]
src/webparts/lmsCourseStudio/                   [lesson-type picker + AI Artefacts tab]
src/services/azureFunctionService.ts            [add doc-* wrappers]
src/stores/lmsAIStore.ts                        [extend OR replace with lmsDocCopilotStore]
CLAUDE.md                                       [add §48: Interactive PDF architecture]
deployment/learniq-manifest.json                [new Azure components + Functions]
config/config.json                              [no change — lesson type is data, not a new web part]
```

### 6.3 Dependencies to install

SPFx side (must use ES5 legacy build of PDF.js):
```
pdfjs-dist@^5.0.0      (use pdfjs-dist/legacy/build/pdf for ES5)
```

No other new front-end deps — existing React 17, Fluent UI v8, Zustand 3.7.2 stack suffices.

Azure Functions side (NuGet):
```
Azure.AI.DocumentIntelligence                   (replaces FormRecognizer)
Microsoft.SemanticKernel                        (core + AzureOpenAI + AzureAISearch)
Microsoft.SemanticKernel.Process                (Process Framework — Build 2025)
Azure.Search.Documents                          (already present)
```

### 6.4 Deployment Manager manifest updates

Per [deployment/learniq-manifest.json](deployment/learniq-manifest.json) and CLAUDE.md §47, add:

- 2 new Azure components: `DocumentIntelligence` (CognitiveServices subtype), `AISearchService` (already may be present — check)
- 6 new Azure Function endpoints to the manifest health-check list
- Config key `docCopilotEnabled` (feature flag, defaults true)
- Guardian tier gate: `docCopilot` requires `Professional` or `Enterprise` (fail-open per §46)

### 6.5 Cost modelling

**One-off per PDF (at ingestion):**
- Document Intelligence prebuilt-layout: ~$0.01 per page → 50-page PDF = $0.50
- Embedding (text-embedding-3-large, ~500 chunks × 512 tokens avg): ~$0.06
- Artefact generation (5 artefacts with gpt-4o/mini mix): ~$0.26
- **Total per 50-page PDF: ~$0.82 one-off**

**Per learner session (60 min average):**
- 5-10 Q&A interactions × gpt-4o ~2k tokens each: ~$0.10
- 2-3 proactive tutor interrupts × gpt-4o-mini: ~$0.005
- Telemetry + audit writes: free
- **Total per learner per 60-min session: ~$0.11**

**For Webber Wentzel (250 lawyers, 20 CPD hours/year ≈ 30 sessions/lawyer/year):**
- Pilot year: 250 × 30 × $0.11 = **~$825/year in OpenAI costs**
- Plus ~100 courses × $0.82 ingestion = **~$82 one-off**
- Total year 1 AI cost for Webber Wentzel: **~$900**

Very acceptable. Most of this would be absorbed into the platform licence price already.

---

## 7. Risks and open questions

### 7.1 Risks

1. **AI hallucination on legal questions** — mitigated by strict source-grounding, refusal when no chunks pass relevance threshold, prominent disclaimer, partner escalation path for high-stakes queries. (CLAUDE.md §7-ish; see SECTION 7 of legal research.)
2. **SPFx + PDF.js bundle size** — PDF.js legacy build is ~900 KB min / 300 KB gzip. Dynamic-import it. Would need to verify no crash on initial page load with other SPFx parts on the page.
3. **Iframe nesting** — PDF.js in SPFx in SharePoint iframe should work (well-trodden), but SharePoint's Content Security Policy has historically broken things (CLAUDE.md §38 on TinyMCE CDN). Must bundle everything; no CDN-loaded resources.
4. **Azure AI Search per-course index proliferation** — if we end up with thousands of courses, this is expensive. Mitigation: use a single index with `documentId` filter field as primary access pattern; keep per-course partitioning as a logical boundary only.
5. **Author review fatigue** — 5 artefacts × author review × every PDF might be too much. Mitigation: one-click "Accept all" option with a clear "these are marked AI-generated" badge on the content; authors can refine later.
6. **Regulated-industry privilege concerns** — if a lawyer pastes client context into the chat, we must detect and block. Handled by PII scrubbing + warning UI + clear policy education in-app.

### 7.2 Open questions for Gary

1. **Persona naming** — reuse DWx Copilot, or introduce LearnIQ-specific persona (e.g. "Iris")?
2. **Guardian tier gating** — which features are Free / Professional / Enterprise?
   - My suggestion: Free = PDF viewer + basic chat with daily limit; Professional = full agentic Q&A + auto-summary + engagement scoring; Enterprise = proactive tutor + scenario generator + regulator critique + workflow embed (Phase 3).
3. **Legacy PDF lesson type** — replace or coexist? My recommendation: coexist. Keep `PDF` for legacy; new lessons default to `InteractivePDF` when enabled.
4. **Voice / audio overview in Phase 2** — is a NotebookLM-style podcast overview worth building, or is it gimmicky for legal content? Worth user-testing with Webber Wentzel.
5. **Marketplace** — does the React 18 marketplace also get the interactive PDF experience, or is it SPFx-only for MVP? My recommendation: SPFx-only for MVP; marketplace in Phase 2 using the same `/api/doc-*` endpoints with a React 18 client.
6. **Demo-readiness** — do we want a slice of this ready to demo to Webber Wentzel at a follow-up before full MVP? If yes, we could prioritise the split-pane chat + citations and skip the proactive tutor for demo purposes.

### 7.3 What I'd do first (concrete next step)

**Before any code:** build an HTML mockup of the split-pane Interactive PDF experience in [docs/interactive-pdf-ux-mockup.html](docs/interactive-pdf-ux-mockup.html), per the CLAUDE.md §45 and `feedback_mockups_first.md` memory. Show you the UX patterns we're committing to, get sign-off, then start cutting code on the PDF.js wrapper.

---

## 8. Sources

Market research:
- [Kotobee Author](https://www.kotobee.com/en/products/author)
- [AISDI / ALMA](https://aisdi.ai/)
- [Docebo](https://www.docebo.com/pricing)
- [Sana Labs](https://www.sanalabs.com/platform)
- [Cornerstone Galaxy AI](https://www.cornerstoneondemand.com/)
- [Litmos](https://www.litmos.com/features)
- [Absorb LMS](https://www.absorblms.com/)
- [ChatPDF](https://www.chatpdf.com/)
- [Humata](https://humata.ai/)
- [NotebookLM](https://notebooklm.google.com/)
- [AskYourPDF](https://www.askyourpdf.com/)
- [Articulate Rise 360](https://articulate.com/360/rise)
- [Easygenerator](https://www.easygenerator.com/)
- [FlipHTML5](https://www.fliphtml5.com/)
- [Publuu](https://publuu.com/)
- [Hotshot Legal](https://www.hotshotlegal.com/)
- [LexisNexis CLE](https://www.lexisnexis.com/en-us/cle/default.page)
- [Thomson Reuters Practical Law Learn](https://legal.thomsonreuters.com/en/products/practical-law/learn)
- [LRN Catalyst / Interactive Services](https://lrn.com/products/ethics-compliance-training)

Technical reference:
- [PDF.js (Mozilla)](https://mozilla.github.io/pdf.js/)
- [Azure AI Search — hybrid search overview](https://learn.microsoft.com/en-us/azure/search/hybrid-search-overview)
- [Azure AI Search — chunking documents](https://learn.microsoft.com/en-us/azure/search/vector-search-how-to-chunk-documents)
- [Azure AI Search — agentic retrieval](https://learn.microsoft.com/en-us/azure/search/search-agentic-retrieval-concept)
- [Azure Document Intelligence — prebuilt-layout](https://learn.microsoft.com/en-us/azure/ai-services/document-intelligence/prebuilt/layout)
- [Semantic Kernel Process Framework](https://learn.microsoft.com/en-us/semantic-kernel/frameworks/process/)
- [Azure-Samples / azure-search-openai-demo](https://github.com/Azure-Samples/azure-search-openai-demo)
- [microsoft / kernel-memory](https://github.com/microsoft/kernel-memory)
- [xAPI cmi5 profile](https://xapi.com/statements-101/)

Legal / CPD:
- LPC CPD framework (South Africa)
- POPIA Act (South Africa)
- UK SRA CPD requirements
- FINRA training evidence-of-comprehension framework

---

*End of document.*
