# DWx Interactive Multimedia Ebook + Document Tutor — Concept & Plan

> **Working title:** LearnIQ Interactive Ebook (lesson type)
> **Companion specs:** [dwx-document-tutor-pattern.md](dwx-document-tutor-pattern.md) · [interactive-pdf-research-and-plan.md](interactive-pdf-research-and-plan.md)
> **Status:** Concept — for alignment before build
> **Date:** 2026-04-17

---

## 1. The expanded vision

The **PDF + Document Tutor** combination is powerful because it transforms inert, existing documents into agentic learning surfaces. But it's still bound by the PDF's limitations: the content is fixed, the interactivity is bolted on *around* the document, and the document itself remains a read-only artefact.

Kotobee, H5P, Articulate Storyline, Adobe Captivate — this whole category of tools proves that **learners engage much more deeply when the content itself is interactive**: embedded video, branching scenarios, drag-and-drop, hotspots, audio narration synced to text, quiz checkpoints, 3D images, timelines. These tools have dominated corporate e-learning for a decade.

**But they are pre-agentic.** None of them have an AI tutor embedded in the content. Their interactivity is *scripted*: the author decides, at authoring time, what happens when the learner clicks. There's no runtime intelligence reading the content, adapting to the learner, answering ad-hoc questions, or generating new interactive moments on the fly.

### The opportunity

Combine the **multimedia interactivity of a Kotobee-style ebook** with the **agentic runtime intelligence of the DWx Document Tutor**.

The result is a **lesson artefact that is simultaneously**:

1. **Authored content** — structured, multimedia, deliberately designed, with all the interactive widget types lawyers actually respond to: scenarios, drag-decision-trees, video with branch points, audio narration, hotspots on diagrams, embedded quizzes.
2. **Agentic content** — every widget, every passage, every embedded media is addressable by the Document Tutor: the learner can pause a video and ask "why did she do that?", highlight a scenario path and ask "what would happen if we picked B instead?", ask for a plain-English explanation of a clause, have the tutor critique their drag-drop answer in context.

We are not replacing the PDF tutor. We are giving it a **richer document to tutor over.**

### Positioning line

> **LearnIQ Interactive Ebooks** — Kotobee-grade multimedia interactivity, with DWx Copilot embedded in every page. An ebook your lawyers can read, scrub through, answer, drag, decide, ask, and be quizzed on — all without leaving the document. Every interaction contributes to an audit-grade comprehension record.

No competitor ships this combination today. Kotobee has no AI. Articulate/Captivate have no grounded tutor. NotebookLM has tutor but no multimedia interactive content. This is a genuine white space.

---

## 2. What the experience looks like

A learner opens lesson **"POPIA for Legal Practitioners — Module 2"**. Instead of a PDF, they see:

- **Page 1** — Hero image with an animated "chapter intro" video clip (30s). A voiceover reads the chapter overview while key terms highlight on screen. The audio is narrated text synced to paragraphs underneath (a Kotobee-style "read along"). The learner can tap any term → glossary pops up.
- **Page 2** — Policy text intercut with a **diagram of the six lawful grounds** (clickable hotspots: click each ground → popover with one-line explanation + jump to source policy).
- **Page 3** — An **embedded scenario widget**: *"A client walks in. You need to process their ID. Which lawful ground applies?"* Five options shown as cards. Learner drags the right card to the "use this ground" zone. Wrong answer → inline explanation; right answer → micro-celebration + unlock next page.
- **Page 4** — An **interactive decision tree** on legitimate interest: *"Does your processing involve special information?"* Click No → next question → Click Yes → branch to "you need a different ground". Every branch logged as an xAPI statement.
- **Page 5** — A **comparison video**: a 2-minute dramatisation of a firm partner explaining the 3-part LIA test, with pause points. At each pause the learner is asked a comprehension question before the video resumes.
- **Page 6** — A **mini case-study simulation**: *"You're acting for ABC Bank. They want to share client info with a vendor for fraud detection. Walk through the LIA test."* Three-step form with AI-graded free-text responses.
- **Page 7** — Chapter recap + audio summary ("Listen to the 90-second recap").

At **every single page** the DWx Copilot panel on the right is live. The learner can:

- Ask anything about the policy, the video, the diagram, the scenario. The tutor has been pre-indexed on every page's content, every video transcript, every widget's authoring data.
- Get a citation back to the exact page **and widget** (e.g. "See p3 · Scenario: Client walks in · option C").
- Be interrupted proactively: if they skip the video, if they answer the drag-drop randomly, if they dwell on a passage with a known misconception hotspot.
- Trigger any auto-generated artefact — summary, audio overview, flashcards, regulator critique — based on the *entire* ebook, not just the text.

At the end, the engagement score aggregates **all** interactions: video watched / paused / re-played, widgets completed, scenarios taken, Q&A asked, proactive checks passed. The comprehension attestation certificate cites every interaction, not just a final quiz.

---

## 3. Widget catalogue — what we build

We don't need to build the full H5P catalogue (50+ content types). We need the **lawyer-appropriate subset** — the widgets that drive real engagement for our audience per the domain research. Ordered by priority.

**Build/buy decision: Option A — build our own.** (Decided 2026-04-18.) First-class React 17 components in `@dwx/document-renderer-ebook`, fully integrated with the Document Tutor event contract. No H5P iframe embedding, no Kotobee licence. The reasoning: the tutor's grounding quality depends on reading widget state directly, and any iframe boundary (H5P, Kotobee) severs that. See the decision rationale in [dwx-document-tutor-pattern.md §3.1](dwx-document-tutor-pattern.md#31-component-1--document-renderer) and the H5P-mapping analysis in §3.4 below.

### 3.1 Core (MVP — Phase 1) · 10 widgets

| Widget | What it does | Lawyer use case |
|--------|-------------|-----------------|
| **Rich text block** | Styled paragraphs, headings, lists, citations | Policy text, explanation |
| **Callout / aside** | Highlighted tip, warning, key point | "Partners note", "Common mistake" |
| **Image with hotspots** | Clickable regions revealing explanation | Flowcharts, diagrams, org-charts |
| **Embedded video** | Video player with chapter markers and pause-for-question triggers | Partner explaining a concept, dramatisation |
| **Audio narration sync** | Kotobee-style text-highlight as audio plays | Read-along for commute, accessibility |
| **Multiple choice** | Standard MCQ with inline explanation (covers True/False as 2-option variant) | Knowledge check |
| **Fill in the blanks** | Text with missing words learner types in | "Complete the consent statement: _______ shall be specific, informed, voluntary and _______" |
| **Scenario card** | Situation + single-choice decision + consequence reveal (flat, non-branching) | "Client walks in" decision practice |
| **Accordion / disclosure** | Collapsible sections for progressive reveal | Detailed sub-clauses |
| **Glossary term chip** | Inline term → popover definition | Legal terminology |

### 3.2 Advanced (Phase 2) · 10 widgets

| Widget | What it does | Lawyer use case |
|--------|-------------|-----------------|
| **Branching scenario** | Multi-step decision tree with divergent paths and consequences | Full conflict-check or LIA walkthrough |
| **Drag-to-zone** | Drag items into correct categories/zones (image-based) | "Sort these clauses by risk" |
| **Mark the words** | Learner highlights target words in a passage | "Highlight every POPIA reference", "Mark every condition in this clause" |
| **Interactive video** | Video with in-stream checkpoints, pause questions, branches | Ethics scenario with decision points |
| **Timeline** | Horizontal interactive timeline with events | POPIA implementation history, regulatory evolution |
| **Comparison slider** | Before/after or two-version comparison | Old vs new policy, standard vs deviated clause |
| **Pop quiz (instant)** | Mid-content MCQ that stops the flow | Rapid check after a concept |
| **Free-text scenario** | Open response graded by tutor agent | "How would you advise?" |
| **Audio recorder** | Learner records their own explanation for AI feedback | Oral argument practice |
| **Drag-the-words** | Drop words into sentence slots (text-only drag) — *only if Fill-in-the-Blanks proves insufficient in practice* | "Complete the clause by dragging the right terms in" |

### 3.3 Frontier (Phase 3+)

| Widget | What it does | Lawyer use case |
|--------|-------------|-----------------|
| **AI-rendered scenario** | Tutor generates a fresh scenario at runtime based on learner's weak areas | Infinite, never-identical practice |
| **Role-play dialogue** | Multi-turn conversation with the tutor playing regulator/opposing counsel | Ethics interview rehearsal |
| **Live case-law link-out** | Inline citation that pulls current case text from LexisNexis/Westlaw | "See [Smith v Jones, 2024]" |
| **Collaborative annotation** | Shared notes visible to cohort, with AI moderation | Reading circle |
| **Live-document side-pane** | The tutor can draft a redline or memo as you read, in a second pane | ContractIQ crossover |

### 3.4 Coverage vs H5P Top 20

How our widget catalogue maps to H5P's most-used content types, validated 2026-04-18.

| H5P widget | Our coverage | Notes |
|------------|-------------|-------|
| Interactive Video | **Phase 2** (via Interactive video widget) | Phase 1 `Embedded video` covers pause-for-question; branching streams are Phase 2 |
| Course Presentation | **Native architecture** | Our `Ebook → Page → Widget` model is the equivalent — not a widget |
| Interactive Book | **Native architecture** | Same as above — the ebook itself is this |
| Multiple Choice | **Phase 1** (MCQ) | 1:1 |
| Quiz / Question Set | **Composition** | Built by sequencing multiple MCQs / scenarios — not a widget |
| Fill in the Blanks | **Phase 1** (added based on coverage analysis) | 1:1 |
| Drag and Drop | **Phase 2** (Drag-to-zone) | 1:1 |
| Branching Scenario | **Phase 2** | Phase 1 `Scenario card` is the flat variant |
| True/False | **Phase 1** (via MCQ with 2 options) | Covered |
| Drag the Words | **Phase 2 (conditional)** | Included only if Fill-in-the-Blanks proves insufficient |
| Mark the Words | **Phase 2** (added based on coverage analysis) | Close-reading tool |
| Dialog Cards / Flashcards | **Artefact, not widget** | Generated as study-artefact deck per [Document Tutor pattern §3.4](dwx-document-tutor-pattern.md#34-component-4--artefact-pipeline) |
| Image Hotspots | **Phase 1** (Image with hotspots) | 1:1 |
| Accordion | **Phase 1** | 1:1 |
| Summary | **Composition** | Chapter-end quiz built from MCQs/scenarios |
| Single Choice Set | **Composition** | Sequence of MCQs |
| Timeline | **Phase 2** | Specialised; on-demand for regulatory-evolution content |
| Image Pairing | **Covered via Drag-to-zone** | Same mechanic, different rendering |
| Essay | **Phase 2** (Free-text scenario — AI-graded) | Superior to H5P's keyword-match approach |

**Widgets we have that H5P under-serves:** Callout (H5P has none as a dedicated widget), Audio narration with text-sync (H5P's Text Track is rarely used), Glossary term chip (H5P has none). These all fit the legal-content pattern particularly well.

**Architectural note**: H5P bundles composition primitives (Course Presentation, Interactive Book, Quiz Set, Summary, Single Choice Set) into its widget catalogue because every H5P widget is a standalone unit. Our cleaner model — `Ebook → Page → Widget` as explicit hierarchy — means each of those five H5P widgets is covered natively without needing a widget counterpart. That's one reason our Phase 1 is 10 widgets and our Phase 2 is another 10 rather than chasing H5P's 50+.

---

## 4. Integrating the tutor — where the magic happens

The DWx Document Tutor pattern already specifies a renderer event contract ([dwx-document-tutor-pattern.md §3.1](dwx-document-tutor-pattern.md#31-component-1--document-renderer)). The interactive ebook is a **new renderer** that emits those same events — plus a richer set specific to multimedia widgets.

### 4.1 Ebook-specific events (extending the renderer contract)

In addition to the base renderer events (docOpened, pageChanged, dwellTick, textSelected, annotationCreated, citationJumped, sessionEnded), the ebook renderer emits:

| Event | Payload |
|-------|---------|
| `videoPlayed` | `{widgetId, page, startMs, endMs, paused, seeked}` |
| `audioPlayed` | `{widgetId, page, startMs, endMs, textSync: boolean}` |
| `hotspotClicked` | `{widgetId, page, hotspotId, label}` |
| `quizAnswered` | `{widgetId, page, questionId, selected, correct, attempts, timeToAnswerMs}` |
| `dragDropCompleted` | `{widgetId, page, correct, attempts, arrangement}` |
| `scenarioCompleted` | `{widgetId, page, pathTaken, terminalNodeId, outcome}` |
| `branchChosen` | `{widgetId, page, nodeId, choice, consequences}` |
| `accordionOpened` | `{widgetId, page, sectionId}` |
| `glossaryOpened` | `{widgetId, page, term}` |
| `freeTextSubmitted` | `{widgetId, page, text, aiGradedScore, rubric}` |

Every one of these is indexed by `widgetId` so the tutor can address any widget in its retrieval.

### 4.2 Tutor addresses widgets, not just pages

When the learner asks the tutor about something they're looking at, the tutor's context is enriched:

```json
{
  "currentPage": 3,
  "currentWidget": {
    "id": "w-scenario-client-walks-in",
    "type": "scenario",
    "state": {
      "questionShown": "Which lawful ground applies?",
      "optionsShown": ["Consent", "Contract", "Legal obligation", "Legitimate interest", "Public law duty"],
      "learnerSelected": "Consent",
      "correct": "Legitimate interest",
      "attempts": 1
    }
  }
}
```

So the user can just say **"Why is my answer wrong?"** and the tutor responds with grounded context:

> "You picked **Consent**. The client is providing their ID to receive legal services they've already contracted for — the processing is necessary to perform that contract. Consent would be redundant (and would imply the client could withdraw and force you to stop the matter). The correct answer is **Legitimate interest** — see Policy 4.2 § 2 {CITE:1}. The three-part LIA test is still required, but consent isn't the right ground here."

This is something no static ebook tool can do. The tutor has the widget state in context. No separate chatbot, no context switch.

### 4.3 The widget authoring side is also agentic

Authors don't build every widget by hand. The **Artefact Pipeline** (Document Tutor pattern §3.4) expands to include **widget auto-generation** at ingestion:

| Input | Auto-generated output |
|-------|----------------------|
| Policy PDF | Summary ebook draft: chapter structure, hotspot diagrams, 5-10 scenario widgets, 1-2 video prompts, glossary |
| DOCX/MD from author | Widget suggestions per section ("this would benefit from a decision-tree widget here") |
| Existing SCORM package | Migration → ebook with widgets preserved, tutor added on top |

**Author review gate is still mandatory.** Every auto-generated widget is draft, author accepts/edits/regenerates/rejects, every decision logged. The pipeline proposes; the author disposes. (CLAUDE.md §18.)

### 4.4 Proactive tutor gains widget-aware triggers

The proactive tutor (Document Tutor pattern §3.6) was already watching for skim signals on text. For ebooks, it adds:

- **Video skipped / fast-forwarded past a checkpoint** → "Want me to summarise what you skipped?"
- **Scenario answered randomly** (< 1s to answer) → "Let's slow down — here's a cleaner way to think about this."
- **Hotspot ignored** on a diagram where the hotspot reveals critical information → nudge
- **Audio narration disabled** on a page known to be dense text → "This page is quite dense — would you like me to read it aloud?"
- **Drag-drop repeatedly wrong** → offer scaffolding ("Let me give you a hint about where these belong")

All still rules-first, LLM-escalated. Still cheap. Still dismissible.

### 4.5 Engagement score gains widget weightings

The canonical formula (pattern §3.7) weights `dwell / scroll / interactions / comprehension` at `0.4 / 0.2 / 0.2 / 0.2`. Ebook-specific refinement:

- **`interactionRate`** now counts widget completions (videos watched to checkpoint, scenarios completed, drag-drops solved, quizzes taken) — dramatically higher signal than "selections per 10min" on a PDF.
- **`comprehensionScore`** now blends proactive checks *and* embedded quiz results. Recommended split: 60% proactive / 40% embedded.
- **`scrollCoverage`** adapts: for widget-heavy pages where there's little text to scroll, weight is automatically reallocated to `interactionRate`.

Result: the engagement score is **more accurate** on interactive ebooks than on PDFs, because there's more signal per page. This makes the comprehension attestation *stronger evidence*, not weaker.

---

## 5. Technical architecture — what changes vs the PDF plan

The good news: **80% of the architecture from the PDF plan is unchanged**. We're adding a new renderer and extending contracts, not re-architecting.

### 5.1 What stays the same

- Document Intelligence ingest pipeline (used for the underlying source policy/content)
- Azure AI Search index with hybrid + reranker
- Azure OpenAI proxy through Azure Functions
- Semantic Kernel agent graph with tool use
- Proactive tutor rules engine
- Engagement telemetry + xAPI + LMS_AuditLog
- Sidecar JSON for coordinate resolution
- PII scrubbing on Function boundary
- Guardian tier gating
- Deployment Manager manifest pattern

### 5.2 What's new

| Layer | New addition |
|-------|-------------|
| **Content model** | New SharePoint list `LMS_Ebooks` OR extend `LMS_Lessons` with a structured `EbookContent` JSON field. See §5.3 below. |
| **Renderer** | New React component tree `@dwx/document-renderer-ebook` — implements the renderer event contract but for HTML+widgets instead of PDF canvas. |
| **Widget library** | React-17-compatible widget components: `<ScenarioWidget>`, `<VideoWidget>`, `<HotspotImage>`, etc. Each widget self-reports its state via the event contract. |
| **Authoring UI** | New web part `lmsEbookBuilder` (or new section in `lmsCourseStudio`) with drag-drop widget authoring. Could reuse TinyMCE for rich-text blocks and add widget-insertion UI around it. |
| **Agent tool extension** | New tools for the tutor: `describe_widget`, `evaluate_widget_answer`, `generate_widget_variant`. |
| **Ingestion** | Extend ingest to index widget text (scenario prompts, video transcripts, hotspot labels) as first-class chunks with `widgetId` metadata. |
| **Sidecar JSON extension** | Includes widget manifests so the renderer can address and highlight specific widgets, not just page coordinates. |
| **Storage** | Video and audio assets go to Azure Blob (existing `LMS_MediaLibrary` or similar). Transcripts are auto-generated on upload via Azure Speech-to-Text (new dependency, cheap). |
| **Accessibility layer** | WCAG 2.2 AA for interactive widgets — keyboard access, aria-labels, screen-reader announcements for widget state changes. Separate hardening pass per widget. |

### 5.3 Content model

Recommended: **extend `LMS_Lessons`**, don't create a new list. The lesson remains the unit of authoring and enrolment; its `LessonType` becomes `'InteractiveEbook'` and its content lives in a new `EbookContent` JSON field.

**Content hierarchy** (decided 2026-04-18): `Ebook → Page → Section → Widget`. A **Section** is an intermediate container with a layout preset (full / 2-col / 2x2 etc.) and a set of widget slots. This gives authors grid layouts without the rope-to-hang-yourself-with of free-form positioning.

```json
{
  "lessonId": "L-POPIA-2026-001-M2-L3",
  "lessonType": "InteractiveEbook",
  "title": "POPIA Consent and Legitimate Interest",
  "ebookContent": {
    "version": "1.1",
    "source": {
      "documentId": "FP-4.2-v3.1",
      "corpusKey": "course:POPIA-2026-001"
    },
    "pages": [
      {
        "id": "p-1",
        "title": "Introduction",
        "sections": [
          {
            "id": "s-1a",
            "layout": "full",
            "widgets": [
              { "id": "w-1", "slot": 0, "type": "video", "config": {...}, "authorReviewed": true }
            ]
          },
          {
            "id": "s-1b",
            "layout": "2-col",
            "widgets": [
              { "id": "w-2", "slot": 0, "type": "richtext", "content": "...", "authorReviewed": true },
              { "id": "w-3", "slot": 1, "type": "audio-narration", "config": {...} }
            ]
          }
        ]
      },
      {
        "id": "p-2",
        "title": "The six lawful grounds",
        "sections": [
          {
            "id": "s-2a",
            "layout": "2x2",
            "widgets": [
              { "id": "w-4", "slot": 0, "type": "mcq", "config": {...} },
              { "id": "w-5", "slot": 1, "type": "mcq", "config": {...} },
              { "id": "w-6", "slot": 2, "type": "mcq", "config": {...} },
              { "id": "w-7", "slot": 3, "type": "mcq", "config": {...} }
            ]
          }
        ]
      }
    ],
    "branches": [...],                 // optional — for branching scenarios (Phase 2)
    "attestationPolicy": {...},        // engagement gate config
    "artefacts": {
      "summary": {...},
      "audioOverview": {"blobUrl": "...", "durationMs": 363000},
      "flashcards": [...],
      "scenarioQuiz": [...]
    }
  }
}
```

The JSON is serialised into a single SharePoint multi-line text field (or split across a few fields if size becomes a concern — SP has a practical ~1 MB limit per item field). Large blobs (video, audio) stay in Blob storage, referenced by URL with SAS.

### 5.3.1 Page-layout presets

Authors choose from a **fixed set of layout presets** per section. No free-form grid editing. Rationale: Confluence, Notion and Kotobee all prove that free-form layouts produce gorgeous work in the hands of experts and disasters in the hands of everyone else — a small, named set gives enough variety without enough rope.

| Preset | Slots | CSS grid | Typical use |
|--------|-------|----------|-------------|
| `full` | 1 | `1fr` | Default. Hero video, long text block, wide diagram, drag-drop widget |
| `2-col` | 2 | `1fr 1fr` | Side-by-side comparison, text + media pairing, policy text + hotspot diagram |
| `2-col-1-2` | 2 | `1fr 2fr` | Narrow label / legend / callout + wide content |
| `2-col-2-1` | 2 | `2fr 1fr` | Wide content + narrow sidebar (tips, glossary, regulator-view micro-callout) |
| `3-col` | 3 | `1fr 1fr 1fr` | Three parallel concepts — Purpose / Necessity / Balancing, past / present / future |
| `2x2` | 4 | `1fr 1fr / 1fr 1fr` | Four-quadrant concepts — the six lawful grounds grouped in pairs, SWOT-style comparisons |

**Responsive rule (enforced by renderer, not authorable):** on viewports < 760px, every preset collapses to a single stacked column in slot order. Authors never choose per-breakpoint layouts — we take that decision away on purpose.

### 5.3.2 Widget-to-layout validity matrix

Some widgets only make sense at full width; forcing them into a narrow slot produces a cramped, broken experience. The renderer validates this at author time and prevents placement in disallowed slots.

| Widget | `full` | `2-col` | `2-col-1-2` wide | `2-col-1-2` narrow | `3-col` | `2x2` |
|--------|:---:|:---:|:---:|:---:|:---:|:---:|
| Rich text | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| Callout | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| Glossary chip | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| Accordion | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| MCQ | ✓ | ✓ | ✓ | · | ✓ | ✓ |
| Fill-in-the-blanks | ✓ | ✓ | ✓ | · | ✓ | · |
| Scenario card | ✓ | ✓ | ✓ | · | · | · |
| Image with hotspots | ✓ | ✓ | ✓ | · | · | · |
| Audio narration (w/ text-sync) | ✓ | ✓ | ✓ | · | · | · |
| Embedded video | ✓ | ✓ | · | · | · | · |
| Drag-to-zone (Phase 2) | ✓ | · | · | · | · | · |
| Branching scenario (Phase 2) | ✓ | · | · | · | · | · |
| Interactive video (Phase 2) | ✓ | · | · | · | · | · |
| Timeline (Phase 2) | ✓ | · | · | · | · | · |
| Comparison slider (Phase 2) | ✓ | · | · | · | · | · |
| Free-text scenario (Phase 2) | ✓ | ✓ | ✓ | · | · | · |
| Audio recorder (Phase 2) | ✓ | ✓ | ✓ | · | · | ✓ |
| Mark the words (Phase 2) | ✓ | ✓ | ✓ | · | · | · |
| Pop quiz (Phase 2) | ✓ | ✓ | ✓ | · | ✓ | ✓ |

Legend: `✓` valid placement · `·` not permitted by renderer.

### 5.3.3 Tutor awareness of sections

Per the [Document Tutor pattern renderer contract](dwx-document-tutor-pattern.md#31-component-1--document-renderer), the renderer always reports `{page, widgetId}`. This gets extended to `{page, sectionId, layout, slot, widgetId}` so the tutor can narrate layouts: *"You're looking at the 2×2 grid on page 2 — slot 3 is the one on the bottom right."* One-line extension to the context payload, no architectural change.

### 5.3.4 Build impact

Adding Section-level layout to Phase 1 costs an additional **~5 dev-days** on top of the prior estimate:

| Workstream | Days |
|------------|-----:|
| Content-model change (JSON schema + SP field + migration path) | 1 |
| Renderer layout logic (CSS grid, responsive collapse, slot placement) | 1 |
| Builder UI — Section toolbar, layout chooser, slot targets | 2 |
| Inspector controls for changing layout (with widget re-slot logic when slot count differs) | 1 |

Phase 1 MVP total revises from ~33-35 dev-days to **~38-40 dev-days** on top of the PDF plan.

### 5.4 Build effort (additional over the PDF plan)

| Workstream | Effort | Notes |
|------------|--------|-------|
| Ebook content model + SP field changes | 1 day | |
| Ebook renderer core (page engine, nav, event contract) | 4 days | |
| Widget library — 9 MVP widgets | 8-10 days | ~1 day each, plus shared primitives |
| Widget authoring UI in Course Studio | 5 days | Drag-drop widget palette, inline edit |
| Ingest extension for widget indexing | 2 days | |
| Sidecar JSON extension for widget manifests | 1 day | |
| Tutor tool extensions (describe_widget, evaluate_widget) | 2 days | |
| Auto-widget suggestion at ingest | 3 days | Optional; Phase 2 candidate |
| Accessibility hardening per widget | 3 days | |
| Author review gate for widgets | 2 days | |
| Playwright + load tests | 2 days | |

**Total ebook extension: ~33-35 dev-days** on top of the ~30 dev-days PDF MVP. Bringing the **combined Interactive PDF + Interactive Ebook MVP to ~60-65 dev-days (~3 months single dev, 6-7 weeks with two in parallel).**

### 5.5 Relationship with Kotobee / H5P — build, buy, or embed?

Three options:

**Option A — Build our own widgets (recommended)**

We already have React 17 + Fluent UI v8 + Zustand and the DWx design language. Building 9 MVP widgets as first-class React components gives us:
- Perfect design-system alignment
- Direct integration with the Document Tutor event contract (no postMessage bridge)
- Every widget addressable by the agent without wrappers
- WCAG hardening owned end-to-end
- No licensing cost
- No CDN dependency (CLAUDE.md §38 — SharePoint CSP blocks CDNs)

**Option B — Embed H5P widgets via iframe**

H5P is open source (MIT). It has 50+ widget types we'd never build ourselves. Embed them as iframes with a postMessage bridge that translates H5P's xAPI statements into our event contract.
- Pros: massive widget library; battle-tested; free
- Cons: iframe-in-iframe-in-SharePoint is CSP hostile; H5P's design language doesn't match ours; agent can only see widget state through the postMessage bridge, which limits grounding quality; maintenance burden of the bridge

**Option C — License Kotobee or similar**

Kotobee ebooks export to HTML5. We could author in Kotobee, publish the ebook, then embed it. Tutor stays in our copilot panel around the iframe.
- Pros: mature authoring UX, extensive widget library, fast to author
- Cons: tutor can't see *inside* the Kotobee iframe (same postMessage problem, amplified); we don't own the authoring story; design-system drift; per-author licence cost; less "DWx product" feel

**Recommendation: Option A for MVP** — build our own 9 MVP widgets as first-class React components. This gives the tutor direct, deep access to every widget's state, which is the whole value proposition. Revisit Option B in Phase 3 *only if* we need a long-tail widget type that's not worth building ourselves (e.g. crossword, AR). Option C we should not pursue — it's the anti-pattern of "AI retrofitted onto a platform that wasn't built for it".

---

## 6. What this unlocks for each app in the suite

### 6.1 LearnIQ
- **Interactive Ebook** becomes the premium lesson type. Positioned as the default format for new CPD courses. Legacy PDF and Video lesson types remain.
- Opens a content services revenue line: First Digital authors ebooks *for* firms (done-for-you CPD content).
- Comprehension attestation gets materially stronger because there's more interaction signal per page.

### 6.2 PolicyIQ (major win)
- Firm policies published as interactive ebooks → mandatory attestation flows with real engagement evidence.
- The "applicability checklist" becomes a widget, not a separate page.
- **Policy-change notifications** can be delivered as a 3-minute interactive mini-ebook with just the changed sections plus a proactive tutor walking through impact on the reader's role.

### 6.3 ContractIQ (distinct but convergent)
- Contracts themselves aren't authored as ebooks (they're legal documents, not training material).
- **But**: contract *training* and *playbook* content is ideal for the ebook format. "How to negotiate an indemnity clause" as an interactive ebook with embedded video of a senior partner, drag-to-zone exercises on risk classification, scenario role-play with the tutor-as-opposing-counsel — this is exactly the content ContractIQ needs for junior-lawyer onboarding.
- The ebook renderer + Document Tutor becomes ContractIQ's **learning layer** while the contract editor becomes its work layer.

### 6.4 Cross-app flywheel
- The 9 MVP widgets live in `@dwx/document-renderer-ebook` — a shared npm package.
- One build across the suite.
- One engagement score formula. One attestation format. One audit export.
- Sales narrative: "Our suite shares a single AI learning substrate — authoring done once benefits every app."

---

## 7. Authoring — "New Interactive Ebook"

The learner experience is half of the product. The **authoring experience** is the other half — and it's where we differentiate from Kotobee most sharply, because our authoring is also agentic: the same Copilot that tutors learners helps authors build.

### 7.1 Where authoring lives

**Recommendation: extend `lmsCourseStudio`** with a new "Ebook Builder" mode. Reuse the left-rail / canvas / right-inspector pattern already in Course Studio. Don't stand up a new web part — fewer moving parts, shared design system, shared permissions model.

Entry points:
- **From Course Builder** → "Add Lesson" → "Interactive Ebook"
- **From Content Library** → "New" → "Interactive Ebook"
- **From a policy PDF** → "Convert to Interactive Ebook (AI-assisted)" — kicks off auto-draft pipeline

### 7.2 The three-pane authoring surface

```
┌──────────────────────────────────────────────────────────────────────┐
│ Header: "Ebook Builder · POPIA for Legal Practitioners · DRAFT v3"  │
│ Save · Preview · Submit for Review · Publish (gated)                 │
├────────┬───────────────────────────────────────┬─────────────────────┤
│        │                                       │                     │
│ PAGES  │         CANVAS (WYSIWYG)              │ INSPECTOR           │
│ list   │                                       │ + AI COPILOT        │
│        │  Page 3: "The six lawful grounds"     │                     │
│ p1 ✓   │  ┌─────────────────────────────────┐  │ Selected: Scenario  │
│ p2 ✓   │  │ [Rich text block]               │  │ Widget #w5          │
│ p3 ●   │  │ Question: Which lawful          │  │                     │
│ p4     │  │ ground applies when...          │  │ • Question          │
│ p5     │  │                                 │  │ • Options           │
│ + Page │  │ [Scenario Widget]               │  │ • Correct answer    │
│        │  │ ┌─────┐ ┌─────┐ ┌─────┐         │  │ • Explanation       │
│ WIDGETS│  │ │ A   │ │ B   │ │ C ✓ │         │  │                     │
│ palette│  │ └─────┘ └─────┘ └─────┘         │  │ ─── AI Copilot ───  │
│        │  │                                 │  │                     │
│ ◉ Text │  │ [Audio narration]               │  │ "Based on Policy    │
│ ◉ Video│  │ 0:34 ━━━━━━○────── 1:12         │  │ 4.2 § 2, I suggest  │
│ ◉ Image│  │                                 │  │ adding these two    │
│ ◉ MCQ  │  └─────────────────────────────────┘  │ scenarios for       │
│ ◉ Scen │                                       │ common mistakes..." │
│ ◉ Drag │                                       │ [+ Generate]        │
│ ◉ Audio│                                       │                     │
│ ◉ Hot  │                                       │                     │
│ ◉ Acc. │                                       │                     │
│        │                                       │                     │
└────────┴───────────────────────────────────────┴─────────────────────┘
```

**Left rail** — pages list + widget palette. Drag widget → canvas to insert.
**Centre** — WYSIWYG canvas showing the ebook page exactly as the learner will see it. Click a widget to select. Double-click to edit.
**Right inspector** — properties of the selected widget (what's editable depends on widget type) plus the authoring Copilot.

### 7.3 The authoring Copilot — "Build with me"

Same DWx Copilot brand, different mode. In Ebook Builder the Copilot is primed with:
- The source document(s) the ebook is based on (policy PDF, prior course content)
- The firm's house style
- The learning objectives the author set at the start
- The ebook's current draft state

It can do four things authors really need:

1. **Suggest widgets** — "This page has a lot of dense text about consent. Consider a scenario widget after paragraph 2 to check understanding, and a hotspot image on the consent-flow diagram." One click inserts the suggestion as a draft widget.
2. **Fill in a widget** — author inserts an empty MCQ widget, clicks "Fill with AI". Copilot reads the surrounding content + source doc, proposes a question + 4 plausible options + correct answer + explanation. Author reviews.
3. **Rewrite for audience** — "Rewrite this page for junior associates who don't have POPIA background yet." Copilot rewrites; author accepts or edits.
4. **Critique the draft** — "Review this ebook from a learner's perspective." Copilot returns a punch list: "Page 4 has no interactivity — consider a quiz checkpoint. Page 6 scenario has two equally-correct answers. Glossary missing 'Information Officer'." Author addresses.

All agentic actions are **proposals**. Nothing is auto-applied. Every accept/reject is logged per the pattern's audit contract.

### 7.4 The "New Interactive Ebook" flow (from blank)

**Step 1 — Start** (one screen, fast):
- Title
- Description
- Learning objectives (3-5 bullet points)
- Estimated duration
- Source material (optional):
  - Upload PDF / DOCX
  - Pick from Content Library
  - Start from blank

**Step 2 — Choose authoring mode**:
- **Author from scratch** → goes straight to the canvas with a blank page 1
- **AI-assisted first draft** (if source material was provided) → shows a progress screen ("Copilot is reading your policy and drafting a 6-page ebook with 8 widgets... ~90 seconds"). Then drops the author into the canvas with the first draft open.
- **Template** → choose from house templates ("POPIA Training Template", "Conflict Check Walkthrough", etc.) — pre-structured pages with widget slots.

**Step 3 — Edit in the three-pane canvas**. Save as Draft automatically. Preview mode shows exactly what a learner sees including tutor panel.

**Step 4 — Submit for Review** (required before publish per CLAUDE.md §18). Enters the Reviewer queue. Reviewer sees the ebook in a special review mode with:
- AI-flagged concerns ("Page 3 scenario: only 62% of plausible-answers were tested")
- Widget-by-widget accept/edit/regenerate/reject controls
- Audit log of every AI-assisted action
- One-click "Approve & Publish"

**Step 5 — Publish** → writes to Course Content via the existing state machine (`/api/update-course-status`), available to enrolled learners.

### 7.5 Author-review gate rules (reiterated)

Per [Document Tutor pattern §3.4](dwx-document-tutor-pattern.md#34-component-4--artefact-pipeline) and CLAUDE.md §18:

- AI-assisted first draft saves as `Status: Draft, AuthorReviewed: false`
- Every individual widget carries its own `authorReviewed` flag
- Learners only see `authorReviewed: true` content
- Publishing is gated on every widget being reviewed (accept/edit, not reject)
- Every review action logged to `LMS_AuditLog` with reviewer ref

### 7.6 Authoring engagement telemetry (yes, really)

Authors also generate engagement signals we should capture — it's useful for:
- Understanding which widgets are hardest to author (helps us improve the builder)
- Measuring author productivity ("time to first publish")
- Identifying authors who over-rely on AI drafts (and conversely, the authors whose drafts need the fewest AI edits — they may be the internal experts to surface for peer-learning)

xAPI verbs for authoring: `authored`, `reviewed`, `regenerated`, `rejected`, `published` — all under a sibling namespace `http://dwx.firstdigital.com/xapi/author-ext/`.

### 7.7 Authoring build effort (additive)

| Workstream | Effort |
|------------|--------|
| Ebook Builder web part shell (3-pane, state management) | 4 days |
| WYSIWYG canvas with widget drag-drop | 5 days |
| Widget inspector panels (per widget type) | 4 days |
| Authoring Copilot (suggest/fill/rewrite/critique tools) | 4 days |
| AI-assisted first draft pipeline (PDF → ebook scaffold) | 3 days |
| Templates (5 house templates) | 2 days |
| Reviewer mode + approve/publish flow | 3 days |
| Preview-as-learner | 1 day |
| Author telemetry | 1 day |
| Total | **~27 dev-days** |

This sits on top of the ~33-35 dev-days for the renderer/widget library. Grand total for full MVP (learner side + authoring side): **~60-70 dev-days — around 3 months of focused single-dev work, or 6-8 weeks with two devs in parallel.**

### 7.8 Open authoring questions

- **Collaborative editing** — two authors on the same ebook simultaneously? Not in MVP; add via SharePoint check-out/check-in initially, live collab in Phase 3.
- **Version history** — SharePoint versioning is adequate for MVP; surface the last 5 versions in the Builder header. Full branch/merge is over-engineering at this stage.
- **Import from SCORM / H5P** — occasionally firms will hand us an existing SCORM package and want it imported. Phase 2 at the earliest.
- **Rich media upload** — where do videos go (Blob) and how large can they be? Suggest 500 MB per video cap for MVP, with Azure Blob SAS URLs for streaming.
- **Transcription** — Azure Speech-to-Text on every uploaded video/audio, stored in the widget config. Essential for accessibility, for tutor grounding, and for search. Should this be automatic-on-upload or opt-in?

---

## 8. Open questions

These need your call before build.

1. **Widget library — own build (A), H5P embed (B), or Kotobee licence (C)?** My recommendation is **A** for all the reasons in §5.5. Want to lock that in?
2. **First widget set** — does my 9-widget MVP list match what you think lawyers will actually engage with, or would you re-order / add / drop?
3. **Authoring UX home** — extend `lmsCourseStudio` (simpler) or stand up a new `lmsEbookBuilder` web part (cleaner separation)?
4. **Content model** — extend `LMS_Lessons.EbookContent` JSON field (my recommendation) or new `LMS_Ebooks` list (more normalised, more plumbing)?
5. **Media library** — do we need a dedicated `LMS_Media` library for video/audio assets with entitlements/licensing tracking, or just dump into `LMS_Documents/Media/`?
6. **Auto-widget generation at ingest** — Phase 1 or Phase 2? Doing it in Phase 1 means authors can paste a policy PDF and get a first-draft ebook they edit down; doing it later means authors build ebooks from scratch which is slower but more controlled.
7. **ContractIQ as sibling product vs feature-of-LearnIQ** — this spec treats them as sibling products. Is that still the plan?
8. **SCORM / xAPI export** — for market credibility with firms that already have an LMS, should we be able to publish our ebooks *out* as SCORM packages that run in competitor LMSs? (Kotobee's entire business model is export-everywhere.) This would be a Phase 3 decision but affects renderer architecture if we want to preserve the option.
9. **Accessibility tier** — WCAG 2.2 AA is baseline. Do any Webber Wentzel contracts require WCAG 2.2 AAA? That affects widget design significantly.
10. **Pricing / Guardian gating** — Interactive Ebook lesson type is an obvious Professional-tier unlock. Pro-tier fair?

---

*End of concept document.*
