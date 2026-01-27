# Policy Manager - Claude Code Context

## Instructions for Claude

1. **Always read CLAUDE.md before you do anything**
2. **Always ask questions if you are unsure of the task or requirement**
3. **Be systematic in your planning, and execution**
4. **After you complete a task, always validate the result**
5. **We are working in https://mf7m.sharepoint.com/sites/PolicyManager**

---

## Project Overview

**Policy Manager** is an enterprise Policy Management Solution being spun off from the integrated JML (Joiner, Mover, Leaver) solution to become a completely standalone application within the **DWx (Digital Workplace Excellence)** suite by First Digital.

### Application Identity
- **App Name**: Policy Manager
- **Suite**: DWx (Digital Workplace Excellence)
- **Company**: First Digital
- **Tagline**: Policy Governance & Compliance
- **Current Version**: 1.2.0

---

## Technology Stack

| Category | Technology | Version |
|----------|------------|---------|
| Framework | SharePoint Framework (SPFx) | 1.20.0 |
| UI Library | React | 17.0.1 |
| Language | TypeScript | 4.7.4 |
| UI Components | Fluent UI v8 + v9 | 8.106.4 / 9.54.0 |
| Data Access | PnP/SP, PnP/Graph | 3.25.0 |
| Build System | Gulp | 4.0.2 |
| Node | Node.js | 18.17.1+, 20.x, or 22.x |

---

## Architecture Overview

### WebParts (6 total)
1. **jmlMyPolicies** - Personal policy dashboard for employees
2. **jmlPolicyHub** - Central policy discovery and search
3. **jmlPolicyAdmin** - Administrative interface
4. **jmlPolicyAuthor** - Policy creation and editing
5. **jmlPolicyDetails** - Detailed policy view
6. **jmlPolicyPackManager** - Policy package bundling

### Directory Structure
```
policy-manager/
├── src/
│   ├── webparts/          # 6 SPFx webparts
│   ├── components/        # Shared components (JmlAppLayout, etc.)
│   ├── services/          # 141+ business logic services
│   ├── models/            # 56+ TypeScript interfaces
│   ├── hooks/             # Custom React hooks
│   ├── styles/            # Centralized styling system
│   └── utils/             # Configuration utilities
├── docs/                  # Brand guide, style guide
├── config/                # SPFx build configurations
└── CLAUDE.md              # This file
```

---

## DWx Brand Guidelines

### Color Palette
| Name | Hex | CSS Variable |
|------|-----|--------------|
| Primary Blue | #1a5a8a | `--dwx-primary` |
| Primary Dark | #0d3a5c | `--dwx-primary-dark` |
| Primary Light | #2d7ab8 | `--dwx-primary-light` |
| Text Primary | #333333 | `--dwx-text-primary` |
| Text Secondary | #6c757d | `--dwx-text-secondary` |
| Text Muted | #a0a0a0 | `--dwx-text-muted` |
| Border | #c8c6c4 | `--dwx-border` |
| Border Light | #e1dfdd | `--dwx-border-light` |
| Background Subtle | #f3f2f1 | `--dwx-bg-subtle` |
| Success | #107c10 | `--dwx-success` |
| Warning | #986f0b | `--dwx-warning` |
| Error | #d13438 | `--dwx-error` |

### Typography
- **Font Family**: Inter (primary), system fallbacks
- **Font Weights**:
  - 400 (Normal) - Default for body text, labels
  - 500 (Medium) - Titles, buttons, badges
  - 700 (Bold) - KPI numbers ONLY
  - **NEVER use 600 (Semibold)**

### Gradient (Headers)
```css
background: linear-gradient(135deg, #1a5a8a 0%, #2d7ab8 100%);
```

### Spacing Scale
- XS: 4px
- SM: 8px
- MD: 12px
- LG: 16px
- XL: 20px
- XXL: 24px

### Reference Documents
- Brand Guide: `docs/DWx-Brand-Guide.pdf`
- Style Guide: `docs/DWx-Style-Guide-Preview.html`

---

## Critical Development Notes

### JML Coupling (HIGH PRIORITY)
The codebase has **3,200+ JML references across 270 files**. This is the primary technical debt to address for the spinoff.

#### JML Integration Services to Refactor:
- `PolicyJMLIntegrationService.ts` - Creates policy tasks in JML processes
- `JMLAssetIntegrationService.ts` - Asset coordination with JML
- `TalentJMLIntegrationService.ts` - Talent management integration

#### Naming Convention Changes Required:
- All "Jml" prefixes need renaming to "Dwx" or "Policy"
- SharePoint lists use "JML_" prefix (e.g., JML_Policies)
- Components: JmlAppLayout, JmlAppHeader, JmlAppFooter

### Current Naming Convention → Target
| Current | Target |
|---------|--------|
| JmlAppLayout | DwxAppLayout or PolicyAppLayout |
| JmlAppHeader | DwxAppHeader |
| JmlAppFooter | DwxAppFooter |
| JML_Policies | PM_Policies |
| jmlMyPolicies | policyMyPolicies |

---

## Codebase Assessment

### Strengths
1. **Comprehensive Service Layer** - 141 services covering all business domains
2. **Type Safety** - Full TypeScript with strict mode, comprehensive interfaces
3. **SharePoint Integration** - Well-abstracted via PnP.js with singleton pattern
4. **Layered Styling** - Multi-layer approach with Fluent UI and design tokens
5. **Modular WebParts** - 6 independent but composable webparts
6. **Workflow Engine** - Sophisticated workflow execution with state machine pattern
7. **Policy Models** - Comprehensive IPolicy interface with 80+ fields
8. **Audit Logging** - Built-in compliance audit trail
9. **Quiz System** - Policy comprehension assessment capability
10. **Gamification** - Points, badges, leaderboards for engagement

### Weaknesses & Technical Debt

#### High Priority
1. **Heavy JML Coupling** - 3,200+ references to decouple
2. **No Centralized State Management** - State scattered across components/hooks/services
3. **Inadequate Testing** - Only 4 test files for 141+ services (~2.8% coverage)
4. **Service Proliferation** - 141 services suggests weak organization

#### Medium Priority
5. **Empty Documentation** - docs/ folder needs architecture guides
6. **Class Components** - Should migrate to functional components
7. **Duplicate Caching** - PolicyCacheService vs CacheService vs hook caching
8. **No Dependency Injection** - Services manually instantiated
9. **Hardcoded Configuration** - List names and URLs hardcoded

#### Lower Priority
10. **No Virtual Scrolling** - Performance risk with large lists
11. **Accessibility Gaps** - Missing ARIA labels in custom components
12. **@ts-nocheck Usage** - Disables type safety in many services

---

## SharePoint Lists

**IMPORTANT: All SharePoint lists and libraries MUST use the `PM_` prefix (Policy Manager), NOT `JML_`.**

### Core Lists
| List Name | Purpose |
|-----------|---------|
| PM_Policies | Core policy records |
| PM_PolicyVersions | Version history |
| PM_PolicyAcknowledgements | User acknowledgements |

### Quiz Lists
| List Name | Purpose |
|-----------|---------|
| PM_PolicyQuizzes | Quiz definitions |
| PM_PolicyQuizQuestions | Quiz questions |
| PM_PolicyQuizResults | Quiz results |

### Workflow Lists
| List Name | Purpose |
|-----------|---------|
| PM_PolicyExemptions | Exemption management |
| PM_PolicyDistributions | Distribution tracking |
| PM_PolicyTemplates | Policy templates library |

### Social/Engagement Lists
| List Name | Purpose |
|-----------|---------|
| PM_PolicyRatings | User ratings |
| PM_PolicyComments | Discussion comments |
| PM_PolicyCommentLikes | Comment likes |
| PM_PolicyShares | Share tracking |
| PM_PolicyFollowers | Policy followers |

### Policy Packs
| List Name | Purpose |
|-----------|---------|
| PM_PolicyPacks | Policy bundle definitions |
| PM_PolicyPackAssignments | Pack assignments |

### Analytics & Audit
| List Name | Purpose |
|-----------|---------|
| PM_PolicyAuditLog | Audit trail |
| PM_PolicyAnalytics | Usage analytics |
| PM_PolicyFeedback | User feedback |
| PM_PolicyDocuments | Supporting documents |

### Provisioning Scripts
Located in `scripts/policy-management/`:
- `01-Core-PolicyLists.ps1` - Core lists (Policies, Versions, Acknowledgements)
- `02-Quiz-Lists.ps1` - Quiz system lists
- `03-Exemption-Distribution-Lists.ps1` - Workflow lists
- `04-Social-Lists.ps1` - Social engagement lists
- `05-PolicyPack-Lists.ps1` - Policy pack lists
- `06-Analytics-Audit-Lists.ps1` - Analytics and audit lists
- `Deploy-AllPolicyLists.ps1` - Master deployment script

---

## Key Models

### IPolicy (src/models/IPolicy.ts)
- 80+ fields covering all policy aspects
- Supports versioning, acknowledgement, quizzes
- Data classification and retention
- Regulatory compliance mapping

### Policy Status Lifecycle
```
Draft → In Review → Pending Approval → Approved → Published → Archived/Retired
                  ↓
               Rejected
```

### Read Timeframes
- Immediate, Day 1, Day 3, Week 1, Week 2, Month 1, Month 3, Month 6, Custom

### Compliance Risk Levels
- Critical, High, Medium, Low, Informational

---

## Build Commands

```bash
# Install dependencies
npm install

# Development build
npm run build

# Production build
npm run ship

# Clean build artifacts
npm run clean

# Run tests
npm run test
```

---

## Development Guidelines

### Styling
1. Always use CSS variables from the DWx design tokens
2. Never hardcode colors - reference `--dwx-*` variables
3. Follow the font weight rules (400, 500, 700 only)
4. Use the gradient for modal/panel headers

### Components
1. Prefer functional components with hooks
2. Use the shared DwxAppLayout for full-page webparts
3. Implement error boundaries for resilience
4. Follow Fluent UI patterns for consistency

### Services
1. Use the singleton SPFI instance via `getSP()`
2. Implement proper caching with TTL
3. Add comprehensive audit logging for compliance
4. Handle errors gracefully with user-friendly messages

### Testing
1. All new code should have unit tests
2. Target 60%+ code coverage
3. Mock PnP calls for isolation
4. Test edge cases and error scenarios

---

## Spinoff Roadmap (Recommended)

### Phase 1: Foundation
- [ ] Rename JML prefixes to DWx/Policy
- [ ] Create configuration service for tenant settings
- [ ] Set up comprehensive testing infrastructure
- [ ] Document architecture and APIs

### Phase 2: Decoupling
- [ ] Abstract JML integration behind interfaces
- [ ] Create standalone process models
- [ ] Remove unused JML-specific services
- [ ] Implement dependency injection

### Phase 3: Enhancement
- [ ] Add centralized state management (Zustand/Redux)
- [ ] Migrate class components to functional
- [ ] Consolidate caching strategy
- [ ] Add virtual scrolling for large lists

### Phase 4: Polish
- [ ] Accessibility audit and fixes
- [ ] Performance optimization
- [ ] Security hardening
- [ ] Documentation completion

---

## Contact & Resources

- **Suite**: DWx (Digital Workplace Excellence)
- **Brand**: First Digital
- **Style Guide**: `docs/DWx-Style-Guide-Preview.html`
- **Brand Guide**: `docs/DWx-Brand-Guide.pdf`
