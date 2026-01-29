# JML_ to PM_ List Name Migration Guide

This document tracks the migration of hardcoded `JML_` list name references to centralized `PM_LISTS` constants.

## Migration Status: ✅ COMPLETED

All Policy Manager code files have been updated to use the `PM_LISTS` constants.

---

## Centralized Constants Location

All list names are defined in:
```
src/constants/SharePointListNames.ts
```

Import and use like:
```typescript
import { PM_LISTS } from '../constants/SharePointListNames';

// Instead of: .getByTitle('JML_PolicyTemplates')
// Use:        .getByTitle(PM_LISTS.POLICY_TEMPLATES)
```

---

## Completed Updates

### 1. PolicyAuthorEnhanced.tsx ✅
**File:** `src/webparts/jmlPolicyAuthor/components/PolicyAuthorEnhanced.tsx`

All 28+ JML_ references updated including:
- `JML_PolicyTemplates` → `PM_LISTS.POLICY_TEMPLATES`
- `JML_PolicyMetadataProfiles` → `PM_LISTS.POLICY_METADATA_PROFILES`
- `JML_PolicySourceDocuments` → `PM_LISTS.POLICY_SOURCE_DOCUMENTS`
- `JML_PolicyReviewers` → `PM_LISTS.POLICY_REVIEWERS`
- `JML_PolicyDelegations` → `PM_LISTS.DELEGATIONS`
- `JML_PolicyPacks` → `PM_LISTS.POLICY_PACKS`
- `JML_PolicyQuizzes` → `PM_LISTS.POLICY_QUIZZES`
- `JML_QuizQuestions` → `PM_LISTS.POLICY_QUIZ_QUESTIONS`
- `JML_Policies` → `PM_LISTS.POLICIES`
- `JML_CorporateTemplates` → `PM_LISTS.CORPORATE_TEMPLATES`

---

### 2. PolicyDetails.tsx ✅
**File:** `src/webparts/jmlPolicyDetails/components/PolicyDetails.tsx`

- `JML_PolicyReadReceipts` → `PM_LISTS.POLICY_READ_RECEIPTS`

---

### 3. IModuleRegistry.ts ✅
**File:** `src/models/IModuleRegistry.ts`

Policy Management module list definitions updated:
- `JML_Policies` → `PM_Policies`
- `JML_PolicyVersions` → `PM_PolicyVersions`
- `JML_PolicyAcknowledgements` → `PM_PolicyAcknowledgements`
- `JML_PolicyPacks` → `PM_PolicyPacks`
- `JML_PolicyCategories` → `PM_PolicyCategories`

---

## Constants Added to SharePointListNames.ts ✅

The following constants were added to `src/constants/SharePointListNames.ts`:

```typescript
// Added to PolicyLists
POLICY_METADATA_PROFILES: 'PM_PolicyMetadataProfiles',
POLICY_REVIEWERS: 'PM_PolicyReviewers',
POLICY_READ_RECEIPTS: 'PM_PolicyReadReceipts',
POLICY_CATEGORIES: 'PM_PolicyCategories',

// Added to TemplateLibraryLists
CORPORATE_TEMPLATES: 'PM_CorporateTemplates',
```

Legacy mappings also added to `LegacyListMapping`.

---

## SharePoint Lists Status

### Existing Lists ✅
These PM_ prefixed lists already exist in SharePoint:
- `PM_PolicyTemplates`
- `PM_Policies`
- `PM_PolicyVersions`
- `PM_PolicyAcknowledgements`
- `PM_PolicyPacks`
- `PM_PolicyQuizzes`
- `PM_PolicySourceDocuments`

### Lists Needing Provisioning ⚠️
These lists need to be created on SharePoint:

| List Name | Description | Script |
|-----------|-------------|--------|
| `PM_PolicyMetadataProfiles` | Metadata profiles for policies | `scripts/Create-PM_PolicyMetadataProfiles.ps1` ✅ |
| `PM_PolicyReviewers` | Policy reviewer assignments | Need to create |
| `PM_PolicyReadReceipts` | Read receipt tracking | Need to create |
| `PM_PolicyCategories` | Policy categorization | Need to create |
| `PM_CorporateTemplates` | Corporate document templates | Need to create |

---

## Remaining Steps

1. ✅ **Add missing constants** to `SharePointListNames.ts`
2. ⚠️ **Create missing lists** on SharePoint using provisioning scripts
3. ✅ **Update hardcoded references** in source files
4. ✅ **Build the solution**: `npm run build`
5. ⬜ **Package for deployment**: `npm run ship`
6. ⬜ **Deploy to App Catalog** and update the site

---

## Verification Commands

To verify no JML_ list references remain in Policy Manager code:

```bash
# PowerShell
Get-ChildItem -Path src -Include *.ts,*.tsx -Recurse | Select-String -Pattern "JML_" | Where-Object { $_.Path -notmatch "SharePointListNames.ts" -and $_.Line -notmatch "JMLIntegrationLists" }

# Bash/grep
grep -rn "JML_" src/ --include="*.ts" --include="*.tsx" | grep -v "SharePointListNames.ts" | grep -v "JMLIntegrationLists"
```

---

## Notes

- The `JMLIntegrationLists` in `SharePointListNames.ts` are intentionally kept with `JML_` prefix as they integrate with the JML system.
- The `LegacyListMapping` in `SharePointListNames.ts` is for reference/migration only and should not be used in production code.
- The `IModuleRegistry.ts` file contains JML_ references for other modules (not Policy Manager) - these are intentional and should remain.
