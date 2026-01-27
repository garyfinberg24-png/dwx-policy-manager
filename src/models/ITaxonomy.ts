/**
 * JML Taxonomy Models
 *
 * Interfaces for SharePoint Managed Metadata (Term Store) integration.
 * Used for the JML Metadata tagging system.
 */

/**
 * Represents a term from the SharePoint Term Store
 */
export interface IJmlTerm {
  /** Unique identifier (GUID) */
  id: string;
  /** Display name of the term */
  name: string;
  /** Description of the term */
  description?: string;
  /** Parent term ID (for hierarchical terms) */
  parentId?: string;
  /** Term set ID this term belongs to */
  termSetId: string;
  /** Child terms */
  children?: IJmlTerm[];
  /** Sort order within parent */
  sortOrder?: number;
  /** Whether the term is available for tagging */
  isAvailableForTagging: boolean;
  /** Whether the term is deprecated */
  isDeprecated: boolean;
  /** Custom properties */
  customProperties?: Record<string, string>;
  /** Labels (including synonyms) */
  labels?: IJmlTermLabel[];
  /** Path from root (e.g., "Business Domains;Human Resources;Onboarding") */
  path?: string;
  /** Depth level in hierarchy (0 = root) */
  level?: number;
}

/**
 * Represents a term label (name and language)
 */
export interface IJmlTermLabel {
  /** The label text */
  name: string;
  /** Whether this is the default label */
  isDefault: boolean;
  /** Language tag (e.g., "en-US") */
  languageTag: string;
}

/**
 * Represents a term set from the SharePoint Term Store
 */
export interface IJmlTermSet {
  /** Unique identifier (GUID) */
  id: string;
  /** Display name of the term set */
  name: string;
  /** Description of the term set */
  description?: string;
  /** Term group ID this set belongs to */
  groupId: string;
  /** Whether users can add terms */
  isOpenForTermCreation: boolean;
  /** Terms within this set */
  terms?: IJmlTerm[];
}

/**
 * Represents a term group from the SharePoint Term Store
 */
export interface IJmlTermGroup {
  /** Unique identifier (GUID) */
  id: string;
  /** Display name of the group */
  name: string;
  /** Description of the group */
  description?: string;
  /** Term sets within this group */
  termSets?: IJmlTermSet[];
}

/**
 * JML-specific term set identifiers
 * These map to the term sets created by the provisioning script
 */
export enum JmlTermSetType {
  BusinessDomains = 'BusinessDomains',
  LifecycleStages = 'LifecycleStages',
  ComplianceRegulatory = 'ComplianceRegulatory',
  PriorityRisk = 'PriorityRisk',
  ContentClassification = 'ContentClassification',
  DocumentTypes = 'DocumentTypes',
  Audience = 'Audience'
}

/**
 * Configuration for JML term sets (loaded from config)
 */
export interface IJmlTermStoreConfig {
  termGroupName: string;
  termGroupId: string;
  termSets: {
    [key in JmlTermSetType]: string;
  };
}

/**
 * A selected tag (term) for assignment to an entity
 */
export interface IJmlTagSelection {
  /** Term ID */
  termId: string;
  /** Term name (for display) */
  termName: string;
  /** Term set ID */
  termSetId: string;
  /** Term set type */
  termSetType: JmlTermSetType;
  /** Full path (for hierarchical display) */
  path?: string;
}

/**
 * Entity types that support tagging
 */
export type TaggableEntityType =
  | 'Policy'
  | 'Document'
  | 'Task'
  | 'Process'
  | 'Asset'
  | 'Training'
  | 'HelpArticle'
  | 'Contract'
  | 'Checklist';

/**
 * Tag assignment record (for tracking who tagged what)
 */
export interface IJmlTagAssignment {
  /** Entity type being tagged */
  entityType: TaggableEntityType;
  /** Entity ID */
  entityId: number;
  /** Selected tags */
  tags: IJmlTagSelection[];
  /** User who assigned the tags */
  assignedBy?: string;
  /** When tags were assigned */
  assignedDate?: Date;
}

/**
 * Filter criteria for searching by tags
 */
export interface IJmlTagFilter {
  /** Term set to filter by */
  termSetType?: JmlTermSetType;
  /** Specific term IDs to match */
  termIds?: string[];
  /** Match mode: 'any' (OR) or 'all' (AND) */
  matchMode?: 'any' | 'all';
  /** Include child terms when filtering by parent */
  includeChildren?: boolean;
}

/**
 * Tag usage statistics
 */
export interface IJmlTagUsageStats {
  /** Term ID */
  termId: string;
  /** Term name */
  termName: string;
  /** Term set type */
  termSetType: JmlTermSetType;
  /** Number of times used */
  usageCount: number;
  /** Entities using this tag */
  entityBreakdown: {
    entityType: TaggableEntityType;
    count: number;
  }[];
}

/**
 * Props for taxonomy picker components
 */
export interface ITaxonomyPickerProps {
  /** Term set type to display */
  termSetType: JmlTermSetType;
  /** Currently selected terms */
  selectedTerms: IJmlTagSelection[];
  /** Callback when selection changes */
  onChange: (terms: IJmlTagSelection[]) => void;
  /** Allow multiple selections */
  allowMultiple?: boolean;
  /** Label for the picker */
  label?: string;
  /** Placeholder text */
  placeholder?: string;
  /** Whether the field is required */
  required?: boolean;
  /** Whether the picker is disabled */
  disabled?: boolean;
  /** Show only specific terms (for filtering) */
  filterTermIds?: string[];
  /** Maximum selections allowed */
  maxSelections?: number;
  /** Error message to display */
  errorMessage?: string;
}

/**
 * Managed Metadata field value (as stored in SharePoint)
 */
export interface IManagedMetadataFieldValue {
  /** Term GUID */
  TermGuid: string;
  /** Term label */
  Label: string;
  /** WssId (SharePoint internal ID) */
  WssId: number;
}

/**
 * Helper to convert SharePoint MM field to IJmlTagSelection
 */
export function convertManagedMetadataToTag(
  fieldValue: IManagedMetadataFieldValue,
  termSetType: JmlTermSetType,
  termSetId: string
): IJmlTagSelection {
  return {
    termId: fieldValue.TermGuid,
    termName: fieldValue.Label,
    termSetId: termSetId,
    termSetType: termSetType
  };
}

/**
 * Helper to convert IJmlTagSelection to SharePoint MM field format
 */
export function convertTagToManagedMetadata(
  tag: IJmlTagSelection
): { TermGuid: string; Label: string } {
  return {
    TermGuid: tag.termId,
    Label: tag.termName
  };
}
