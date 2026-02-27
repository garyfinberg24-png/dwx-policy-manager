// ============================================================================
// DWx Policy Manager - Audience Targeting Interfaces
// Models for custom audience definitions used in policy distribution
// ============================================================================

/**
 * A saved audience definition stored in PM_Audiences
 */
export interface IAudience {
  Id?: number;
  Title: string;
  Description: string;
  Criteria: IAudienceCriteria;
  MemberCount: number;
  IsActive: boolean;
  LastEvaluated?: string;
}

/**
 * Audience criteria — a set of filters combined with AND/OR logic
 */
export interface IAudienceCriteria {
  filters: IAudienceFilter[];
  operator: 'AND' | 'OR';
}

/**
 * A single filter condition within an audience definition
 */
export interface IAudienceFilter {
  field: AudienceFilterField;
  operator: 'equals' | 'contains' | 'startsWith' | 'in';
  value: string | string[];
}

/**
 * Fields available for audience filtering (matching PM_Employees columns)
 */
export type AudienceFilterField =
  | 'Department'
  | 'JobTitle'
  | 'Location'
  | 'EmploymentType'
  | 'PMRole'
  | 'Status';

/**
 * Result of evaluating an audience against PM_Employees
 */
export interface IAudienceEvalResult {
  count: number;
  preview: Array<{ Title: string; Email: string; Department?: string; JobTitle?: string }>;
}
