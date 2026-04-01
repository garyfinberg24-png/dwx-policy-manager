// @ts-nocheck
/**
 * SchemaValidatorService — Compares expected SP list schema (from provisioning
 * scripts) against actual list columns at runtime. Reports missing columns,
 * type mismatches, and missing lists.
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';

// ============================================================================
// TYPES
// ============================================================================

export interface ISchemaIssue {
  list: string;
  issue: string;
  severity: 'error' | 'warning';
  detail: string;
}

export interface ISchemaValidationResult {
  listName: string;
  exists: boolean;
  expectedColumns: number;
  matchedColumns: number;
  issues: ISchemaIssue[];
}

export interface ISchemaValidationSummary {
  results: ISchemaValidationResult[];
  totalLists: number;
  healthyLists: number;
  totalIssues: number;
  durationMs: number;
}

/** Column definition: internal name → expected SP field type string */
interface IExpectedColumn {
  internalName: string;
  /** SP field TypeAsString values: 'Text', 'Note', 'Choice', 'Number', 'Boolean', 'DateTime', 'User', 'UserMulti', 'URL', 'Counter' */
  type: string;
}

// ============================================================================
// EXPECTED SCHEMAS — derived from provisioning scripts
// Only includes custom columns (Title is implicit on all lists)
// ============================================================================

const EXPECTED_SCHEMAS: Record<string, IExpectedColumn[]> = {
  PM_Policies: [
    { internalName: 'PolicyNumber', type: 'Text' },
    { internalName: 'PolicyDescription', type: 'Note' },
    { internalName: 'PolicyCategory', type: 'Choice' },
    { internalName: 'PolicyStatus', type: 'Choice' },
    { internalName: 'ComplianceRisk', type: 'Choice' },
    { internalName: 'VersionNumber', type: 'Text' },
    { internalName: 'EffectiveDate', type: 'DateTime' },
    { internalName: 'ExpiryDate', type: 'DateTime' },
    { internalName: 'IsActive', type: 'Boolean' },
    { internalName: 'PolicyOwner', type: 'User' },
    { internalName: 'Visibility', type: 'Choice' },
    { internalName: 'TargetSecurityGroups', type: 'Note' },
    { internalName: 'SubCategory', type: 'Text' },
    { internalName: 'LinkedQuizId', type: 'Number' },
    { internalName: 'SourceRequestId', type: 'Number' },
    { internalName: 'CreationMethod', type: 'Text' },
  ],
  PM_PolicyVersions: [
    { internalName: 'PolicyId', type: 'Number' },
    { internalName: 'VersionNumber', type: 'Text' },
    { internalName: 'VersionType', type: 'Choice' },
    { internalName: 'VersionDescription', type: 'Note' },
    { internalName: 'HTMLContent', type: 'Note' },
    { internalName: 'CreatedByEmail', type: 'Text' },
  ],
  PM_PolicyAcknowledgements: [
    { internalName: 'PolicyId', type: 'Number' },
    { internalName: 'AckUserId', type: 'Number' },
    { internalName: 'AckStatus', type: 'Choice' },
    { internalName: 'AcknowledgedDate', type: 'DateTime' },
    { internalName: 'DueDate', type: 'DateTime' },
  ],
  PM_PolicyTemplates: [
    { internalName: 'TemplateType', type: 'Choice' },
    { internalName: 'TemplateCategory', type: 'Choice' },
    { internalName: 'TemplateDescription', type: 'Note' },
    { internalName: 'TemplateContent', type: 'Note' },
    { internalName: 'HTMLTemplate', type: 'Note' },
    { internalName: 'DocumentTemplateURL', type: 'Note' },
    { internalName: 'IsActive', type: 'Boolean' },
    { internalName: 'UsageCount', type: 'Number' },
    { internalName: 'Tags', type: 'Note' },
  ],
  PM_PolicyAuditLog: [
    { internalName: 'AuditAction', type: 'Choice' },
    { internalName: 'EntityType', type: 'Choice' },
    { internalName: 'EntityId', type: 'Number' },
    { internalName: 'ActionDescription', type: 'Note' },
    { internalName: 'PerformedByEmail', type: 'Text' },
  ],
  PM_Configuration: [
    { internalName: 'ConfigKey', type: 'Text' },
    { internalName: 'ConfigValue', type: 'Note' },
    { internalName: 'Category', type: 'Text' },
    { internalName: 'IsActive', type: 'Boolean' },
    { internalName: 'IsSystemConfig', type: 'Boolean' },
  ],
  PM_Approvals: [
    { internalName: 'PolicyId', type: 'Number' },
    { internalName: 'ApprovalStatus', type: 'Choice' },
    { internalName: 'ApproverEmail', type: 'Text' },
    { internalName: 'RequestedByEmail', type: 'Text' },
    { internalName: 'ApprovalLevel', type: 'Number' },
  ],
  PM_ApprovalHistory: [
    { internalName: 'ApprovalId', type: 'Number' },
    { internalName: 'AuditAction', type: 'Choice' },
    { internalName: 'ActionDescription', type: 'Note' },
    { internalName: 'PerformedByEmail', type: 'Text' },
  ],
  PM_ApprovalDelegations: [
    { internalName: 'DelegatedById', type: 'Number' },
    { internalName: 'DelegatedToId', type: 'Number' },
    { internalName: 'DelegatedByEmail', type: 'Text' },
    { internalName: 'DelegatedToEmail', type: 'Text' },
    { internalName: 'StartDate', type: 'DateTime' },
    { internalName: 'EndDate', type: 'DateTime' },
    { internalName: 'IsActive', type: 'Boolean' },
  ],
  PM_Notifications: [
    { internalName: 'Type', type: 'Choice' },
    { internalName: 'RecipientEmail', type: 'Text' },
    { internalName: 'ActionUrl', type: 'Note' },
    { internalName: 'IsRead', type: 'Boolean' },
  ],
  PM_NotificationQueue: [
    { internalName: 'To', type: 'Text' },
    { internalName: 'Subject', type: 'Text' },
    { internalName: 'Message', type: 'Note' },
    { internalName: 'QueueStatus', type: 'Choice' },
    { internalName: 'RetryCount', type: 'Number' },
  ],
  PM_PolicyQuizzes: [
    { internalName: 'PolicyId', type: 'Number' },
    { internalName: 'PassingScore', type: 'Number' },
    { internalName: 'MaxAttempts', type: 'Number' },
    { internalName: 'IsActive', type: 'Boolean' },
  ],
  PM_PolicyQuizQuestions: [
    { internalName: 'QuizId', type: 'Number' },
    { internalName: 'QuestionType', type: 'Choice' },
    { internalName: 'QuestionText', type: 'Note' },
    { internalName: 'CorrectAnswer', type: 'Note' },
    { internalName: 'Points', type: 'Number' },
  ],
  PM_PolicyQuizResults: [
    { internalName: 'QuizId', type: 'Number' },
    { internalName: 'UserId', type: 'Number' },
    { internalName: 'Score', type: 'Number' },
    { internalName: 'Passed', type: 'Boolean' },
  ],
  PM_PolicyPacks: [
    { internalName: 'PackDescription', type: 'Note' },
    { internalName: 'PolicyIds', type: 'Note' },
    { internalName: 'IsActive', type: 'Boolean' },
  ],
  PM_PolicyPackAssignments: [
    { internalName: 'PackId', type: 'Number' },
    { internalName: 'AssignedToEmail', type: 'Text' },
    { internalName: 'AssignmentStatus', type: 'Choice' },
  ],
  PM_EventLog: [
    { internalName: 'EventSeverity', type: 'Choice' },
    { internalName: 'EventChannel', type: 'Choice' },
    { internalName: 'EventCode', type: 'Text' },
    { internalName: 'EventSource', type: 'Text' },
    { internalName: 'EventMessage', type: 'Note' },
    { internalName: 'StackTrace', type: 'Note' },
    { internalName: 'SessionId', type: 'Text' },
    { internalName: 'PageUrl', type: 'Text' },
  ],
  PM_UserProfiles: [
    { internalName: 'Email', type: 'Text' },
    { internalName: 'DisplayName', type: 'Text' },
    { internalName: 'Department', type: 'Text' },
    { internalName: 'JobTitle', type: 'Text' },
    { internalName: 'EmployeeStatus', type: 'Choice' },
  ],
  PM_ReminderSchedule: [
    { internalName: 'PolicyId', type: 'Number' },
    { internalName: 'RecipientEmail', type: 'Text' },
    { internalName: 'ReminderType', type: 'Choice' },
    { internalName: 'ScheduledDate', type: 'DateTime' },
    { internalName: 'ReminderStatus', type: 'Choice' },
  ],
};

// ============================================================================
// SERVICE
// ============================================================================

export class SchemaValidatorService {
  private _sp: SPFI;

  constructor(sp: SPFI) {
    this._sp = sp;
  }

  /**
   * Validate all lists with defined schemas.
   */
  public async validateAll(): Promise<ISchemaValidationSummary> {
    const startTime = Date.now();
    const results: ISchemaValidationResult[] = [];

    const listNames = Object.keys(EXPECTED_SCHEMAS);

    for (const listName of listNames) {
      const result = await this._validateList(listName, EXPECTED_SCHEMAS[listName]);
      results.push(result);
    }

    const healthyLists = results.filter(r => r.exists && r.issues.length === 0).length;
    const totalIssues = results.reduce((sum, r) => sum + r.issues.length, 0);

    return {
      results,
      totalLists: listNames.length,
      healthyLists,
      totalIssues,
      durationMs: Date.now() - startTime,
    };
  }

  private async _validateList(listName: string, expectedColumns: IExpectedColumn[]): Promise<ISchemaValidationResult> {
    const issues: ISchemaIssue[] = [];

    try {
      // Get all fields from the list
      const fields = await this._sp.web.lists.getByTitle(listName)
        .fields
        .filter("Hidden eq false and ReadOnlyField eq false")
        .select('InternalName', 'TypeAsString', 'Title')();

      const fieldMap = new Map<string, string>();
      for (const f of fields) {
        fieldMap.set(f.InternalName, f.TypeAsString);
      }

      let matchedColumns = 0;

      for (const expected of expectedColumns) {
        const actualType = fieldMap.get(expected.internalName);

        if (!actualType) {
          issues.push({
            list: listName,
            issue: 'Missing column',
            severity: 'error',
            detail: `Column "${expected.internalName}" (${expected.type}) not found. Run the provisioning script for ${listName}.`,
          });
        } else {
          // Normalize type comparison — SP returns various type strings
          const normalizedActual = this._normalizeType(actualType);
          const normalizedExpected = this._normalizeType(expected.type);

          if (normalizedActual !== normalizedExpected) {
            issues.push({
              list: listName,
              issue: 'Type mismatch',
              severity: 'warning',
              detail: `Column "${expected.internalName}" is "${actualType}" but expected "${expected.type}".`,
            });
          } else {
            matchedColumns++;
          }
        }
      }

      return {
        listName,
        exists: true,
        expectedColumns: expectedColumns.length,
        matchedColumns,
        issues,
      };
    } catch (err: any) {
      const is404 = err?.status === 404 || err?.message?.includes('does not exist');
      return {
        listName,
        exists: false,
        expectedColumns: expectedColumns.length,
        matchedColumns: 0,
        issues: [{
          list: listName,
          issue: is404 ? 'List not found' : 'Access error',
          severity: 'error',
          detail: is404
            ? `${listName} does not exist. Run Deploy-AllPolicyLists.ps1 to provision.`
            : `Error accessing ${listName}: ${err?.message || 'Unknown'}`,
        }],
      };
    }
  }

  /**
   * Normalize SP type strings for comparison.
   * SP returns many variants: 'Text' vs 'Note' vs 'MultiLineText', etc.
   */
  private _normalizeType(type: string): string {
    const t = type.toLowerCase();
    if (t === 'note' || t === 'multilinetext') return 'note';
    if (t === 'text') return 'text';
    if (t === 'choice' || t === 'multichoice') return 'choice';
    if (t === 'number' || t === 'currency') return 'number';
    if (t === 'boolean') return 'boolean';
    if (t === 'datetime') return 'datetime';
    if (t === 'user' || t === 'usermulti') return 'user';
    if (t === 'url') return 'url';
    if (t === 'counter') return 'counter';
    if (t === 'lookup' || t === 'lookupmulti') return 'lookup';
    return t;
  }
}
