// @ts-nocheck
/**
 * AudienceRuleService
 * Evaluates audience rules against PM_UserProfiles to resolve targeting.
 * Used by Policy Builder (audience selection) and Distribution (user resolution).
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

// ============================================================================
// TYPES
// ============================================================================

export interface IAudienceRule {
  field: string;      // PM_UserProfiles field name (Department, JobTitle, Office, etc.)
  operator: string;   // equals, not_equals, contains, starts_with, within_days
  value: string;      // Target value
}

export interface IAudience {
  Id: number;
  Title: string;
  Description: string;
  Rules: IAudienceRule[];
  Combinator: 'AND' | 'OR';
  Category: string;
  IsActive: boolean;
  IsSystem: boolean;
  EstimatedCount: number;
  LastEvaluated: string;
}

export interface IResolvedUser {
  Id: number;
  Title: string;
  Email: string;
  Department: string;
  JobTitle: string;
  Office: string;
}

// Fields available for audience targeting
export const AUDIENCE_FIELDS = [
  { key: 'Department', label: 'Department', type: 'text' },
  { key: 'JobTitle', label: 'Job Title', type: 'text' },
  { key: 'Office', label: 'Office / Location', type: 'text' },
  { key: 'Location', label: 'City / Region', type: 'text' },
  { key: 'PMRole', label: 'Policy Manager Role', type: 'choice', choices: ['User', 'Author', 'Manager', 'Admin'] },
  { key: 'EmployeeType', label: 'Employee Type', type: 'choice', choices: ['Employee', 'Contractor', 'Intern', 'Consultant'] },
  { key: 'Company', label: 'Company', type: 'text' },
  { key: 'IsActive', label: 'Active Status', type: 'boolean' },
  { key: 'StartDate', label: 'Start Date', type: 'date' },
  { key: 'ManagerEmail', label: 'Manager', type: 'text' }
];

export const AUDIENCE_OPERATORS = [
  { key: 'equals', label: 'equals', types: ['text', 'choice', 'boolean'] },
  { key: 'not_equals', label: 'does not equal', types: ['text', 'choice'] },
  { key: 'contains', label: 'contains', types: ['text'] },
  { key: 'starts_with', label: 'starts with', types: ['text'] },
  { key: 'within_days', label: 'within last N days', types: ['date'] }
];

// ============================================================================
// SERVICE
// ============================================================================

export class AudienceRuleService {
  private sp: SPFI;
  private readonly AUDIENCES_LIST = 'PM_Audiences';
  private readonly USER_PROFILES_LIST = 'PM_UserProfiles';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ============================================================================
  // AUDIENCE CRUD
  // ============================================================================

  /** Load all active audiences */
  public async getAudiences(): Promise<IAudience[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.AUDIENCES_LIST)
        .items.filter('IsActive eq 1')
        .select('Id', 'Title', 'Description', 'Rules', 'Combinator', 'Category', 'IsActive', 'IsSystem', 'EstimatedCount', 'LastEvaluated')
        .orderBy('Title')
        .top(100)();

      return items.map((item: any) => ({
        Id: item.Id,
        Title: item.Title,
        Description: item.Description || '',
        Rules: this.parseRules(item.Rules),
        Combinator: item.Combinator || 'AND',
        Category: item.Category || 'Custom',
        IsActive: item.IsActive !== false,
        IsSystem: item.IsSystem || false,
        EstimatedCount: item.EstimatedCount || 0,
        LastEvaluated: item.LastEvaluated || ''
      }));
    } catch (err) {
      console.warn('[AudienceRuleService] Failed to load audiences:', err);
      return [];
    }
  }

  /** Create a new audience */
  public async createAudience(audience: Partial<IAudience>): Promise<number> {
    const result = await this.sp.web.lists
      .getByTitle(this.AUDIENCES_LIST)
      .items.add({
        Title: audience.Title,
        Description: audience.Description || '',
        Rules: JSON.stringify(audience.Rules || []),
        Combinator: audience.Combinator || 'AND',
        Category: audience.Category || 'Custom',
        IsActive: true,
        IsSystem: false
      });
    return result.data.Id;
  }

  /** Update an audience */
  public async updateAudience(id: number, updates: Partial<IAudience>): Promise<void> {
    const data: any = {};
    if (updates.Title !== undefined) data.Title = updates.Title;
    if (updates.Description !== undefined) data.Description = updates.Description;
    if (updates.Rules !== undefined) data.Rules = JSON.stringify(updates.Rules);
    if (updates.Combinator !== undefined) data.Combinator = updates.Combinator;
    if (updates.Category !== undefined) data.Category = updates.Category;
    if (updates.IsActive !== undefined) data.IsActive = updates.IsActive;
    if (updates.EstimatedCount !== undefined) data.EstimatedCount = updates.EstimatedCount;
    if (updates.LastEvaluated !== undefined) data.LastEvaluated = updates.LastEvaluated;

    await this.sp.web.lists
      .getByTitle(this.AUDIENCES_LIST)
      .items.getById(id).update(data);
  }

  /** Delete an audience (non-system only) */
  public async deleteAudience(id: number): Promise<void> {
    await this.sp.web.lists
      .getByTitle(this.AUDIENCES_LIST)
      .items.getById(id).delete();
  }

  // ============================================================================
  // RULE EVALUATION ENGINE
  // ============================================================================

  /**
   * Resolve an audience to a list of matching users.
   * Evaluates the rules against PM_UserProfiles.
   */
  public async resolveAudience(audience: IAudience): Promise<IResolvedUser[]> {
    return this.evaluateRules(audience.Rules, audience.Combinator);
  }

  /**
   * Resolve audience by ID
   */
  public async resolveAudienceById(audienceId: number): Promise<IResolvedUser[]> {
    const audiences = await this.getAudiences();
    const audience = audiences.find(a => a.Id === audienceId);
    if (!audience) return [];
    return this.resolveAudience(audience);
  }

  /**
   * Evaluate rules against PM_UserProfiles and return matching users.
   * Rules are evaluated client-side for flexibility (SP OData can't do "contains" on text fields).
   */
  public async evaluateRules(rules: IAudienceRule[], combinator: 'AND' | 'OR'): Promise<IResolvedUser[]> {
    if (!rules || rules.length === 0) return [];

    try {
      // Load all active user profiles
      let users: any[];
      try {
        users = await this.sp.web.lists
          .getByTitle(this.USER_PROFILES_LIST)
          .items.filter('IsActive eq 1')
          .select('Id', 'Title', 'Email', 'Department', 'JobTitle', 'Office', 'Location', 'PMRole', 'PMRoles', 'EmployeeType', 'Company', 'ManagerEmail', 'StartDate', 'IsActive')
          .top(5000)();
      } catch {
        // Fallback without optional columns
        users = await this.sp.web.lists
          .getByTitle(this.USER_PROFILES_LIST)
          .items.filter('IsActive eq 1')
          .select('Id', 'Title', 'Department', 'JobTitle', 'PMRole', 'IsActive')
          .top(5000)();
      }

      // Evaluate each user against the rules
      const matched = users.filter((user: any) => {
        const results = rules.map(rule => this.evaluateRule(user, rule));
        return combinator === 'AND'
          ? results.every(r => r)
          : results.some(r => r);
      });

      return matched.map((user: any) => ({
        Id: user.Id,
        Title: user.Title || '',
        Email: user.Email || '',
        Department: user.Department || '',
        JobTitle: user.JobTitle || '',
        Office: user.Office || user.Location || ''
      }));
    } catch (err) {
      console.error('[AudienceRuleService] Failed to evaluate rules:', err);
      return [];
    }
  }

  /**
   * Evaluate a single rule against a user record.
   */
  private evaluateRule(user: any, rule: IAudienceRule): boolean {
    const fieldValue = String(user[rule.field] || '').toLowerCase();
    const ruleValue = String(rule.value || '').toLowerCase();

    switch (rule.operator) {
      case 'equals':
        return fieldValue === ruleValue;

      case 'not_equals':
        return fieldValue !== ruleValue;

      case 'contains':
        return fieldValue.includes(ruleValue);

      case 'starts_with':
        return fieldValue.startsWith(ruleValue);

      case 'within_days': {
        // For date fields: check if the date is within the last N days
        const fieldDate = user[rule.field] ? new Date(user[rule.field]) : null;
        if (!fieldDate || isNaN(fieldDate.getTime())) return false;
        const daysAgo = parseInt(rule.value, 10) || 90;
        const threshold = new Date();
        threshold.setDate(threshold.getDate() - daysAgo);
        return fieldDate >= threshold;
      }

      default:
        return false;
    }
  }

  /**
   * Get estimated count for an audience (without returning full user list)
   */
  public async getEstimatedCount(rules: IAudienceRule[], combinator: 'AND' | 'OR'): Promise<number> {
    const users = await this.evaluateRules(rules, combinator);
    return users.length;
  }

  /**
   * Get distinct values for a field (for autocomplete in rule builder)
   */
  public async getDistinctValues(fieldName: string): Promise<string[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.USER_PROFILES_LIST)
        .items.filter('IsActive eq 1')
        .select(fieldName)
        .top(5000)();

      const values = new Set<string>();
      for (const item of items) {
        const val = item[fieldName];
        if (val && String(val).trim()) {
          values.add(String(val).trim());
        }
      }
      return Array.from(values).sort();
    } catch {
      return [];
    }
  }

  // ============================================================================
  // HELPERS
  // ============================================================================

  private parseRules(rulesJson: string): IAudienceRule[] {
    try {
      if (!rulesJson) return [];
      const parsed = JSON.parse(rulesJson);
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  }
}
