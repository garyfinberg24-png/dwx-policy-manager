// ============================================================================
// DWx Policy Manager - Audience Service
// CRUD for PM_Audiences + evaluation against PM_Employees
// ============================================================================

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import {
  IAudience,
  IAudienceCriteria,
  IAudienceFilter,
  IAudienceEvalResult,
} from '../models/IAudience';

// ============================================================================
// SERVICE
// ============================================================================

export class AudienceService {
  private readonly sp: SPFI;
  private readonly AUDIENCES_LIST = 'PM_Audiences';
  private readonly EMPLOYEES_LIST = 'PM_Employees';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ==========================================================================
  // AUDIENCE CRUD
  // ==========================================================================

  /**
   * Load all audiences
   */
  public async getAudiences(): Promise<IAudience[]> {
    try {
      const items: any[] = await this.sp.web.lists.getByTitle(this.AUDIENCES_LIST).items
        .select('Id', 'Title', 'Description', 'Criteria', 'MemberCount', 'IsActive', 'LastEvaluated')
        .orderBy('Title', true)
        .top(200)();

      return items.map(item => ({
        Id: item.Id,
        Title: item.Title,
        Description: item.Description || '',
        Criteria: this.parseCriteria(item.Criteria),
        MemberCount: item.MemberCount || 0,
        IsActive: item.IsActive !== false,
        LastEvaluated: item.LastEvaluated || undefined,
      }));
    } catch (err) {
      console.warn('AudienceService.getAudiences failed:', err);
      return [];
    }
  }

  /**
   * Get a single audience by ID
   */
  public async getAudience(id: number): Promise<IAudience | null> {
    try {
      const item: any = await this.sp.web.lists.getByTitle(this.AUDIENCES_LIST).items
        .getById(id)
        .select('Id', 'Title', 'Description', 'Criteria', 'MemberCount', 'IsActive', 'LastEvaluated')();

      return {
        Id: item.Id,
        Title: item.Title,
        Description: item.Description || '',
        Criteria: this.parseCriteria(item.Criteria),
        MemberCount: item.MemberCount || 0,
        IsActive: item.IsActive !== false,
        LastEvaluated: item.LastEvaluated || undefined,
      };
    } catch (err) {
      console.warn('AudienceService.getAudience failed:', err);
      return null;
    }
  }

  /**
   * Create a new audience
   */
  public async createAudience(audience: Omit<IAudience, 'Id'>): Promise<IAudience> {
    const result = await this.sp.web.lists.getByTitle(this.AUDIENCES_LIST).items.add({
      Title: audience.Title,
      Description: audience.Description,
      Criteria: JSON.stringify(audience.Criteria),
      MemberCount: audience.MemberCount,
      IsActive: audience.IsActive,
    });

    return {
      ...audience,
      Id: result.data.Id,
    };
  }

  /**
   * Update an existing audience
   */
  public async updateAudience(id: number, updates: Partial<IAudience>): Promise<void> {
    const payload: any = {};

    if (updates.Title !== undefined) payload.Title = updates.Title;
    if (updates.Description !== undefined) payload.Description = updates.Description;
    if (updates.Criteria !== undefined) payload.Criteria = JSON.stringify(updates.Criteria);
    if (updates.MemberCount !== undefined) payload.MemberCount = updates.MemberCount;
    if (updates.IsActive !== undefined) payload.IsActive = updates.IsActive;
    if (updates.LastEvaluated !== undefined) payload.LastEvaluated = updates.LastEvaluated;

    await this.sp.web.lists.getByTitle(this.AUDIENCES_LIST).items
      .getById(id)
      .update(payload);
  }

  /**
   * Delete an audience
   */
  public async deleteAudience(id: number): Promise<void> {
    await this.sp.web.lists.getByTitle(this.AUDIENCES_LIST).items
      .getById(id)
      .delete();
  }

  // ==========================================================================
  // AUDIENCE EVALUATION
  // ==========================================================================

  /**
   * Evaluate audience criteria against PM_Employees.
   * Returns matching count and a preview of the first 10 users.
   */
  public async evaluateAudience(criteria: IAudienceCriteria): Promise<IAudienceEvalResult> {
    try {
      const filterStr = this.buildODataFilter(criteria);

      // Get count
      let countQuery = this.sp.web.lists.getByTitle(this.EMPLOYEES_LIST).items
        .select('Id')
        .filter("Status eq 'Active'");
      if (filterStr) {
        countQuery = countQuery.filter(`${filterStr} and Status eq 'Active'`);
      }
      const allIds = await countQuery.top(5000)();
      const count = allIds.length;

      // Get preview (first 10)
      let previewQuery = this.sp.web.lists.getByTitle(this.EMPLOYEES_LIST).items
        .select('Title', 'Email', 'Department', 'JobTitle')
        .orderBy('Title', true);
      if (filterStr) {
        previewQuery = previewQuery.filter(`${filterStr} and Status eq 'Active'`);
      } else {
        previewQuery = previewQuery.filter("Status eq 'Active'");
      }
      const preview = await previewQuery.top(10)();

      return {
        count,
        preview: preview.map((p: any) => ({
          Title: p.Title,
          Email: p.Email,
          Department: p.Department || undefined,
          JobTitle: p.JobTitle || undefined,
        })),
      };
    } catch (err) {
      console.warn('AudienceService.evaluateAudience failed:', err);
      return { count: 0, preview: [] };
    }
  }

  /**
   * Evaluate and persist — updates the audience's MemberCount and LastEvaluated
   */
  public async evaluateAndSave(audienceId: number, criteria: IAudienceCriteria): Promise<IAudienceEvalResult> {
    const result = await this.evaluateAudience(criteria);

    await this.updateAudience(audienceId, {
      MemberCount: result.count,
      LastEvaluated: new Date().toISOString(),
    });

    return result;
  }

  // ==========================================================================
  // ODATA FILTER BUILDER
  // ==========================================================================

  /**
   * Translate IAudienceCriteria into an OData $filter string.
   * Combines filters with the criteria's AND/OR operator.
   */
  private buildODataFilter(criteria: IAudienceCriteria): string {
    if (!criteria.filters || criteria.filters.length === 0) {
      return '';
    }

    const parts: string[] = [];

    for (const filter of criteria.filters) {
      const clause = this.buildFilterClause(filter);
      if (clause) {
        parts.push(clause);
      }
    }

    if (parts.length === 0) return '';
    if (parts.length === 1) return parts[0];

    const joiner = criteria.operator === 'OR' ? ' or ' : ' and ';
    return `(${parts.join(joiner)})`;
  }

  /**
   * Build a single OData filter clause from an IAudienceFilter
   */
  private buildFilterClause(filter: IAudienceFilter): string {
    const field = this.sanitizeFieldName(filter.field);
    if (!field) return '';

    switch (filter.operator) {
      case 'equals': {
        const val = this.sanitizeValue(String(filter.value));
        return `${field} eq '${val}'`;
      }
      case 'contains': {
        const val = this.sanitizeValue(String(filter.value));
        return `substringof('${val}',${field})`;
      }
      case 'startsWith': {
        const val = this.sanitizeValue(String(filter.value));
        return `startswith(${field},'${val}')`;
      }
      case 'in': {
        const values = Array.isArray(filter.value) ? filter.value : [filter.value];
        if (values.length === 0) return '';
        const clauses = values.map(v => `${field} eq '${this.sanitizeValue(v)}'`);
        return `(${clauses.join(' or ')})`;
      }
      default:
        return '';
    }
  }

  // ==========================================================================
  // SANITIZATION
  // ==========================================================================

  /**
   * Whitelist field names to prevent OData injection
   */
  private sanitizeFieldName(field: string): string {
    const allowed = ['Department', 'JobTitle', 'Location', 'EmploymentType', 'PMRole', 'Status'];
    return allowed.includes(field) ? field : '';
  }

  /**
   * Sanitize string values for OData filter (escape single quotes)
   */
  private sanitizeValue(value: string): string {
    return value.replace(/'/g, "''");
  }

  // ==========================================================================
  // JSON PARSING
  // ==========================================================================

  /**
   * Safely parse criteria JSON from SP Note field
   */
  private parseCriteria(json: string | null | undefined): IAudienceCriteria {
    if (!json) {
      return { filters: [], operator: 'AND' };
    }
    try {
      const parsed = JSON.parse(json);
      return {
        filters: Array.isArray(parsed.filters) ? parsed.filters : [],
        operator: parsed.operator === 'OR' ? 'OR' : 'AND',
      };
    } catch {
      console.warn('AudienceService: Failed to parse criteria JSON');
      return { filters: [], operator: 'AND' };
    }
  }
}
