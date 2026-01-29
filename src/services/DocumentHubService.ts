// @ts-nocheck
// Document Hub Service
// Core service for Document Hub operations including configuration, taxonomy, and statistics

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import {
  IDocumentHubConfig,
  IDocumentTaxonomy,
  IRetentionPolicy,
  ILegalHold,
  IDocumentHubStats,
  IUserDocumentContext,
  DocumentStatus,
  LegalHoldStatus,
  RetentionUnit,
  ConfidentialityLevel
} from '../models';
import { logger } from './LoggingService';

/**
 * Service for Document Hub core operations
 */
export class DocumentHubService {
  private sp: SPFI;

  // List names
  private readonly CONFIG_LIST = 'PM_DocumentHub_Config';
  private readonly TAXONOMY_LIST = 'PM_DocumentTaxonomy';
  private readonly RETENTION_LIST = 'PM_RetentionPolicies';
  private readonly LEGAL_HOLDS_LIST = 'PM_LegalHolds';
  private readonly REGISTRY_LIST = 'PM_DocumentRegistry';
  private readonly WORKFLOWS_LIST = 'PM_DocumentWorkflows';
  private readonly ACTIVITY_LIST = 'PM_DocumentActivity';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ============================================================================
  // CONFIGURATION
  // ============================================================================

  /**
   * Get all configuration settings
   */
  public async getConfiguration(): Promise<IDocumentHubConfig[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.CONFIG_LIST)
        .items
        .select('Id', 'Title', 'Category', 'SettingKey', 'SettingValue', 'DataType', 'Description', 'IsEncrypted')
        .orderBy('Category')
        .orderBy('SettingKey')();

      return items.map(this.mapToConfig);
    } catch (error) {
      logger.error('DocumentHubService', 'Failed to get configuration:', error);
      throw error;
    }
  }

  /**
   * Get a specific configuration setting
   */
  public async getConfigValue(key: string): Promise<string | null> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.CONFIG_LIST)
        .items
        .filter(`SettingKey eq '${key}'`)
        .select('SettingValue')
        .top(1)();

      return items.length > 0 ? items[0].SettingValue : null;
    } catch (error) {
      logger.error('DocumentHubService', `Failed to get config value for ${key}:`, error);
      return null;
    }
  }

  /**
   * Update a configuration setting
   */
  public async updateConfigValue(key: string, value: string): Promise<void> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.CONFIG_LIST)
        .items
        .filter(`SettingKey eq '${key}'`)
        .select('Id')
        .top(1)();

      if (items.length > 0) {
        await this.sp.web.lists
          .getByTitle(this.CONFIG_LIST)
          .items
          .getById(items[0].Id)
          .update({ SettingValue: value });
      }
    } catch (error) {
      logger.error('DocumentHubService', `Failed to update config value for ${key}:`, error);
      throw error;
    }
  }

  /**
   * Save multiple configuration settings
   */
  public async saveConfiguration(config: Record<string, any>): Promise<void> {
    try {
      // Update each setting
      for (const [key, value] of Object.entries(config)) {
        if (value !== undefined && value !== null) {
          await this.updateConfigValue(key, typeof value === 'string' ? value : JSON.stringify(value));
        }
      }
    } catch (error) {
      logger.error('DocumentHubService', 'Failed to save configuration:', error);
      throw error;
    }
  }

  // ============================================================================
  // TAXONOMY
  // ============================================================================

  /**
   * Get all taxonomy terms
   */
  public async getTaxonomy(): Promise<IDocumentTaxonomy[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.TAXONOMY_LIST)
        .items
        .select(
          'Id', 'Title', 'TermCode', 'ParentTermId', 'Level', 'TaxonomyPath',
          'Description', 'Icon', 'Color', 'IsActive', 'SortOrder',
          'DefaultRetentionPolicyId', 'DefaultConfidentiality', 'RequiresApproval',
          'AllowedFileTypes', 'MaxFileSizeMB'
        )
        .filter('IsActive eq 1')
        .orderBy('Level')
        .orderBy('SortOrder')();

      return items.map(this.mapToTaxonomy);
    } catch (error) {
      logger.error('DocumentHubService', 'Failed to get taxonomy:', error);
      throw error;
    }
  }

  /**
   * Get taxonomy as hierarchical tree
   */
  public async getTaxonomyTree(): Promise<IDocumentTaxonomy[]> {
    const flatList = await this.getTaxonomy();
    return this.buildTaxonomyTree(flatList);
  }

  /**
   * Get taxonomy term by ID
   */
  public async getTaxonomyById(id: number): Promise<IDocumentTaxonomy | null> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.TAXONOMY_LIST)
        .items
        .getById(id)
        .select(
          'Id', 'Title', 'TermCode', 'ParentTermId', 'Level', 'TaxonomyPath',
          'Description', 'Icon', 'Color', 'IsActive', 'SortOrder',
          'DefaultRetentionPolicyId', 'DefaultConfidentiality', 'RequiresApproval',
          'AllowedFileTypes', 'MaxFileSizeMB'
        )();

      return this.mapToTaxonomy(item);
    } catch (error) {
      logger.error('DocumentHubService', `Failed to get taxonomy ${id}:`, error);
      return null;
    }
  }

  // ============================================================================
  // RETENTION POLICIES
  // ============================================================================

  /**
   * Get all active retention policies
   */
  public async getRetentionPolicies(): Promise<IRetentionPolicy[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.RETENTION_LIST)
        .items
        .select(
          'Id', 'Title', 'PolicyCode', 'Description', 'RetentionPeriod', 'RetentionUnit',
          'RetentionTrigger', 'DispositionAction', 'ReviewRequired', 'ReviewerIds',
          'NotifyDaysBefore', 'NotifyOwner', 'NotifyAdditional', 'IsActive',
          'EffectiveDate', 'ExpiryDate', 'ApplicableClassifications',
          'ApplicableDepartments', 'LegalBasis', 'Notes'
        )
        .filter('IsActive eq 1')
        .orderBy('Title')();

      return items.map(this.mapToRetentionPolicy);
    } catch (error) {
      logger.error('DocumentHubService', 'Failed to get retention policies:', error);
      throw error;
    }
  }

  /**
   * Get retention policy by ID
   */
  public async getRetentionPolicyById(id: number): Promise<IRetentionPolicy | null> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.RETENTION_LIST)
        .items
        .getById(id)
        .select(
          'Id', 'Title', 'PolicyCode', 'Description', 'RetentionPeriod', 'RetentionUnit',
          'RetentionTrigger', 'DispositionAction', 'ReviewRequired', 'ReviewerIds',
          'NotifyDaysBefore', 'NotifyOwner', 'NotifyAdditional', 'IsActive',
          'EffectiveDate', 'ExpiryDate', 'ApplicableClassifications',
          'ApplicableDepartments', 'LegalBasis', 'Notes'
        )();

      return this.mapToRetentionPolicy(item);
    } catch (error) {
      logger.error('DocumentHubService', `Failed to get retention policy ${id}:`, error);
      return null;
    }
  }

  /**
   * Calculate retention expiry date based on policy
   */
  public calculateRetentionExpiry(policy: IRetentionPolicy, startDate: Date): Date {
    const expiry = new Date(startDate);

    switch (policy.RetentionUnit) {
      case RetentionUnit.Days:
        expiry.setDate(expiry.getDate() + policy.RetentionPeriod);
        break;
      case RetentionUnit.Months:
        expiry.setMonth(expiry.getMonth() + policy.RetentionPeriod);
        break;
      case RetentionUnit.Years:
        expiry.setFullYear(expiry.getFullYear() + policy.RetentionPeriod);
        break;
      case RetentionUnit.Permanent:
        expiry.setFullYear(9999); // Far future date
        break;
    }

    return expiry;
  }

  // ============================================================================
  // LEGAL HOLDS
  // ============================================================================

  /**
   * Get all legal holds (alias for getActiveLegalHolds)
   */
  public async getLegalHolds(): Promise<ILegalHold[]> {
    return this.getActiveLegalHolds();
  }

  /**
   * Get all active legal holds
   */
  public async getActiveLegalHolds(): Promise<ILegalHold[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.LEGAL_HOLDS_LIST)
        .items
        .select(
          'Id', 'Title', 'HoldCode', 'HoldDescription', 'MatterReference',
          'Custodians', 'HoldDepartments', 'Keywords', 'DateRangeStart', 'DateRangeEnd',
          'HoldStatus', 'IssuedDate', 'IssuedById', 'ReleasedDate', 'ReleasedById',
          'ReleaseReason', 'ExternalCounsel', 'CaseNumber', 'DocumentCount', 'Notes',
          'NotifyOnNewDocuments', 'NotificationRecipients'
        )
        .filter('HoldStatus eq \'Active\'')
        .orderBy('IssuedDate', false)();

      return items.map(this.mapToLegalHold);
    } catch (error) {
      logger.error('DocumentHubService', 'Failed to get legal holds:', error);
      throw error;
    }
  }

  /**
   * Get legal hold by ID
   */
  public async getLegalHoldById(id: number): Promise<ILegalHold | null> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.LEGAL_HOLDS_LIST)
        .items
        .getById(id)
        .select(
          'Id', 'Title', 'HoldCode', 'HoldDescription', 'MatterReference',
          'Custodians', 'HoldDepartments', 'Keywords', 'DateRangeStart', 'DateRangeEnd',
          'HoldStatus', 'IssuedDate', 'IssuedById', 'ReleasedDate', 'ReleasedById',
          'ReleaseReason', 'ExternalCounsel', 'CaseNumber', 'DocumentCount', 'Notes',
          'NotifyOnNewDocuments', 'NotificationRecipients'
        )();

      return this.mapToLegalHold(item);
    } catch (error) {
      logger.error('DocumentHubService', `Failed to get legal hold ${id}:`, error);
      return null;
    }
  }

  /**
   * Create a new legal hold
   */
  public async createLegalHold(hold: Partial<ILegalHold>): Promise<ILegalHold> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.LEGAL_HOLDS_LIST)
        .items
        .add({
          Title: hold.Title,
          HoldCode: hold.HoldCode,
          HoldDescription: hold.HoldDescription,
          MatterReference: hold.MatterReference,
          HoldDepartments: hold.HoldDepartments ? { results: hold.HoldDepartments } : undefined,
          Keywords: hold.Keywords,
          DateRangeStart: hold.DateRangeStart?.toISOString(),
          DateRangeEnd: hold.DateRangeEnd?.toISOString(),
          HoldStatus: LegalHoldStatus.Active,
          IssuedDate: new Date().toISOString(),
          ExternalCounsel: hold.ExternalCounsel,
          CaseNumber: hold.CaseNumber,
          DocumentCount: 0,
          Notes: hold.Notes,
          NotifyOnNewDocuments: hold.NotifyOnNewDocuments || false
        });

      return await this.getLegalHoldById(item.data.Id) as ILegalHold;
    } catch (error) {
      logger.error('DocumentHubService', 'Failed to create legal hold:', error);
      throw error;
    }
  }

  /**
   * Release a legal hold
   */
  public async releaseLegalHold(id: number, reason: string, userId: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.LEGAL_HOLDS_LIST)
        .items
        .getById(id)
        .update({
          HoldStatus: LegalHoldStatus.Released,
          ReleasedDate: new Date().toISOString(),
          ReleasedById: userId,
          ReleaseReason: reason
        });
    } catch (error) {
      logger.error('DocumentHubService', `Failed to release legal hold ${id}:`, error);
      throw error;
    }
  }

  // ============================================================================
  // STATISTICS & DASHBOARD
  // ============================================================================

  /**
   * Get Document Hub statistics for dashboard
   */
  public async getStatistics(): Promise<IDocumentHubStats> {
    try {
      // Get document counts by status
      const statusCounts = await this.getDocumentCountsByStatus();

      // Get documents on legal hold
      const legalHoldCount = await this.getDocumentsOnLegalHoldCount();

      // Get documents expiring soon (next 30 days)
      const expiringCount = await this.getDocumentsExpiringSoonCount(30);

      // Get records count
      const recordsCount = await this.getRecordsCount();

      // Get active workflows count
      const activeWorkflows = await this.getActiveWorkflowsCount();

      // Get recent activity
      const recentActivity = await this.getRecentActivity(10);

      // Calculate totals
      const totalDocuments = Object.values(statusCounts).reduce((sum, count) => sum + count, 0);

      return {
        totalDocuments,
        documentsByStatus: statusCounts,
        documentsByClassification: {},
        documentsByDepartment: {},
        documentsOnLegalHold: legalHoldCount,
        documentsExpiringSoon: expiringCount,
        recordsCount,
        activeWorkflows,
        pendingApprovals: 0,
        recentActivity,
        topAccessedDocuments: [],
        storageUsedBytes: 0
      };
    } catch (error) {
      logger.error('DocumentHubService', 'Failed to get statistics:', error);
      throw error;
    }
  }

  /**
   * Get user's document context
   */
  public async getUserContext(userId: number): Promise<IUserDocumentContext> {
    try {
      // Get counts for user
      const myDocuments = await this.getUserDocumentCount(userId);
      const sharedWithMe = await this.getSharedWithUserCount(userId);
      const pendingActions = await this.getUserPendingActionsCount(userId);

      return {
        myDocuments,
        sharedWithMe,
        pendingActions,
        recentDocuments: [],
        savedSearches: [],
        favoriteDocuments: []
      };
    } catch (error) {
      logger.error('DocumentHubService', 'Failed to get user context:', error);
      throw error;
    }
  }

  // ============================================================================
  // PRIVATE HELPER METHODS
  // ============================================================================

  private async getDocumentCountsByStatus(): Promise<Record<DocumentStatus, number>> {
    const counts: Record<string, number> = {};

    try {
      const items = await this.sp.web.lists
        .getByTitle(this.REGISTRY_LIST)
        .items
        .select('DocumentStatus')();

      items.forEach(item => {
        const status = item.DocumentStatus || 'Unknown';
        counts[status] = (counts[status] || 0) + 1;
      });
    } catch (error) {
      logger.error('DocumentHubService', 'Failed to get document counts by status:', error);
    }

    return counts as Record<DocumentStatus, number>;
  }

  private async getDocumentsOnLegalHoldCount(): Promise<number> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.REGISTRY_LIST)
        .items
        .filter('OnLegalHold eq 1')
        .select('Id')();

      return items.length;
    } catch (error) {
      logger.error('DocumentHubService', 'Failed to get legal hold count:', error);
      return 0;
    }
  }

  private async getDocumentsExpiringSoonCount(days: number): Promise<number> {
    try {
      const futureDate = new Date();
      futureDate.setDate(futureDate.getDate() + days);

      const items = await this.sp.web.lists
        .getByTitle(this.REGISTRY_LIST)
        .items
        .filter(`RetentionExpiryDate le '${futureDate.toISOString()}'`)
        .select('Id')();

      return items.length;
    } catch (error) {
      logger.error('DocumentHubService', 'Failed to get expiring documents count:', error);
      return 0;
    }
  }

  private async getRecordsCount(): Promise<number> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.REGISTRY_LIST)
        .items
        .filter('IsRecord eq 1')
        .select('Id')();

      return items.length;
    } catch (error) {
      logger.error('DocumentHubService', 'Failed to get records count:', error);
      return 0;
    }
  }

  private async getActiveWorkflowsCount(): Promise<number> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.WORKFLOWS_LIST)
        .items
        .filter("WorkflowStatus eq 'In Progress' or WorkflowStatus eq 'Pending Approval'")
        .select('Id')();

      return items.length;
    } catch (error) {
      logger.error('DocumentHubService', 'Failed to get active workflows count:', error);
      return 0;
    }
  }

  private async getRecentActivity(count: number): Promise<any[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.ACTIVITY_LIST)
        .items
        .select('Id', 'Title', 'DocumentTitle', 'ActivityType', 'ActivityDate', 'ActivityById')
        .orderBy('ActivityDate', false)
        .top(count)();

      return items;
    } catch (error) {
      logger.error('DocumentHubService', 'Failed to get recent activity:', error);
      return [];
    }
  }

  private async getUserDocumentCount(userId: number): Promise<number> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.REGISTRY_LIST)
        .items
        .filter(`DocumentOwnerId eq ${userId}`)
        .select('Id')();

      return items.length;
    } catch (error) {
      logger.error('DocumentHubService', 'Failed to get user document count:', error);
      return 0;
    }
  }

  private async getSharedWithUserCount(userId: number): Promise<number> {
    // This would require querying the sharing list
    return 0;
  }

  private async getUserPendingActionsCount(userId: number): Promise<number> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.WORKFLOWS_LIST)
        .items
        .filter(`CurrentAssigneesId eq ${userId} and (WorkflowStatus eq 'In Progress' or WorkflowStatus eq 'Pending Approval')`)
        .select('Id')();

      return items.length;
    } catch (error) {
      logger.error('DocumentHubService', 'Failed to get user pending actions count:', error);
      return 0;
    }
  }

  private buildTaxonomyTree(flatList: IDocumentTaxonomy[]): IDocumentTaxonomy[] {
    const map = new Map<number, IDocumentTaxonomy & { children?: IDocumentTaxonomy[] }>();
    const roots: IDocumentTaxonomy[] = [];

    // First pass: create map
    flatList.forEach(item => {
      map.set(item.Id!, { ...item, children: [] });
    });

    // Second pass: build tree
    flatList.forEach(item => {
      const node = map.get(item.Id!);
      if (node) {
        if (item.ParentTermId && map.has(item.ParentTermId)) {
          const parent = map.get(item.ParentTermId);
          if (parent && parent.children) {
            parent.children.push(node);
          }
        } else {
          roots.push(node);
        }
      }
    });

    return roots;
  }

  // ============================================================================
  // MAPPING FUNCTIONS
  // ============================================================================

  private mapToConfig(item: any): IDocumentHubConfig {
    return {
      Id: item.Id,
      Title: item.Title,
      Category: item.Category,
      SettingKey: item.SettingKey,
      SettingValue: item.SettingValue,
      DataType: item.DataType,
      Description: item.Description,
      IsEncrypted: item.IsEncrypted
    };
  }

  private mapToTaxonomy(item: any): IDocumentTaxonomy {
    return {
      Id: item.Id,
      Title: item.Title,
      TermCode: item.TermCode,
      ParentTermId: item.ParentTermId,
      Level: item.Level,
      TaxonomyPath: item.TaxonomyPath,
      Description: item.Description,
      Icon: item.Icon,
      Color: item.Color,
      IsActive: item.IsActive,
      SortOrder: item.SortOrder,
      DefaultRetentionPolicyId: item.DefaultRetentionPolicyId,
      DefaultConfidentiality: item.DefaultConfidentiality as ConfidentialityLevel,
      RequiresApproval: item.RequiresApproval,
      AllowedFileTypes: item.AllowedFileTypes,
      MaxFileSizeMB: item.MaxFileSizeMB
    };
  }

  private mapToRetentionPolicy(item: any): IRetentionPolicy {
    return {
      Id: item.Id,
      Title: item.Title,
      PolicyCode: item.PolicyCode,
      Description: item.Description,
      RetentionPeriod: item.RetentionPeriod,
      RetentionUnit: item.RetentionUnit as RetentionUnit,
      RetentionTrigger: item.RetentionTrigger,
      DispositionAction: item.DispositionAction,
      ReviewRequired: item.ReviewRequired,
      ReviewerIds: item.ReviewerIds,
      NotifyDaysBefore: item.NotifyDaysBefore,
      NotifyOwner: item.NotifyOwner,
      NotifyAdditionalIds: item.NotifyAdditional,
      IsActive: item.IsActive,
      EffectiveDate: item.EffectiveDate ? new Date(item.EffectiveDate) : undefined,
      ExpiryDate: item.ExpiryDate ? new Date(item.ExpiryDate) : undefined,
      ApplicableClassifications: item.ApplicableClassifications,
      ApplicableDepartments: item.ApplicableDepartments,
      LegalBasis: item.LegalBasis,
      Notes: item.Notes
    };
  }

  private mapToLegalHold(item: any): ILegalHold {
    return {
      Id: item.Id,
      Title: item.Title,
      HoldCode: item.HoldCode,
      HoldDescription: item.HoldDescription,
      MatterReference: item.MatterReference,
      CustodianIds: item.CustodiansId,
      HoldDepartments: item.HoldDepartments?.results,
      Keywords: item.Keywords,
      DateRangeStart: item.DateRangeStart ? new Date(item.DateRangeStart) : undefined,
      DateRangeEnd: item.DateRangeEnd ? new Date(item.DateRangeEnd) : undefined,
      HoldStatus: item.HoldStatus as LegalHoldStatus,
      IssuedDate: item.IssuedDate ? new Date(item.IssuedDate) : undefined,
      IssuedById: item.IssuedById,
      ReleasedDate: item.ReleasedDate ? new Date(item.ReleasedDate) : undefined,
      ReleasedById: item.ReleasedById,
      ReleaseReason: item.ReleaseReason,
      ExternalCounsel: item.ExternalCounsel,
      CaseNumber: item.CaseNumber,
      DocumentCount: item.DocumentCount || 0,
      Notes: item.Notes,
      NotifyOnNewDocuments: item.NotifyOnNewDocuments,
      NotificationRecipientIds: item.NotificationRecipientsId
    };
  }
}
