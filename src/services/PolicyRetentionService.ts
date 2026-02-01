// @ts-nocheck
/**
 * Policy Retention Service
 * Manages retention policies for policy documents and acknowledgement records
 * Ensures compliance with regulatory requirements
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';
import {
  IPolicy,
  IPolicyAcknowledgement,
  DataClassification,
  RetentionCategory,
  PolicyStatus,
  AcknowledgementStatus
} from '../models/IPolicy';
import { PolicyAuditService, AuditEventType, AuditSeverity } from './PolicyAuditService';
import { logger } from './LoggingService';
import { PolicyLists, RetentionLists, SystemLists } from '../constants/SharePointListNames';

// ============================================================================
// INTERFACES
// ============================================================================

/**
 * Retention policy configuration
 */
export interface IRetentionPolicy {
  Id?: number;
  Name: string;
  Description: string;

  // Scope
  AppliesTo: 'Policy' | 'Acknowledgement' | 'AuditLog' | 'All';
  DataClassifications?: DataClassification[];
  PolicyCategories?: string[];
  RegulatoryFrameworks?: string[];

  // Retention Period
  RetentionCategory: RetentionCategory;
  RetentionPeriodDays: number;
  RetentionStartEvent: 'Created' | 'Modified' | 'Published' | 'Archived' | 'Acknowledged';

  // Actions
  ActionOnExpiry: 'Delete' | 'Archive' | 'Review' | 'Notify';
  NotifyBeforeDays?: number;
  NotifyUserIds?: number[];
  NotifyEmails?: string[];

  // Exceptions
  ExcludeOnLegalHold: boolean;
  ExcludeComplianceRelevant: boolean;

  // Status
  IsActive: boolean;
  Priority: number; // Higher priority policies override lower

  // Audit
  CreatedById?: number;
  CreatedDate?: Date;
  ModifiedById?: number;
  ModifiedDate?: Date;
}

/**
 * Retention schedule entry
 */
export interface IRetentionScheduleEntry {
  EntityType: 'Policy' | 'Acknowledgement' | 'AuditLog';
  EntityId: number;
  EntityName?: string;
  PolicyId?: number;
  PolicyName?: string;

  // Retention details
  RetentionPolicyId: number;
  RetentionPolicyName: string;
  RetentionCategory: RetentionCategory;
  RetentionPeriodDays: number;

  // Dates
  CreatedDate: Date;
  RetentionStartDate: Date;
  RetentionExpiryDate: Date;
  DaysUntilExpiry: number;
  IsExpired: boolean;

  // Status
  IsOnLegalHold: boolean;
  LegalHoldReason?: string;
  ActionRequired: string;

  // Classification
  DataClassification?: DataClassification;
  RegulatoryFrameworks?: string[];
}

/**
 * Retention action result
 */
export interface IRetentionActionResult {
  totalProcessed: number;
  archived: number;
  deleted: number;
  reviewRequired: number;
  notificationsSent: number;
  skippedLegalHold: number;
  errors: string[];
}

/**
 * Legal hold request
 */
export interface ILegalHoldRequest {
  entityType: 'Policy' | 'Acknowledgement';
  entityIds: number[];
  reason: string;
  requestedBy: string;
  startDate: Date;
  endDate?: Date;
  caseReference?: string;
  notes?: string;
}

/**
 * Legal hold record
 */
export interface ILegalHold {
  Id?: number;
  EntityType: 'Policy' | 'Acknowledgement';
  EntityId: number;
  EntityName?: string;
  Reason: string;
  CaseReference?: string;
  RequestedById: number;
  RequestedByName?: string;
  StartDate: Date;
  EndDate?: Date;
  Status: 'Active' | 'Released' | 'Expired';
  ReleasedById?: number;
  ReleasedByName?: string;
  ReleasedDate?: Date;
  ReleaseReason?: string;
  Notes?: string;
}

// ============================================================================
// SERVICE CLASS
// ============================================================================

export class PolicyRetentionService {
  private sp: SPFI;
  private auditService: PolicyAuditService;
  private currentUserId: number = 0;
  private currentUserEmail: string = '';
  private currentUserName: string = '';
  private initialized: boolean = false;

  private readonly POLICIES_LIST = PolicyLists.POLICIES;
  private readonly ACKNOWLEDGEMENTS_LIST = PolicyLists.POLICY_ACKNOWLEDGEMENTS;
  private readonly RETENTION_POLICIES_LIST = RetentionLists.RETENTION_POLICIES;
  private readonly LEGAL_HOLDS_LIST = RetentionLists.LEGAL_HOLDS;
  private readonly RETENTION_ARCHIVE_LIST = RetentionLists.RETENTION_ARCHIVE;

  // Default retention periods in days
  private readonly DEFAULT_RETENTION_PERIODS: Record<RetentionCategory, number> = {
    [RetentionCategory.Standard]: 1095, // 3 years
    [RetentionCategory.Extended]: 2555, // 7 years
    [RetentionCategory.Regulatory]: 2555, // 7 years (default, can be overridden)
    [RetentionCategory.Legal]: -1, // Indefinite
    [RetentionCategory.Permanent]: -1 // Never delete
  };

  constructor(sp: SPFI) {
    this.sp = sp;
    this.auditService = new PolicyAuditService(sp);
  }

  // ============================================================================
  // INITIALIZATION
  // ============================================================================

  /**
   * Initialize the service
   */
  public async initialize(): Promise<void> {
    if (this.initialized) return;

    try {
      const user = await this.sp.web.currentUser();
      this.currentUserId = user.Id;
      this.currentUserEmail = user.Email;
      this.currentUserName = user.Title;
      await this.auditService.initialize();
      this.initialized = true;
    } catch (error) {
      logger.error('PolicyRetentionService', 'Failed to initialize:', error);
      throw error;
    }
  }

  // ============================================================================
  // HELPER METHODS
  // ============================================================================

  /**
   * Get retention period in days for a given category
   * @param category The retention category
   * @returns Number of days, or -1 for indefinite retention
   */
  public getRetentionPeriodDays(category: RetentionCategory): number {
    return this.DEFAULT_RETENTION_PERIODS[category] || this.DEFAULT_RETENTION_PERIODS[RetentionCategory.Standard];
  }

  /**
   * Calculate retention expiry date based on category
   * @param category The retention category
   * @param startDate The date to calculate from (defaults to now)
   * @returns Expiry date, or null for indefinite retention
   */
  public calculateRetentionExpiry(category: RetentionCategory, startDate?: Date): Date | null {
    const days = this.getRetentionPeriodDays(category);
    if (days === -1) return null; // Indefinite

    const expiry = new Date(startDate || new Date());
    expiry.setDate(expiry.getDate() + days);
    return expiry;
  }

  /**
   * Check if a retention category requires indefinite retention
   */
  public isIndefiniteRetention(category: RetentionCategory): boolean {
    return this.DEFAULT_RETENTION_PERIODS[category] === -1;
  }

  // ============================================================================
  // RETENTION POLICY MANAGEMENT
  // ============================================================================

  /**
   * Create a retention policy
   */
  public async createRetentionPolicy(policy: Partial<IRetentionPolicy>): Promise<IRetentionPolicy> {
    await this.initialize();

    try {
      const result = await this.sp.web.lists
        .getByTitle(this.RETENTION_POLICIES_LIST)
        .items.add({
          Title: policy.Name,
          Name: policy.Name,
          Description: policy.Description,
          AppliesTo: policy.AppliesTo,
          DataClassifications: policy.DataClassifications ? JSON.stringify(policy.DataClassifications) : undefined,
          PolicyCategories: policy.PolicyCategories ? JSON.stringify(policy.PolicyCategories) : undefined,
          RegulatoryFrameworks: policy.RegulatoryFrameworks ? JSON.stringify(policy.RegulatoryFrameworks) : undefined,
          RetentionCategory: policy.RetentionCategory,
          RetentionPeriodDays: policy.RetentionPeriodDays || this.DEFAULT_RETENTION_PERIODS[policy.RetentionCategory || RetentionCategory.Standard],
          RetentionStartEvent: policy.RetentionStartEvent || 'Created',
          ActionOnExpiry: policy.ActionOnExpiry || 'Review',
          NotifyBeforeDays: policy.NotifyBeforeDays,
          NotifyUserIds: policy.NotifyUserIds ? JSON.stringify(policy.NotifyUserIds) : undefined,
          NotifyEmails: policy.NotifyEmails ? JSON.stringify(policy.NotifyEmails) : undefined,
          ExcludeOnLegalHold: policy.ExcludeOnLegalHold ?? true,
          ExcludeComplianceRelevant: policy.ExcludeComplianceRelevant ?? false,
          IsActive: policy.IsActive ?? true,
          Priority: policy.Priority || 10,
          CreatedById: this.currentUserId,
          CreatedDate: new Date().toISOString()
        });

      await this.auditService.logEvent({
        EventType: AuditEventType.SettingsChanged,
        Severity: AuditSeverity.Info,
        EntityType: 'System',
        EntityId: result.data.Id,
        ActionDescription: `Retention policy "${policy.Name}" created`,
        ComplianceRelevant: true
      });

      return this.mapRetentionPolicy(result.data);

    } catch (error) {
      logger.error('PolicyRetentionService', 'Failed to create retention policy:', error);
      throw error;
    }
  }

  /**
   * Get all retention policies
   */
  public async getRetentionPolicies(): Promise<IRetentionPolicy[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.RETENTION_POLICIES_LIST)
        .items.filter('IsActive eq true')
        .orderBy('Priority', false)
        .top(100)();

      return items.map(item => this.mapRetentionPolicy(item));

    } catch (error) {
      logger.error('PolicyRetentionService', 'Failed to get retention policies:', error);
      return [];
    }
  }

  /**
   * Get applicable retention policy for an entity
   */
  public async getApplicableRetentionPolicy(
    entityType: 'Policy' | 'Acknowledgement',
    entity: IPolicy | IPolicyAcknowledgement
  ): Promise<IRetentionPolicy | null> {
    const policies = await this.getRetentionPolicies();

    // Sort by priority (highest first) and find first match
    for (const policy of policies) {
      if (policy.AppliesTo !== entityType && policy.AppliesTo !== 'All') continue;

      // Check data classification match
      if (policy.DataClassifications?.length && entityType === 'Policy') {
        const policyEntity = entity as IPolicy;
        if (!policyEntity.DataClassification ||
            !policy.DataClassifications.includes(policyEntity.DataClassification)) {
          continue;
        }
      }

      // Check policy category match
      if (policy.PolicyCategories?.length && entityType === 'Policy') {
        const policyEntity = entity as IPolicy;
        if (!policy.PolicyCategories.includes(policyEntity.PolicyCategory)) {
          continue;
        }
      }

      // Check regulatory framework match
      if (policy.RegulatoryFrameworks?.length && entityType === 'Policy') {
        const policyEntity = entity as IPolicy;
        const entityFrameworks = policyEntity.RegulatoryFrameworks || [];
        if (!policy.RegulatoryFrameworks.some(f => entityFrameworks.includes(f))) {
          continue;
        }
      }

      return policy;
    }

    return null;
  }

  // ============================================================================
  // RETENTION SCHEDULE
  // ============================================================================

  /**
   * Generate retention schedule for all policies
   */
  public async generateRetentionSchedule(): Promise<IRetentionScheduleEntry[]> {
    await this.initialize();

    const schedule: IRetentionScheduleEntry[] = [];
    const now = new Date();

    try {
      // Get all policies
      const policies = await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.select('*')
        .top(500)() as IPolicy[];

      // Get all legal holds
      const legalHolds = await this.getActiveLegalHolds();
      const policyHolds = new Map(legalHolds.filter(h => h.EntityType === 'Policy').map(h => [h.EntityId, h]));

      for (const policy of policies) {
        const retentionPolicy = await this.getApplicableRetentionPolicy('Policy', policy);
        if (!retentionPolicy) continue;

        const startDate = this.getRetentionStartDate(policy, retentionPolicy.RetentionStartEvent);
        const expiryDate = this.calculateExpiryDate(startDate, retentionPolicy.RetentionPeriodDays);
        const daysUntilExpiry = expiryDate ? Math.ceil((expiryDate.getTime() - now.getTime()) / (1000 * 60 * 60 * 24)) : -1;
        const isOnLegalHold = policyHolds.has(policy.Id!);
        const legalHold = policyHolds.get(policy.Id!);

        schedule.push({
          EntityType: 'Policy',
          EntityId: policy.Id!,
          EntityName: policy.PolicyName,
          PolicyId: policy.Id,
          PolicyName: policy.PolicyName,
          RetentionPolicyId: retentionPolicy.Id!,
          RetentionPolicyName: retentionPolicy.Name,
          RetentionCategory: retentionPolicy.RetentionCategory,
          RetentionPeriodDays: retentionPolicy.RetentionPeriodDays,
          CreatedDate: new Date(policy.Created!),
          RetentionStartDate: startDate,
          RetentionExpiryDate: expiryDate!,
          DaysUntilExpiry: daysUntilExpiry,
          IsExpired: daysUntilExpiry !== -1 && daysUntilExpiry <= 0,
          IsOnLegalHold: isOnLegalHold,
          LegalHoldReason: legalHold?.Reason,
          ActionRequired: this.determineAction(daysUntilExpiry, isOnLegalHold, retentionPolicy),
          DataClassification: policy.DataClassification,
          RegulatoryFrameworks: policy.RegulatoryFrameworks
        });
      }

      // Sort by days until expiry (most urgent first)
      schedule.sort((a, b) => {
        if (a.DaysUntilExpiry === -1) return 1;
        if (b.DaysUntilExpiry === -1) return -1;
        return a.DaysUntilExpiry - b.DaysUntilExpiry;
      });

      return schedule;

    } catch (error) {
      logger.error('PolicyRetentionService', 'Failed to generate retention schedule:', error);
      return [];
    }
  }

  /**
   * Get items expiring within specified days
   */
  public async getExpiringItems(withinDays: number): Promise<IRetentionScheduleEntry[]> {
    const schedule = await this.generateRetentionSchedule();
    return schedule.filter(entry =>
      entry.DaysUntilExpiry !== -1 &&
      entry.DaysUntilExpiry >= 0 &&
      entry.DaysUntilExpiry <= withinDays &&
      !entry.IsOnLegalHold
    );
  }

  /**
   * Get expired items
   */
  public async getExpiredItems(): Promise<IRetentionScheduleEntry[]> {
    const schedule = await this.generateRetentionSchedule();
    return schedule.filter(entry =>
      entry.IsExpired &&
      !entry.IsOnLegalHold
    );
  }

  // ============================================================================
  // RETENTION ACTIONS
  // ============================================================================

  /**
   * Process retention actions for expired items
   */
  public async processRetentionActions(
    dryRun: boolean = true
  ): Promise<IRetentionActionResult> {
    await this.initialize();

    const result: IRetentionActionResult = {
      totalProcessed: 0,
      archived: 0,
      deleted: 0,
      reviewRequired: 0,
      notificationsSent: 0,
      skippedLegalHold: 0,
      errors: []
    };

    try {
      const expiredItems = await this.getExpiredItems();
      result.totalProcessed = expiredItems.length;

      for (const item of expiredItems) {
        try {
          // Skip legal holds
          if (item.IsOnLegalHold) {
            result.skippedLegalHold++;
            continue;
          }

          const retentionPolicy = (await this.getRetentionPolicies())
            .find(p => p.Id === item.RetentionPolicyId);

          if (!retentionPolicy) continue;

          switch (retentionPolicy.ActionOnExpiry) {
            case 'Archive':
              if (!dryRun) {
                await this.archiveItem(item);
              }
              result.archived++;
              break;

            case 'Delete':
              if (!dryRun) {
                await this.deleteItem(item);
              }
              result.deleted++;
              break;

            case 'Review':
              result.reviewRequired++;
              break;

            case 'Notify':
              if (!dryRun && retentionPolicy.NotifyEmails?.length) {
                // Notification would be sent here
                result.notificationsSent++;
              }
              break;
          }

        } catch (error) {
          result.errors.push(`Failed to process ${item.EntityType} ${item.EntityId}: ${error}`);
        }
      }

      // Log the retention action
      await this.auditService.logEvent({
        EventType: AuditEventType.RetentionApplied,
        Severity: AuditSeverity.Info,
        EntityType: 'System',
        EntityId: 0,
        ActionDescription: `Retention processing ${dryRun ? '(dry run)' : ''}: ${result.archived} archived, ${result.deleted} deleted, ${result.reviewRequired} for review`,
        ComplianceRelevant: true,
        Metadata: JSON.stringify(result)
      });

      return result;

    } catch (error) {
      logger.error('PolicyRetentionService', 'Failed to process retention actions:', error);
      throw error;
    }
  }

  /**
   * Archive an item
   */
  private async archiveItem(entry: IRetentionScheduleEntry): Promise<void> {
    try {
      // Get the item data
      let itemData: any;
      if (entry.EntityType === 'Policy') {
        itemData = await this.sp.web.lists
          .getByTitle(this.POLICIES_LIST)
          .items.getById(entry.EntityId)();
      } else {
        itemData = await this.sp.web.lists
          .getByTitle(this.ACKNOWLEDGEMENTS_LIST)
          .items.getById(entry.EntityId)();
      }

      // Copy to archive
      await this.sp.web.lists
        .getByTitle(this.RETENTION_ARCHIVE_LIST)
        .items.add({
          Title: itemData.Title,
          OriginalEntityType: entry.EntityType,
          OriginalEntityId: entry.EntityId,
          OriginalData: JSON.stringify(itemData),
          ArchivedDate: new Date().toISOString(),
          ArchivedById: this.currentUserId,
          RetentionPolicyId: entry.RetentionPolicyId,
          RetentionPolicyName: entry.RetentionPolicyName
        });

      // Update original to archived status (don't delete)
      if (entry.EntityType === 'Policy') {
        await this.sp.web.lists
          .getByTitle(this.POLICIES_LIST)
          .items.getById(entry.EntityId)
          .update({
            Status: PolicyStatus.Archived,
            ArchivedDate: new Date().toISOString(),
            IsActive: false
          });
      }

      await this.auditService.logEvent({
        EventType: AuditEventType.PolicyArchived,
        Severity: AuditSeverity.Info,
        EntityType: entry.EntityType,
        EntityId: entry.EntityId,
        PolicyId: entry.PolicyId,
        PolicyName: entry.PolicyName,
        ActionDescription: `${entry.EntityType} archived due to retention policy "${entry.RetentionPolicyName}"`,
        ComplianceRelevant: true
      });

    } catch (error) {
      logger.error('PolicyRetentionService', `Failed to archive ${entry.EntityType} ${entry.EntityId}:`, error);
      throw error;
    }
  }

  /**
   * Delete an item (after archiving metadata)
   */
  private async deleteItem(entry: IRetentionScheduleEntry): Promise<void> {
    try {
      // First archive the metadata
      await this.archiveItem(entry);

      // Then delete if it's an acknowledgement (policies should only be archived)
      if (entry.EntityType === 'Acknowledgement') {
        await this.sp.web.lists
          .getByTitle(this.ACKNOWLEDGEMENTS_LIST)
          .items.getById(entry.EntityId)
          .delete();

        await this.auditService.logEvent({
          EventType: AuditEventType.DataPurged,
          Severity: AuditSeverity.Warning,
          EntityType: 'Acknowledgement',
          EntityId: entry.EntityId,
          PolicyId: entry.PolicyId,
          PolicyName: entry.PolicyName,
          ActionDescription: `Acknowledgement record deleted due to retention policy "${entry.RetentionPolicyName}"`,
          ComplianceRelevant: true
        });
      }

    } catch (error) {
      logger.error('PolicyRetentionService', `Failed to delete ${entry.EntityType} ${entry.EntityId}:`, error);
      throw error;
    }
  }

  // ============================================================================
  // LEGAL HOLD MANAGEMENT
  // ============================================================================

  /**
   * Place items on legal hold
   */
  public async placeLegalHold(request: ILegalHoldRequest): Promise<ILegalHold[]> {
    await this.initialize();

    const holds: ILegalHold[] = [];

    try {
      for (const entityId of request.entityIds) {
        // Check if already on hold
        const existingHold = await this.sp.web.lists
          .getByTitle(this.LEGAL_HOLDS_LIST)
          .items.filter(`EntityType eq '${request.entityType}' and EntityId eq ${entityId} and Status eq 'Active'`)
          .top(1)();

        if (existingHold.length > 0) {
          logger.warn('PolicyRetentionService', `${request.entityType} ${entityId} is already on legal hold`);
          continue;
        }

        // Get entity name
        let entityName = '';
        if (request.entityType === 'Policy') {
          const policy = await this.sp.web.lists
            .getByTitle(this.POLICIES_LIST)
            .items.getById(entityId)
            .select('PolicyName')() as IPolicy;
          entityName = policy.PolicyName;

          // Update policy's legal hold flags
          await this.sp.web.lists
            .getByTitle(this.POLICIES_LIST)
            .items.getById(entityId)
            .update({
              IsLegalHold: true,
              LegalHoldReason: request.reason,
              LegalHoldStartDate: request.startDate.toISOString(),
              LegalHoldEndDate: request.endDate?.toISOString()
            });
        }

        // Create legal hold record
        const result = await this.sp.web.lists
          .getByTitle(this.LEGAL_HOLDS_LIST)
          .items.add({
            Title: `${request.entityType} - ${entityId}`,
            EntityType: request.entityType,
            EntityId: entityId,
            EntityName: entityName,
            Reason: request.reason,
            CaseReference: request.caseReference,
            RequestedById: this.currentUserId,
            RequestedByName: this.currentUserName,
            StartDate: request.startDate.toISOString(),
            EndDate: request.endDate?.toISOString(),
            Status: 'Active',
            Notes: request.notes
          });

        holds.push({
          Id: result.data.Id,
          EntityType: request.entityType,
          EntityId: entityId,
          EntityName: entityName,
          Reason: request.reason,
          CaseReference: request.caseReference,
          RequestedById: this.currentUserId,
          RequestedByName: this.currentUserName,
          StartDate: request.startDate,
          EndDate: request.endDate,
          Status: 'Active',
          Notes: request.notes
        });

        await this.auditService.logSecurityEvent(
          AuditEventType.SettingsChanged,
          `Legal hold placed on ${request.entityType} ${entityId}: ${request.reason}`,
          entityId,
          { caseReference: request.caseReference }
        );
      }

      return holds;

    } catch (error) {
      logger.error('PolicyRetentionService', 'Failed to place legal hold:', error);
      throw error;
    }
  }

  /**
   * Release legal hold
   */
  public async releaseLegalHold(
    holdId: number,
    releaseReason: string
  ): Promise<void> {
    await this.initialize();

    try {
      const hold = await this.sp.web.lists
        .getByTitle(this.LEGAL_HOLDS_LIST)
        .items.getById(holdId)() as ILegalHold;

      // Update hold record
      await this.sp.web.lists
        .getByTitle(this.LEGAL_HOLDS_LIST)
        .items.getById(holdId)
        .update({
          Status: 'Released',
          ReleasedById: this.currentUserId,
          ReleasedByName: this.currentUserName,
          ReleasedDate: new Date().toISOString(),
          ReleaseReason: releaseReason
        });

      // Update entity
      if (hold.EntityType === 'Policy') {
        await this.sp.web.lists
          .getByTitle(this.POLICIES_LIST)
          .items.getById(hold.EntityId)
          .update({
            IsLegalHold: false,
            LegalHoldEndDate: new Date().toISOString()
          });
      }

      await this.auditService.logEvent({
        EventType: AuditEventType.SettingsChanged,
        Severity: AuditSeverity.Info,
        EntityType: hold.EntityType,
        EntityId: hold.EntityId,
        ActionDescription: `Legal hold released: ${releaseReason}`,
        ComplianceRelevant: true
      });

    } catch (error) {
      logger.error('PolicyRetentionService', 'Failed to release legal hold:', error);
      throw error;
    }
  }

  /**
   * Get active legal holds
   */
  public async getActiveLegalHolds(): Promise<ILegalHold[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.LEGAL_HOLDS_LIST)
        .items.filter("Status eq 'Active'")
        .orderBy('StartDate', false)
        .top(1000)();

      return items.map(item => ({
        Id: item.Id,
        EntityType: item.EntityType,
        EntityId: item.EntityId,
        EntityName: item.EntityName,
        Reason: item.Reason,
        CaseReference: item.CaseReference,
        RequestedById: item.RequestedById,
        RequestedByName: item.RequestedByName,
        StartDate: new Date(item.StartDate),
        EndDate: item.EndDate ? new Date(item.EndDate) : undefined,
        Status: item.Status,
        Notes: item.Notes
      }));

    } catch (error) {
      logger.error('PolicyRetentionService', 'Failed to get active legal holds:', error);
      return [];
    }
  }

  /**
   * Check if entity is on legal hold
   */
  public async isOnLegalHold(
    entityType: 'Policy' | 'Acknowledgement',
    entityId: number
  ): Promise<boolean> {
    try {
      const holds = await this.sp.web.lists
        .getByTitle(this.LEGAL_HOLDS_LIST)
        .items.filter(`EntityType eq '${entityType}' and EntityId eq ${entityId} and Status eq 'Active'`)
        .top(1)();

      return holds.length > 0;

    } catch (error) {
      logger.error('PolicyRetentionService', 'Failed to check legal hold:', error);
      return false;
    }
  }

  // ============================================================================
  // DATA CLASSIFICATION
  // ============================================================================

  /**
   * Apply data classification to a policy
   */
  public async applyDataClassification(
    policyId: number,
    classification: DataClassification,
    justification: string,
    regulatoryFrameworks?: string[]
  ): Promise<void> {
    await this.initialize();

    try {
      // Determine retention based on classification
      const retentionCategory = this.getRetentionForClassification(classification);

      await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.getById(policyId)
        .update({
          DataClassification: classification,
          ClassificationJustification: justification,
          ClassifiedById: this.currentUserId,
          ClassifiedDate: new Date().toISOString(),
          ClassificationReviewDate: this.calculateClassificationReviewDate(classification).toISOString(),
          RetentionCategory: retentionCategory,
          RegulatoryFrameworks: regulatoryFrameworks ? JSON.stringify(regulatoryFrameworks) : undefined
        });

      await this.auditService.logEvent({
        EventType: AuditEventType.PolicyUpdated,
        Severity: classification === DataClassification.Restricted ? AuditSeverity.Warning : AuditSeverity.Info,
        EntityType: 'Policy',
        EntityId: policyId,
        ActionDescription: `Data classification set to "${classification}"`,
        ComplianceRelevant: true,
        Metadata: JSON.stringify({
          classification,
          justification,
          regulatoryFrameworks
        })
      });

    } catch (error) {
      logger.error('PolicyRetentionService', 'Failed to apply data classification:', error);
      throw error;
    }
  }

  /**
   * Get suggested classification based on policy content
   */
  public suggestClassification(policy: Partial<IPolicy>): {
    suggestedClassification: DataClassification;
    reasons: string[];
    handlingInstructions: string[];
  } {
    const reasons: string[] = [];
    let classification = DataClassification.Internal;

    // Check for PII indicators
    if (policy.ContainsPII) {
      classification = DataClassification.Confidential;
      reasons.push('Contains personally identifiable information (PII)');
    }

    // Check for PHI
    if (policy.ContainsPHI) {
      classification = DataClassification.Regulated;
      reasons.push('Contains protected health information (PHI)');
    }

    // Check for financial data
    if (policy.ContainsFinancialData) {
      classification = DataClassification.Confidential;
      reasons.push('Contains financial data');
    }

    // Check compliance risk
    if (policy.ComplianceRisk === 'Critical' || policy.ComplianceRisk === 'High') {
      if (classification !== DataClassification.Regulated) {
        classification = DataClassification.Confidential;
      }
      reasons.push(`High compliance risk level (${policy.ComplianceRisk})`);
    }

    // Check regulatory frameworks
    if (policy.RegulatoryFrameworks?.length) {
      const highRegFrameworks = ['HIPAA', 'PCI-DSS', 'SOX'];
      if (policy.RegulatoryFrameworks.some(f => highRegFrameworks.includes(f))) {
        classification = DataClassification.Regulated;
        reasons.push(`Subject to regulatory framework: ${policy.RegulatoryFrameworks.join(', ')}`);
      }
    }

    // Check for restricted distribution
    if (policy.TargetUserIds?.length || policy.TargetRoles?.length === 1) {
      if (classification === DataClassification.Internal) {
        classification = DataClassification.Confidential;
      }
      reasons.push('Has restricted distribution targeting');
    }

    // If no specific indicators, default to Internal
    if (reasons.length === 0) {
      reasons.push('No sensitive data indicators detected');
    }

    // Determine handling instructions
    const handlingInstructions = this.getHandlingInstructionsForClassification(classification);

    return {
      suggestedClassification: classification,
      reasons,
      handlingInstructions
    };
  }

  // ============================================================================
  // UTILITY METHODS
  // ============================================================================

  /**
   * Get retention start date based on event type
   */
  private getRetentionStartDate(
    policy: IPolicy,
    startEvent: string
  ): Date {
    switch (startEvent) {
      case 'Published':
        return policy.PublishedDate ? new Date(policy.PublishedDate) : new Date(policy.Created!);
      case 'Modified':
        return policy.Modified ? new Date(policy.Modified) : new Date(policy.Created!);
      case 'Archived':
        return policy.ArchivedDate ? new Date(policy.ArchivedDate) : new Date();
      case 'Created':
      default:
        return new Date(policy.Created!);
    }
  }

  /**
   * Calculate expiry date
   */
  private calculateExpiryDate(startDate: Date, periodDays: number): Date | null {
    if (periodDays === -1) return null; // Indefinite

    const expiry = new Date(startDate);
    expiry.setDate(expiry.getDate() + periodDays);
    return expiry;
  }

  /**
   * Determine action based on expiry and status
   */
  private determineAction(
    daysUntilExpiry: number,
    isOnLegalHold: boolean,
    retentionPolicy: IRetentionPolicy
  ): string {
    if (isOnLegalHold) return 'On Legal Hold - No Action';
    if (daysUntilExpiry === -1) return 'Permanent Retention';
    if (daysUntilExpiry <= 0) return retentionPolicy.ActionOnExpiry;
    if (retentionPolicy.NotifyBeforeDays && daysUntilExpiry <= retentionPolicy.NotifyBeforeDays) {
      return 'Notify - Approaching Expiry';
    }
    return 'No Action Required';
  }

  /**
   * Get retention category for classification
   */
  private getRetentionForClassification(classification: DataClassification): RetentionCategory {
    switch (classification) {
      case DataClassification.Regulated:
        return RetentionCategory.Regulatory;
      case DataClassification.Restricted:
        return RetentionCategory.Extended;
      case DataClassification.Confidential:
        return RetentionCategory.Extended;
      case DataClassification.Internal:
        return RetentionCategory.Standard;
      case DataClassification.Public:
      default:
        return RetentionCategory.Standard;
    }
  }

  /**
   * Calculate classification review date
   */
  private calculateClassificationReviewDate(classification: DataClassification): Date {
    const reviewDate = new Date();
    switch (classification) {
      case DataClassification.Restricted:
      case DataClassification.Regulated:
        reviewDate.setMonth(reviewDate.getMonth() + 6); // 6 months
        break;
      case DataClassification.Confidential:
        reviewDate.setFullYear(reviewDate.getFullYear() + 1); // 1 year
        break;
      default:
        reviewDate.setFullYear(reviewDate.getFullYear() + 2); // 2 years
    }
    return reviewDate;
  }

  /**
   * Get handling instructions for classification
   */
  private getHandlingInstructionsForClassification(classification: DataClassification): string[] {
    switch (classification) {
      case DataClassification.Restricted:
        return [
          'Encrypt at Rest',
          'Encrypt in Transit',
          'No External Sharing',
          'Audit All Access',
          'Approval Required for Access'
        ];
      case DataClassification.Regulated:
        return [
          'Encrypt at Rest',
          'Encrypt in Transit',
          'Audit All Access',
          'Regulatory Compliance Required'
        ];
      case DataClassification.Confidential:
        return [
          'No External Sharing',
          'Watermark on Print/Download'
        ];
      case DataClassification.Internal:
        return ['Internal Use Only'];
      case DataClassification.Public:
      default:
        return ['No Restrictions'];
    }
  }

  /**
   * Map retention policy from SharePoint item
   */
  private mapRetentionPolicy(item: any): IRetentionPolicy {
    return {
      Id: item.Id,
      Name: item.Name || item.Title,
      Description: item.Description,
      AppliesTo: item.AppliesTo,
      DataClassifications: item.DataClassifications ? JSON.parse(item.DataClassifications) : undefined,
      PolicyCategories: item.PolicyCategories ? JSON.parse(item.PolicyCategories) : undefined,
      RegulatoryFrameworks: item.RegulatoryFrameworks ? JSON.parse(item.RegulatoryFrameworks) : undefined,
      RetentionCategory: item.RetentionCategory,
      RetentionPeriodDays: item.RetentionPeriodDays,
      RetentionStartEvent: item.RetentionStartEvent,
      ActionOnExpiry: item.ActionOnExpiry,
      NotifyBeforeDays: item.NotifyBeforeDays,
      NotifyUserIds: item.NotifyUserIds ? JSON.parse(item.NotifyUserIds) : undefined,
      NotifyEmails: item.NotifyEmails ? JSON.parse(item.NotifyEmails) : undefined,
      ExcludeOnLegalHold: item.ExcludeOnLegalHold,
      ExcludeComplianceRelevant: item.ExcludeComplianceRelevant,
      IsActive: item.IsActive,
      Priority: item.Priority,
      CreatedById: item.CreatedById,
      CreatedDate: item.CreatedDate ? new Date(item.CreatedDate) : undefined,
      ModifiedById: item.ModifiedById,
      ModifiedDate: item.ModifiedDate ? new Date(item.ModifiedDate) : undefined
    };
  }
}
