// @ts-nocheck
/**
 * RetentionService
 *
 * Manages retention policies, legal holds, and policy archival.
 *
 * This service powers:
 *   - Admin Centre > Legal Holds section
 *   - PolicyDetails legal hold banner
 *   - PolicyAuthorView action button disabling for held policies
 *   - Retention policy management and archival workflows
 *
 * SharePoint Lists:
 *   - PM_RetentionPolicies: Retention rule definitions per entity type
 *   - PM_LegalHolds: Active/released legal hold records
 *   - PM_RetentionArchive: Archived policy data snapshots
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { logger } from './LoggingService';

// ─── Interfaces ────────────────────────────────────────────────

export interface IRetentionPolicy {
  Id?: number;
  Title: string;
  EntityType: 'Policy' | 'Document' | 'AuditLog' | 'Quiz';
  RetentionDays: number;
  RetentionAction: 'Archive' | 'Delete' | 'Review';
  IsActive: boolean;
  Description?: string;
  CreatedByUser?: string;
  CreatedDate?: string;
}

export interface ILegalHold {
  Id?: number;
  Title: string;
  PolicyId: number;
  PolicyTitle: string;
  HoldReason: string;
  PlacedBy: string;
  PlacedByEmail: string;
  PlacedDate: string;
  ExpiryDate?: string;
  IsActive: boolean;
  Status: 'Active' | 'Released' | 'Expired';
  ReleasedBy?: string;
  ReleasedDate?: string;
  ReleaseReason?: string;
  CaseReference?: string;
  ComplianceRelevant: boolean;
}

export interface IRetentionArchiveItem {
  Id?: number;
  Title: string;
  OriginalPolicyId: number;
  PolicyTitle: string;
  PolicyNumber: string;
  PolicyCategory: string;
  ArchivedDate: string;
  ArchivedBy: string;
  RetentionRuleId?: number;
  OriginalContent?: string;
  OriginalMetadata?: string;
  ArchiveReason?: string;
}

// ─── List Names ────────────────────────────────────────────────

const LIST_RETENTION_POLICIES = 'PM_RetentionPolicies';
const LIST_LEGAL_HOLDS = 'PM_LegalHolds';
const LIST_RETENTION_ARCHIVE = 'PM_RetentionArchive';
const LIST_POLICIES = 'PM_Policies';

// ─── Service ───────────────────────────────────────────────────

export class RetentionService {
  private sp: SPFI;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ═══════════════════════════════════════════════════════════════
  // RETENTION POLICIES
  // ═══════════════════════════════════════════════════════════════

  /**
   * Load all active retention rules from PM_RetentionPolicies.
   */
  public async getRetentionPolicies(): Promise<IRetentionPolicy[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(LIST_RETENTION_POLICIES)
        .items.select(
          'Id', 'Title', 'EntityType', 'RetentionDays', 'RetentionAction',
          'IsActive', 'Description', 'CreatedByUser', 'CreatedDate'
        )
        .filter("IsActive eq 1")
        .orderBy('Title')
        .top(200)();

      return items.map((item: any) => ({
        Id: item.Id,
        Title: item.Title || '',
        EntityType: item.EntityType || 'Policy',
        RetentionDays: item.RetentionDays || 0,
        RetentionAction: item.RetentionAction || 'Review',
        IsActive: item.IsActive !== false,
        Description: item.Description || '',
        CreatedByUser: item.CreatedByUser || '',
        CreatedDate: item.CreatedDate || ''
      }));
    } catch (err) {
      logger.warn('RetentionService', 'Failed to load retention policies:', err);
      return [];
    }
  }

  /**
   * Create a new retention rule.
   */
  public async createRetentionPolicy(rule: Partial<IRetentionPolicy>): Promise<number> {
    try {
      const result = await this.sp.web.lists
        .getByTitle(LIST_RETENTION_POLICIES)
        .items.add({
          Title: rule.Title || `${rule.EntityType} Retention`,
          EntityType: rule.EntityType,
          RetentionDays: rule.RetentionDays,
          RetentionAction: rule.RetentionAction,
          IsActive: rule.IsActive !== false,
          Description: rule.Description || '',
          CreatedByUser: rule.CreatedByUser || '',
          CreatedDate: new Date().toISOString()
        });
      const newId = result?.data?.Id || result?.data?.id || 0;
      logger.info('RetentionService', `Created retention policy: ${rule.Title} (ID: ${newId})`);
      return newId;
    } catch (err) {
      logger.error('RetentionService', 'Failed to create retention policy:', err);
      throw err;
    }
  }

  /**
   * Update an existing retention rule.
   */
  public async updateRetentionPolicy(id: number, rule: Partial<IRetentionPolicy>): Promise<void> {
    try {
      const updates: Record<string, any> = {};
      if (rule.Title !== undefined) updates.Title = rule.Title;
      if (rule.EntityType !== undefined) updates.EntityType = rule.EntityType;
      if (rule.RetentionDays !== undefined) updates.RetentionDays = rule.RetentionDays;
      if (rule.RetentionAction !== undefined) updates.RetentionAction = rule.RetentionAction;
      if (rule.IsActive !== undefined) updates.IsActive = rule.IsActive;
      if (rule.Description !== undefined) updates.Description = rule.Description;

      await this.sp.web.lists
        .getByTitle(LIST_RETENTION_POLICIES)
        .items.getById(id)
        .update(updates);

      logger.info('RetentionService', `Updated retention policy ID: ${id}`);
    } catch (err) {
      logger.error('RetentionService', 'Failed to update retention policy:', err);
      throw err;
    }
  }

  /**
   * Soft-delete a retention rule (IsActive=false).
   */
  public async deleteRetentionPolicy(id: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(LIST_RETENTION_POLICIES)
        .items.getById(id)
        .update({ IsActive: false });

      logger.info('RetentionService', `Soft-deleted retention policy ID: ${id}`);
    } catch (err) {
      logger.error('RetentionService', 'Failed to delete retention policy:', err);
      throw err;
    }
  }

  // ═══════════════════════════════════════════════════════════════
  // LEGAL HOLDS
  // ═══════════════════════════════════════════════════════════════

  /**
   * Load all legal holds (active, released, expired).
   */
  public async getLegalHolds(): Promise<ILegalHold[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(LIST_LEGAL_HOLDS)
        .items.select(
          'Id', 'Title', 'PolicyId', 'PolicyTitle', 'HoldReason',
          'PlacedBy', 'PlacedByEmail', 'PlacedDate', 'ExpiryDate',
          'IsActive', 'Status', 'ReleasedBy', 'ReleasedDate',
          'ReleaseReason', 'CaseReference', 'ComplianceRelevant'
        )
        .orderBy('PlacedDate', false)
        .top(500)();

      return items.map((item: any) => ({
        Id: item.Id,
        Title: item.Title || '',
        PolicyId: item.PolicyId || 0,
        PolicyTitle: item.PolicyTitle || '',
        HoldReason: item.HoldReason || '',
        PlacedBy: item.PlacedBy || '',
        PlacedByEmail: item.PlacedByEmail || '',
        PlacedDate: item.PlacedDate || '',
        ExpiryDate: item.ExpiryDate || '',
        IsActive: item.IsActive !== false,
        Status: item.Status || 'Active',
        ReleasedBy: item.ReleasedBy || '',
        ReleasedDate: item.ReleasedDate || '',
        ReleaseReason: item.ReleaseReason || '',
        CaseReference: item.CaseReference || '',
        ComplianceRelevant: item.ComplianceRelevant !== false
      }));
    } catch (err) {
      logger.warn('RetentionService', 'Failed to load legal holds:', err);
      return [];
    }
  }

  /**
   * Place a new legal hold on a policy.
   */
  public async placeLegalHold(
    policyId: number,
    reason: string,
    caseRef: string,
    placedBy: string,
    placedByEmail?: string,
    expiryDate?: string,
    policyTitle?: string
  ): Promise<number> {
    try {
      // If policyTitle not provided, look it up
      let title = policyTitle || '';
      if (!title) {
        try {
          const policy = await this.sp.web.lists
            .getByTitle(LIST_POLICIES)
            .items.getById(policyId)
            .select('Title', 'PolicyName')();
          title = policy.PolicyName || policy.Title || `Policy #${policyId}`;
        } catch { title = `Policy #${policyId}`; }
      }

      const holdTitle = `LH-${Date.now().toString(36).toUpperCase()}`;

      const result = await this.sp.web.lists
        .getByTitle(LIST_LEGAL_HOLDS)
        .items.add({
          Title: holdTitle,
          PolicyId: policyId,
          PolicyTitle: title,
          HoldReason: reason,
          PlacedBy: placedBy,
          PlacedByEmail: placedByEmail || '',
          PlacedDate: new Date().toISOString(),
          ExpiryDate: expiryDate || null,
          IsActive: true,
          Status: 'Active',
          CaseReference: caseRef || '',
          ComplianceRelevant: true
        });

      const newId = result?.data?.Id || result?.data?.id || 0;
      logger.info('RetentionService', `Legal hold placed on policy ${policyId}: ${holdTitle} (ID: ${newId})`);
      return newId;
    } catch (err) {
      logger.error('RetentionService', 'Failed to place legal hold:', err);
      throw err;
    }
  }

  /**
   * Release a legal hold.
   */
  public async releaseLegalHold(holdId: number, releasedBy: string, reason: string): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(LIST_LEGAL_HOLDS)
        .items.getById(holdId)
        .update({
          IsActive: false,
          Status: 'Released',
          ReleasedBy: releasedBy,
          ReleasedDate: new Date().toISOString(),
          ReleaseReason: reason
        });

      logger.info('RetentionService', `Legal hold ${holdId} released by ${releasedBy}`);
    } catch (err) {
      logger.error('RetentionService', 'Failed to release legal hold:', err);
      throw err;
    }
  }

  /**
   * Check if a specific policy has an active legal hold.
   */
  public async isPolicyOnHold(policyId: number): Promise<boolean> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(LIST_LEGAL_HOLDS)
        .items.select('Id')
        .filter(`PolicyId eq ${policyId} and IsActive eq 1 and Status eq 'Active'`)
        .top(1)();

      return items.length > 0;
    } catch (err) {
      logger.warn('RetentionService', `Failed to check legal hold for policy ${policyId}:`, err);
      return false;
    }
  }

  // ═══════════════════════════════════════════════════════════════
  // ARCHIVAL
  // ═══════════════════════════════════════════════════════════════

  /**
   * Archive a policy: snapshot data to PM_RetentionArchive, then set status to Archived.
   */
  public async archivePolicy(policyId: number, retentionRuleId: number, archivedBy: string): Promise<void> {
    try {
      // 1. Load the policy data
      const policy = await this.sp.web.lists
        .getByTitle(LIST_POLICIES)
        .items.getById(policyId)
        .select(
          'Id', 'Title', 'PolicyName', 'PolicyNumber', 'PolicyCategory',
          'PolicyContent', 'HTMLContent', 'Description', 'PolicyStatus',
          'ComplianceRisk', 'Departments', 'VersionNumber'
        )();

      // 2. Check for active legal hold — block archival
      const isHeld = await this.isPolicyOnHold(policyId);
      if (isHeld) {
        throw new Error('Cannot archive policy: active legal hold exists');
      }

      // 3. Build metadata snapshot
      const metadata = JSON.stringify({
        PolicyStatus: policy.PolicyStatus,
        ComplianceRisk: policy.ComplianceRisk,
        Departments: policy.Departments,
        VersionNumber: policy.VersionNumber
      });

      // 4. Create archive record
      await this.sp.web.lists
        .getByTitle(LIST_RETENTION_ARCHIVE)
        .items.add({
          Title: `ARCH-${policy.PolicyNumber || policyId}`,
          OriginalPolicyId: policyId,
          PolicyTitle: policy.PolicyName || policy.Title || '',
          PolicyNumber: policy.PolicyNumber || '',
          PolicyCategory: policy.PolicyCategory || '',
          ArchivedDate: new Date().toISOString(),
          ArchivedBy: archivedBy,
          RetentionRuleId: retentionRuleId,
          OriginalContent: policy.HTMLContent || policy.PolicyContent || policy.Description || '',
          OriginalMetadata: metadata,
          ArchiveReason: 'Retention policy applied'
        });

      // 5. Update PM_Policies status to Archived
      await this.sp.web.lists
        .getByTitle(LIST_POLICIES)
        .items.getById(policyId)
        .update({ PolicyStatus: 'Archived' });

      logger.info('RetentionService', `Policy ${policyId} archived by ${archivedBy}`);
    } catch (err) {
      logger.error('RetentionService', `Failed to archive policy ${policyId}:`, err);
      throw err;
    }
  }

  /**
   * Load archived policies from PM_RetentionArchive.
   */
  public async getArchive(): Promise<IRetentionArchiveItem[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(LIST_RETENTION_ARCHIVE)
        .items.select(
          'Id', 'Title', 'OriginalPolicyId', 'PolicyTitle', 'PolicyNumber',
          'PolicyCategory', 'ArchivedDate', 'ArchivedBy', 'RetentionRuleId',
          'OriginalContent', 'OriginalMetadata', 'ArchiveReason'
        )
        .orderBy('ArchivedDate', false)
        .top(500)();

      return items.map((item: any) => ({
        Id: item.Id,
        Title: item.Title || '',
        OriginalPolicyId: item.OriginalPolicyId || 0,
        PolicyTitle: item.PolicyTitle || '',
        PolicyNumber: item.PolicyNumber || '',
        PolicyCategory: item.PolicyCategory || '',
        ArchivedDate: item.ArchivedDate || '',
        ArchivedBy: item.ArchivedBy || '',
        RetentionRuleId: item.RetentionRuleId || 0,
        OriginalContent: item.OriginalContent || '',
        OriginalMetadata: item.OriginalMetadata || '',
        ArchiveReason: item.ArchiveReason || ''
      }));
    } catch (err) {
      logger.warn('RetentionService', 'Failed to load archive:', err);
      return [];
    }
  }
}
