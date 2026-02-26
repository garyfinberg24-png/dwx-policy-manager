// @ts-nocheck
// PolicyDistributionService.ts
// Service layer for Policy Distribution campaigns — CRUD against PM_PolicyDistributions,
// plus helper queries for policies and policy packs used in the create/edit form.

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { PolicyLists, PolicyPackLists } from '../constants/SharePointListNames';
import { logger } from './LoggingService';
import { ValidationUtils } from '../utils/ValidationUtils';

// ============================================================================
// Interfaces
// ============================================================================

/** Shape of a distribution item as stored in PM_PolicyDistributions */
export interface ISPDistributionItem {
  Id: number;
  Title: string;
  DistributionName: string;
  CampaignName?: string;
  ContentType?: string;
  PolicyId?: number;
  PolicyTitle?: string;
  PolicyPackId?: number;
  PolicyPackName?: string;
  DistributionScope: string;
  TargetUsers?: string;
  TargetGroups?: string;
  ScheduledDate?: string;
  DistributedDate?: string;
  DueDate?: string;
  TargetCount: number;
  TotalSent: number;
  TotalDelivered: number;
  TotalOpened: number;
  TotalAcknowledged: number;
  TotalOverdue: number;
  TotalExempted: number;
  TotalFailed: number;
  EscalationEnabled?: boolean;
  ReminderSchedule?: string;
  Status?: string;
  IsActive: boolean;
  CompletedDate?: string;
  Created: string;
  Modified: string;
  Author?: { Title: string };
}

/** Shape of a recipient/acknowledgement row from PM_PolicyAcknowledgements */
export interface ISPRecipientItem {
  Id: number;
  Title: string;
  UserEmail?: string;
  Department?: string;
  AckStatus: string;
  DueDate?: string;
  SentDate?: string;
  OpenedDate?: string;
  AcknowledgedDate?: string;
}

/** Minimal policy record for dropdowns */
export interface ISPPolicyOption {
  Id: number;
  Title: string;
  PolicyName?: string;
  PolicyCategory?: string;
}

/** Minimal policy pack record for dropdowns */
export interface ISPPolicyPackOption {
  Id: number;
  Title: string;
  PackName?: string;
}

// ============================================================================
// Service
// ============================================================================

export class PolicyDistributionService {
  private sp: SPFI;

  private readonly DISTRIBUTIONS_LIST = PolicyLists.POLICY_DISTRIBUTIONS;
  private readonly ACKNOWLEDGEMENTS_LIST = PolicyLists.POLICY_ACKNOWLEDGEMENTS;
  private readonly POLICIES_LIST = PolicyLists.POLICIES;
  private readonly POLICY_PACKS_LIST = PolicyPackLists.POLICY_PACKS;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ──────────── Distribution CRUD ────────────

  /**
   * Fetch all distribution campaigns, most-recently-modified first.
   */
  public async getDistributions(): Promise<ISPDistributionItem[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.DISTRIBUTIONS_LIST)
        .items.select(
          'Id', 'Title', 'DistributionName', 'CampaignName', 'ContentType',
          'PolicyId', 'PolicyTitle', 'PolicyPackId', 'PolicyPackName',
          'DistributionScope', 'TargetUsers', 'TargetGroups',
          'ScheduledDate', 'DistributedDate', 'DueDate',
          'TargetCount', 'TotalSent', 'TotalDelivered', 'TotalOpened',
          'TotalAcknowledged', 'TotalOverdue', 'TotalExempted', 'TotalFailed',
          'EscalationEnabled', 'ReminderSchedule', 'Status',
          'IsActive', 'CompletedDate', 'Created', 'Modified',
          'Author/Title'
        )
        .expand('Author')
        .orderBy('Modified', false)
        .top(200)();

      logger.info('PolicyDistributionService', `Loaded ${items.length} distributions from SharePoint`);
      return items as ISPDistributionItem[];
    } catch (error) {
      logger.error('PolicyDistributionService', 'getDistributions failed:', error);
      throw error;
    }
  }

  /**
   * Create a new distribution campaign.
   */
  public async createDistribution(data: Record<string, unknown>): Promise<ISPDistributionItem> {
    try {
      const result = await this.sp.web.lists
        .getByTitle(this.DISTRIBUTIONS_LIST)
        .items.add(data);

      logger.info('PolicyDistributionService', `Created distribution id=${result.Id}`);
      return result as unknown as ISPDistributionItem;
    } catch (error) {
      logger.error('PolicyDistributionService', 'createDistribution failed:', error);
      throw error;
    }
  }

  /**
   * Update an existing distribution campaign.
   */
  public async updateDistribution(id: number, data: Record<string, unknown>): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.DISTRIBUTIONS_LIST)
        .items.getById(id)
        .update(data);

      logger.info('PolicyDistributionService', `Updated distribution id=${id}`);
    } catch (error) {
      logger.error('PolicyDistributionService', `updateDistribution id=${id} failed:`, error);
      throw error;
    }
  }

  /**
   * Delete a distribution campaign.
   */
  public async deleteDistribution(id: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.DISTRIBUTIONS_LIST)
        .items.getById(id)
        .delete();

      logger.info('PolicyDistributionService', `Deleted distribution id=${id}`);
    } catch (error) {
      logger.error('PolicyDistributionService', `deleteDistribution id=${id} failed:`, error);
      throw error;
    }
  }

  // ──────────── Recipients ────────────

  /**
   * Fetch recipients (acknowledgement rows) for a specific distribution.
   * Uses OData sanitization on the distributionId to prevent injection.
   */
  public async getDistributionRecipients(distributionId: number): Promise<ISPRecipientItem[]> {
    try {
      const safeId = ValidationUtils.sanitizeForOData(String(distributionId));
      const items = await this.sp.web.lists
        .getByTitle(this.ACKNOWLEDGEMENTS_LIST)
        .items.filter(`DistributionId eq ${safeId}`)
        .select(
          'Id', 'Title', 'UserEmail', 'Department',
          'AckStatus', 'DueDate', 'SentDate', 'OpenedDate', 'AcknowledgedDate'
        )
        .top(500)();

      logger.info('PolicyDistributionService', `Loaded ${items.length} recipients for distribution id=${distributionId}`);
      return items as ISPRecipientItem[];
    } catch (error) {
      logger.error('PolicyDistributionService', `getDistributionRecipients id=${distributionId} failed:`, error);
      throw error;
    }
  }

  // ──────────── Dropdown helpers ────────────

  /**
   * Fetch published policies for the form dropdown.
   */
  public async getPolicies(): Promise<ISPPolicyOption[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.filter("PolicyStatus eq 'Published'")
        .select('Id', 'Title', 'PolicyName', 'PolicyCategory')
        .orderBy('PolicyName')
        .top(200)();

      return items as ISPPolicyOption[];
    } catch (error) {
      logger.error('PolicyDistributionService', 'getPolicies failed:', error);
      throw error;
    }
  }

  /**
   * Fetch policy packs for the form dropdown.
   */
  public async getPolicyPacks(): Promise<ISPPolicyPackOption[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.POLICY_PACKS_LIST)
        .items.select('Id', 'Title', 'PackName')
        .orderBy('Title')
        .top(100)();

      return items as ISPPolicyPackOption[];
    } catch (error) {
      logger.error('PolicyDistributionService', 'getPolicyPacks failed:', error);
      throw error;
    }
  }
}
