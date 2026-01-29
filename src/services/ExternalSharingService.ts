// @ts-nocheck
// External Sharing Service
// Main orchestrator for cross-tenant collaboration and B2B guest management

import { SPFI } from '@pnp/sp';
import { GraphFI } from '@pnp/graph';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';

import {
  ITrustedOrganization,
  IExternalGuestUser,
  ISharedResource,
  IExternalSharingAuditLog,
  IAccessReview,
  ITrustConfig,
  ITrustHealthStatus,
  IExternalSharingKPIs,
  IActivityFeedItem,
  IExternalRecipient,
  IValidationResult,
  TrustLevel,
  TrustStatus,
  GuestStatus,
  SharedResourceStatus,
  RiskLevel,
  AuditActionType,
  AuditResult,
  SharingLevel,
  ReviewType,
  ReviewStatus,
  ReviewDecision
} from '../models/IExternalSharing';
import { logger } from './LoggingService';
import { GuestUserService } from './GuestUserService';
import { SharedResourceService } from './SharedResourceService';
import { CrossTenantAccessService } from './CrossTenantAccessService';
import { ExternalSharingAuditService } from './ExternalSharingAuditService';

export class ExternalSharingService {
  private sp: SPFI;
  private graph: GraphFI;

  // SharePoint list names
  private readonly TRUSTED_ORGS_LIST = 'PM_TrustedOrganizations';
  private readonly GUEST_USERS_LIST = 'PM_ExternalGuestUsers';
  private readonly SHARED_RESOURCES_LIST = 'PM_ExternalSharedResources';
  private readonly AUDIT_LOG_LIST = 'PM_ExternalSharingAuditLog';
  private readonly POLICIES_LIST = 'PM_ExternalSharingPolicies';
  private readonly ACCESS_REVIEWS_LIST = 'PM_ExternalAccessReviews';

  // Current user context
  private currentUserId: number = 0;
  private currentUserEmail: string = '';
  private currentUserName: string = '';

  // Sub-services
  private guestUserService: GuestUserService | null = null;
  private sharedResourceService: SharedResourceService | null = null;
  private crossTenantService: CrossTenantAccessService | null = null;
  private auditService: ExternalSharingAuditService | null = null;

  constructor(sp: SPFI, graph: GraphFI) {
    this.sp = sp;
    this.graph = graph;
  }

  /**
   * Initialize service with current user context
   */
  public async initialize(): Promise<void> {
    try {
      const user = await this.sp.web.currentUser();
      this.currentUserId = user.Id;
      this.currentUserEmail = user.Email;
      this.currentUserName = user.Title;

      // Initialize sub-services
      this.guestUserService = new GuestUserService(this.sp, this.graph);
      this.sharedResourceService = new SharedResourceService(this.sp);
      this.crossTenantService = new CrossTenantAccessService(this.graph);
      this.auditService = new ExternalSharingAuditService(this.sp);

      await this.guestUserService.initialize();
      await this.sharedResourceService.initialize();
      await this.auditService.initialize();

      logger.info('ExternalSharingService', 'Service initialized successfully');
    } catch (error) {
      logger.error('ExternalSharingService', 'Failed to initialize service:', error);
      throw error;
    }
  }

  /**
   * Get sub-services for direct access
   */
  public getGuestUserService(): GuestUserService {
    if (!this.guestUserService) {
      throw new Error('Service not initialized. Call initialize() first.');
    }
    return this.guestUserService;
  }

  public getSharedResourceService(): SharedResourceService {
    if (!this.sharedResourceService) {
      throw new Error('Service not initialized. Call initialize() first.');
    }
    return this.sharedResourceService;
  }

  public getCrossTenantService(): CrossTenantAccessService {
    if (!this.crossTenantService) {
      throw new Error('Service not initialized. Call initialize() first.');
    }
    return this.crossTenantService;
  }

  public getAuditService(): ExternalSharingAuditService {
    if (!this.auditService) {
      throw new Error('Service not initialized. Call initialize() first.');
    }
    return this.auditService;
  }

  // ============================================================================
  // KPIs AND DASHBOARD
  // ============================================================================

  /**
   * Get External Sharing KPIs for dashboard
   */
  public async getKPIs(): Promise<IExternalSharingKPIs> {
    try {
      const today = new Date();
      const thirtyDaysFromNow = new Date();
      thirtyDaysFromNow.setDate(today.getDate() + 30);

      // Get counts in parallel
      const [
        trustedOrgs,
        guestUsers,
        sharedResources,
        accessReviews,
        alerts
      ] = await Promise.all([
        this.getTrustedOrganizations(),
        this.guestUserService?.getAllGuests() || [],
        this.sharedResourceService?.getAllResources() || [],
        this.getAccessReviews(),
        this.auditService?.getSecurityAlerts() || []
      ]);

      const activeOrgs = trustedOrgs.filter(o => o.Status === TrustStatus.Active);
      const pendingOrgs = trustedOrgs.filter(o => o.Status === TrustStatus.Pending);
      const activeGuests = guestUsers.filter(g => g.Status === GuestStatus.Active);
      const expiringGuests = guestUsers.filter(g =>
        g.AccessExpirationDate && new Date(g.AccessExpirationDate) <= thirtyDaysFromNow && !g.IsExpired
      );
      const activeResources = sharedResources.filter(r => r.Status === SharedResourceStatus.Active);
      const expiringResources = sharedResources.filter(r =>
        r.ExpirationDate && new Date(r.ExpirationDate) <= thirtyDaysFromNow && !r.IsExpired
      );
      const pendingReviews = accessReviews.filter(r => r.Status === 'Pending');
      const overdueReviews = accessReviews.filter(r =>
        r.Status === 'Pending' && new Date(r.DueDate) < today
      );

      // Calculate risk and compliance scores
      const riskScore = this.calculateRiskScore(guestUsers, sharedResources, alerts);
      const complianceScore = this.calculateComplianceScore(accessReviews, overdueReviews.length);

      return {
        activeTrustedOrganizations: activeOrgs.length,
        pendingTrustedOrganizations: pendingOrgs.length,
        activeGuestUsers: activeGuests.length,
        expiringGuestUsers: expiringGuests.length,
        activeSharedResources: activeResources.length,
        expiringResources: expiringResources.length,
        pendingAccessReviews: pendingReviews.length,
        overdueAccessReviews: overdueReviews.length,
        securityAlerts: alerts.filter(a => !a.isResolved).length,
        complianceScore,
        riskScore
      };
    } catch (error) {
      logger.error('ExternalSharingService', 'Failed to get KPIs:', error);
      throw error;
    }
  }

  /**
   * Get recent activity feed
   */
  public async getActivityFeed(count: number = 20): Promise<IActivityFeedItem[]> {
    try {
      const auditLogs = await this.auditService?.getRecentLogs(count) || [];

      return auditLogs.map(log => ({
        id: log.Id?.toString() || '',
        type: log.ActionType,
        title: log.Title,
        description: this.formatAuditDescription(log),
        timestamp: log.PerformedDate,
        performedBy: log.PerformedBy?.Title || 'System',
        organizationName: log.TargetOrganization?.Title,
        resourceName: log.TargetResource?.Title,
        isHighRisk: (log.RiskScore || 0) >= 70
      }));
    } catch (error) {
      logger.error('ExternalSharingService', 'Failed to get activity feed:', error);
      return [];
    }
  }

  // ============================================================================
  // TRUSTED ORGANIZATIONS
  // ============================================================================

  /**
   * Get all trusted organizations
   */
  public async getTrustedOrganizations(): Promise<ITrustedOrganization[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.TRUSTED_ORGS_LIST)
        .items
        .select(
          'Id', 'Title', 'TenantId', 'TenantDomain', 'TrustLevel', 'Status',
          'TrustMFAClaims', 'TrustDeviceClaims', 'TrustHybridJoinedDevices',
          'AllowedDomains', 'AllowedUserGroups', 'DefaultGuestExpiration',
          'MaxSharingLevel', 'InboundAccessEnabled', 'OutboundAccessEnabled',
          'B2BDirectConnectEnabled', 'ContactName', 'ContactEmail', 'Notes',
          'EstablishedDate', 'LastVerifiedDate', 'VerifiedBy/Id', 'VerifiedBy/Title',
          'Created', 'Modified'
        )
        .expand('VerifiedBy')
        .orderBy('Title')
        .getAll();

      return items.map(item => this.mapToTrustedOrganization(item));
    } catch (error) {
      logger.error('ExternalSharingService', 'Failed to get trusted organizations:', error);
      throw error;
    }
  }

  /**
   * Get a single trusted organization by ID
   */
  public async getTrustedOrganizationById(id: number): Promise<ITrustedOrganization | null> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.TRUSTED_ORGS_LIST)
        .items
        .getById(id)
        .select(
          'Id', 'Title', 'TenantId', 'TenantDomain', 'TrustLevel', 'Status',
          'TrustMFAClaims', 'TrustDeviceClaims', 'TrustHybridJoinedDevices',
          'AllowedDomains', 'AllowedUserGroups', 'DefaultGuestExpiration',
          'MaxSharingLevel', 'InboundAccessEnabled', 'OutboundAccessEnabled',
          'B2BDirectConnectEnabled', 'ContactName', 'ContactEmail', 'Notes',
          'EstablishedDate', 'LastVerifiedDate', 'VerifiedBy/Id', 'VerifiedBy/Title',
          'Created', 'Modified'
        )
        .expand('VerifiedBy')();

      return this.mapToTrustedOrganization(item);
    } catch (error) {
      logger.error('ExternalSharingService', `Failed to get organization ${id}:`, error);
      return null;
    }
  }

  /**
   * Get trusted organization by tenant ID
   */
  public async getTrustedOrganizationByTenantId(tenantId: string): Promise<ITrustedOrganization | null> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.TRUSTED_ORGS_LIST)
        .items
        .filter(`TenantId eq '${tenantId}'`)
        .top(1)();

      if (items.length === 0) return null;
      return this.mapToTrustedOrganization(items[0]);
    } catch (error) {
      logger.error('ExternalSharingService', `Failed to get organization by tenant ${tenantId}:`, error);
      return null;
    }
  }

  /**
   * Create a new trust relationship
   */
  public async createTrustRelationship(config: ITrustConfig): Promise<ITrustedOrganization> {
    try {
      // Validate tenant exists
      const tenantValid = await this.crossTenantService?.validateTenantExists(config.tenantId);
      if (!tenantValid) {
        throw new Error(`Tenant ${config.tenantDomain} does not exist or is not accessible`);
      }

      // Check if trust already exists
      const existing = await this.getTrustedOrganizationByTenantId(config.tenantId);
      if (existing) {
        throw new Error(`Trust relationship already exists with ${config.tenantDomain}`);
      }

      // Create SharePoint list item
      const itemData = {
        Title: config.organizationName,
        TenantId: config.tenantId,
        TenantDomain: config.tenantDomain,
        TrustLevel: config.trustLevel,
        Status: TrustStatus.Pending,
        TrustMFAClaims: config.trustMFAClaims,
        TrustDeviceClaims: config.trustDeviceClaims,
        TrustHybridJoinedDevices: config.trustHybridJoinedDevices,
        AllowedDomains: config.allowedDomains ? JSON.stringify(config.allowedDomains) : null,
        AllowedUserGroups: config.allowedUserGroups ? JSON.stringify(config.allowedUserGroups) : null,
        DefaultGuestExpiration: config.defaultGuestExpiration,
        MaxSharingLevel: config.maxSharingLevel,
        InboundAccessEnabled: config.inboundAccessEnabled,
        OutboundAccessEnabled: config.outboundAccessEnabled,
        B2BDirectConnectEnabled: config.b2bDirectConnectEnabled,
        ContactName: config.contactName,
        ContactEmail: config.contactEmail,
        Notes: config.notes,
        EstablishedDate: new Date().toISOString()
      };

      const result = await this.sp.web.lists
        .getByTitle(this.TRUSTED_ORGS_LIST)
        .items
        .add(itemData);

      // Configure cross-tenant access policy via Graph
      try {
        await this.crossTenantService?.addPartnerConfiguration(config.tenantId, {
          b2bCollaborationInbound: config.inboundAccessEnabled,
          b2bCollaborationOutbound: config.outboundAccessEnabled,
          b2bDirectConnect: config.b2bDirectConnectEnabled,
          trustMfa: config.trustMFAClaims,
          trustDeviceCompliance: config.trustDeviceClaims,
          trustHybridJoined: config.trustHybridJoinedDevices
        });

        // Update status to Active after Graph configuration succeeds
        await this.sp.web.lists
          .getByTitle(this.TRUSTED_ORGS_LIST)
          .items
          .getById(result.data.Id)
          .update({ Status: TrustStatus.Active });

      } catch (graphError) {
        logger.warn('ExternalSharingService', 'Failed to configure Graph policy, trust remains pending:', graphError);
      }

      // Log audit event
      await this.auditService?.logAction({
        actionType: AuditActionType.TrustEstablished,
        targetOrganizationId: result.data.Id,
        newValue: JSON.stringify(config),
        result: AuditResult.Success
      });

      return await this.getTrustedOrganizationById(result.data.Id) as ITrustedOrganization;
    } catch (error) {
      logger.error('ExternalSharingService', 'Failed to create trust relationship:', error);
      throw error;
    }
  }

  /**
   * Update trust relationship
   */
  public async updateTrustRelationship(orgId: number, updates: Partial<ITrustConfig>): Promise<void> {
    try {
      const existing = await this.getTrustedOrganizationById(orgId);
      if (!existing) {
        throw new Error('Trust relationship not found');
      }

      const updateData: Record<string, unknown> = {};
      if (updates.organizationName) updateData.Title = updates.organizationName;
      if (updates.trustLevel) updateData.TrustLevel = updates.trustLevel;
      if (updates.trustMFAClaims !== undefined) updateData.TrustMFAClaims = updates.trustMFAClaims;
      if (updates.trustDeviceClaims !== undefined) updateData.TrustDeviceClaims = updates.trustDeviceClaims;
      if (updates.trustHybridJoinedDevices !== undefined) updateData.TrustHybridJoinedDevices = updates.trustHybridJoinedDevices;
      if (updates.allowedDomains) updateData.AllowedDomains = JSON.stringify(updates.allowedDomains);
      if (updates.allowedUserGroups) updateData.AllowedUserGroups = JSON.stringify(updates.allowedUserGroups);
      if (updates.defaultGuestExpiration !== undefined) updateData.DefaultGuestExpiration = updates.defaultGuestExpiration;
      if (updates.maxSharingLevel) updateData.MaxSharingLevel = updates.maxSharingLevel;
      if (updates.inboundAccessEnabled !== undefined) updateData.InboundAccessEnabled = updates.inboundAccessEnabled;
      if (updates.outboundAccessEnabled !== undefined) updateData.OutboundAccessEnabled = updates.outboundAccessEnabled;
      if (updates.b2bDirectConnectEnabled !== undefined) updateData.B2BDirectConnectEnabled = updates.b2bDirectConnectEnabled;
      if (updates.contactName) updateData.ContactName = updates.contactName;
      if (updates.contactEmail) updateData.ContactEmail = updates.contactEmail;
      if (updates.notes) updateData.Notes = updates.notes;

      await this.sp.web.lists
        .getByTitle(this.TRUSTED_ORGS_LIST)
        .items
        .getById(orgId)
        .update(updateData);

      // Update Graph policy if needed
      if (updates.inboundAccessEnabled !== undefined ||
          updates.outboundAccessEnabled !== undefined ||
          updates.trustMFAClaims !== undefined ||
          updates.trustDeviceClaims !== undefined) {
        try {
          await this.crossTenantService?.updatePartnerConfiguration(existing.TenantId, {
            b2bCollaborationInbound: updates.inboundAccessEnabled ?? existing.InboundAccessEnabled,
            b2bCollaborationOutbound: updates.outboundAccessEnabled ?? existing.OutboundAccessEnabled,
            trustMfa: updates.trustMFAClaims ?? existing.TrustMFAClaims,
            trustDeviceCompliance: updates.trustDeviceClaims ?? existing.TrustDeviceClaims
          });
        } catch (graphError) {
          logger.warn('ExternalSharingService', 'Failed to update Graph policy:', graphError);
        }
      }

      // Log audit event
      await this.auditService?.logAction({
        actionType: AuditActionType.TrustModified,
        targetOrganizationId: orgId,
        previousValue: JSON.stringify(existing),
        newValue: JSON.stringify(updates),
        result: AuditResult.Success
      });
    } catch (error) {
      logger.error('ExternalSharingService', `Failed to update trust relationship ${orgId}:`, error);
      throw error;
    }
  }

  /**
   * Revoke trust relationship
   */
  public async revokeTrust(orgId: number, reason: string): Promise<void> {
    try {
      const existing = await this.getTrustedOrganizationById(orgId);
      if (!existing) {
        throw new Error('Trust relationship not found');
      }

      // Update status to Revoked
      await this.sp.web.lists
        .getByTitle(this.TRUSTED_ORGS_LIST)
        .items
        .getById(orgId)
        .update({
          Status: TrustStatus.Revoked,
          Notes: `${existing.Notes || ''}\n\n[REVOKED ${new Date().toISOString()}] Reason: ${reason}`
        });

      // Remove Graph policy
      try {
        await this.crossTenantService?.removePartnerConfiguration(existing.TenantId);
      } catch (graphError) {
        logger.warn('ExternalSharingService', 'Failed to remove Graph policy:', graphError);
      }

      // Suspend all guest users from this organization
      const guests = await this.guestUserService?.getGuestsByOrganization(orgId) || [];
      for (const guest of guests) {
        if (guest.Status === GuestStatus.Active && guest.Id) {
          await this.guestUserService?.suspendGuestUser(guest.Id, `Trust relationship revoked: ${reason}`);
        }
      }

      // Revoke all shared resources
      const resources = await this.sharedResourceService?.getResourcesByOrganization(orgId) || [];
      for (const resource of resources) {
        if (resource.Status === SharedResourceStatus.Active && resource.Id) {
          await this.sharedResourceService?.revokeSharing(resource.Id, `Trust relationship revoked: ${reason}`);
        }
      }

      // Log audit event
      await this.auditService?.logAction({
        actionType: AuditActionType.TrustRevoked,
        targetOrganizationId: orgId,
        previousValue: JSON.stringify({ status: existing.Status }),
        newValue: JSON.stringify({ status: TrustStatus.Revoked, reason }),
        result: AuditResult.Success
      });
    } catch (error) {
      logger.error('ExternalSharingService', `Failed to revoke trust ${orgId}:`, error);
      throw error;
    }
  }

  /**
   * Validate trust health
   */
  public async validateTrustHealth(orgId: number): Promise<ITrustHealthStatus> {
    try {
      const org = await this.getTrustedOrganizationById(orgId);
      if (!org) {
        throw new Error('Trust relationship not found');
      }

      const issues: string[] = [];
      let isHealthy = true;

      // Check Graph policy sync
      let policyInSync = false;
      try {
        const graphPolicy = await this.crossTenantService?.getPartnerConfiguration(org.TenantId);
        policyInSync = !!graphPolicy;
        if (!policyInSync) {
          issues.push('Cross-tenant access policy not configured in Azure AD');
          isHealthy = false;
        }
      } catch {
        issues.push('Unable to verify cross-tenant access policy');
        isHealthy = false;
      }

      // Get related data
      const guests = await this.guestUserService?.getGuestsByOrganization(orgId) || [];
      const resources = await this.sharedResourceService?.getResourcesByOrganization(orgId) || [];
      const reviews = await this.getAccessReviewsByOrganization(orgId);

      const activeGuests = guests.filter(g => g.Status === GuestStatus.Active);
      const activeResources = resources.filter(r => r.Status === SharedResourceStatus.Active);
      const pendingReviews = reviews.filter(r => r.Status === 'Pending');

      // Check for expired guests
      const expiredGuests = guests.filter(g => g.IsExpired && g.Status === GuestStatus.Active);
      if (expiredGuests.length > 0) {
        issues.push(`${expiredGuests.length} guest(s) with expired access still active`);
        isHealthy = false;
      }

      // Check for overdue reviews
      const overdueReviews = reviews.filter(r =>
        r.Status === 'Pending' && new Date(r.DueDate) < new Date()
      );
      if (overdueReviews.length > 0) {
        issues.push(`${overdueReviews.length} access review(s) overdue`);
        isHealthy = false;
      }

      // Calculate risk level
      const riskLevel = this.calculateOrganizationRisk(guests, resources, issues.length);

      // Update last verified date
      await this.sp.web.lists
        .getByTitle(this.TRUSTED_ORGS_LIST)
        .items
        .getById(orgId)
        .update({
          LastVerifiedDate: new Date().toISOString(),
          VerifiedById: this.currentUserId
        });

      return {
        organizationId: orgId,
        organizationName: org.Title,
        tenantDomain: org.TenantDomain,
        isHealthy,
        lastVerifiedDate: new Date(),
        policyInSync,
        activeGuestCount: activeGuests.length,
        activeResourceCount: activeResources.length,
        pendingReviews: pendingReviews.length,
        riskLevel,
        issues
      };
    } catch (error) {
      logger.error('ExternalSharingService', `Failed to validate trust health ${orgId}:`, error);
      throw error;
    }
  }

  // ============================================================================
  // EXTERNAL RECIPIENT VALIDATION
  // ============================================================================

  /**
   * Validate recipients are from trusted organizations
   */
  public async validateRecipients(recipients: IExternalRecipient[]): Promise<IValidationResult> {
    const errors: string[] = [];
    const warnings: string[] = [];

    for (const recipient of recipients) {
      // Check if organization is trusted and active
      const org = await this.getTrustedOrganizationById(recipient.organizationId);
      if (!org) {
        errors.push(`Organization not found for ${recipient.email}`);
        continue;
      }

      if (org.Status !== TrustStatus.Active) {
        errors.push(`Organization ${org.Title} is not active (status: ${org.Status})`);
        continue;
      }

      // Check if email domain is allowed
      if (org.AllowedDomains) {
        const allowedDomains = JSON.parse(org.AllowedDomains) as string[];
        const emailDomain = recipient.email.split('@')[1];
        if (allowedDomains.length > 0 && !allowedDomains.includes(emailDomain)) {
          errors.push(`Email domain ${emailDomain} is not in allowed domains for ${org.Title}`);
        }
      }

      // Check if it's a new user that needs invitation
      if (recipient.isNew) {
        warnings.push(`${recipient.email} will need to be invited as a guest user`);
      }
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }

  // ============================================================================
  // ACCESS REVIEWS
  // ============================================================================

  /**
   * Get all access reviews (public)
   */
  public async getAllAccessReviews(showOnlyMyReviews: boolean = false): Promise<IAccessReview[]> {
    try {
      let items: any[];

      if (showOnlyMyReviews) {
        // Get current user's email
        const currentUser = await this.sp.web.currentUser();
        items = await this.sp.web.lists
          .getByTitle(this.ACCESS_REVIEWS_LIST)
          .items
          .filter(`ReviewerEmail eq '${currentUser.Email}'`)
          .select('Id', 'Title', 'ReviewType', 'TargetId', 'TargetType', 'ReviewerEmail', 'ReviewerId', 'Status', 'DueDate', 'CompletedDate', 'Decision', 'Justification', 'ActionTaken', 'ActionDate', 'NextReviewDate', 'AutoRevokeOnExpiry')
          .orderBy('DueDate', true)
          .getAll();
      } else {
        items = await this.sp.web.lists
          .getByTitle(this.ACCESS_REVIEWS_LIST)
          .items
          .select('Id', 'Title', 'ReviewType', 'TargetId', 'TargetType', 'ReviewerEmail', 'ReviewerId', 'Status', 'DueDate', 'CompletedDate', 'Decision', 'Justification', 'ActionTaken', 'ActionDate', 'NextReviewDate', 'AutoRevokeOnExpiry')
          .orderBy('DueDate', true)
          .getAll();
      }

      return items.map(item => this.mapToAccessReview(item));
    } catch (error) {
      logger.error('ExternalSharingService', 'Failed to get access reviews:', error);
      // Return empty array instead of throwing to gracefully handle list not existing
      return [];
    }
  }

  /**
   * Complete an access review
   */
  public async completeAccessReview(reviewId: number, decision: ReviewDecision, justification: string): Promise<void> {
    try {
      const actionTaken = decision === ReviewDecision.Revoke ? 'AccessRevoked' :
                          decision === ReviewDecision.Modify ? 'PermissionsModified' : 'Approved';

      await this.sp.web.lists
        .getByTitle(this.ACCESS_REVIEWS_LIST)
        .items
        .getById(reviewId)
        .update({
          Status: ReviewStatus.Completed,
          CompletedDate: new Date().toISOString(),
          Decision: decision,
          Justification: justification,
          ActionTaken: actionTaken,
          ActionDate: new Date().toISOString()
        });

      logger.info('ExternalSharingService', `Access review ${reviewId} completed with decision: ${decision}`);
    } catch (error) {
      logger.error('ExternalSharingService', `Failed to complete access review ${reviewId}:`, error);
      throw error;
    }
  }

  /**
   * Map SharePoint item to IAccessReview
   */
  private mapToAccessReview(item: any): IAccessReview {
    return {
      Id: item.Id,
      Title: item.Title,
      ReviewType: item.ReviewType as ReviewType,
      TargetId: item.TargetId,
      TargetType: item.TargetType,
      ReviewerEmail: item.ReviewerEmail,
      ReviewerId: item.ReviewerId,
      Status: item.Status as ReviewStatus,
      DueDate: item.DueDate,
      CompletedDate: item.CompletedDate,
      Decision: item.Decision as ReviewDecision,
      Justification: item.Justification,
      ActionTaken: item.ActionTaken,
      ActionDate: item.ActionDate,
      NextReviewDate: item.NextReviewDate,
      AutoRevokeOnExpiry: item.AutoRevokeOnExpiry || false
    };
  }

  /**
   * Get all access reviews (internal use)
   */
  private async getAccessReviews(): Promise<any[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.ACCESS_REVIEWS_LIST)
        .items
        .select('Id', 'Title', 'ReviewType', 'Status', 'DueDate', 'CompletedDate')
        .getAll();
      return items;
    } catch (error) {
      logger.warn('ExternalSharingService', 'Failed to get access reviews:', error);
      return [];
    }
  }

  /**
   * Get access reviews for an organization
   */
  private async getAccessReviewsByOrganization(orgId: number): Promise<any[]> {
    try {
      // This would filter by organization - simplified for now
      return await this.getAccessReviews();
    } catch (error) {
      logger.warn('ExternalSharingService', 'Failed to get organization access reviews:', error);
      return [];
    }
  }

  // ============================================================================
  // HELPER METHODS
  // ============================================================================

  /**
   * Map SharePoint item to ITrustedOrganization
   */
  private mapToTrustedOrganization(item: any): ITrustedOrganization {
    return {
      Id: item.Id,
      Title: item.Title,
      TenantId: item.TenantId,
      TenantDomain: item.TenantDomain,
      TrustLevel: item.TrustLevel as TrustLevel,
      Status: item.Status as TrustStatus,
      TrustMFAClaims: item.TrustMFAClaims || false,
      TrustDeviceClaims: item.TrustDeviceClaims || false,
      TrustHybridJoinedDevices: item.TrustHybridJoinedDevices || false,
      AllowedDomains: item.AllowedDomains,
      AllowedUserGroups: item.AllowedUserGroups,
      DefaultGuestExpiration: item.DefaultGuestExpiration || 90,
      MaxSharingLevel: item.MaxSharingLevel as SharingLevel || SharingLevel.View,
      InboundAccessEnabled: item.InboundAccessEnabled || false,
      OutboundAccessEnabled: item.OutboundAccessEnabled || false,
      B2BDirectConnectEnabled: item.B2BDirectConnectEnabled || false,
      ContactName: item.ContactName,
      ContactEmail: item.ContactEmail,
      Notes: item.Notes,
      EstablishedDate: item.EstablishedDate ? new Date(item.EstablishedDate) : undefined,
      LastVerifiedDate: item.LastVerifiedDate ? new Date(item.LastVerifiedDate) : undefined,
      VerifiedById: item.VerifiedBy?.Id,
      VerifiedBy: item.VerifiedBy ? { Id: item.VerifiedBy.Id, Title: item.VerifiedBy.Title } : undefined,
      Created: item.Created ? new Date(item.Created) : undefined,
      Modified: item.Modified ? new Date(item.Modified) : undefined
    };
  }

  /**
   * Format audit log entry description
   */
  private formatAuditDescription(log: IExternalSharingAuditLog): string {
    const action = log.ActionType.replace(/([A-Z])/g, ' $1').trim();
    let desc = action;

    if (log.TargetUser) {
      desc += ` for ${log.TargetUser}`;
    }
    if (log.TargetResource?.Title) {
      desc += ` - ${log.TargetResource.Title}`;
    }
    if (log.TargetOrganization?.Title) {
      desc += ` (${log.TargetOrganization.Title})`;
    }

    return desc;
  }

  /**
   * Calculate overall risk score
   */
  private calculateRiskScore(
    guests: IExternalGuestUser[],
    resources: ISharedResource[],
    alerts: any[]
  ): number {
    let score = 0;

    // High risk guests
    const highRiskGuests = guests.filter(g => g.RiskLevel === RiskLevel.High);
    score += highRiskGuests.length * 10;

    // Expired but active items
    const expiredGuests = guests.filter(g => g.IsExpired && g.Status === GuestStatus.Active);
    score += expiredGuests.length * 15;

    const expiredResources = resources.filter(r => r.IsExpired && r.Status === SharedResourceStatus.Active);
    score += expiredResources.length * 10;

    // Unresolved alerts
    const unresolvedAlerts = alerts.filter(a => !a.isResolved);
    score += unresolvedAlerts.length * 20;

    // Cap at 100
    return Math.min(score, 100);
  }

  /**
   * Calculate compliance score
   */
  private calculateComplianceScore(reviews: any[], overdueCount: number): number {
    if (reviews.length === 0) return 100;

    const completedReviews = reviews.filter(r => r.Status === 'Completed');
    const completionRate = (completedReviews.length / reviews.length) * 100;

    // Deduct for overdue reviews
    const overdueDeduction = overdueCount * 5;

    return Math.max(0, Math.min(100, completionRate - overdueDeduction));
  }

  /**
   * Calculate organization-specific risk level
   */
  private calculateOrganizationRisk(
    guests: IExternalGuestUser[],
    resources: ISharedResource[],
    issueCount: number
  ): RiskLevel {
    let riskPoints = 0;

    // High risk guests
    riskPoints += guests.filter(g => g.RiskLevel === RiskLevel.High).length * 3;
    riskPoints += guests.filter(g => g.RiskLevel === RiskLevel.Medium).length;

    // Expired items still active
    riskPoints += guests.filter(g => g.IsExpired && g.Status === GuestStatus.Active).length * 2;
    riskPoints += resources.filter(r => r.IsExpired && r.Status === SharedResourceStatus.Active).length * 2;

    // Issues found
    riskPoints += issueCount * 2;

    if (riskPoints >= 10) return RiskLevel.High;
    if (riskPoints >= 5) return RiskLevel.Medium;
    return RiskLevel.Low;
  }
}
