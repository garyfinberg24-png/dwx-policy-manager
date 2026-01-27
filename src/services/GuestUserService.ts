// @ts-nocheck
// Guest User Service
// Manages B2B guest user lifecycle including invitations, permissions, and access control

import { SPFI } from '@pnp/sp';
import { GraphFI } from '@pnp/graph';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';
import '@pnp/graph/users';

import {
  IExternalGuestUser,
  IGuestFilter,
  IInvitation,
  IInvitationResult,
  IBulkInviteResult,
  IGuestAccessDetails,
  GuestStatus,
  InvitationStatus,
  RiskLevel
} from '../models/IExternalSharing';
import { logger } from './LoggingService';

export class GuestUserService {
  private sp: SPFI;
  private graph: GraphFI;

  private readonly GUEST_USERS_LIST = 'JML_ExternalGuestUsers';
  private readonly SHARED_RESOURCES_LIST = 'JML_ExternalSharedResources';
  private readonly AUDIT_LOG_LIST = 'JML_ExternalSharingAuditLog';
  private readonly ACCESS_REVIEWS_LIST = 'JML_ExternalAccessReviews';

  private currentUserId: number = 0;
  private currentUserEmail: string = '';

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
      logger.info('GuestUserService', 'Service initialized');
    } catch (error) {
      logger.error('GuestUserService', 'Failed to initialize:', error);
      throw error;
    }
  }

  // ============================================================================
  // GUEST USER CRUD
  // ============================================================================

  /**
   * Get all guest users
   */
  public async getAllGuests(): Promise<IExternalGuestUser[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.GUEST_USERS_LIST)
        .items
        .select(
          'Id', 'Title', 'Email', 'UserPrincipalName', 'AzureADObjectId',
          'SourceOrganizationId', 'InvitedBy/Id', 'InvitedBy/Title', 'InvitedBy/EMail',
          'InvitationDate', 'InvitationStatus', 'FirstAccessDate', 'LastAccessDate',
          'AccessExpirationDate', 'IsExpired', 'AccessLevel', 'AssignedSites',
          'AssignedGroups', 'TotalResourcesAccessed', 'MFARegistered',
          'DeviceCompliant', 'RiskLevel', 'Status', 'SuspensionReason', 'Notes',
          'Created', 'Modified'
        )
        .expand('InvitedBy')
        .orderBy('Title')
        .getAll();

      return items.map(item => this.mapToGuestUser(item));
    } catch (error) {
      logger.error('GuestUserService', 'Failed to get all guests:', error);
      throw error;
    }
  }

  /**
   * Get guests with filter
   */
  public async getGuests(filter: IGuestFilter): Promise<IExternalGuestUser[]> {
    try {
      let filterParts: string[] = [];

      if (filter.organizationId) {
        filterParts.push(`SourceOrganizationId eq ${filter.organizationId}`);
      }
      if (filter.status) {
        filterParts.push(`Status eq '${filter.status}'`);
      }
      if (filter.invitationStatus) {
        filterParts.push(`InvitationStatus eq '${filter.invitationStatus}'`);
      }
      if (filter.riskLevel) {
        filterParts.push(`RiskLevel eq '${filter.riskLevel}'`);
      }
      if (filter.isExpired !== undefined) {
        filterParts.push(`IsExpired eq ${filter.isExpired}`);
      }

      let query = this.sp.web.lists
        .getByTitle(this.GUEST_USERS_LIST)
        .items
        .select(
          'Id', 'Title', 'Email', 'UserPrincipalName', 'AzureADObjectId',
          'SourceOrganizationId', 'InvitedBy/Id', 'InvitedBy/Title',
          'InvitationDate', 'InvitationStatus', 'FirstAccessDate', 'LastAccessDate',
          'AccessExpirationDate', 'IsExpired', 'AccessLevel', 'TotalResourcesAccessed',
          'MFARegistered', 'DeviceCompliant', 'RiskLevel', 'Status', 'Notes'
        )
        .expand('InvitedBy');

      if (filterParts.length > 0) {
        query = query.filter(filterParts.join(' and '));
      }

      const items = await query.orderBy('Title').getAll();

      // Apply search filter client-side if provided
      let results = items.map(item => this.mapToGuestUser(item));
      if (filter.searchText) {
        const searchLower = filter.searchText.toLowerCase();
        results = results.filter(g =>
          g.Title.toLowerCase().includes(searchLower) ||
          g.Email.toLowerCase().includes(searchLower)
        );
      }

      return results;
    } catch (error) {
      logger.error('GuestUserService', 'Failed to get guests with filter:', error);
      throw error;
    }
  }

  /**
   * Get guests by organization
   */
  public async getGuestsByOrganization(organizationId: number): Promise<IExternalGuestUser[]> {
    return this.getGuests({ organizationId });
  }

  /**
   * Get guest by ID
   */
  public async getGuestById(id: number): Promise<IExternalGuestUser | null> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.GUEST_USERS_LIST)
        .items
        .getById(id)
        .select(
          'Id', 'Title', 'Email', 'UserPrincipalName', 'AzureADObjectId',
          'SourceOrganizationId', 'InvitedBy/Id', 'InvitedBy/Title', 'InvitedBy/EMail',
          'InvitationDate', 'InvitationStatus', 'FirstAccessDate', 'LastAccessDate',
          'AccessExpirationDate', 'IsExpired', 'AccessLevel', 'AssignedSites',
          'AssignedGroups', 'TotalResourcesAccessed', 'MFARegistered',
          'DeviceCompliant', 'RiskLevel', 'Status', 'SuspensionReason', 'Notes'
        )
        .expand('InvitedBy')();

      return this.mapToGuestUser(item);
    } catch (error) {
      logger.error('GuestUserService', `Failed to get guest ${id}:`, error);
      return null;
    }
  }

  /**
   * Get guest by email
   */
  public async getGuestByEmail(email: string): Promise<IExternalGuestUser | null> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.GUEST_USERS_LIST)
        .items
        .filter(`Email eq '${email}'`)
        .top(1)();

      if (items.length === 0) return null;
      return this.mapToGuestUser(items[0]);
    } catch (error) {
      logger.error('GuestUserService', `Failed to get guest by email ${email}:`, error);
      return null;
    }
  }

  // ============================================================================
  // GUEST INVITATION
  // ============================================================================

  /**
   * Invite a guest user
   */
  public async inviteGuestUser(invitation: IInvitation): Promise<IInvitationResult> {
    try {
      // Check if guest already exists
      const existing = await this.getGuestByEmail(invitation.email);
      if (existing) {
        return {
          success: false,
          error: `User ${invitation.email} is already a guest`,
          guestId: existing.Id
        };
      }

      // Send invitation via Graph API
      let azureAdObjectId: string | undefined;
      let invitationUrl: string | undefined;

      try {
        const inviteResponse = await (this.graph as any).api('/invitations').post({
          invitedUserEmailAddress: invitation.email,
          invitedUserDisplayName: invitation.displayName,
          sendInvitationMessage: invitation.sendInvitationMessage,
          inviteRedirectUrl: 'https://myapps.microsoft.com',
          invitedUserMessageInfo: invitation.invitationMessage ? {
            customizedMessageBody: invitation.invitationMessage
          } : undefined
        });

        azureAdObjectId = inviteResponse.invitedUser?.id;
        invitationUrl = inviteResponse.inviteRedeemUrl;
      } catch (graphError) {
        logger.warn('GuestUserService', 'Graph invitation failed, creating local record only:', graphError);
      }

      // Calculate expiration date
      const expirationDate = new Date();
      expirationDate.setDate(expirationDate.getDate() + (invitation.accessExpirationDays || 90));

      // Create SharePoint record
      const itemData = {
        Title: invitation.displayName,
        Email: invitation.email,
        AzureADObjectId: azureAdObjectId,
        SourceOrganizationId: invitation.sourceOrganizationId,
        InvitedById: this.currentUserId,
        InvitationDate: new Date().toISOString(),
        InvitationStatus: azureAdObjectId ? InvitationStatus.PendingAcceptance : InvitationStatus.Failed,
        AccessExpirationDate: expirationDate.toISOString(),
        IsExpired: false,
        AccessLevel: invitation.accessLevel || 'Guest',
        TotalResourcesAccessed: 0,
        MFARegistered: false,
        DeviceCompliant: false,
        RiskLevel: RiskLevel.Low,
        Status: GuestStatus.Active
      };

      const result = await this.sp.web.lists
        .getByTitle(this.GUEST_USERS_LIST)
        .items
        .add(itemData);

      return {
        success: true,
        guestId: result.data.Id,
        azureAdObjectId,
        invitationUrl
      };
    } catch (error) {
      logger.error('GuestUserService', 'Failed to invite guest:', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to invite guest'
      };
    }
  }

  /**
   * Bulk invite guests
   */
  public async bulkInviteGuests(invitations: IInvitation[]): Promise<IBulkInviteResult> {
    const results: IInvitationResult[] = [];

    for (const invitation of invitations) {
      const result = await this.inviteGuestUser(invitation);
      results.push(result);
    }

    return {
      total: invitations.length,
      succeeded: results.filter(r => r.success).length,
      failed: results.filter(r => !r.success).length,
      results
    };
  }

  // ============================================================================
  // GUEST MANAGEMENT
  // ============================================================================

  /**
   * Remove guest user
   */
  public async removeGuestUser(guestId: number, reason?: string): Promise<void> {
    try {
      const guest = await this.getGuestById(guestId);
      if (!guest) {
        throw new Error('Guest user not found');
      }

      // Remove from Azure AD if we have the object ID
      if (guest.AzureADObjectId) {
        try {
          await (this.graph as any).api(`/users/${guest.AzureADObjectId}`).delete();
        } catch (graphError) {
          logger.warn('GuestUserService', 'Failed to remove guest from Azure AD:', graphError);
        }
      }

      // Update status in SharePoint
      await this.sp.web.lists
        .getByTitle(this.GUEST_USERS_LIST)
        .items
        .getById(guestId)
        .update({
          Status: GuestStatus.Removed,
          Notes: `${guest.Notes || ''}\n\n[REMOVED ${new Date().toISOString()}] ${reason || 'No reason provided'}`
        });
    } catch (error) {
      logger.error('GuestUserService', `Failed to remove guest ${guestId}:`, error);
      throw error;
    }
  }

  /**
   * Suspend guest user
   */
  public async suspendGuestUser(guestId: number, reason: string): Promise<void> {
    try {
      const guest = await this.getGuestById(guestId);
      if (!guest) {
        throw new Error('Guest user not found');
      }

      // Block sign-in in Azure AD if we have the object ID
      if (guest.AzureADObjectId) {
        try {
          await (this.graph as any).api(`/users/${guest.AzureADObjectId}`).patch({
            accountEnabled: false
          });
        } catch (graphError) {
          logger.warn('GuestUserService', 'Failed to disable guest in Azure AD:', graphError);
        }
      }

      // Update status in SharePoint
      await this.sp.web.lists
        .getByTitle(this.GUEST_USERS_LIST)
        .items
        .getById(guestId)
        .update({
          Status: GuestStatus.Suspended,
          SuspensionReason: reason
        });
    } catch (error) {
      logger.error('GuestUserService', `Failed to suspend guest ${guestId}:`, error);
      throw error;
    }
  }

  /**
   * Reactivate suspended guest user
   */
  public async reactivateGuestUser(guestId: number): Promise<void> {
    try {
      const guest = await this.getGuestById(guestId);
      if (!guest) {
        throw new Error('Guest user not found');
      }

      if (guest.Status !== GuestStatus.Suspended) {
        throw new Error('Guest is not suspended');
      }

      // Re-enable sign-in in Azure AD
      if (guest.AzureADObjectId) {
        try {
          await (this.graph as any).api(`/users/${guest.AzureADObjectId}`).patch({
            accountEnabled: true
          });
        } catch (graphError) {
          logger.warn('GuestUserService', 'Failed to enable guest in Azure AD:', graphError);
        }
      }

      // Update status in SharePoint
      await this.sp.web.lists
        .getByTitle(this.GUEST_USERS_LIST)
        .items
        .getById(guestId)
        .update({
          Status: GuestStatus.Active,
          SuspensionReason: null
        });
    } catch (error) {
      logger.error('GuestUserService', `Failed to reactivate guest ${guestId}:`, error);
      throw error;
    }
  }

  /**
   * Update guest expiration date
   */
  public async updateGuestExpiration(guestId: number, newExpirationDate: Date): Promise<void> {
    try {
      const isExpired = newExpirationDate < new Date();

      await this.sp.web.lists
        .getByTitle(this.GUEST_USERS_LIST)
        .items
        .getById(guestId)
        .update({
          AccessExpirationDate: newExpirationDate.toISOString(),
          IsExpired: isExpired
        });
    } catch (error) {
      logger.error('GuestUserService', `Failed to update guest expiration ${guestId}:`, error);
      throw error;
    }
  }

  /**
   * Update guest risk level
   */
  public async updateGuestRiskLevel(guestId: number, riskLevel: RiskLevel): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.GUEST_USERS_LIST)
        .items
        .getById(guestId)
        .update({
          RiskLevel: riskLevel
        });
    } catch (error) {
      logger.error('GuestUserService', `Failed to update guest risk level ${guestId}:`, error);
      throw error;
    }
  }

  /**
   * Get detailed access information for a guest
   */
  public async getGuestAccessDetails(guestId: number): Promise<IGuestAccessDetails | null> {
    try {
      const guest = await this.getGuestById(guestId);
      if (!guest) return null;

      // Get shared resources accessible by this guest
      const resources = await this.sp.web.lists
        .getByTitle(this.SHARED_RESOURCES_LIST)
        .items
        .filter(`substringof('${guest.Email}', SharedWithUsers) or SourceOrganizationId eq ${guest.SourceOrganizationId}`)
        .top(50)();

      // Get recent activity
      const auditLogs = await this.sp.web.lists
        .getByTitle(this.AUDIT_LOG_LIST)
        .items
        .filter(`TargetUser eq '${guest.Email}'`)
        .orderBy('PerformedDate', false)
        .top(20)();

      // Get pending reviews
      const reviews = await this.sp.web.lists
        .getByTitle(this.ACCESS_REVIEWS_LIST)
        .items
        .filter(`TargetId eq '${guestId}' and Status eq 'Pending'`)
        .top(10)();

      return {
        guest,
        resources: resources as any[],
        recentActivity: auditLogs as any[],
        pendingReviews: reviews as any[]
      };
    } catch (error) {
      logger.error('GuestUserService', `Failed to get access details for guest ${guestId}:`, error);
      return null;
    }
  }

  /**
   * Record guest first access
   */
  public async recordFirstAccess(guestId: number): Promise<void> {
    try {
      const now = new Date().toISOString();
      await this.sp.web.lists
        .getByTitle(this.GUEST_USERS_LIST)
        .items
        .getById(guestId)
        .update({
          FirstAccessDate: now,
          LastAccessDate: now,
          InvitationStatus: InvitationStatus.Accepted
        });
    } catch (error) {
      logger.error('GuestUserService', `Failed to record first access for guest ${guestId}:`, error);
    }
  }

  /**
   * Update last access date
   */
  public async updateLastAccess(guestId: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.GUEST_USERS_LIST)
        .items
        .getById(guestId)
        .update({
          LastAccessDate: new Date().toISOString()
        });
    } catch (error) {
      logger.error('GuestUserService', `Failed to update last access for guest ${guestId}:`, error);
    }
  }

  /**
   * Increment resource access count
   */
  public async incrementResourceAccessCount(guestId: number): Promise<void> {
    try {
      const guest = await this.getGuestById(guestId);
      if (guest) {
        await this.sp.web.lists
          .getByTitle(this.GUEST_USERS_LIST)
          .items
          .getById(guestId)
          .update({
            TotalResourcesAccessed: (guest.TotalResourcesAccessed || 0) + 1,
            LastAccessDate: new Date().toISOString()
          });
      }
    } catch (error) {
      logger.error('GuestUserService', `Failed to increment access count for guest ${guestId}:`, error);
    }
  }

  // ============================================================================
  // HELPER METHODS
  // ============================================================================

  /**
   * Map SharePoint item to IExternalGuestUser
   */
  private mapToGuestUser(item: any): IExternalGuestUser {
    return {
      Id: item.Id,
      Title: item.Title,
      Email: item.Email,
      UserPrincipalName: item.UserPrincipalName,
      AzureADObjectId: item.AzureADObjectId,
      SourceOrganizationId: item.SourceOrganizationId,
      InvitedById: item.InvitedBy?.Id,
      InvitedBy: item.InvitedBy ? { Id: item.InvitedBy.Id, Title: item.InvitedBy.Title, EMail: item.InvitedBy.EMail } : undefined,
      InvitationDate: item.InvitationDate ? new Date(item.InvitationDate) : undefined,
      InvitationStatus: item.InvitationStatus as InvitationStatus || InvitationStatus.PendingAcceptance,
      FirstAccessDate: item.FirstAccessDate ? new Date(item.FirstAccessDate) : undefined,
      LastAccessDate: item.LastAccessDate ? new Date(item.LastAccessDate) : undefined,
      AccessExpirationDate: item.AccessExpirationDate ? new Date(item.AccessExpirationDate) : undefined,
      IsExpired: item.IsExpired || false,
      AccessLevel: item.AccessLevel || 'Guest',
      AssignedSites: item.AssignedSites,
      AssignedGroups: item.AssignedGroups,
      TotalResourcesAccessed: item.TotalResourcesAccessed || 0,
      MFARegistered: item.MFARegistered || false,
      DeviceCompliant: item.DeviceCompliant || false,
      RiskLevel: item.RiskLevel as RiskLevel || RiskLevel.Low,
      Status: item.Status as GuestStatus || GuestStatus.Active,
      SuspensionReason: item.SuspensionReason,
      Notes: item.Notes,
      Created: item.Created ? new Date(item.Created) : undefined,
      Modified: item.Modified ? new Date(item.Modified) : undefined
    };
  }
}
