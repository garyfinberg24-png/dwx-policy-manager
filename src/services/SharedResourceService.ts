// @ts-nocheck
// Shared Resource Service
// Manages external resource sharing, tracking, and access control

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users/web';

import {
  ISharedResource,
  IResourceFilter,
  IShareRequest,
  IBulkShareResult,
  IAccessLogEntry,
  ResourceType,
  SharingLevel,
  SharedResourceStatus,
  DataClassification,
  AcknowledgmentStatus,
  RelatedModule
} from '../models/IExternalSharing';
import { logger } from './LoggingService';

export class SharedResourceService {
  private sp: SPFI;

  private readonly SHARED_RESOURCES_LIST = 'JML_ExternalSharedResources';
  private readonly AUDIT_LOG_LIST = 'JML_ExternalSharingAuditLog';

  private currentUserId: number = 0;
  private currentUserEmail: string = '';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Initialize service with current user context
   */
  public async initialize(): Promise<void> {
    try {
      const user = await this.sp.web.currentUser();
      this.currentUserId = user.Id;
      this.currentUserEmail = user.Email;
      logger.info('SharedResourceService', 'Service initialized');
    } catch (error) {
      logger.error('SharedResourceService', 'Failed to initialize:', error);
      throw error;
    }
  }

  // ============================================================================
  // RESOURCE CRUD
  // ============================================================================

  /**
   * Get all shared resources
   */
  public async getAllResources(): Promise<ISharedResource[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.SHARED_RESOURCES_LIST)
        .items
        .select(
          'Id', 'Title', 'ResourceType', 'ResourceUrl', 'ResourceId',
          'SharedWithOrganizationId', 'SharedWithUsers', 'SharingLevel',
          'SharedDate', 'SharedBy/Id', 'SharedBy/Title', 'SharedBy/EMail',
          'ExpirationDate', 'IsExpired', 'DataClassification',
          'RequiresAcknowledgment', 'AcknowledgmentStatus',
          'RelatedModule', 'RelatedItemId', 'AccessCount',
          'LastAccessedDate', 'LastAccessedBy', 'Status',
          'Created', 'Modified'
        )
        .expand('SharedBy')
        .orderBy('SharedDate', false)
        .getAll();

      return items.map(item => this.mapToSharedResource(item));
    } catch (error) {
      logger.error('SharedResourceService', 'Failed to get all resources:', error);
      throw error;
    }
  }

  /**
   * Get resources with filter
   */
  public async getResources(filter: IResourceFilter): Promise<ISharedResource[]> {
    try {
      let filterParts: string[] = [];

      if (filter.organizationId) {
        filterParts.push(`SharedWithOrganizationId eq ${filter.organizationId}`);
      }
      if (filter.resourceType) {
        filterParts.push(`ResourceType eq '${filter.resourceType}'`);
      }
      if (filter.status) {
        filterParts.push(`Status eq '${filter.status}'`);
      }
      if (filter.relatedModule) {
        filterParts.push(`RelatedModule eq '${filter.relatedModule}'`);
      }
      if (filter.isExpired !== undefined) {
        filterParts.push(`IsExpired eq ${filter.isExpired}`);
      }
      if (filter.sharedById) {
        filterParts.push(`SharedById eq ${filter.sharedById}`);
      }

      let query = this.sp.web.lists
        .getByTitle(this.SHARED_RESOURCES_LIST)
        .items
        .select(
          'Id', 'Title', 'ResourceType', 'ResourceUrl', 'ResourceId',
          'SharedWithOrganizationId', 'SharedWithUsers', 'SharingLevel',
          'SharedDate', 'SharedBy/Id', 'SharedBy/Title',
          'ExpirationDate', 'IsExpired', 'DataClassification',
          'RequiresAcknowledgment', 'AcknowledgmentStatus',
          'RelatedModule', 'RelatedItemId', 'AccessCount',
          'LastAccessedDate', 'LastAccessedBy', 'Status'
        )
        .expand('SharedBy');

      if (filterParts.length > 0) {
        query = query.filter(filterParts.join(' and '));
      }

      const items = await query.orderBy('SharedDate', false).getAll();

      // Apply search filter client-side
      let results = items.map(item => this.mapToSharedResource(item));
      if (filter.searchText) {
        const searchLower = filter.searchText.toLowerCase();
        results = results.filter(r =>
          r.Title.toLowerCase().includes(searchLower) ||
          r.ResourceUrl.toLowerCase().includes(searchLower)
        );
      }

      return results;
    } catch (error) {
      logger.error('SharedResourceService', 'Failed to get resources with filter:', error);
      throw error;
    }
  }

  /**
   * Get resources by organization
   */
  public async getResourcesByOrganization(organizationId: number): Promise<ISharedResource[]> {
    return this.getResources({ organizationId });
  }

  /**
   * Get resource by ID
   */
  public async getResourceById(id: number): Promise<ISharedResource | null> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.SHARED_RESOURCES_LIST)
        .items
        .getById(id)
        .select(
          'Id', 'Title', 'ResourceType', 'ResourceUrl', 'ResourceId',
          'SharedWithOrganizationId', 'SharedWithUsers', 'SharingLevel',
          'SharedDate', 'SharedBy/Id', 'SharedBy/Title', 'SharedBy/EMail',
          'ExpirationDate', 'IsExpired', 'DataClassification',
          'RequiresAcknowledgment', 'AcknowledgmentStatus',
          'RelatedModule', 'RelatedItemId', 'AccessCount',
          'LastAccessedDate', 'LastAccessedBy', 'Status'
        )
        .expand('SharedBy')();

      return this.mapToSharedResource(item);
    } catch (error) {
      logger.error('SharedResourceService', `Failed to get resource ${id}:`, error);
      return null;
    }
  }

  // ============================================================================
  // SHARE OPERATIONS
  // ============================================================================

  /**
   * Share a resource
   */
  public async shareResource(request: IShareRequest): Promise<ISharedResource> {
    try {
      // Calculate expiration date
      const expirationDate = request.expirationDays
        ? new Date(Date.now() + request.expirationDays * 24 * 60 * 60 * 1000)
        : null;

      const itemData = {
        Title: request.title,
        ResourceType: request.resourceType,
        ResourceUrl: request.resourceUrl,
        ResourceId: request.resourceId,
        SharedWithOrganizationId: request.sharedWithOrganizationId,
        SharedWithUsers: request.sharedWithUsers ? JSON.stringify(request.sharedWithUsers) : null,
        SharingLevel: request.sharingLevel,
        SharedDate: new Date().toISOString(),
        SharedById: this.currentUserId,
        ExpirationDate: expirationDate?.toISOString() || null,
        IsExpired: false,
        DataClassification: request.dataClassification,
        RequiresAcknowledgment: request.requiresAcknowledgment,
        AcknowledgmentStatus: request.requiresAcknowledgment
          ? AcknowledgmentStatus.Pending
          : AcknowledgmentStatus.Acknowledged,
        RelatedModule: request.relatedModule,
        RelatedItemId: request.relatedItemId,
        AccessCount: 0,
        Status: SharedResourceStatus.Active
      };

      const result = await this.sp.web.lists
        .getByTitle(this.SHARED_RESOURCES_LIST)
        .items
        .add(itemData);

      logger.info('SharedResourceService', `Resource shared: ${request.title} (ID: ${result.data.Id})`);

      return await this.getResourceById(result.data.Id) as ISharedResource;
    } catch (error) {
      logger.error('SharedResourceService', 'Failed to share resource:', error);
      throw error;
    }
  }

  /**
   * Bulk share resources
   */
  public async bulkShare(requests: IShareRequest[]): Promise<IBulkShareResult> {
    const results: { resourceUrl: string; success: boolean; resourceId?: number; error?: string }[] = [];

    for (const request of requests) {
      try {
        const resource = await this.shareResource(request);
        results.push({
          resourceUrl: request.resourceUrl,
          success: true,
          resourceId: resource.Id
        });
      } catch (error) {
        results.push({
          resourceUrl: request.resourceUrl,
          success: false,
          error: error instanceof Error ? error.message : 'Unknown error'
        });
      }
    }

    return {
      total: requests.length,
      succeeded: results.filter(r => r.success).length,
      failed: results.filter(r => !r.success).length,
      results
    };
  }

  /**
   * Update sharing settings
   */
  public async updateSharing(resourceId: number, updates: Partial<IShareRequest>): Promise<void> {
    try {
      const updateData: Record<string, unknown> = {};

      if (updates.title) updateData.Title = updates.title;
      if (updates.sharingLevel) updateData.SharingLevel = updates.sharingLevel;
      if (updates.sharedWithUsers) updateData.SharedWithUsers = JSON.stringify(updates.sharedWithUsers);
      if (updates.dataClassification) updateData.DataClassification = updates.dataClassification;
      if (updates.requiresAcknowledgment !== undefined) {
        updateData.RequiresAcknowledgment = updates.requiresAcknowledgment;
      }
      if (updates.expirationDays) {
        const newExpiration = new Date(Date.now() + updates.expirationDays * 24 * 60 * 60 * 1000);
        updateData.ExpirationDate = newExpiration.toISOString();
        updateData.IsExpired = false;
      }

      await this.sp.web.lists
        .getByTitle(this.SHARED_RESOURCES_LIST)
        .items
        .getById(resourceId)
        .update(updateData);

      logger.info('SharedResourceService', `Resource ${resourceId} sharing updated`);
    } catch (error) {
      logger.error('SharedResourceService', `Failed to update sharing for resource ${resourceId}:`, error);
      throw error;
    }
  }

  /**
   * Revoke sharing
   */
  public async revokeSharing(resourceId: number, reason?: string): Promise<void> {
    try {
      const resource = await this.getResourceById(resourceId);
      if (!resource) {
        throw new Error('Resource not found');
      }

      await this.sp.web.lists
        .getByTitle(this.SHARED_RESOURCES_LIST)
        .items
        .getById(resourceId)
        .update({
          Status: SharedResourceStatus.Revoked
        });

      logger.info('SharedResourceService', `Resource ${resourceId} sharing revoked. Reason: ${reason || 'Not specified'}`);
    } catch (error) {
      logger.error('SharedResourceService', `Failed to revoke sharing for resource ${resourceId}:`, error);
      throw error;
    }
  }

  // ============================================================================
  // ACCESS TRACKING
  // ============================================================================

  /**
   * Record resource access
   */
  public async recordAccess(
    resourceId: number,
    userEmail: string,
    action: 'Viewed' | 'Downloaded' | 'Edited' | 'Shared',
    ipAddress?: string,
    deviceInfo?: string
  ): Promise<void> {
    try {
      const resource = await this.getResourceById(resourceId);
      if (!resource) return;

      // Update access count and last accessed info
      await this.sp.web.lists
        .getByTitle(this.SHARED_RESOURCES_LIST)
        .items
        .getById(resourceId)
        .update({
          AccessCount: (resource.AccessCount || 0) + 1,
          LastAccessedDate: new Date().toISOString(),
          LastAccessedBy: userEmail
        });

      // Log to audit trail
      await this.sp.web.lists
        .getByTitle(this.AUDIT_LOG_LIST)
        .items
        .add({
          Title: `Resource ${action}: ${resource.Title}`,
          ActionType: action === 'Viewed' ? 'ResourceAccessed' : `Resource${action}`,
          TargetResourceId: resourceId,
          TargetUser: userEmail,
          PerformedDate: new Date().toISOString(),
          IPAddress: ipAddress,
          UserAgent: deviceInfo,
          Result: 'Success'
        });
    } catch (error) {
      logger.error('SharedResourceService', `Failed to record access for resource ${resourceId}:`, error);
    }
  }

  /**
   * Get access log for a resource
   */
  public async getResourceAccessLog(resourceId: number): Promise<IAccessLogEntry[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.AUDIT_LOG_LIST)
        .items
        .filter(`TargetResourceId eq ${resourceId}`)
        .select('Id', 'Title', 'ActionType', 'TargetUser', 'PerformedDate', 'IPAddress', 'UserAgent')
        .orderBy('PerformedDate', false)
        .top(100)();

      return items.map(item => ({
        date: new Date(item.PerformedDate),
        userEmail: item.TargetUser,
        action: this.mapActionType(item.ActionType),
        ipAddress: item.IPAddress,
        deviceInfo: item.UserAgent
      }));
    } catch (error) {
      logger.error('SharedResourceService', `Failed to get access log for resource ${resourceId}:`, error);
      return [];
    }
  }

  // ============================================================================
  // ACKNOWLEDGMENT
  // ============================================================================

  /**
   * Record acknowledgment for a resource
   */
  public async recordAcknowledgment(resourceId: number, userEmail: string, acknowledged: boolean): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.SHARED_RESOURCES_LIST)
        .items
        .getById(resourceId)
        .update({
          AcknowledgmentStatus: acknowledged
            ? AcknowledgmentStatus.Acknowledged
            : AcknowledgmentStatus.Declined
        });

      logger.info('SharedResourceService', `Resource ${resourceId} ${acknowledged ? 'acknowledged' : 'declined'} by ${userEmail}`);
    } catch (error) {
      logger.error('SharedResourceService', `Failed to record acknowledgment for resource ${resourceId}:`, error);
      throw error;
    }
  }

  // ============================================================================
  // EXPIRATION MANAGEMENT
  // ============================================================================

  /**
   * Get expiring resources (within specified days)
   */
  public async getExpiringResources(withinDays: number = 30): Promise<ISharedResource[]> {
    try {
      const futureDate = new Date();
      futureDate.setDate(futureDate.getDate() + withinDays);

      const items = await this.sp.web.lists
        .getByTitle(this.SHARED_RESOURCES_LIST)
        .items
        .filter(`ExpirationDate le datetime'${futureDate.toISOString()}' and IsExpired eq false and Status eq 'Active'`)
        .select(
          'Id', 'Title', 'ResourceType', 'ResourceUrl',
          'SharedWithOrganizationId', 'SharingLevel', 'SharedDate',
          'ExpirationDate', 'IsExpired', 'Status'
        )
        .orderBy('ExpirationDate')
        .getAll();

      return items.map(item => this.mapToSharedResource(item));
    } catch (error) {
      logger.error('SharedResourceService', 'Failed to get expiring resources:', error);
      return [];
    }
  }

  /**
   * Extend resource expiration
   */
  public async extendExpiration(resourceId: number, additionalDays: number): Promise<void> {
    try {
      const resource = await this.getResourceById(resourceId);
      if (!resource) {
        throw new Error('Resource not found');
      }

      const currentExpiration = resource.ExpirationDate || new Date();
      const newExpiration = new Date(currentExpiration);
      newExpiration.setDate(newExpiration.getDate() + additionalDays);

      await this.sp.web.lists
        .getByTitle(this.SHARED_RESOURCES_LIST)
        .items
        .getById(resourceId)
        .update({
          ExpirationDate: newExpiration.toISOString(),
          IsExpired: false
        });

      logger.info('SharedResourceService', `Resource ${resourceId} expiration extended by ${additionalDays} days`);
    } catch (error) {
      logger.error('SharedResourceService', `Failed to extend expiration for resource ${resourceId}:`, error);
      throw error;
    }
  }

  /**
   * Mark expired resources
   */
  public async processExpiredResources(): Promise<number> {
    try {
      const now = new Date().toISOString();

      // Get expired but not yet marked resources
      const expiredItems = await this.sp.web.lists
        .getByTitle(this.SHARED_RESOURCES_LIST)
        .items
        .filter(`ExpirationDate le datetime'${now}' and IsExpired eq false and Status eq 'Active'`)
        .select('Id')
        .getAll();

      // Mark each as expired
      for (const item of expiredItems) {
        await this.sp.web.lists
          .getByTitle(this.SHARED_RESOURCES_LIST)
          .items
          .getById(item.Id)
          .update({
            IsExpired: true,
            Status: SharedResourceStatus.Expired
          });
      }

      logger.info('SharedResourceService', `Marked ${expiredItems.length} resources as expired`);
      return expiredItems.length;
    } catch (error) {
      logger.error('SharedResourceService', 'Failed to process expired resources:', error);
      return 0;
    }
  }

  // ============================================================================
  // HELPER METHODS
  // ============================================================================

  /**
   * Map SharePoint item to ISharedResource
   */
  private mapToSharedResource(item: any): ISharedResource {
    return {
      Id: item.Id,
      Title: item.Title,
      ResourceType: item.ResourceType as ResourceType || ResourceType.Document,
      ResourceUrl: item.ResourceUrl,
      ResourceId: item.ResourceId,
      SharedWithOrganizationId: item.SharedWithOrganizationId,
      SharedWithUsers: item.SharedWithUsers,
      SharingLevel: item.SharingLevel as SharingLevel || SharingLevel.View,
      SharedDate: new Date(item.SharedDate),
      SharedById: item.SharedBy?.Id,
      SharedBy: item.SharedBy ? { Id: item.SharedBy.Id, Title: item.SharedBy.Title, EMail: item.SharedBy.EMail } : undefined,
      ExpirationDate: item.ExpirationDate ? new Date(item.ExpirationDate) : undefined,
      IsExpired: item.IsExpired || false,
      DataClassification: item.DataClassification as DataClassification || DataClassification.Internal,
      RequiresAcknowledgment: item.RequiresAcknowledgment || false,
      AcknowledgmentStatus: item.AcknowledgmentStatus as AcknowledgmentStatus || AcknowledgmentStatus.Pending,
      RelatedModule: item.RelatedModule as RelatedModule || RelatedModule.Other,
      RelatedItemId: item.RelatedItemId,
      AccessCount: item.AccessCount || 0,
      LastAccessedDate: item.LastAccessedDate ? new Date(item.LastAccessedDate) : undefined,
      LastAccessedBy: item.LastAccessedBy,
      Status: item.Status as SharedResourceStatus || SharedResourceStatus.Active,
      Created: item.Created ? new Date(item.Created) : undefined,
      Modified: item.Modified ? new Date(item.Modified) : undefined
    };
  }

  /**
   * Map action type string to access log action
   */
  private mapActionType(actionType: string): 'Viewed' | 'Downloaded' | 'Edited' | 'Shared' {
    if (actionType.includes('Download')) return 'Downloaded';
    if (actionType.includes('Edit')) return 'Edited';
    if (actionType.includes('Share')) return 'Shared';
    return 'Viewed';
  }
}
