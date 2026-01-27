// @ts-nocheck
/**
 * AzureADActionHandler
 * Handles Azure AD / Entra ID operations within workflow execution
 *
 * Phase 2 Features:
 * - Add/Remove users from groups
 * - Enable/Disable user accounts
 * - Update user profiles
 *
 * Phase 4 Features (Automation & Intelligence):
 * - User provisioning (create new users)
 * - License assignment/revocation
 * - Bulk license operations
 *
 * Requires Graph API permissions:
 * - User.ReadWrite.All (for account operations and profile updates)
 * - GroupMember.ReadWrite.All (for group membership management)
 * - Directory.ReadWrite.All (alternative broader permission)
 * - LicenseAssignment.ReadWrite.All (for license management)
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import {
  IActionContext,
  IActionResult,
  IActionConfig,
  IAzureADProfileUpdate
} from '../../../models/IWorkflow';
import { logger } from '../../LoggingService';

/**
 * Result of Azure AD operation
 */
export interface IAzureADOperationResult {
  success: boolean;
  userId?: string;
  groupId?: string;
  error?: string;
  details?: Record<string, unknown>;
}

/**
 * New user provisioning request
 */
export interface IUserProvisioningRequest {
  displayName: string;
  userPrincipalName: string;
  mailNickname: string;
  password?: string;
  generatePassword?: boolean;
  forceChangePasswordNextSignIn?: boolean;
  accountEnabled?: boolean;
  department?: string;
  jobTitle?: string;
  officeLocation?: string;
  manager?: string;  // Manager's user ID or UPN
  usageLocation?: string;  // Required for license assignment
}

/**
 * User provisioning result
 */
export interface IUserProvisioningResult {
  success: boolean;
  userId?: string;
  userPrincipalName?: string;
  temporaryPassword?: string;
  error?: string;
}

/**
 * License assignment request
 */
export interface ILicenseAssignmentRequest {
  userId: string;
  skuIds: string[];  // License SKU IDs to assign
  disabledPlans?: string[];  // Service plans to disable within the license
}

/**
 * License revocation request
 */
export interface ILicenseRevocationRequest {
  userId: string;
  skuIds: string[];  // License SKU IDs to remove
}

/**
 * License operation result
 */
export interface ILicenseOperationResult {
  success: boolean;
  userId: string;
  licensesAssigned?: string[];
  licensesRevoked?: string[];
  error?: string;
}

/**
 * Available license information
 */
export interface IAvailableLicense {
  skuId: string;
  skuPartNumber: string;
  displayName: string;
  totalUnits: number;
  consumedUnits: number;
  availableUnits: number;
}

export class AzureADActionHandler {
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  // ============================================================================
  // GROUP MEMBERSHIP OPERATIONS
  // ============================================================================

  /**
   * Add a user to one or more Azure AD groups
   */
  public async addUserToGroup(config: IActionConfig, context: IActionContext): Promise<IActionResult> {
    try {
      const userId = this.resolveUserId(config, context);
      const groupIds = this.resolveGroupIds(config, context);

      if (!userId) {
        return { success: false, error: 'User ID not specified or could not be resolved' };
      }

      if (groupIds.length === 0) {
        return { success: false, error: 'No group IDs specified or could not be resolved' };
      }

      const graphClient = await this.getGraphClient();
      const results: IAzureADOperationResult[] = [];

      for (const groupId of groupIds) {
        try {
          // Check if user is already a member
          const isMember = await this.isUserMemberOfGroup(graphClient, userId, groupId);

          if (isMember) {
            logger.info('AzureADActionHandler', `User ${userId} is already a member of group ${groupId}`);
            results.push({ success: true, userId, groupId, details: { alreadyMember: true } });
            continue;
          }

          // Add user to group
          await graphClient
            .api(`/groups/${groupId}/members/$ref`)
            .post({
              '@odata.id': `https://graph.microsoft.com/v1.0/directoryObjects/${userId}`
            });

          logger.info('AzureADActionHandler', `Added user ${userId} to group ${groupId}`);
          results.push({ success: true, userId, groupId });
        } catch (error) {
          const errorMsg = error instanceof Error ? error.message : 'Unknown error';
          logger.error('AzureADActionHandler', `Failed to add user ${userId} to group ${groupId}`, error);
          results.push({ success: false, userId, groupId, error: errorMsg });
        }
      }

      const failedCount = results.filter(r => !r.success).length;

      return {
        success: failedCount === 0,
        error: failedCount > 0 ? `Failed to add user to ${failedCount} group(s)` : undefined,
        outputVariables: {
          groupsAdded: results.filter(r => r.success).map(r => r.groupId),
          groupsFailed: results.filter(r => !r.success).map(r => r.groupId),
          operationResults: results
        },
        nextAction: 'continue'
      };
    } catch (error) {
      logger.error('AzureADActionHandler', 'Error in addUserToGroup', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to add user to group'
      };
    }
  }

  /**
   * Remove a user from one or more Azure AD groups
   */
  public async removeUserFromGroup(config: IActionConfig, context: IActionContext): Promise<IActionResult> {
    try {
      const userId = this.resolveUserId(config, context);
      const groupIds = this.resolveGroupIds(config, context);

      if (!userId) {
        return { success: false, error: 'User ID not specified or could not be resolved' };
      }

      if (groupIds.length === 0) {
        return { success: false, error: 'No group IDs specified or could not be resolved' };
      }

      const graphClient = await this.getGraphClient();
      const results: IAzureADOperationResult[] = [];

      for (const groupId of groupIds) {
        try {
          // Check if user is a member
          const isMember = await this.isUserMemberOfGroup(graphClient, userId, groupId);

          if (!isMember) {
            logger.info('AzureADActionHandler', `User ${userId} is not a member of group ${groupId}`);
            results.push({ success: true, userId, groupId, details: { wasNotMember: true } });
            continue;
          }

          // Remove user from group
          await graphClient
            .api(`/groups/${groupId}/members/${userId}/$ref`)
            .delete();

          logger.info('AzureADActionHandler', `Removed user ${userId} from group ${groupId}`);
          results.push({ success: true, userId, groupId });
        } catch (error) {
          const errorMsg = error instanceof Error ? error.message : 'Unknown error';
          logger.error('AzureADActionHandler', `Failed to remove user ${userId} from group ${groupId}`, error);
          results.push({ success: false, userId, groupId, error: errorMsg });
        }
      }

      const failedCount = results.filter(r => !r.success).length;

      return {
        success: failedCount === 0,
        error: failedCount > 0 ? `Failed to remove user from ${failedCount} group(s)` : undefined,
        outputVariables: {
          groupsRemoved: results.filter(r => r.success).map(r => r.groupId),
          groupsFailed: results.filter(r => !r.success).map(r => r.groupId),
          operationResults: results
        },
        nextAction: 'continue'
      };
    } catch (error) {
      logger.error('AzureADActionHandler', 'Error in removeUserFromGroup', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to remove user from group'
      };
    }
  }

  // ============================================================================
  // ACCOUNT MANAGEMENT OPERATIONS
  // ============================================================================

  /**
   * Disable a user account (for leavers)
   */
  public async disableUserAccount(config: IActionConfig, context: IActionContext): Promise<IActionResult> {
    try {
      const userId = this.resolveUserId(config, context);

      if (!userId) {
        return { success: false, error: 'User ID not specified or could not be resolved' };
      }

      const graphClient = await this.getGraphClient();

      // Get current account status
      const user = await graphClient
        .api(`/users/${userId}`)
        .select('id,displayName,accountEnabled')
        .get();

      if (!user.accountEnabled) {
        logger.info('AzureADActionHandler', `User ${userId} account is already disabled`);
        return {
          success: true,
          outputVariables: {
            userId,
            userDisplayName: user.displayName,
            wasAlreadyDisabled: true
          },
          nextAction: 'continue'
        };
      }

      // Disable account
      await graphClient
        .api(`/users/${userId}`)
        .patch({ accountEnabled: false });

      logger.info('AzureADActionHandler', `Disabled account for user ${userId} (${user.displayName})`);

      return {
        success: true,
        outputVariables: {
          userId,
          userDisplayName: user.displayName,
          accountDisabled: true,
          disabledAt: new Date().toISOString()
        },
        nextAction: 'continue'
      };
    } catch (error) {
      logger.error('AzureADActionHandler', 'Error in disableUserAccount', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to disable user account'
      };
    }
  }

  /**
   * Enable a user account
   */
  public async enableUserAccount(config: IActionConfig, context: IActionContext): Promise<IActionResult> {
    try {
      const userId = this.resolveUserId(config, context);

      if (!userId) {
        return { success: false, error: 'User ID not specified or could not be resolved' };
      }

      const graphClient = await this.getGraphClient();

      // Get current account status
      const user = await graphClient
        .api(`/users/${userId}`)
        .select('id,displayName,accountEnabled')
        .get();

      if (user.accountEnabled) {
        logger.info('AzureADActionHandler', `User ${userId} account is already enabled`);
        return {
          success: true,
          outputVariables: {
            userId,
            userDisplayName: user.displayName,
            wasAlreadyEnabled: true
          },
          nextAction: 'continue'
        };
      }

      // Enable account
      await graphClient
        .api(`/users/${userId}`)
        .patch({ accountEnabled: true });

      logger.info('AzureADActionHandler', `Enabled account for user ${userId} (${user.displayName})`);

      return {
        success: true,
        outputVariables: {
          userId,
          userDisplayName: user.displayName,
          accountEnabled: true,
          enabledAt: new Date().toISOString()
        },
        nextAction: 'continue'
      };
    } catch (error) {
      logger.error('AzureADActionHandler', 'Error in enableUserAccount', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to enable user account'
      };
    }
  }

  // ============================================================================
  // PROFILE MANAGEMENT OPERATIONS
  // ============================================================================

  /**
   * Update user profile properties
   */
  public async updateUserProfile(config: IActionConfig, context: IActionContext): Promise<IActionResult> {
    try {
      const userId = this.resolveUserId(config, context);
      const profileUpdates = config.profileUpdates || [];

      if (!userId) {
        return { success: false, error: 'User ID not specified or could not be resolved' };
      }

      if (profileUpdates.length === 0) {
        return { success: false, error: 'No profile updates specified' };
      }

      const graphClient = await this.getGraphClient();

      // Build update payload
      const updatePayload: Record<string, string> = {};
      for (const update of profileUpdates) {
        const value = this.resolveProfileUpdateValue(update, context);
        if (value !== undefined) {
          updatePayload[update.property] = value;
        }
      }

      if (Object.keys(updatePayload).length === 0) {
        return { success: false, error: 'No valid profile updates to apply' };
      }

      // Get current user info for logging
      const user = await graphClient
        .api(`/users/${userId}`)
        .select('id,displayName')
        .get();

      // Apply updates
      await graphClient
        .api(`/users/${userId}`)
        .patch(updatePayload);

      logger.info('AzureADActionHandler', `Updated profile for user ${userId} (${user.displayName})`, {
        updatedProperties: Object.keys(updatePayload)
      });

      return {
        success: true,
        outputVariables: {
          userId,
          userDisplayName: user.displayName,
          updatedProperties: Object.keys(updatePayload),
          updatedAt: new Date().toISOString()
        },
        nextAction: 'continue'
      };
    } catch (error) {
      logger.error('AzureADActionHandler', 'Error in updateUserProfile', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to update user profile'
      };
    }
  }

  // ============================================================================
  // PHASE 4: USER PROVISIONING OPERATIONS
  // ============================================================================

  /**
   * Provision a new Azure AD user (for Joiner workflows)
   */
  public async provisionUser(config: IActionConfig, context: IActionContext): Promise<IActionResult> {
    try {
      const request = this.buildProvisioningRequest(config, context);

      if (!request.displayName || !request.userPrincipalName) {
        return { success: false, error: 'Display name and user principal name are required' };
      }

      const graphClient = await this.getGraphClient();

      // Generate password if not provided
      const password = request.password || (request.generatePassword !== false ? this.generateSecurePassword() : undefined);

      if (!password) {
        return { success: false, error: 'Password is required for user provisioning' };
      }

      // Build user creation payload
      const userPayload: Record<string, unknown> = {
        accountEnabled: request.accountEnabled !== false,
        displayName: request.displayName,
        mailNickname: request.mailNickname || request.displayName.replace(/\s/g, '').toLowerCase(),
        userPrincipalName: request.userPrincipalName,
        passwordProfile: {
          password,
          forceChangePasswordNextSignIn: request.forceChangePasswordNextSignIn !== false
        }
      };

      // Add optional properties
      if (request.department) userPayload.department = request.department;
      if (request.jobTitle) userPayload.jobTitle = request.jobTitle;
      if (request.officeLocation) userPayload.officeLocation = request.officeLocation;
      if (request.usageLocation) userPayload.usageLocation = request.usageLocation;

      // Create the user
      const createdUser = await graphClient
        .api('/users')
        .post(userPayload);

      logger.info('AzureADActionHandler', `Provisioned new user: ${createdUser.userPrincipalName} (${createdUser.id})`);

      // Set manager if specified
      if (request.manager) {
        try {
          await this.setUserManager(graphClient, createdUser.id, request.manager);
        } catch (managerError) {
          logger.warn('AzureADActionHandler', `Failed to set manager for ${createdUser.id}`, managerError);
        }
      }

      return {
        success: true,
        outputVariables: {
          provisionedUserId: createdUser.id,
          provisionedUserPrincipalName: createdUser.userPrincipalName,
          temporaryPassword: password,
          provisionedAt: new Date().toISOString()
        },
        nextAction: 'continue'
      };
    } catch (error) {
      logger.error('AzureADActionHandler', 'Error provisioning user', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to provision user'
      };
    }
  }

  /**
   * Provision user from process data (convenience method for workflows)
   */
  public async provisionUserFromProcess(
    displayName: string,
    email: string,
    department: string,
    jobTitle: string,
    managerId?: string,
    usageLocation: string = 'GB'
  ): Promise<IUserProvisioningResult> {
    try {
      const graphClient = await this.getGraphClient();
      const password = this.generateSecurePassword();
      const mailNickname = email.split('@')[0];

      const userPayload = {
        accountEnabled: true,
        displayName,
        mailNickname,
        userPrincipalName: email,
        department,
        jobTitle,
        usageLocation,
        passwordProfile: {
          password,
          forceChangePasswordNextSignIn: true
        }
      };

      const createdUser = await graphClient
        .api('/users')
        .post(userPayload);

      // Set manager if provided
      if (managerId) {
        try {
          await this.setUserManager(graphClient, createdUser.id, managerId);
        } catch {
          logger.warn('AzureADActionHandler', `Failed to set manager for new user ${createdUser.id}`);
        }
      }

      logger.info('AzureADActionHandler', `Provisioned user ${email} with ID ${createdUser.id}`);

      return {
        success: true,
        userId: createdUser.id,
        userPrincipalName: createdUser.userPrincipalName,
        temporaryPassword: password
      };
    } catch (error) {
      logger.error('AzureADActionHandler', 'Error in provisionUserFromProcess', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to provision user'
      };
    }
  }

  /**
   * Set user's manager
   */
  private async setUserManager(graphClient: MSGraphClientV3, userId: string, managerId: string): Promise<void> {
    await graphClient
      .api(`/users/${userId}/manager/$ref`)
      .put({
        '@odata.id': `https://graph.microsoft.com/v1.0/users/${managerId}`
      });

    logger.info('AzureADActionHandler', `Set manager ${managerId} for user ${userId}`);
  }

  // ============================================================================
  // PHASE 4: LICENSE MANAGEMENT OPERATIONS
  // ============================================================================

  /**
   * Assign licenses to a user
   */
  public async assignLicenses(config: IActionConfig, context: IActionContext): Promise<IActionResult> {
    try {
      const userId = this.resolveUserId(config, context);
      const skuIds = this.resolveLicenseSkuIds(config, context);

      if (!userId) {
        return { success: false, error: 'User ID not specified or could not be resolved' };
      }

      if (skuIds.length === 0) {
        return { success: false, error: 'No license SKU IDs specified' };
      }

      const graphClient = await this.getGraphClient();

      // Build license assignment payload
      const addLicenses = skuIds.map(skuId => ({
        skuId,
        disabledPlans: config.disabledPlans || []
      }));

      await graphClient
        .api(`/users/${userId}/assignLicense`)
        .post({
          addLicenses,
          removeLicenses: []
        });

      logger.info('AzureADActionHandler', `Assigned ${skuIds.length} license(s) to user ${userId}`);

      return {
        success: true,
        outputVariables: {
          licensesAssigned: skuIds,
          licenseAssignedAt: new Date().toISOString()
        },
        nextAction: 'continue'
      };
    } catch (error) {
      logger.error('AzureADActionHandler', 'Error assigning licenses', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to assign licenses'
      };
    }
  }

  /**
   * Revoke licenses from a user (for Leaver workflows)
   */
  public async revokeLicenses(config: IActionConfig, context: IActionContext): Promise<IActionResult> {
    try {
      const userId = this.resolveUserId(config, context);
      const skuIds = this.resolveLicenseSkuIds(config, context);

      if (!userId) {
        return { success: false, error: 'User ID not specified or could not be resolved' };
      }

      const graphClient = await this.getGraphClient();

      // If no specific SKUs provided, revoke all licenses
      let licensesToRevoke = skuIds;

      if (licensesToRevoke.length === 0 && config.revokeAllLicenses) {
        const currentLicenses = await this.getUserLicenses(userId);
        licensesToRevoke = currentLicenses.map(l => l.skuId);
      }

      if (licensesToRevoke.length === 0) {
        return {
          success: true,
          outputVariables: {
            licensesRevoked: [],
            message: 'No licenses to revoke'
          },
          nextAction: 'continue'
        };
      }

      await graphClient
        .api(`/users/${userId}/assignLicense`)
        .post({
          addLicenses: [],
          removeLicenses: licensesToRevoke
        });

      logger.info('AzureADActionHandler', `Revoked ${licensesToRevoke.length} license(s) from user ${userId}`);

      return {
        success: true,
        outputVariables: {
          licensesRevoked: licensesToRevoke,
          licenseRevokedAt: new Date().toISOString()
        },
        nextAction: 'continue'
      };
    } catch (error) {
      logger.error('AzureADActionHandler', 'Error revoking licenses', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to revoke licenses'
      };
    }
  }

  /**
   * Get user's current licenses
   */
  public async getUserLicenses(userId: string): Promise<Array<{ skuId: string; skuPartNumber: string }>> {
    try {
      const graphClient = await this.getGraphClient();

      const result = await graphClient
        .api(`/users/${userId}/licenseDetails`)
        .select('skuId,skuPartNumber')
        .get();

      return (result.value || []).map((license: Record<string, unknown>) => ({
        skuId: license.skuId as string,
        skuPartNumber: license.skuPartNumber as string
      }));
    } catch (error) {
      logger.error('AzureADActionHandler', `Error getting licenses for user ${userId}`, error);
      return [];
    }
  }

  /**
   * Get available licenses in the tenant
   */
  public async getAvailableLicenses(): Promise<IAvailableLicense[]> {
    try {
      const graphClient = await this.getGraphClient();

      const result = await graphClient
        .api('/subscribedSkus')
        .select('skuId,skuPartNumber,prepaidUnits,consumedUnits')
        .get();

      return (result.value || []).map((sku: Record<string, unknown>) => {
        const prepaidUnits = sku.prepaidUnits as { enabled?: number } || {};
        const totalUnits = prepaidUnits.enabled || 0;
        const consumedUnits = (sku.consumedUnits as number) || 0;

        return {
          skuId: sku.skuId as string,
          skuPartNumber: sku.skuPartNumber as string,
          displayName: this.getLicenseDisplayName(sku.skuPartNumber as string),
          totalUnits,
          consumedUnits,
          availableUnits: totalUnits - consumedUnits
        };
      });
    } catch (error) {
      logger.error('AzureADActionHandler', 'Error getting available licenses', error);
      return [];
    }
  }

  /**
   * Assign standard license package to a new user
   * Convenience method for Joiner workflows
   */
  public async assignStandardLicensePackage(
    userId: string,
    packageType: 'basic' | 'standard' | 'premium' = 'standard'
  ): Promise<ILicenseOperationResult> {
    try {
      const graphClient = await this.getGraphClient();

      // Get available licenses
      const availableLicenses = await this.getAvailableLicenses();

      // Define license packages (these SKU IDs are examples - actual IDs depend on tenant)
      const licensePackages: Record<string, string[]> = {
        basic: ['EXCHANGESTANDARD'],  // Exchange Online Plan 1
        standard: ['O365_BUSINESS_ESSENTIALS', 'EXCHANGESTANDARD'],  // M365 Business Basic + Exchange
        premium: ['SPE_E3', 'ENTERPRISEPACK']  // M365 E3
      };

      const requestedSkuPartNumbers = licensePackages[packageType] || licensePackages.standard;

      // Find matching SKU IDs from available licenses
      const skuIdsToAssign = availableLicenses
        .filter(l => requestedSkuPartNumbers.includes(l.skuPartNumber) && l.availableUnits > 0)
        .map(l => l.skuId);

      if (skuIdsToAssign.length === 0) {
        return {
          success: false,
          userId,
          error: `No available licenses for package type: ${packageType}`
        };
      }

      await graphClient
        .api(`/users/${userId}/assignLicense`)
        .post({
          addLicenses: skuIdsToAssign.map(skuId => ({ skuId, disabledPlans: [] })),
          removeLicenses: []
        });

      logger.info('AzureADActionHandler', `Assigned ${packageType} license package to user ${userId}`);

      return {
        success: true,
        userId,
        licensesAssigned: skuIdsToAssign
      };
    } catch (error) {
      logger.error('AzureADActionHandler', 'Error assigning license package', error);
      return {
        success: false,
        userId,
        error: error instanceof Error ? error.message : 'Failed to assign license package'
      };
    }
  }

  // ============================================================================
  // UTILITY METHODS
  // ============================================================================

  /**
   * Get user by email address
   */
  public async getUserByEmail(email: string): Promise<{ id: string; displayName: string } | null> {
    try {
      const graphClient = await this.getGraphClient();

      const users = await graphClient
        .api('/users')
        .filter(`mail eq '${email}' or userPrincipalName eq '${email}'`)
        .select('id,displayName,mail,userPrincipalName')
        .get();

      if (users.value && users.value.length > 0) {
        return {
          id: users.value[0].id,
          displayName: users.value[0].displayName
        };
      }

      return null;
    } catch (error) {
      logger.error('AzureADActionHandler', `Error looking up user by email: ${email}`, error);
      return null;
    }
  }

  /**
   * Get group by name
   */
  public async getGroupByName(groupName: string): Promise<{ id: string; displayName: string } | null> {
    try {
      const graphClient = await this.getGraphClient();

      const groups = await graphClient
        .api('/groups')
        .filter(`displayName eq '${groupName}'`)
        .select('id,displayName')
        .get();

      if (groups.value && groups.value.length > 0) {
        return {
          id: groups.value[0].id,
          displayName: groups.value[0].displayName
        };
      }

      return null;
    } catch (error) {
      logger.error('AzureADActionHandler', `Error looking up group by name: ${groupName}`, error);
      return null;
    }
  }

  /**
   * Get user's current group memberships
   */
  public async getUserGroups(userId: string): Promise<Array<{ id: string; displayName: string }>> {
    try {
      const graphClient = await this.getGraphClient();

      const groups = await graphClient
        .api(`/users/${userId}/memberOf`)
        .select('id,displayName')
        .get();

      return (groups.value || [])
        .filter((g: Record<string, unknown>) => g['@odata.type'] === '#microsoft.graph.group')
        .map((g: Record<string, unknown>) => ({
          id: g.id as string,
          displayName: g.displayName as string
        }));
    } catch (error) {
      logger.error('AzureADActionHandler', `Error getting groups for user: ${userId}`, error);
      return [];
    }
  }

  // ============================================================================
  // PRIVATE HELPER METHODS
  // ============================================================================

  private async getGraphClient(): Promise<MSGraphClientV3> {
    return this.context.msGraphClientFactory.getClient('3');
  }

  private async isUserMemberOfGroup(graphClient: MSGraphClientV3, userId: string, groupId: string): Promise<boolean> {
    try {
      const result = await graphClient
        .api(`/groups/${groupId}/members`)
        .filter(`id eq '${userId}'`)
        .select('id')
        .get();

      return result.value && result.value.length > 0;
    } catch {
      return false;
    }
  }

  private resolveUserId(config: IActionConfig, context: IActionContext): string | undefined {
    // Direct user ID
    if (config.userId) {
      return config.userId;
    }

    // From process field
    if (config.userIdField) {
      const fieldValue = context.process[config.userIdField];
      if (typeof fieldValue === 'string') {
        return fieldValue;
      }
    }

    // From user email field (needs lookup - return email for now, caller should handle)
    if (config.userEmailField) {
      const email = context.process[config.userEmailField];
      if (typeof email === 'string') {
        // Note: In a real implementation, this would need to be async to look up the user
        // For now, we'll use the email directly if it looks like an Azure AD object ID
        return email;
      }
    }

    // Direct user email
    if (config.userEmail) {
      return config.userEmail;
    }

    // Default to employee's Azure AD ID from process
    const employeeAzureId = context.process['EmployeeAzureId'] || context.process['EmployeeEntraId'];
    if (typeof employeeAzureId === 'string') {
      return employeeAzureId;
    }

    return undefined;
  }

  private resolveGroupIds(config: IActionConfig, context: IActionContext): string[] {
    const groupIds: string[] = [];

    // Direct group ID
    if (config.groupId) {
      groupIds.push(config.groupId);
    }

    // Multiple group IDs
    if (config.groupIds && config.groupIds.length > 0) {
      groupIds.push(...config.groupIds);
    }

    // From process field
    if (config.groupIdField) {
      const fieldValue = context.process[config.groupIdField];
      if (typeof fieldValue === 'string') {
        groupIds.push(fieldValue);
      } else if (Array.isArray(fieldValue)) {
        groupIds.push(...fieldValue.filter((v): v is string => typeof v === 'string'));
      }
    }

    // Note: Group names would need async lookup - not supported in this sync method
    // Use groupId or groupIds instead

    return Array.from(new Set(groupIds)); // Remove duplicates
  }

  private resolveProfileUpdateValue(update: IAzureADProfileUpdate, context: IActionContext): string | undefined {
    // Direct value
    if (update.value !== undefined) {
      return update.value;
    }

    // From process field
    if (update.valueField) {
      const fieldValue = context.process[update.valueField];
      if (typeof fieldValue === 'string') {
        return fieldValue;
      }
      if (fieldValue !== undefined && fieldValue !== null) {
        return String(fieldValue);
      }
    }

    return undefined;
  }

  /**
   * Build user provisioning request from config and context
   */
  private buildProvisioningRequest(config: IActionConfig, context: IActionContext): IUserProvisioningRequest {
    return {
      displayName: config.displayName || context.process['EmployeeName'] as string || '',
      userPrincipalName: config.userPrincipalName || context.process['EmployeeEmail'] as string || '',
      mailNickname: config.mailNickname || (context.process['EmployeeEmail'] as string)?.split('@')[0] || '',
      password: config.password,
      generatePassword: config.generatePassword !== false,
      forceChangePasswordNextSignIn: config.forceChangePasswordNextSignIn !== false,
      accountEnabled: config.accountEnabled !== false,
      department: config.department || context.process['Department'] as string,
      jobTitle: config.jobTitle || context.process['JobTitle'] as string,
      officeLocation: config.officeLocation || context.process['Location'] as string,
      manager: config.manager || context.process['ManagerId'] as string,
      usageLocation: config.usageLocation || context.process['UsageLocation'] as string || 'GB'
    };
  }

  /**
   * Generate a secure random password
   */
  private generateSecurePassword(): string {
    const length = 16;
    const uppercase = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    const lowercase = 'abcdefghijklmnopqrstuvwxyz';
    const numbers = '0123456789';
    const symbols = '!@#$%^&*()_+-=[]{}|;:,.<>?';
    const allChars = uppercase + lowercase + numbers + symbols;

    let password = '';

    // Ensure at least one character from each category
    password += uppercase[Math.floor(Math.random() * uppercase.length)];
    password += lowercase[Math.floor(Math.random() * lowercase.length)];
    password += numbers[Math.floor(Math.random() * numbers.length)];
    password += symbols[Math.floor(Math.random() * symbols.length)];

    // Fill the rest randomly
    for (let i = password.length; i < length; i++) {
      password += allChars[Math.floor(Math.random() * allChars.length)];
    }

    // Shuffle the password
    return password.split('').sort(() => Math.random() - 0.5).join('');
  }

  /**
   * Resolve license SKU IDs from config and context
   */
  private resolveLicenseSkuIds(config: IActionConfig, context: IActionContext): string[] {
    const skuIds: string[] = [];

    // Direct SKU ID
    if (config.licenseSkuId) {
      skuIds.push(config.licenseSkuId);
    }

    // Multiple SKU IDs
    if (config.licenseSkuIds && config.licenseSkuIds.length > 0) {
      skuIds.push(...config.licenseSkuIds);
    }

    // From process field
    if (config.licenseSkuIdField) {
      const fieldValue = context.process[config.licenseSkuIdField];
      if (typeof fieldValue === 'string') {
        skuIds.push(fieldValue);
      } else if (Array.isArray(fieldValue)) {
        skuIds.push(...fieldValue.filter((v): v is string => typeof v === 'string'));
      }
    }

    return Array.from(new Set(skuIds)); // Remove duplicates
  }

  /**
   * Get human-readable license display name
   */
  private getLicenseDisplayName(skuPartNumber: string): string {
    const licenseNames: Record<string, string> = {
      'ENTERPRISEPACK': 'Office 365 E3',
      'ENTERPRISEPREMIUM': 'Office 365 E5',
      'SPE_E3': 'Microsoft 365 E3',
      'SPE_E5': 'Microsoft 365 E5',
      'O365_BUSINESS_ESSENTIALS': 'Microsoft 365 Business Basic',
      'O365_BUSINESS_PREMIUM': 'Microsoft 365 Business Standard',
      'SMB_BUSINESS_PREMIUM': 'Microsoft 365 Business Premium',
      'EXCHANGESTANDARD': 'Exchange Online (Plan 1)',
      'EXCHANGEENTERPRISE': 'Exchange Online (Plan 2)',
      'POWER_BI_STANDARD': 'Power BI (free)',
      'POWER_BI_PRO': 'Power BI Pro',
      'PROJECTPROFESSIONAL': 'Project Plan 3',
      'VISIOCLIENT': 'Visio Plan 2',
      'TEAMS_EXPLORATORY': 'Microsoft Teams Exploratory',
      'FLOW_FREE': 'Power Automate Free',
      'POWERAPPS_VIRAL': 'Power Apps Plan 2 Trial',
      'AAD_PREMIUM': 'Azure AD Premium P1',
      'AAD_PREMIUM_P2': 'Azure AD Premium P2',
      'EMS': 'Enterprise Mobility + Security E3',
      'EMSPREMIUM': 'Enterprise Mobility + Security E5'
    };

    return licenseNames[skuPartNumber] || skuPartNumber;
  }
}

export default AzureADActionHandler;
