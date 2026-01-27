// @ts-nocheck
// IT Provisioning Service
// Orchestrates IT provisioning/deprovisioning for JML processes
// Integrates with Azure AD/Entra ID, M365 Licenses, Groups, and Teams

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { GraphFI } from '@pnp/graph';
import '@pnp/graph/users';
import '@pnp/graph/groups';
import '@pnp/graph/teams';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  IProvisioningRequest,
  IProvisioningResult,
  IProvisioningStep,
  IProvisioningAuditLog,
  IEntraUserCreateRequest,
  IEntraUser,
  IEntraUserUpdateRequest,
  ILicenseAssignmentRequest,
  IGroupMembershipChange,
  ITeamMembershipChange,
  ProvisioningActionType,
  ProvisioningStatus,
  IProvisioningConfig,
  IDepartmentProvisioningConfig
} from '../models/IITProvisioning';
import { IJmlProcess, ProcessType } from '../models';
import { logger } from './LoggingService';
import { EmailQueueService, EmailPriority } from './EmailQueueService';
import { v4 as uuidv4 } from 'uuid';

export class ITProvisioningService {
  private sp: SPFI;
  private graph: GraphFI;
  private context: WebPartContext;
  private emailQueueService: EmailQueueService;

  private readonly PROVISIONING_LOG_LIST = 'JML_ITProvisioningLog';
  private readonly PROVISIONING_AUDIT_LIST = 'JML_ITProvisioningAuditLog';
  private readonly PROVISIONING_CONFIG_LIST = 'JML_ITProvisioningConfig';

  // Default configuration (can be overridden by SharePoint config)
  private defaultConfig: IProvisioningConfig = {
    tenantId: '',
    defaultUsageLocation: 'US',
    passwordLength: 16,
    forcePasswordChange: true,
    sendWelcomeEmail: true,
    departmentConfigs: [],
    roleConfigs: [],
    leaverGracePeriodDays: 30,
    autoDisableOnLeave: true
  };

  constructor(sp: SPFI, graph: GraphFI, context: WebPartContext) {
    this.sp = sp;
    this.graph = graph;
    this.context = context;
    this.emailQueueService = new EmailQueueService(sp);
  }

  // ============================================================================
  // Main Provisioning Workflows
  // ============================================================================

  /**
   * Provision a new Joiner - create account, assign licenses, add to groups
   */
  public async provisionJoiner(
    process: IJmlProcess,
    config?: Partial<IProvisioningConfig>
  ): Promise<IProvisioningResult> {
    const mergedConfig = { ...this.defaultConfig, ...config };
    const requestId = await this.createProvisioningRequest(process, 'Joiner');

    const steps: IProvisioningStep[] = [];
    const completedSteps: IProvisioningStep[] = [];

    try {
      // Step 1: Create user in Entra ID
      const createUserStep = this.createStep('Create User Account', 'CreateUser', process.EmployeeEmail);
      steps.push(createUserStep);

      const userRequest = this.buildUserCreateRequest(process, mergedConfig);
      const user = await this.createEntraUser(userRequest);

      createUserStep.status = 'Completed';
      createUserStep.responsePayload = JSON.stringify({ id: user.id, upn: user.userPrincipalName });
      createUserStep.completedAt = new Date();
      completedSteps.push(createUserStep);

      await this.updateProvisioningStep(requestId, createUserStep);

      // Step 2: Assign licenses based on department/role
      const departmentConfig = this.getDepartmentConfig(process.Department, mergedConfig);

      // Log department config status for visibility
      if (departmentConfig.defaultLicenses.length === 0 &&
          departmentConfig.securityGroups.length === 0 &&
          departmentConfig.teams.length === 0) {
        logger.warn(
          'ITProvisioningService',
          `Department "${process.Department}" has no licenses, groups, or teams configured. User account created but no access assigned.`
        );
      }

      if (departmentConfig.defaultLicenses.length > 0) {
        const licenseStep = this.createStep('Assign Licenses', 'AssignLicense', user.id);
        steps.push(licenseStep);

        await this.assignLicenses(user.id, departmentConfig.defaultLicenses);

        licenseStep.status = 'Completed';
        licenseStep.responsePayload = JSON.stringify({ licenses: departmentConfig.defaultLicenses });
        licenseStep.completedAt = new Date();
        completedSteps.push(licenseStep);

        await this.updateProvisioningStep(requestId, licenseStep);
      }

      // Step 3: Add to security groups
      if (departmentConfig.securityGroups.length > 0) {
        for (const groupId of departmentConfig.securityGroups) {
          const groupStep = this.createStep(`Add to Security Group`, 'AddToGroup', groupId);
          steps.push(groupStep);

          await this.addUserToGroup(user.id, groupId);

          groupStep.status = 'Completed';
          groupStep.completedAt = new Date();
          completedSteps.push(groupStep);

          await this.updateProvisioningStep(requestId, groupStep);
        }
      }

      // Step 4: Add to Teams
      if (departmentConfig.teams.length > 0) {
        for (const teamId of departmentConfig.teams) {
          const teamStep = this.createStep(`Add to Team`, 'AddToTeam', teamId);
          steps.push(teamStep);

          await this.addUserToTeam(user.id, teamId);

          teamStep.status = 'Completed';
          teamStep.completedAt = new Date();
          completedSteps.push(teamStep);

          await this.updateProvisioningStep(requestId, teamStep);
        }
      }

      // Step 5: Send welcome email
      if (mergedConfig.sendWelcomeEmail) {
        const emailStep = this.createStep('Send Welcome Email', 'SendWelcomeEmail', user.mail || user.userPrincipalName);
        steps.push(emailStep);

        await this.sendWelcomeEmail(user, userRequest.passwordProfile.password, process);

        emailStep.status = 'Completed';
        emailStep.completedAt = new Date();
        completedSteps.push(emailStep);

        await this.updateProvisioningStep(requestId, emailStep);
      }

      // Update request as completed
      await this.updateProvisioningRequestStatus(requestId, 'Completed');

      // Create audit log
      await this.createAuditLog({
        ProcessId: process.Id!,
        RequestId: requestId,
        EmployeeId: process.EmployeeID,
        EmployeeName: process.EmployeeName,
        ActionType: 'CreateUser',
        ActionStatus: 'Success',
        TargetResource: user.id,
        TargetResourceName: user.userPrincipalName,
        ExecutedById: this.context.pageContext.legacyPageContext?.userId,
        ExecutedByName: this.context.pageContext.user?.displayName,
        ExecutedAt: new Date()
      });

      return {
        success: true,
        requestId,
        status: 'Completed',
        completedSteps: completedSteps.length,
        totalSteps: steps.length,
        userCreated: user,
        licensesAssigned: departmentConfig.defaultLicenses,
        groupsAdded: departmentConfig.securityGroups,
        teamsAdded: departmentConfig.teams
      };

    } catch (error) {
      logger.error('ITProvisioningService', 'Error provisioning Joiner:', error);

      const errorMessage = error instanceof Error ? error.message : 'Unknown error';

      // Rollback completed steps
      await this.rollbackSteps(completedSteps.reverse());

      await this.updateProvisioningRequestStatus(requestId, 'Failed');

      // Create audit log for failure
      await this.createAuditLog({
        ProcessId: process.Id!,
        RequestId: requestId,
        EmployeeName: process.EmployeeName,
        ActionType: 'CreateUser',
        ActionStatus: 'Failed',
        TargetResource: process.EmployeeEmail,
        ErrorDetails: errorMessage,
        ExecutedById: this.context.pageContext.legacyPageContext?.userId,
        ExecutedAt: new Date()
      });

      // Send admin notification for provisioning failure
      await this.notifyAdminsOfFailure(
        'Joiner',
        process,
        errorMessage,
        steps.find(s => s.status === 'Failed')?.name || 'Unknown step',
        requestId
      );

      return {
        success: false,
        requestId,
        status: 'Failed',
        completedSteps: completedSteps.length,
        totalSteps: steps.length,
        failedStep: steps.find(s => s.status === 'Failed')?.name,
        errorMessage
      };
    }
  }

  /**
   * Provision a Mover - update groups, licenses, and access
   */
  public async provisionMover(
    process: IJmlProcess,
    previousDepartment: string,
    config?: Partial<IProvisioningConfig>
  ): Promise<IProvisioningResult> {
    const mergedConfig = { ...this.defaultConfig, ...config };
    const requestId = await this.createProvisioningRequest(process, 'Mover');

    const steps: IProvisioningStep[] = [];
    const completedSteps: IProvisioningStep[] = [];

    try {
      // Get user by email
      const user = await this.getEntraUserByEmail(process.EmployeeEmail);
      if (!user) {
        throw new Error(`User not found: ${process.EmployeeEmail}`);
      }

      // Step 1: Update user profile
      const updateStep = this.createStep('Update User Profile', 'UpdateUser', user.id);
      steps.push(updateStep);

      await this.updateEntraUser(user.id, {
        department: process.Department,
        jobTitle: process.JobTitle,
        officeLocation: process.Location
      });

      updateStep.status = 'Completed';
      updateStep.completedAt = new Date();
      completedSteps.push(updateStep);

      // Step 2: Remove from old department groups
      const oldDeptConfig = this.getDepartmentConfig(previousDepartment, mergedConfig);
      if (oldDeptConfig) {
        for (const groupId of oldDeptConfig.securityGroups) {
          const removeStep = this.createStep('Remove from Old Group', 'RemoveFromGroup', groupId);
          steps.push(removeStep);

          await this.removeUserFromGroup(user.id, groupId);

          removeStep.status = 'Completed';
          removeStep.completedAt = new Date();
          completedSteps.push(removeStep);
        }
      }

      // Step 3: Add to new department groups
      const newDeptConfig = this.getDepartmentConfig(process.Department, mergedConfig);
      if (newDeptConfig) {
        for (const groupId of newDeptConfig.securityGroups) {
          const addStep = this.createStep('Add to New Group', 'AddToGroup', groupId);
          steps.push(addStep);

          await this.addUserToGroup(user.id, groupId);

          addStep.status = 'Completed';
          addStep.completedAt = new Date();
          completedSteps.push(addStep);
        }

        // Step 4: License rotation - remove old, add new
        const oldLicenses = oldDeptConfig?.defaultLicenses || [];
        const newLicenses = newDeptConfig.defaultLicenses || [];

        // Remove licenses that are in old dept but not in new dept
        const licensesToRemove = oldLicenses.filter(lic => !newLicenses.includes(lic));
        if (licensesToRemove.length > 0) {
          const removeLicenseStep = this.createStep('Remove Old Licenses', 'RemoveLicense', user.id);
          steps.push(removeLicenseStep);

          try {
            await this.removeLicenses(user.id, licensesToRemove);
            removeLicenseStep.status = 'Completed';
            removeLicenseStep.responsePayload = JSON.stringify({ removed: licensesToRemove });
            removeLicenseStep.completedAt = new Date();
            completedSteps.push(removeLicenseStep);
            logger.info('ITProvisioningService', `Removed licenses for Mover: ${licensesToRemove.join(', ')}`);
          } catch (licError) {
            // Log but continue - license removal failure shouldn't block the move
            logger.warn('ITProvisioningService', 'Failed to remove old licenses:', licError);
            removeLicenseStep.status = 'Failed';
            removeLicenseStep.errorMessage = licError instanceof Error ? licError.message : 'License removal failed';
          }
        }

        // Add licenses that are in new dept but not in old dept
        const licensesToAdd = newLicenses.filter(lic => !oldLicenses.includes(lic));
        if (licensesToAdd.length > 0) {
          const addLicenseStep = this.createStep('Assign New Licenses', 'AssignLicense', user.id);
          steps.push(addLicenseStep);

          try {
            await this.assignLicenses(user.id, licensesToAdd);
            addLicenseStep.status = 'Completed';
            addLicenseStep.responsePayload = JSON.stringify({ added: licensesToAdd });
            addLicenseStep.completedAt = new Date();
            completedSteps.push(addLicenseStep);
            logger.info('ITProvisioningService', `Assigned new licenses for Mover: ${licensesToAdd.join(', ')}`);
          } catch (licError) {
            logger.warn('ITProvisioningService', 'Failed to assign new licenses:', licError);
            addLicenseStep.status = 'Failed';
            addLicenseStep.errorMessage = licError instanceof Error ? licError.message : 'License assignment failed';
          }
        }

        // Step 5: Update Teams memberships
        if (oldDeptConfig) {
          for (const teamId of oldDeptConfig.teams) {
            if (!newDeptConfig.teams.includes(teamId)) {
              await this.removeUserFromTeam(user.id, teamId);
            }
          }
        }

        for (const teamId of newDeptConfig.teams) {
          if (!oldDeptConfig?.teams.includes(teamId)) {
            await this.addUserToTeam(user.id, teamId);
          }
        }
      }

      await this.updateProvisioningRequestStatus(requestId, 'Completed');

      return {
        success: true,
        requestId,
        status: 'Completed',
        completedSteps: completedSteps.length,
        totalSteps: steps.length
      };

    } catch (error) {
      logger.error('ITProvisioningService', 'Error provisioning Mover:', error);
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';

      await this.updateProvisioningRequestStatus(requestId, 'Failed');

      // Send admin notification for provisioning failure
      await this.notifyAdminsOfFailure(
        'Mover',
        process,
        errorMessage,
        steps.find(s => s.status === 'Failed')?.name || 'Unknown step',
        requestId
      );

      return {
        success: false,
        requestId,
        status: 'Failed',
        completedSteps: completedSteps.length,
        totalSteps: steps.length,
        errorMessage
      };
    }
  }

  /**
   * Deprovision a Leaver - disable account, revoke access
   */
  public async deprovisionLeaver(
    process: IJmlProcess,
    config?: Partial<IProvisioningConfig>
  ): Promise<IProvisioningResult> {
    const mergedConfig = { ...this.defaultConfig, ...config };
    const requestId = await this.createProvisioningRequest(process, 'Leaver');

    const steps: IProvisioningStep[] = [];
    const completedSteps: IProvisioningStep[] = [];

    try {
      // Get user by email
      const user = await this.getEntraUserByEmail(process.EmployeeEmail);
      if (!user) {
        throw new Error(`User not found: ${process.EmployeeEmail}`);
      }

      // Step 1: Disable account (block sign-in)
      if (mergedConfig.autoDisableOnLeave) {
        const disableStep = this.createStep('Disable User Account', 'DisableUser', user.id);
        steps.push(disableStep);

        await this.disableEntraUser(user.id);

        disableStep.status = 'Completed';
        disableStep.completedAt = new Date();
        completedSteps.push(disableStep);
      }

      // Step 2: Revoke all sessions
      const revokeStep = this.createStep('Revoke Sessions', 'RevokeSession', user.id);
      steps.push(revokeStep);

      await this.revokeUserSessions(user.id);

      revokeStep.status = 'Completed';
      revokeStep.completedAt = new Date();
      completedSteps.push(revokeStep);

      // Step 3: Remove from all groups (get current memberships first)
      const userGroups = await this.getUserGroupMemberships(user.id);
      for (const group of userGroups) {
        const removeStep = this.createStep(`Remove from Group: ${group.displayName}`, 'RemoveFromGroup', group.id);
        steps.push(removeStep);

        await this.removeUserFromGroup(user.id, group.id);

        removeStep.status = 'Completed';
        removeStep.completedAt = new Date();
        completedSteps.push(removeStep);
      }

      // Step 4: Remove from all Teams
      const userTeams = await this.getUserTeamMemberships(user.id);
      for (const team of userTeams) {
        const removeStep = this.createStep(`Remove from Team: ${team.displayName}`, 'RemoveFromTeam', team.id);
        steps.push(removeStep);

        await this.removeUserFromTeam(user.id, team.id);

        removeStep.status = 'Completed';
        removeStep.completedAt = new Date();
        completedSteps.push(removeStep);
      }

      // Step 5: Schedule license removal (grace period)
      // Note: Actual license removal should be scheduled via Power Automate or timer job
      const licenseStep = this.createStep('Schedule License Removal', 'RemoveLicense', user.id);
      licenseStep.status = 'Completed';
      licenseStep.responsePayload = JSON.stringify({
        scheduledFor: new Date(Date.now() + mergedConfig.leaverGracePeriodDays * 24 * 60 * 60 * 1000)
      });
      steps.push(licenseStep);

      await this.updateProvisioningRequestStatus(requestId, 'Completed');

      // Create audit log
      await this.createAuditLog({
        ProcessId: process.Id!,
        RequestId: requestId,
        EmployeeName: process.EmployeeName,
        ActionType: 'DisableUser',
        ActionStatus: 'Success',
        TargetResource: user.id,
        TargetResourceName: user.userPrincipalName,
        ExecutedById: this.context.pageContext.legacyPageContext?.userId,
        ExecutedAt: new Date()
      });

      return {
        success: true,
        requestId,
        status: 'Completed',
        completedSteps: completedSteps.length,
        totalSteps: steps.length
      };

    } catch (error) {
      logger.error('ITProvisioningService', 'Error deprovisioning Leaver:', error);
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';

      await this.updateProvisioningRequestStatus(requestId, 'Failed');

      // Send admin notification for deprovisioning failure
      await this.notifyAdminsOfFailure(
        'Leaver',
        process,
        errorMessage,
        steps.find(s => s.status === 'Failed')?.name || 'Unknown step',
        requestId
      );

      return {
        success: false,
        requestId,
        status: 'Failed',
        completedSteps: completedSteps.length,
        totalSteps: steps.length,
        errorMessage
      };
    }
  }

  // ============================================================================
  // Entra ID / Azure AD Operations
  // ============================================================================

  /**
   * Create a new user in Entra ID
   */
  public async createEntraUser(request: IEntraUserCreateRequest): Promise<IEntraUser> {
    try {
      const user = await (this.graph.users as any).add({
        accountEnabled: request.accountEnabled,
        displayName: request.displayName,
        givenName: request.givenName,
        surname: request.surname,
        mailNickname: request.mailNickname,
        userPrincipalName: request.userPrincipalName,
        passwordProfile: {
          password: request.passwordProfile.password,
          forceChangePasswordNextSignIn: request.passwordProfile.forceChangePasswordNextSignIn
        },
        usageLocation: request.usageLocation,
        jobTitle: request.jobTitle,
        department: request.department,
        officeLocation: request.officeLocation,
        mobilePhone: request.mobilePhone,
        companyName: request.companyName,
        employeeId: request.employeeId
      });

      // Set manager if provided
      if (request.managerId) {
        await this.setUserManager(user.id, request.managerId);
      }

      logger.info('ITProvisioningService', `Created user: ${user.userPrincipalName}`);

      return {
        id: user.id,
        displayName: user.displayName,
        givenName: user.givenName,
        surname: user.surname,
        userPrincipalName: user.userPrincipalName,
        mail: user.mail,
        jobTitle: user.jobTitle,
        department: user.department,
        accountEnabled: user.accountEnabled
      };
    } catch (error) {
      logger.error('ITProvisioningService', 'Error creating Entra user:', error);
      throw error;
    }
  }

  /**
   * Get user by email/UPN
   */
  public async getEntraUserByEmail(email: string): Promise<IEntraUser | null> {
    try {
      const user = await this.graph.users.getById(email)();
      return {
        id: user.id!,
        displayName: user.displayName!,
        userPrincipalName: user.userPrincipalName!,
        mail: user.mail || undefined,
        jobTitle: user.jobTitle || undefined,
        department: user.department || undefined,
        accountEnabled: user.accountEnabled || false
      };
    } catch (error: any) {
      if (error?.status === 404 || error?.message?.includes('does not exist')) {
        return null;
      }
      throw error;
    }
  }

  /**
   * Update user in Entra ID
   */
  public async updateEntraUser(userId: string, updates: IEntraUserUpdateRequest): Promise<void> {
    try {
      await this.graph.users.getById(userId).update(updates as any);
      logger.info('ITProvisioningService', `Updated user: ${userId}`);
    } catch (error) {
      logger.error('ITProvisioningService', 'Error updating Entra user:', error);
      throw error;
    }
  }

  /**
   * Disable user account (block sign-in)
   */
  public async disableEntraUser(userId: string): Promise<void> {
    try {
      await this.graph.users.getById(userId).update({
        accountEnabled: false
      } as any);
      logger.info('ITProvisioningService', `Disabled user: ${userId}`);
    } catch (error) {
      logger.error('ITProvisioningService', 'Error disabling Entra user:', error);
      throw error;
    }
  }

  /**
   * Enable user account
   */
  public async enableEntraUser(userId: string): Promise<void> {
    try {
      await this.graph.users.getById(userId).update({
        accountEnabled: true
      } as any);
      logger.info('ITProvisioningService', `Enabled user: ${userId}`);
    } catch (error) {
      logger.error('ITProvisioningService', 'Error enabling Entra user:', error);
      throw error;
    }
  }

  /**
   * Set user's manager
   */
  public async setUserManager(userId: string, managerId: string): Promise<void> {
    try {
      await (this.graph.users.getById(userId) as any).manager.set({
        '@odata.id': `https://graph.microsoft.com/v1.0/users/${managerId}`
      });
      logger.info('ITProvisioningService', `Set manager for user: ${userId}`);
    } catch (error) {
      logger.error('ITProvisioningService', 'Error setting user manager:', error);
      throw error;
    }
  }

  /**
   * Revoke all active sessions for a user
   */
  public async revokeUserSessions(userId: string): Promise<void> {
    try {
      await (this.graph.users.getById(userId) as any).revokeSignInSessions();
      logger.info('ITProvisioningService', `Revoked sessions for user: ${userId}`);
    } catch (error) {
      logger.error('ITProvisioningService', 'Error revoking user sessions:', error);
      throw error;
    }
  }

  // ============================================================================
  // License Operations
  // ============================================================================

  /**
   * Assign licenses to a user
   */
  public async assignLicenses(userId: string, skuIds: string[]): Promise<void> {
    try {
      const addLicenses = skuIds.map(skuId => ({ skuId }));

      await (this.graph.users.getById(userId) as any).assignLicense({
        addLicenses,
        removeLicenses: []
      });

      logger.info('ITProvisioningService', `Assigned licenses to user ${userId}: ${skuIds.join(', ')}`);
    } catch (error) {
      logger.error('ITProvisioningService', 'Error assigning licenses:', error);
      throw error;
    }
  }

  /**
   * Remove licenses from a user
   */
  public async removeLicenses(userId: string, skuIds: string[]): Promise<void> {
    try {
      await (this.graph.users.getById(userId) as any).assignLicense({
        addLicenses: [],
        removeLicenses: skuIds
      });

      logger.info('ITProvisioningService', `Removed licenses from user ${userId}: ${skuIds.join(', ')}`);
    } catch (error) {
      logger.error('ITProvisioningService', 'Error removing licenses:', error);
      throw error;
    }
  }

  /**
   * Get available licenses in the tenant
   */
  public async getAvailableLicenses(): Promise<any[]> {
    try {
      const licenses = await (this.graph as any).subscribedSkus();
      return licenses;
    } catch (error) {
      logger.error('ITProvisioningService', 'Error getting available licenses:', error);
      throw error;
    }
  }

  // ============================================================================
  // Group Operations
  // ============================================================================

  /**
   * Add user to a group
   */
  public async addUserToGroup(userId: string, groupId: string): Promise<void> {
    try {
      await (this.graph.groups.getById(groupId) as any).members.add({
        '@odata.id': `https://graph.microsoft.com/v1.0/users/${userId}`
      });
      logger.info('ITProvisioningService', `Added user ${userId} to group ${groupId}`);
    } catch (error: any) {
      // Ignore if already a member
      if (error?.message?.includes('already exist')) {
        logger.warn('ITProvisioningService', `User ${userId} already in group ${groupId}`);
        return;
      }
      logger.error('ITProvisioningService', 'Error adding user to group:', error);
      throw error;
    }
  }

  /**
   * Remove user from a group
   */
  public async removeUserFromGroup(userId: string, groupId: string): Promise<void> {
    try {
      await (this.graph.groups.getById(groupId) as any).members.getById(userId).remove();
      logger.info('ITProvisioningService', `Removed user ${userId} from group ${groupId}`);
    } catch (error: any) {
      // Ignore if not a member
      if (error?.status === 404) {
        logger.warn('ITProvisioningService', `User ${userId} not in group ${groupId}`);
        return;
      }
      logger.error('ITProvisioningService', 'Error removing user from group:', error);
      throw error;
    }
  }

  /**
   * Get user's group memberships
   */
  public async getUserGroupMemberships(userId: string): Promise<Array<{ id: string; displayName: string }>> {
    try {
      const groups = await (this.graph.users.getById(userId) as any).memberOf();
      return groups
        .filter((g: any) => g['@odata.type'] === '#microsoft.graph.group')
        .map((g: any) => ({ id: g.id, displayName: g.displayName }));
    } catch (error) {
      logger.error('ITProvisioningService', 'Error getting user group memberships:', error);
      throw error;
    }
  }

  // ============================================================================
  // Teams Operations
  // ============================================================================

  /**
   * Add user to a Team
   */
  public async addUserToTeam(userId: string, teamId: string, role: 'member' | 'owner' = 'member'): Promise<void> {
    try {
      await (this.graph.teams.getById(teamId) as any).members.add({
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${userId}')`,
        roles: role === 'owner' ? ['owner'] : []
      });
      logger.info('ITProvisioningService', `Added user ${userId} to team ${teamId}`);
    } catch (error: any) {
      if (error?.message?.includes('already exist')) {
        logger.warn('ITProvisioningService', `User ${userId} already in team ${teamId}`);
        return;
      }
      logger.error('ITProvisioningService', 'Error adding user to team:', error);
      throw error;
    }
  }

  /**
   * Remove user from a Team
   */
  public async removeUserFromTeam(userId: string, teamId: string): Promise<void> {
    try {
      // Get membership ID first
      const members = await (this.graph.teams.getById(teamId) as any).members();
      const membership = members.find((m: any) => m.userId === userId);

      if (membership) {
        await (this.graph.teams.getById(teamId) as any).members.getById(membership.id).delete();
        logger.info('ITProvisioningService', `Removed user ${userId} from team ${teamId}`);
      }
    } catch (error: any) {
      if (error?.status === 404) {
        logger.warn('ITProvisioningService', `User ${userId} not in team ${teamId}`);
        return;
      }
      logger.error('ITProvisioningService', 'Error removing user from team:', error);
      throw error;
    }
  }

  /**
   * Get user's team memberships
   */
  public async getUserTeamMemberships(userId: string): Promise<Array<{ id: string; displayName: string }>> {
    try {
      const teams = await (this.graph.users.getById(userId) as any).joinedTeams();
      return teams.map((t: any) => ({ id: t.id, displayName: t.displayName }));
    } catch (error) {
      logger.error('ITProvisioningService', 'Error getting user team memberships:', error);
      return [];
    }
  }

  // ============================================================================
  // Helper Methods
  // ============================================================================

  private createStep(
    name: string,
    actionType: ProvisioningActionType,
    targetResource: string
  ): IProvisioningStep {
    return {
      id: uuidv4(),
      name,
      actionType,
      status: 'InProgress',
      order: 0,
      targetResource,
      startedAt: new Date(),
      canRollback: true
    };
  }

  private buildUserCreateRequest(
    process: IJmlProcess,
    config: IProvisioningConfig
  ): IEntraUserCreateRequest {
    const nameParts = process.EmployeeName.split(' ');
    const givenName = nameParts[0];
    const surname = nameParts.slice(1).join(' ') || nameParts[0];
    const mailNickname = process.EmployeeEmail.split('@')[0];

    return {
      displayName: process.EmployeeName,
      givenName,
      surname,
      mailNickname,
      userPrincipalName: process.EmployeeEmail,
      mail: process.EmployeeEmail,
      jobTitle: process.JobTitle,
      department: process.Department,
      officeLocation: process.Location,
      usageLocation: config.defaultUsageLocation,
      accountEnabled: true,
      passwordProfile: {
        password: this.generateSecurePassword(config.passwordLength),
        forceChangePasswordNextSignIn: config.forcePasswordChange
      }
    };
  }

  private generateSecurePassword(length: number = 16): string {
    const uppercase = 'ABCDEFGHJKLMNPQRSTUVWXYZ';
    const lowercase = 'abcdefghjkmnpqrstuvwxyz';
    const numbers = '23456789';
    const special = '!@#$%^&*';
    const all = uppercase + lowercase + numbers + special;

    let password = '';
    // Ensure at least one of each type
    password += uppercase[Math.floor(Math.random() * uppercase.length)];
    password += lowercase[Math.floor(Math.random() * lowercase.length)];
    password += numbers[Math.floor(Math.random() * numbers.length)];
    password += special[Math.floor(Math.random() * special.length)];

    // Fill rest randomly
    for (let i = 4; i < length; i++) {
      password += all[Math.floor(Math.random() * all.length)];
    }

    // Shuffle
    return password.split('').sort(() => Math.random() - 0.5).join('');
  }

  /**
   * Get department-specific provisioning config with fallback to default
   * Prevents silent failures when department config is missing
   */
  private getDepartmentConfig(
    department: string,
    config: IProvisioningConfig
  ): IDepartmentProvisioningConfig {
    // Try to find exact department match
    const exactMatch = config.departmentConfigs.find(
      dc => dc.department.toLowerCase() === department.toLowerCase()
    );

    if (exactMatch) {
      return exactMatch;
    }

    // Try to find 'Default' or 'General' department config
    const defaultConfig = config.departmentConfigs.find(
      dc => dc.department.toLowerCase() === 'default' || dc.department.toLowerCase() === 'general'
    );

    if (defaultConfig) {
      logger.warn(
        'ITProvisioningService',
        `No config found for department "${department}", using default config`
      );
      return defaultConfig;
    }

    // Return a minimal fallback config to prevent silent failures
    logger.warn(
      'ITProvisioningService',
      `No department config found for "${department}" and no default exists. Using minimal fallback.`
    );

    return {
      department: department,
      defaultLicenses: [], // No licenses - admin needs to configure
      securityGroups: [],  // No groups - admin needs to configure
      teams: []            // No teams - admin needs to configure
    };
  }

  private async rollbackSteps(steps: IProvisioningStep[]): Promise<void> {
    for (const step of steps) {
      if (!step.canRollback) continue;

      try {
        switch (step.actionType) {
          case 'CreateUser':
            // Don't delete user on rollback - just disable
            if (step.responsePayload) {
              const { id } = JSON.parse(step.responsePayload);
              await this.disableEntraUser(id);
            }
            break;
          case 'AddToGroup':
            // Remove from group
            if (step.responsePayload) {
              const { userId, groupId } = JSON.parse(step.responsePayload);
              await this.removeUserFromGroup(userId, groupId);
            }
            break;
          case 'AddToTeam':
            // Remove from team
            if (step.responsePayload) {
              const { userId, teamId } = JSON.parse(step.responsePayload);
              await this.removeUserFromTeam(userId, teamId);
            }
            break;
          case 'AssignLicense':
            // Remove assigned licenses
            if (step.responsePayload) {
              const { userId, skuIds } = JSON.parse(step.responsePayload);
              await this.removeLicenses(userId, skuIds);
            }
            break;
        }
        step.rollbackCompleted = true;
        logger.info('ITProvisioningService', `Rolled back step: ${step.name}`);
      } catch (error) {
        logger.error('ITProvisioningService', `Failed to rollback step ${step.name}:`, error);
      }
    }
  }

  /**
   * Notify IT administrators of provisioning/deprovisioning failure
   * Sends urgent email to admin group for immediate action
   */
  private async notifyAdminsOfFailure(
    processType: 'Joiner' | 'Mover' | 'Leaver',
    process: IJmlProcess,
    errorMessage: string,
    failedStep: string,
    requestId: number
  ): Promise<void> {
    try {
      // Get IT Admin group members from config or use default
      const adminEmails = await this.getITAdminEmails();

      if (adminEmails.length === 0) {
        logger.warn('ITProvisioningService', 'No IT Admin emails configured for failure notifications');
        return;
      }

      const siteUrl = this.context.pageContext.web.absoluteUrl;
      const subject = `⚠️ IT Provisioning FAILED: ${processType} - ${process.EmployeeName}`;

      const htmlBody = this.buildAdminFailureNotificationHtml(
        processType,
        process,
        errorMessage,
        failedStep,
        requestId,
        siteUrl
      );

      // Queue urgent notification to IT admins
      const result = await this.emailQueueService.queueEmail({
        to: adminEmails,
        subject,
        htmlBody,
        priority: EmailPriority.Urgent,
        processId: process.Id,
        notificationType: 'ProvisioningFailure'
      });

      if (result.success) {
        logger.info(
          'ITProvisioningService',
          `Admin failure notification queued (Queue ID: ${result.queueItemId})`
        );
      } else {
        logger.warn(
          'ITProvisioningService',
          `Failed to queue admin failure notification: ${result.error}`
        );
      }

      // Also create high-priority in-app notification
      for (const adminEmail of adminEmails) {
        try {
          await this.sp.web.lists.getByTitle('JML_Notifications').items.add({
            Title: `Provisioning Failed: ${process.EmployeeName}`,
            NotificationType: 'ProvisioningFailure',
            RecipientEmail: adminEmail,
            MessageBody: `${processType} provisioning failed for ${process.EmployeeName}. Error: ${errorMessage}. Failed at step: ${failedStep}`,
            Priority: 'Critical',
            Status: 'Pending',
            ProcessId: process.Id?.toString(),
            IsRead: false
          });
        } catch (notifError) {
          // Continue even if one notification fails
          logger.warn('ITProvisioningService', `Failed to create in-app notification for ${adminEmail}`, notifError);
        }
      }
    } catch (error) {
      // Don't throw - failure notification shouldn't break the main flow
      logger.error('ITProvisioningService', 'Failed to send admin failure notification:', error);
    }
  }

  /**
   * Get IT Administrator email addresses from config
   */
  private async getITAdminEmails(): Promise<string[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.PROVISIONING_CONFIG_LIST).items
        .select('ConfigKey', 'ConfigValue', 'IsActive')
        .filter(`ConfigKey eq 'ITAdminEmails' and IsActive eq 1`)();

      if (items.length > 0 && items[0].ConfigValue) {
        // Config stores emails as semicolon-separated string
        return items[0].ConfigValue.split(';').map((e: string) => e.trim()).filter((e: string) => e);
      }

      // Fallback: Try to get from SharePoint group
      // This uses the default "IT Admin" group if configured
      return [];
    } catch (error) {
      logger.warn('ITProvisioningService', 'Could not load IT Admin emails:', error);
      return [];
    }
  }

  /**
   * Build admin failure notification email HTML
   */
  private buildAdminFailureNotificationHtml(
    processType: 'Joiner' | 'Mover' | 'Leaver',
    process: IJmlProcess,
    errorMessage: string,
    failedStep: string,
    requestId: number,
    siteUrl: string
  ): string {
    const processColor = processType === 'Joiner' ? '#107C10' : processType === 'Mover' ? '#0078d4' : '#d13438';

    return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { font-family: 'Segoe UI', sans-serif; margin: 0; padding: 0; background: #f3f2f1; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .card { background: #fff; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); overflow: hidden; }
    .header { background: #d13438; padding: 24px; }
    .header h1 { margin: 0; font-size: 20px; color: #fff; display: flex; align-items: center; gap: 8px; }
    .content { padding: 24px; }
    .alert-box { background: #fde7e9; border: 1px solid #d13438; border-radius: 8px; padding: 16px; margin-bottom: 20px; }
    .alert-box h3 { margin: 0 0 8px; color: #a80000; font-size: 14px; }
    .alert-box p { margin: 0; color: #323130; font-size: 14px; }
    .details-table { width: 100%; border-collapse: collapse; margin: 20px 0; }
    .details-table td { padding: 12px 0; border-bottom: 1px solid #edebe9; font-size: 14px; }
    .details-table td:first-child { color: #605e5c; width: 140px; }
    .details-table td:last-child { color: #323130; font-weight: 500; }
    .process-badge { display: inline-block; padding: 4px 12px; border-radius: 16px; font-size: 12px; font-weight: 600; color: #fff; background: ${processColor}; }
    .error-code { background: #faf9f8; border: 1px solid #edebe9; border-radius: 4px; padding: 12px; font-family: monospace; font-size: 13px; color: #a80000; margin: 16px 0; white-space: pre-wrap; word-break: break-word; }
    .button { display: inline-block; background: #0078d4; color: #fff; padding: 12px 24px; border-radius: 4px; text-decoration: none; font-weight: 600; margin-top: 8px; }
    .footer { padding: 16px 24px; background: #faf9f8; text-align: center; font-size: 12px; color: #605e5c; }
  </style>
</head>
<body>
  <div class="container">
    <div class="card">
      <div class="header">
        <h1>⚠️ IT Provisioning Failed</h1>
      </div>
      <div class="content">
        <div class="alert-box">
          <h3>Action Required</h3>
          <p>IT provisioning has failed and requires immediate attention. The user account may be in an inconsistent state.</p>
        </div>

        <table class="details-table">
          <tr>
            <td>Process Type</td>
            <td><span class="process-badge">${processType}</span></td>
          </tr>
          <tr>
            <td>Employee</td>
            <td>${process.EmployeeName}</td>
          </tr>
          <tr>
            <td>Email</td>
            <td>${process.EmployeeEmail}</td>
          </tr>
          <tr>
            <td>Department</td>
            <td>${process.Department}</td>
          </tr>
          <tr>
            <td>Failed Step</td>
            <td style="color: #d13438; font-weight: 600;">${failedStep}</td>
          </tr>
          <tr>
            <td>Request ID</td>
            <td>#${requestId}</td>
          </tr>
          <tr>
            <td>Process ID</td>
            <td>#${process.Id}</td>
          </tr>
          <tr>
            <td>Time</td>
            <td>${new Date().toLocaleString()}</td>
          </tr>
        </table>

        <h4 style="margin: 20px 0 8px; color: #323130;">Error Details</h4>
        <div class="error-code">${errorMessage}</div>

        <h4 style="margin: 20px 0 8px; color: #323130;">Recommended Actions</h4>
        <ol style="margin: 0; padding-left: 20px; color: #323130; font-size: 14px; line-height: 1.8;">
          <li>Check the IT Provisioning Audit Log for detailed step-by-step status</li>
          <li>Verify the user account state in Azure AD / Entra ID</li>
          <li>Review any partial rollback actions that may have occurred</li>
          <li>Manually complete any remaining provisioning steps if needed</li>
          <li>Update the process status once resolved</li>
        </ol>

        <div style="margin-top: 24px;">
          <a href="${siteUrl}/SitePages/AdminPanel.aspx?view=provisioning" class="button">View Provisioning Logs</a>
        </div>
      </div>
      <div class="footer">
        This is an automated alert from the JML IT Provisioning System.
      </div>
    </div>
  </div>
</body>
</html>`;
  }

  /**
   * Send welcome email to new employee using EmailQueueService
   * This queues the email for background sending via Power Automate or Azure Function
   */
  private async sendWelcomeEmail(
    user: IEntraUser,
    tempPassword: string,
    process: IJmlProcess
  ): Promise<void> {
    const recipientEmail = user.mail || user.userPrincipalName;
    const siteUrl = this.context.pageContext.web.absoluteUrl;

    // Build welcome email HTML
    const htmlBody = this.buildWelcomeEmailHtml(
      user.displayName,
      user.userPrincipalName,
      tempPassword,
      process.Department,
      process.JobTitle,
      siteUrl
    );

    try {
      // Queue the welcome email for sending
      const result = await this.emailQueueService.queueEmail({
        to: [recipientEmail],
        subject: `Welcome to the Company, ${user.givenName || user.displayName}!`,
        htmlBody,
        priority: EmailPriority.High,
        processId: process.Id,
        notificationType: 'WelcomeEmail'
      });

      if (result.success) {
        logger.info(
          'ITProvisioningService',
          `Welcome email queued for ${recipientEmail} (Queue ID: ${result.queueItemId})`
        );
      } else {
        logger.warn(
          'ITProvisioningService',
          `Failed to queue welcome email for ${recipientEmail}: ${result.error}`
        );
      }

      // Also create in-app notification as backup
      await this.sp.web.lists.getByTitle('JML_Notifications').items.add({
        Title: `Welcome - ${user.displayName}`,
        NotificationType: 'WelcomeEmail',
        RecipientId: null, // External recipient
        RecipientEmail: recipientEmail,
        MessageBody: `Welcome to the company! Your account has been created. Please sign in with your email and temporary password to get started.`,
        Priority: 'High',
        Status: result.success ? 'EmailQueued' : 'Pending',
        ProcessId: process.Id?.toString()
      });

    } catch (error) {
      logger.warn('ITProvisioningService', 'Could not send welcome email:', error);

      // Fallback: Create notification for manual sending
      try {
        await this.sp.web.lists.getByTitle('JML_Notifications').items.add({
          Title: `Welcome - ${user.displayName}`,
          NotificationType: 'WelcomeEmail',
          RecipientId: null,
          RecipientEmail: recipientEmail,
          MessageBody: `Welcome to the company! Your account has been created. Please sign in with your email and temporary password to get started.`,
          Priority: 'High',
          Status: 'Pending',
          ProcessId: process.Id?.toString()
        });
      } catch (notifError) {
        logger.error('ITProvisioningService', 'Could not create fallback notification:', notifError);
      }
    }
  }

  /**
   * Build welcome email HTML template
   */
  private buildWelcomeEmailHtml(
    displayName: string,
    userPrincipalName: string,
    tempPassword: string,
    department: string,
    jobTitle: string,
    siteUrl: string
  ): string {
    return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { font-family: 'Segoe UI', sans-serif; margin: 0; padding: 0; background: #f3f2f1; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .card { background: #fff; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); overflow: hidden; }
    .header { background: linear-gradient(135deg, #107C10 0%, #0078d4 100%); padding: 32px 24px; text-align: center; }
    .header h1 { margin: 0; font-size: 28px; color: #fff; font-weight: 600; }
    .header p { margin: 8px 0 0; font-size: 16px; color: rgba(255,255,255,0.9); }
    .content { padding: 32px 24px; }
    .welcome-text { font-size: 16px; color: #323130; line-height: 1.6; margin-bottom: 24px; }
    .credentials-box { background: #faf9f8; border: 1px solid #edebe9; border-radius: 8px; padding: 20px; margin: 24px 0; }
    .credentials-box h3 { margin: 0 0 16px; font-size: 16px; color: #323130; display: flex; align-items: center; gap: 8px; }
    .credential-row { display: flex; justify-content: space-between; padding: 12px 0; border-bottom: 1px solid #edebe9; }
    .credential-row:last-child { border-bottom: none; }
    .credential-label { font-size: 13px; color: #605e5c; }
    .credential-value { font-size: 14px; color: #323130; font-weight: 600; font-family: monospace; }
    .info-card { background: #e7f3ff; border-left: 4px solid #0078d4; padding: 16px; margin: 24px 0; border-radius: 0 8px 8px 0; }
    .info-card p { margin: 0; font-size: 14px; color: #004578; }
    .button { display: inline-block; background: #0078d4; color: #fff; padding: 14px 32px; border-radius: 4px; text-decoration: none; font-weight: 600; margin-top: 16px; }
    .button:hover { background: #106ebe; }
    .steps { margin: 24px 0; }
    .step { display: flex; align-items: flex-start; margin-bottom: 16px; }
    .step-number { background: #0078d4; color: #fff; width: 28px; height: 28px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-weight: 600; font-size: 14px; margin-right: 12px; flex-shrink: 0; }
    .step-content h4 { margin: 0 0 4px; font-size: 15px; color: #323130; }
    .step-content p { margin: 0; font-size: 13px; color: #605e5c; }
    .footer { padding: 24px; background: #faf9f8; text-align: center; }
    .footer p { margin: 0; font-size: 12px; color: #605e5c; }
    .footer a { color: #0078d4; text-decoration: none; }
  </style>
</head>
<body>
  <div class="container">
    <div class="card">
      <div class="header">
        <h1>Welcome to the Team!</h1>
        <p>${displayName}</p>
      </div>
      <div class="content">
        <p class="welcome-text">
          We're thrilled to have you join us as <strong>${jobTitle}</strong> in the <strong>${department}</strong> department.
          Your account has been created and you're ready to get started!
        </p>

        <div class="credentials-box">
          <h3>🔐 Your Login Credentials</h3>
          <div class="credential-row">
            <span class="credential-label">Email / Username</span>
            <span class="credential-value">${userPrincipalName}</span>
          </div>
          <div class="credential-row">
            <span class="credential-label">Temporary Password</span>
            <span class="credential-value">${tempPassword}</span>
          </div>
        </div>

        <div class="info-card">
          <p><strong>Important:</strong> You will be prompted to change your password when you first sign in. Please choose a strong, unique password.</p>
        </div>

        <h3 style="margin-top: 32px;">Getting Started</h3>
        <div class="steps">
          <div class="step">
            <div class="step-number">1</div>
            <div class="step-content">
              <h4>Sign In to Microsoft 365</h4>
              <p>Go to office.com and sign in with your credentials above</p>
            </div>
          </div>
          <div class="step">
            <div class="step-number">2</div>
            <div class="step-content">
              <h4>Set Your New Password</h4>
              <p>Create a secure password that's at least 12 characters</p>
            </div>
          </div>
          <div class="step">
            <div class="step-number">3</div>
            <div class="step-content">
              <h4>Complete Your Onboarding Tasks</h4>
              <p>Visit the Employee Portal to see your personalized checklist</p>
            </div>
          </div>
        </div>

        <div style="text-align: center; margin-top: 32px;">
          <a href="${siteUrl}" class="button">Go to Employee Portal</a>
        </div>
      </div>
      <div class="footer">
        <p>Need help? Contact your IT Help Desk or reach out to your manager.</p>
        <p style="margin-top: 8px;">This is an automated message from the JML Onboarding System.</p>
      </div>
    </div>
  </div>
</body>
</html>`;
  }

  // ============================================================================
  // SharePoint List Operations
  // ============================================================================

  private async createProvisioningRequest(
    process: IJmlProcess,
    processType: 'Joiner' | 'Mover' | 'Leaver'
  ): Promise<number> {
    try {
      const result = await this.sp.web.lists.getByTitle(this.PROVISIONING_LOG_LIST).items.add({
        Title: `${processType} - ${process.EmployeeName}`,
        ProcessId: process.Id,
        EmployeeId: process.EmployeeID,
        EmployeeName: process.EmployeeName,
        EmployeeEmail: process.EmployeeEmail,
        ProcessType: processType,
        Department: process.Department,
        JobTitle: process.JobTitle,
        Status: 'InProgress',
        CreatedById: this.context.pageContext.legacyPageContext?.userId
      });

      return result.data.Id;
    } catch (error) {
      logger.error('ITProvisioningService', 'Error creating provisioning request:', error);
      throw error;
    }
  }

  private async updateProvisioningRequestStatus(
    requestId: number,
    status: ProvisioningStatus
  ): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(this.PROVISIONING_LOG_LIST).items
        .getById(requestId)
        .update({
          Status: status,
          CompletedAt: status === 'Completed' || status === 'Failed' ? new Date() : null
        });
    } catch (error) {
      logger.error('ITProvisioningService', 'Error updating provisioning request:', error);
    }
  }

  private async updateProvisioningStep(
    requestId: number,
    step: IProvisioningStep
  ): Promise<void> {
    // Store step details in audit log or separate tracking list
    // This is simplified - in production, you might want a separate steps list
    try {
      await this.sp.web.lists.getByTitle(this.PROVISIONING_LOG_LIST).items
        .getById(requestId)
        .update({
          LastStepCompleted: step.name,
          LastStepStatus: step.status
        });
    } catch (error) {
      logger.warn('ITProvisioningService', 'Error updating provisioning step:', error);
    }
  }

  private async createAuditLog(log: Partial<IProvisioningAuditLog>): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(this.PROVISIONING_AUDIT_LIST).items.add({
        Title: `${log.ActionType} - ${log.EmployeeName}`,
        ProcessId: log.ProcessId,
        RequestId: log.RequestId,
        EmployeeId: log.EmployeeId,
        EmployeeName: log.EmployeeName,
        ActionType: log.ActionType,
        ActionStatus: log.ActionStatus,
        TargetResource: log.TargetResource,
        TargetResourceName: log.TargetResourceName,
        RequestPayload: log.RequestPayload,
        ResponsePayload: log.ResponsePayload,
        ErrorDetails: log.ErrorDetails,
        ExecutedById: log.ExecutedById,
        ExecutedByName: log.ExecutedByName,
        ExecutedAt: log.ExecutedAt || new Date()
      });
    } catch (error) {
      logger.error('ITProvisioningService', 'Error creating audit log:', error);
      // Don't throw - audit logging shouldn't break main flow
    }
  }

  /**
   * Get provisioning configuration from SharePoint
   */
  public async getProvisioningConfig(): Promise<IProvisioningConfig> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.PROVISIONING_CONFIG_LIST).items
        .select('ConfigKey', 'ConfigValue', 'IsActive')
        .filter('IsActive eq 1')();

      const config = { ...this.defaultConfig };

      for (const item of items) {
        switch (item.ConfigKey) {
          case 'DefaultUsageLocation':
            config.defaultUsageLocation = item.ConfigValue;
            break;
          case 'PasswordLength':
            config.passwordLength = parseInt(item.ConfigValue, 10);
            break;
          case 'ForcePasswordChange':
            config.forcePasswordChange = item.ConfigValue === 'true';
            break;
          case 'SendWelcomeEmail':
            config.sendWelcomeEmail = item.ConfigValue === 'true';
            break;
          case 'LeaverGracePeriodDays':
            config.leaverGracePeriodDays = parseInt(item.ConfigValue, 10);
            break;
          case 'AutoDisableOnLeave':
            config.autoDisableOnLeave = item.ConfigValue === 'true';
            break;
          case 'DepartmentConfigs':
            config.departmentConfigs = JSON.parse(item.ConfigValue);
            break;
          case 'RoleConfigs':
            config.roleConfigs = JSON.parse(item.ConfigValue);
            break;
        }
      }

      return config;
    } catch (error) {
      logger.warn('ITProvisioningService', 'Could not load config, using defaults:', error);
      return this.defaultConfig;
    }
  }
}
