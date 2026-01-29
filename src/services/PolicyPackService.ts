// @ts-nocheck
// Policy Pack Service
// Manages bundled policies for onboarding and deployment

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';
import {
  IPolicyPack,
  IPolicyPackAssignment,
  IPolicyPackProgress,
  ICreatePolicyPackRequest,
  IAssignPolicyPackRequest,
  IPolicyPackDeploymentResult,
  IPolicy,
  IPolicyAcknowledgement,
  IPersonalPolicyView,
  ReadTimeframe,
  AcknowledgementStatus
} from '../models/IPolicy';
import { logger } from './LoggingService';
import { PolicyLists, PolicyPackLists, SocialLists } from '../constants/SharePointListNames';

export class PolicyPackService {
  private sp: SPFI;
  private readonly POLICY_PACKS_LIST = PolicyPackLists.POLICY_PACKS;
  private readonly POLICY_PACK_ASSIGNMENTS_LIST = PolicyPackLists.POLICY_PACK_ASSIGNMENTS;
  private readonly POLICIES_LIST = PolicyLists.POLICIES;
  private readonly POLICY_ACKNOWLEDGEMENTS_LIST = PolicyLists.POLICY_ACKNOWLEDGEMENTS;
  private currentUserId: number = 0;
  private currentUserEmail: string = '';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Initialize service
   */
  public async initialize(): Promise<void> {
    try {
      const user = await this.sp.web.currentUser();
      this.currentUserId = user.Id;
      this.currentUserEmail = user.Email;
    } catch (error) {
      logger.error('PolicyPackService', 'Failed to initialize:', error);
      throw error;
    }
  }

  // ============================================================================
  // POLICY PACK CRUD
  // ============================================================================

  /**
   * Create a policy pack
   */
  public async createPolicyPack(request: ICreatePolicyPackRequest): Promise<IPolicyPack> {
    try {
      const packData = {
        Title: request.packName,
        PackName: request.packName,
        PackDescription: request.packDescription,
        PackCategory: request.packType,
        PackType: request.packType,
        IsActive: true,
        IsMandatory: true,
        TargetDepartments: request.targetDepartments ? JSON.stringify(request.targetDepartments) : undefined,
        TargetRoles: request.targetRoles ? JSON.stringify(request.targetRoles) : undefined,
        TargetLocations: request.targetLocations ? JSON.stringify(request.targetLocations) : undefined,
        TargetProcessType: request.targetProcessType,
        PolicyIds: JSON.stringify(request.policyIds),
        PolicyCount: request.policyIds.length,
        RequireAllAcknowledged: request.requireAllAcknowledged ?? true,
        AcknowledgementDeadlineDays: request.acknowledgementDeadlineDays,
        ReadTimeframe: request.readTimeframe,
        IsSequential: request.isSequential ?? false,
        PolicySequence: request.isSequential ? JSON.stringify(request.policyIds) : undefined,
        SendWelcomeEmail: request.sendWelcomeEmail ?? true,
        SendTeamsNotification: request.sendTeamsNotification ?? true,
        TotalAssignments: 0,
        TotalCompleted: 0,
        AverageCompletionDays: 0,
        CompletionRate: 0,
        CreatedById: this.currentUserId,
        CreatedDate: new Date().toISOString(),
        Version: '1.0'
      };

      const result = await this.sp.web.lists
        .getByTitle(this.POLICY_PACKS_LIST)
        .items.add(packData);

      return await this.getPolicyPackById(result.data.Id);
    } catch (error) {
      logger.error('PolicyPackService', 'Failed to create policy pack:', error);
      throw error;
    }
  }

  /**
   * Get policy pack by ID
   */
  public async getPolicyPackById(packId: number): Promise<IPolicyPack> {
    try {
      const pack = await this.sp.web.lists
        .getByTitle(this.POLICY_PACKS_LIST)
        .items.getById(packId)();

      return this.mapPolicyPack(pack);
    } catch (error) {
      logger.error('PolicyPackService', 'Failed to get policy pack:', error);
      throw error;
    }
  }

  /**
   * Get all policy packs
   */
  public async getPolicyPacks(packType?: string): Promise<IPolicyPack[]> {
    try {
      let query = this.sp.web.lists
        .getByTitle(this.POLICY_PACKS_LIST)
        .items.filter('IsActive eq true');

      if (packType) {
        query = query.filter(`PackType eq '${packType}'`);
      }

      const packs = await query.top(1000)();
      return packs.map(p => this.mapPolicyPack(p));
    } catch (error) {
      logger.error('PolicyPackService', 'Failed to get policy packs:', error);
      throw error;
    }
  }

  /**
   * Update an existing policy pack
   */
  public async updatePolicyPack(packId: number, request: ICreatePolicyPackRequest): Promise<IPolicyPack> {
    try {
      const packData = {
        Title: request.packName,
        PackName: request.packName,
        PackDescription: request.packDescription,
        PackCategory: request.packType,
        PackType: request.packType,
        TargetDepartments: request.targetDepartments ? JSON.stringify(request.targetDepartments) : undefined,
        TargetRoles: request.targetRoles ? JSON.stringify(request.targetRoles) : undefined,
        TargetLocations: request.targetLocations ? JSON.stringify(request.targetLocations) : undefined,
        TargetProcessType: request.targetProcessType,
        PolicyIds: JSON.stringify(request.policyIds),
        PolicyCount: request.policyIds.length,
        RequireAllAcknowledged: request.requireAllAcknowledged ?? true,
        AcknowledgementDeadlineDays: request.acknowledgementDeadlineDays,
        ReadTimeframe: request.readTimeframe,
        IsSequential: request.isSequential ?? false,
        PolicySequence: request.isSequential ? JSON.stringify(request.policyIds) : undefined,
        SendWelcomeEmail: request.sendWelcomeEmail ?? true,
        SendTeamsNotification: request.sendTeamsNotification ?? true,
        ModifiedById: this.currentUserId,
        ModifiedDate: new Date().toISOString()
      };

      await this.sp.web.lists
        .getByTitle(this.POLICY_PACKS_LIST)
        .items.getById(packId)
        .update(packData);

      logger.info('PolicyPackService', `Updated policy pack ${packId}`);
      return await this.getPolicyPackById(packId);
    } catch (error) {
      logger.error('PolicyPackService', 'Failed to update policy pack:', error);
      throw error;
    }
  }

  /**
   * Delete a policy pack (soft delete by setting IsActive to false)
   */
  public async deletePolicyPack(packId: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.POLICY_PACKS_LIST)
        .items.getById(packId)
        .update({
          IsActive: false
        });
      logger.info('PolicyPackService', `Deleted policy pack ${packId}`);
    } catch (error) {
      logger.error('PolicyPackService', 'Failed to delete policy pack:', error);
      throw error;
    }
  }

  // ============================================================================
  // POLICY PACK ASSIGNMENT
  // ============================================================================

  /**
   * Assign policy pack to users
   */
  public async assignPolicyPack(request: IAssignPolicyPackRequest): Promise<IPolicyPackDeploymentResult> {
    try {
      const pack = await this.getPolicyPackById(request.packId);
      const result: IPolicyPackDeploymentResult = {
        packId: request.packId,
        packName: pack.PackName,
        totalUsers: request.userIds.length,
        successfulAssignments: 0,
        failedAssignments: 0,
        emailsSent: 0,
        teamsNotificationsSent: 0,
        errors: [],
        assignmentIds: []
      };

      for (const userId of request.userIds) {
        try {
          const assignmentId = await this.assignPackToUser(
            request.packId,
            userId,
            request.dueDate,
            request.assignmentReason,
            pack
          );

          result.assignmentIds.push(assignmentId);
          result.successfulAssignments++;

          // Send welcome email
          if (pack.SendWelcomeEmail) {
            await this.sendWelcomeEmail(assignmentId);
            result.emailsSent++;
          }

          // Send Teams notification
          if (pack.SendTeamsNotification) {
            await this.sendTeamsNotification(assignmentId);
            result.teamsNotificationsSent++;
          }
        } catch (error) {
          result.failedAssignments++;
          result.errors.push(`Failed to assign to user ${userId}: ${error}`);
        }
      }

      // Update pack analytics
      await this.updatePackAnalytics(request.packId);

      return result;
    } catch (error) {
      logger.error('PolicyPackService', 'Failed to assign policy pack:', error);
      throw error;
    }
  }

  /**
   * Assign pack to individual user
   */
  private async assignPackToUser(
    packId: number,
    userId: number,
    dueDate: Date | undefined,
    assignmentReason: string,
    pack: IPolicyPack
  ): Promise<number> {
    // Get user details
    const user = await this.sp.web.siteUsers.getById(userId)();

    // Calculate due date if not provided
    let calculatedDueDate = dueDate;
    if (!calculatedDueDate && pack.AcknowledgementDeadlineDays) {
      calculatedDueDate = new Date();
      calculatedDueDate.setDate(calculatedDueDate.getDate() + pack.AcknowledgementDeadlineDays);
    }

    // Create assignment
    const assignmentData = {
      Title: `${pack.PackName} - ${user.Title}`,
      PackId: packId,
      UserId: userId,
      UserEmail: user.Email,
      AssignedDate: new Date().toISOString(),
      AssignedById: this.currentUserId,
      AssignmentReason: assignmentReason,
      DueDate: calculatedDueDate?.toISOString(),
      ReadTimeframe: pack.ReadTimeframe,
      TotalPolicies: pack.PolicyCount,
      AcknowledgedPolicies: 0,
      PendingPolicies: pack.PolicyCount,
      OverduePolicies: 0,
      ProgressPercentage: 0,
      Status: 'Not Started',
      WelcomeEmailSent: false,
      TeamsNotificationSent: false,
      RemindersSent: 0,
      PersonalViewURL: this.generatePersonalViewURL(userId)
    };

    const result = await this.sp.web.lists
      .getByTitle(this.POLICY_PACK_ASSIGNMENTS_LIST)
      .items.add(assignmentData);

    // Create individual policy acknowledgements
    await this.createAcknowledgementsForPack(pack, userId);

    return result.data.Id;
  }

  /**
   * Create individual policy acknowledgements for pack
   */
  private async createAcknowledgementsForPack(
    pack: IPolicyPack,
    userId: number
  ): Promise<void> {
    const policyIds = JSON.parse(pack.PolicyIds as any);
    const user = await this.sp.web.siteUsers.getById(userId)();

    for (let i = 0; i < policyIds.length; i++) {
      const policyId = policyIds[i];
      const policy = await this.sp.web.lists
        .getByTitle(this.POLICIES_LIST)
        .items.getById(policyId)() as IPolicy;

      // Calculate due date based on policy's read timeframe
      const dueDate = this.calculateDueDate(policy.ReadTimeframe, policy.ReadTimeframeDays);

      const ackData = {
        Title: `${policy.PolicyName} - ${user.Title}`,
        PolicyId: policyId,
        PolicyVersionNumber: policy.VersionNumber,
        AckUserId: userId,
        UserEmail: user.Email,
        AckStatus: AcknowledgementStatus.Sent,
        AssignedDate: new Date().toISOString(),
        DueDate: dueDate?.toISOString(),
        QuizRequired: policy.RequiresQuiz,
        DocumentOpenCount: 0,
        TotalReadTimeSeconds: 0,
        IsDelegated: false,
        RemindersSent: 0,
        IsExempted: false,
        IsCompliant: false
      };

      // Check if acknowledgement already exists
      const existing = await this.sp.web.lists
        .getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)
        .items.filter(`PolicyId eq ${policyId} and AckUserId eq ${userId}`)
        .top(1)();

      if (existing.length === 0) {
        await this.sp.web.lists
          .getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)
          .items.add(ackData);
      }
    }
  }

  // ============================================================================
  // PROGRESS TRACKING
  // ============================================================================

  /**
   * Get policy pack progress for a user
   */
  public async getPolicyPackProgress(assignmentId: number): Promise<IPolicyPackProgress> {
    try {
      const assignment = await this.sp.web.lists
        .getByTitle(this.POLICY_PACK_ASSIGNMENTS_LIST)
        .items.getById(assignmentId)() as IPolicyPackAssignment;

      const pack = await this.getPolicyPackById(assignment.PackId);
      const policyIds = JSON.parse(pack.PolicyIds as any);

      // Get acknowledgements for all policies in pack
      const acks = await Promise.all(
        policyIds.map(async (policyId: number) => {
          const items = await this.sp.web.lists
            .getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)
            .items.filter(`PolicyId eq ${policyId} and AckUserId eq ${assignment.UserId}`)
            .top(1)();
          return items.length > 0 ? items[0] as IPolicyAcknowledgement : null;
        })
      );

      // Get policy details
      const policies = await Promise.all(
        policyIds.map(async (policyId: number) => {
          return await this.sp.web.lists
            .getByTitle(this.POLICIES_LIST)
            .items.getById(policyId)() as IPolicy;
        })
      );

      // Build policy status array
      const policyStatus = acks.map((ack, index) => {
        if (!ack) return null;

        const policy = policies[index];
        const assignedDate = new Date(ack.AssignedDate);
        const now = new Date();
        const daysSinceAssigned = Math.floor((now.getTime() - assignedDate.getTime()) / (1000 * 60 * 60 * 24));
        const dueDate = ack.DueDate ? new Date(ack.DueDate) : undefined;
        const isOverdue = dueDate ? now > dueDate && ack.AckStatus !== AcknowledgementStatus.Acknowledged : false;

        return {
          policyId: policy.Id!,
          policyName: policy.PolicyName,
          readTimeframe: policy.ReadTimeframe || ReadTimeframe.Month1,
          status: ack.AckStatus,
          dueDate,
          acknowledgedDate: ack.AcknowledgedDate ? new Date(ack.AcknowledgedDate) : undefined,
          isOverdue,
          daysSinceAssigned,
          sequenceOrder: pack.IsSequential ? index + 1 : undefined
        };
      }).filter(p => p !== null) as any[];

      // Calculate metrics
      const acknowledged = policyStatus.filter(p => p.status === AcknowledgementStatus.Acknowledged).length;
      const pending = policyStatus.filter(p =>
        p.status !== AcknowledgementStatus.Acknowledged && !p.isOverdue
      ).length;
      const overdue = policyStatus.filter(p => p.isOverdue).length;

      return {
        packAssignmentId: assignmentId,
        userId: assignment.UserId,
        packName: pack.PackName,
        totalPolicies: policyIds.length,
        acknowledgedPolicies: acknowledged,
        pendingPolicies: pending,
        overduePolicies: overdue,
        progressPercentage: policyIds.length > 0 ? (acknowledged / policyIds.length) * 100 : 0,
        policyStatus,
        estimatedCompletionDate: this.estimateCompletionDate(policyStatus),
        daysUntilDue: assignment.DueDate ? this.calculateDaysUntilDue(new Date(assignment.DueDate)) : undefined,
        isOnTrack: this.isOnTrack(policyStatus, assignment.DueDate)
      };
    } catch (error) {
      logger.error('PolicyPackService', 'Failed to get policy pack progress:', error);
      throw error;
    }
  }

  // ============================================================================
  // PERSONAL POLICY VIEW
  // ============================================================================

  /**
   * Get personal policy view for user
   */
  public async getPersonalPolicyView(userId?: number): Promise<IPersonalPolicyView> {
    try {
      const targetUserId = userId || this.currentUserId;

      // Validate user ID
      if (!targetUserId || targetUserId <= 0) {
        throw new Error('Invalid user ID. Please ensure you are logged in.');
      }

      // Get user info
      let user: any;
      try {
        user = await this.sp.web.siteUsers.getById(targetUserId)();
      } catch (userError) {
        logger.error('PolicyPackService', 'Failed to get user info:', userError);
        throw new Error('Unable to retrieve your user profile. Please ensure you have access to this site.');
      }

      // Check if required lists exist and get acknowledgements
      let allAcks: IPolicyAcknowledgement[] = [];
      try {
        // First try to access the list
        const list = await this.sp.web.lists.getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)();
        if (!list) {
          throw new Error(`List '${this.POLICY_ACKNOWLEDGEMENTS_LIST}' does not exist`);
        }

        // Now get the items
        allAcks = await this.sp.web.lists
          .getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)
          .items.filter(`AckUserId eq ${targetUserId}`)
          .top(1000)() as IPolicyAcknowledgement[];
      } catch (listError: any) {
        const errorMsg = listError?.message || String(listError);
        if (errorMsg.includes('404') || errorMsg.includes('does not exist')) {
          throw new Error(`Policy lists have not been provisioned. The list '${this.POLICY_ACKNOWLEDGEMENTS_LIST}' is missing. Please contact your administrator.`);
        }
        logger.error('PolicyPackService', `Failed to access ${this.POLICY_ACKNOWLEDGEMENTS_LIST}:`, listError);
        throw new Error(`Unable to access policy acknowledgements. Error: ${errorMsg}`);
      }

      const pending = allAcks.filter(a =>
        a.AckStatus === AcknowledgementStatus.Sent || a.AckStatus === AcknowledgementStatus.Opened
      );
      const overdue = allAcks.filter(a => a.AckStatus === AcknowledgementStatus.Overdue);
      const completed = allAcks.filter(a => a.AckStatus === AcknowledgementStatus.Acknowledged);

      // Categorize policies
      const now = new Date();
      const urgent = allAcks.filter(a => {
        if (!a.DueDate || a.AckStatus === AcknowledgementStatus.Acknowledged) return false;
        const due = new Date(a.DueDate);
        const hoursUntilDue = (due.getTime() - now.getTime()) / (1000 * 60 * 60);
        return hoursUntilDue <= 24 && hoursUntilDue > 0;
      });

      const dueSoon = allAcks.filter(a => {
        if (!a.DueDate || a.AckStatus === AcknowledgementStatus.Acknowledged) return false;
        const due = new Date(a.DueDate);
        const daysUntilDue = (due.getTime() - now.getTime()) / (1000 * 60 * 60 * 24);
        return daysUntilDue <= 7 && daysUntilDue > 1;
      });

      const sevenDaysAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
      const newPolicies = allAcks.filter(a => {
        const assigned = new Date(a.AssignedDate);
        return assigned >= sevenDaysAgo;
      });

      // Get policy packs - wrapped in try/catch since this list may not be provisioned
      let packAssignments: IPolicyPackAssignment[] = [];
      let activePacks: IPolicyPackProgress[] = [];
      let completedPacks: IPolicyPackProgress[] = [];

      try {
        packAssignments = await this.sp.web.lists
          .getByTitle(this.POLICY_PACK_ASSIGNMENTS_LIST)
          .items.filter(`UserId eq ${targetUserId}`)
          .top(100)() as IPolicyPackAssignment[];

        // Get progress for active packs (handle individual failures gracefully)
        const activePromises = packAssignments
          .filter(a => a.Status !== 'Completed')
          .map(async a => {
            try {
              return await this.getPolicyPackProgress(a.Id!);
            } catch (e) {
              logger.warn('PolicyPackService', `Failed to get progress for pack ${a.Id}:`, e);
              return null;
            }
          });
        activePacks = (await Promise.all(activePromises)).filter((p): p is IPolicyPackProgress => p !== null);

        // Get progress for completed packs (handle individual failures gracefully)
        const completedPromises = packAssignments
          .filter(a => a.Status === 'Completed')
          .map(async a => {
            try {
              return await this.getPolicyPackProgress(a.Id!);
            } catch (e) {
              logger.warn('PolicyPackService', `Failed to get progress for completed pack ${a.Id}:`, e);
              return null;
            }
          });
        completedPacks = (await Promise.all(completedPromises)).filter((p): p is IPolicyPackProgress => p !== null);
      } catch (packError: any) {
        // Policy pack assignments list may not exist - this is not fatal, just means no packs assigned
        logger.warn('PolicyPackService', `Failed to get policy pack assignments: ${packError?.message || packError}`);
        // Continue with empty arrays - this is a non-critical feature
      }

      // Get followed policies (need to fetch full policy data)
      let followedPolicies: IPolicy[] = [];
      try {
        const follows = await this.sp.web.lists
          .getByTitle(SocialLists.POLICY_FOLLOWERS)
          .items.filter(`UserId eq ${targetUserId}`)
          .select('PolicyId')
          .top(50)();
        const followedIds = follows.map((f: Record<string, unknown>) => f.PolicyId as number);
        if (followedIds.length > 0) {
          const followedFilter = followedIds.map(id => `Id eq ${id}`).join(' or ');
          const policies = await this.sp.web.lists
            .getByTitle(this.POLICIES_LIST)
            .items.filter(followedFilter)
            .select('Id', 'Title', 'PolicyName', 'PolicyCategory', 'ComplianceRisk', 'PolicyStatus')
            .top(50)();
          followedPolicies = policies as IPolicy[];
        }
      } catch {
        // Policy follows list may not exist
      }

      // Get recent updates (policies updated in last 30 days)
      let recentUpdates: {
        policyId: number;
        policyName: string;
        updateType: 'NewVersion' | 'NewComment' | 'Updated';
        updateDate: Date;
        updateDescription: string;
      }[] = [];
      try {
        const thirtyDaysAgo = new Date();
        thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
        const updatedPolicies = await this.sp.web.lists
          .getByTitle(this.POLICIES_LIST)
          .items.filter(`Modified ge datetime'${thirtyDaysAgo.toISOString()}'`)
          .select('Id', 'Title', 'PolicyName', 'Modified', 'VersionNumber')
          .orderBy('Modified', false)
          .top(20)();
        recentUpdates = updatedPolicies.map((p: Record<string, unknown>) => ({
          policyId: p.Id as number,
          policyName: (p.PolicyName as string) || (p.Title as string) || `Policy ${p.Id}`,
          updateType: 'Updated' as const,
          updateDate: new Date(p.Modified as string),
          updateDescription: `Policy updated on ${new Date(p.Modified as string).toLocaleDateString()}`
        }));
      } catch {
        // Policies list may not have expected columns
      }

      // Get recommended policies based on department/role
      let recommendedPolicies: IPolicy[] = [];
      try {
        const userDept = (user as any).Department || '';
        const userRole = (user as any).JobTitle || '';
        const assignedPolicyIds = allAcks.map(a => a.PolicyId);

        // Get policies for user's department that haven't been assigned
        const deptPolicies = await this.sp.web.lists
          .getByTitle(this.POLICIES_LIST)
          .items.filter(`PolicyStatus eq 'Published' and (TargetDepartments eq '${userDept}' or TargetDepartments eq 'All')`)
          .select('Id', 'Title', 'PolicyName', 'PolicyCategory', 'ComplianceRisk')
          .top(10)();

        recommendedPolicies = (deptPolicies as IPolicy[])
          .filter(p => !assignedPolicyIds.includes(p.Id!));
      } catch {
        // Policies list may not have expected columns
      }

      // Get related policies (same category as assigned policies)
      let relatedPolicies: IPolicy[] = [];
      try {
        // Get unique categories (ES5 compatible)
        const allCategories = allAcks.map(a => a.PolicyCategory).filter(Boolean);
        const assignedCategories = allCategories.filter((cat, index) => allCategories.indexOf(cat) === index);
        const assignedPolicyIds = allAcks.map(a => a.PolicyId);

        if (assignedCategories.length > 0) {
          const categoryFilter = assignedCategories.map(c => `PolicyCategory eq '${c}'`).join(' or ');
          const catPolicies = await this.sp.web.lists
            .getByTitle(this.POLICIES_LIST)
            .items.filter(`PolicyStatus eq 'Published' and (${categoryFilter})`)
            .select('Id', 'Title', 'PolicyName', 'PolicyCategory', 'ComplianceRisk')
            .top(10)();

          relatedPolicies = (catPolicies as IPolicy[])
            .filter(p => !assignedPolicyIds.includes(p.Id!));
        }
      } catch {
        // Policies list may not have expected columns
      }

      return {
        userId: targetUserId,
        userEmail: user.Email,
        userName: user.Title,
        department: (user as any).Department || '',
        role: (user as any).JobTitle || '',
        location: (user as any).Location || '',
        totalAssigned: allAcks.length,
        pending: pending.length,
        overdue: overdue.length,
        completed: completed.length,
        complianceScore: allAcks.length > 0 ? (completed.length / allAcks.length) * 100 : 100,
        urgentPolicies: urgent,
        dueSoon,
        newPolicies,
        overduePolicies: overdue,
        activePolicyPacks: activePacks,
        completedPolicyPacks: completedPacks,
        followedPolicies,
        recentUpdates,
        recommendedPolicies,
        relatedPolicies
      };
    } catch (error) {
      logger.error('PolicyPackService', 'Failed to get personal policy view:', error);
      throw error;
    }
  }

  // ============================================================================
  // POLICY ACKNOWLEDGEMENT ACTIONS
  // ============================================================================

  /**
   * Acknowledge a policy
   */
  public async acknowledgePolicy(
    acknowledgementId: number,
    data: {
      acknowledgedDate: Date;
      readDuration: number;
      comments?: string;
    }
  ): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)
        .items.getById(acknowledgementId)
        .update({
          AckStatus: AcknowledgementStatus.Acknowledged,
          AcknowledgedDate: data.acknowledgedDate.toISOString(),
          ReadDuration: data.readDuration,
          Comments: data.comments || ''
        });

      logger.info('PolicyPackService', `Policy ${acknowledgementId} acknowledged`);
    } catch (error) {
      logger.error('PolicyPackService', 'Failed to acknowledge policy:', error);
      throw new Error('Failed to acknowledge policy. Please try again.');
    }
  }

  /**
   * Rate a policy acknowledgement
   */
  public async ratePolicyAcknowledgement(acknowledgementId: number, rating: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.POLICY_ACKNOWLEDGEMENTS_LIST)
        .items.getById(acknowledgementId)
        .update({
          UserRating: rating
        });

      logger.info('PolicyPackService', `Policy ${acknowledgementId} rated ${rating}`);
    } catch (error) {
      logger.error('PolicyPackService', 'Failed to rate policy:', error);
      // Don't throw - rating is optional
    }
  }

  // ============================================================================
  // HELPER METHODS
  // ============================================================================

  /**
   * Calculate due date based on timeframe
   */
  private calculateDueDate(timeframe?: ReadTimeframe, customDays?: number): Date | undefined {
    if (!timeframe) return undefined;

    const now = new Date();
    const timeframeMap: Record<ReadTimeframe, number> = {
      [ReadTimeframe.Immediate]: 0,
      [ReadTimeframe.Day1]: 1,
      [ReadTimeframe.Day3]: 3,
      [ReadTimeframe.Week1]: 7,
      [ReadTimeframe.Week2]: 14,
      [ReadTimeframe.Month1]: 30,
      [ReadTimeframe.Month3]: 90,
      [ReadTimeframe.Month6]: 180,
      [ReadTimeframe.Custom]: customDays || 30
    };

    const days = timeframeMap[timeframe];
    const dueDate = new Date(now.getTime() + days * 24 * 60 * 60 * 1000);
    return dueDate;
  }

  /**
   * Generate personal view URL
   */
  private generatePersonalViewURL(userId: number): string {
    // TODO: Generate actual URL to personal policy view page
    return `/sites/PolicyManager/SitePages/MyPolicies.aspx?userId=${userId}`;
  }

  /**
   * Estimate completion date
   */
  private estimateCompletionDate(policyStatus: any[]): Date | undefined {
    // Simple estimation: average days per policy * remaining policies
    const acknowledged = policyStatus.filter(p => p.acknowledgedDate);
    if (acknowledged.length === 0) return undefined;

    const avgDaysPerPolicy = acknowledged.reduce((sum, p) => {
      const days = (new Date(p.acknowledgedDate).getTime() - new Date().getTime()) / (1000 * 60 * 60 * 24);
      return sum + Math.abs(days);
    }, 0) / acknowledged.length;

    const remaining = policyStatus.length - acknowledged.length;
    const estimatedDays = avgDaysPerPolicy * remaining;

    return new Date(Date.now() + estimatedDays * 24 * 60 * 60 * 1000);
  }

  /**
   * Calculate days until due
   */
  private calculateDaysUntilDue(dueDate: Date): number {
    const now = new Date();
    return Math.floor((dueDate.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));
  }

  /**
   * Check if on track
   */
  private isOnTrack(policyStatus: any[], dueDate?: Date): boolean {
    if (!dueDate) return true;

    const acknowledged = policyStatus.filter(p => p.acknowledgedDate).length;
    const total = policyStatus.length;
    const percentComplete = (acknowledged / total) * 100;

    const now = new Date();
    const totalDays = (new Date(dueDate).getTime() - new Date(policyStatus[0]?.assignedDate || now).getTime()) / (1000 * 60 * 60 * 24);
    const daysPassed = (now.getTime() - new Date(policyStatus[0]?.assignedDate || now).getTime()) / (1000 * 60 * 60 * 24);
    const expectedProgress = (daysPassed / totalDays) * 100;

    return percentComplete >= expectedProgress;
  }

  /**
   * Update pack analytics
   */
  private async updatePackAnalytics(packId: number): Promise<void> {
    try {
      const assignments = await this.sp.web.lists
        .getByTitle(this.POLICY_PACK_ASSIGNMENTS_LIST)
        .items.filter(`PackId eq ${packId}`)
        .top(5000)() as IPolicyPackAssignment[];

      const completed = assignments.filter(a => a.Status === 'Completed');
      const completionDays = completed
        .filter(a => a.CompletionDays)
        .map(a => a.CompletionDays!);

      const avgCompletionDays = completionDays.length > 0
        ? completionDays.reduce((sum, days) => sum + days, 0) / completionDays.length
        : 0;

      await this.sp.web.lists
        .getByTitle(this.POLICY_PACKS_LIST)
        .items.getById(packId)
        .update({
          TotalAssignments: assignments.length,
          TotalCompleted: completed.length,
          AverageCompletionDays: avgCompletionDays,
          CompletionRate: assignments.length > 0 ? (completed.length / assignments.length) * 100 : 0
        });
    } catch (error) {
      logger.error('PolicyPackService', 'Failed to update pack analytics:', error);
    }
  }

  /**
   * Send welcome email for policy pack assignment
   */
  private async sendWelcomeEmail(assignmentId: number): Promise<void> {
    try {
      // Get assignment details
      const assignment = await this.sp.web.lists
        .getByTitle(this.POLICY_PACK_ASSIGNMENTS_LIST)
        .items.getById(assignmentId)
        .select('*', 'User/EMail', 'User/Title', 'Pack/Title', 'Pack/Description')
        .expand('User', 'Pack')();

      const userEmail = assignment.User?.EMail;
      const userName = assignment.User?.Title || 'User';
      const packName = assignment.Pack?.Title || 'Policy Pack';
      const packDescription = assignment.Pack?.Description || '';
      const viewUrl = this.generatePersonalViewURL(assignment.UserId);

      if (userEmail) {
        // Create notification record for email queue
        await this.sp.web.lists
          .getByTitle('PM_NotificationQueue')
          .items.add({
            Title: `Policy Pack Assignment: ${packName}`,
            RecipientEmail: userEmail,
            RecipientName: userName,
            NotificationType: 'PolicyPackAssigned',
            Subject: `Action Required: Complete ${packName}`,
            Body: `
              <p>Dear ${userName},</p>
              <p>You have been assigned the <strong>${packName}</strong> policy pack.</p>
              ${packDescription ? `<p>${packDescription}</p>` : ''}
              <p>Please review and acknowledge the policies in this pack.</p>
              <p><a href="${viewUrl}">View Your Policies</a></p>
              <p>Best regards,<br/>Policy Management Team</p>
            `,
            Status: 'Pending',
            Priority: 'Normal'
          });
      }

      // Mark email as sent
      await this.sp.web.lists
        .getByTitle(this.POLICY_PACK_ASSIGNMENTS_LIST)
        .items.getById(assignmentId)
        .update({
          WelcomeEmailSent: true,
          WelcomeEmailSentDate: new Date().toISOString()
        });

      logger.info('PolicyPackService', `Sent welcome email for assignment ${assignmentId}`);
    } catch (error) {
      logger.error('PolicyPackService', 'Failed to send welcome email:', error);
      // Still mark as sent to prevent retries - errors are logged
      try {
        await this.sp.web.lists
          .getByTitle(this.POLICY_PACK_ASSIGNMENTS_LIST)
          .items.getById(assignmentId)
          .update({
            WelcomeEmailSent: true,
            WelcomeEmailSentDate: new Date().toISOString()
          });
      } catch { /* ignore */ }
    }
  }

  /**
   * Send Teams notification for policy pack assignment
   */
  private async sendTeamsNotification(assignmentId: number): Promise<void> {
    try {
      // Get assignment details
      const assignment = await this.sp.web.lists
        .getByTitle(this.POLICY_PACK_ASSIGNMENTS_LIST)
        .items.getById(assignmentId)
        .select('*', 'User/EMail', 'User/Title', 'Pack/Title')
        .expand('User', 'Pack')();

      const userEmail = assignment.User?.EMail;
      const userName = assignment.User?.Title || 'User';
      const packName = assignment.Pack?.Title || 'Policy Pack';
      const viewUrl = this.generatePersonalViewURL(assignment.UserId);

      if (userEmail) {
        // Create Teams notification record (to be processed by Power Automate or Graph API)
        await this.sp.web.lists
          .getByTitle('PM_NotificationQueue')
          .items.add({
            Title: `Teams: Policy Pack - ${packName}`,
            RecipientEmail: userEmail,
            RecipientName: userName,
            NotificationType: 'TeamsNotification',
            Subject: `New Policy Pack: ${packName}`,
            Body: JSON.stringify({
              type: 'AdaptiveCard',
              body: [
                { type: 'TextBlock', text: `📋 New Policy Pack Assigned`, weight: 'bolder', size: 'medium' },
                { type: 'TextBlock', text: `Hi ${userName}, you have been assigned the "${packName}" policy pack.`, wrap: true },
                { type: 'TextBlock', text: 'Please review and acknowledge the policies.', wrap: true }
              ],
              actions: [
                { type: 'Action.OpenUrl', title: 'View Policies', url: viewUrl }
              ]
            }),
            Status: 'Pending',
            Priority: 'Normal',
            Channel: 'Teams'
          });
      }

      // Mark notification as sent
      await this.sp.web.lists
        .getByTitle(this.POLICY_PACK_ASSIGNMENTS_LIST)
        .items.getById(assignmentId)
        .update({
          TeamsNotificationSent: true,
          TeamsNotificationSentDate: new Date().toISOString()
        });

      logger.info('PolicyPackService', `Sent Teams notification for assignment ${assignmentId}`);
    } catch (error) {
      logger.error('PolicyPackService', 'Failed to send Teams notification:', error);
      // Still mark as sent to prevent retries
      try {
        await this.sp.web.lists
          .getByTitle(this.POLICY_PACK_ASSIGNMENTS_LIST)
          .items.getById(assignmentId)
          .update({
            TeamsNotificationSent: true,
            TeamsNotificationSentDate: new Date().toISOString()
          });
      } catch { /* ignore */ }
    }
  }

  /**
   * Map policy pack
   */
  private mapPolicyPack(item: any): IPolicyPack {
    return {
      ...item,
      PolicyIds: item.PolicyIds ? JSON.parse(item.PolicyIds) : [],
      TargetDepartments: item.TargetDepartments ? JSON.parse(item.TargetDepartments) : [],
      TargetRoles: item.TargetRoles ? JSON.parse(item.TargetRoles) : [],
      TargetLocations: item.TargetLocations ? JSON.parse(item.TargetLocations) : [],
      PolicySequence: item.PolicySequence ? JSON.parse(item.PolicySequence) : undefined
    } as IPolicyPack;
  }
}
