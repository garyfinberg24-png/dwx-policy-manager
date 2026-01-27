// @ts-nocheck
/**
 * ApprovalActionHandler
 * Handles approval-related workflow actions
 * Creates and manages approval requests within workflow execution
 *
 * INTEGRATION: Now fully integrated with ApprovalNotificationService for:
 * - New approval email notifications
 * - Completion notifications to requesters
 * - Escalation notifications
 * - Multi-level approval progression
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import {
  IStepConfig,
  IActionContext,
  IActionResult
} from '../../../models/IWorkflow';
import { IJmlApproval, ApprovalType } from '../../../models/IJmlApproval';
import { logger } from '../../LoggingService';
import { ApprovalNotificationService } from '../../ApprovalNotificationService';

// Approval status enum matching existing ApprovalService
export enum ApprovalStatus {
  Pending = 'Pending',
  Approved = 'Approved',
  Rejected = 'Rejected',
  Cancelled = 'Cancelled',
  Delegated = 'Delegated',
  Escalated = 'Escalated'
}

/**
 * Result of multi-level approval progression
 */
export interface IApprovalProgressionResult {
  approvalId: number;
  currentLevel: number;
  totalLevels: number;
  isComplete: boolean;
  nextApproverId?: number;
  createdNextLevelApprovalId?: number;
}

export class ApprovalActionHandler {
  private sp: SPFI;
  private notificationService: ApprovalNotificationService | null = null;
  private siteUrl: string = '';

  constructor(sp: SPFI, siteUrl?: string) {
    this.sp = sp;
    if (siteUrl) {
      this.siteUrl = siteUrl;
      this.notificationService = new ApprovalNotificationService(sp, siteUrl);
    }
  }

  /**
   * Initialize notification service (can be called after construction)
   */
  public initializeNotificationService(siteUrl: string): void {
    this.siteUrl = siteUrl;
    this.notificationService = new ApprovalNotificationService(this.sp, siteUrl);
  }

  /**
   * Create an approval request
   */
  public async createApproval(config: IStepConfig, context: IActionContext): Promise<IActionResult> {
    try {
      const approverId = config.approverId;

      if (!approverId) {
        return { success: false, error: 'Approver not specified' };
      }

      // Build approval title
      const title = `Approval Required: ${context.currentStep.name}`;

      // Calculate due date if SLA configured
      let dueDate: Date | undefined;
      if (context.currentStep.sla?.breachHours) {
        dueDate = new Date();
        dueDate.setHours(dueDate.getHours() + context.currentStep.sla.breachHours);
      }

      // Create approval request
      const approvalData = {
        Title: title,
        ProcessId: context.workflowInstance.ProcessId,
        ApproverId: approverId,
        Status: ApprovalStatus.Pending,
        RequestedDate: new Date().toISOString(),
        DueDate: dueDate?.toISOString(),
        WorkflowInstanceId: context.workflowInstance.Id,
        WorkflowStepId: context.currentStep.id,
        ApprovalTemplateId: config.approvalTemplateId,
        Comments: `Workflow approval for step: ${context.currentStep.name}`,
        ApprovalLevel: 1,
        CanDelegate: true,
        CanEscalate: true
      };

      const result = await this.sp.web.lists.getByTitle('JML_Approvals').items.add(approvalData);

      logger.info('ApprovalActionHandler', `Created approval request: ${result.data.Id}`);

      // Create in-app notification for approver
      await this.notifyApprover(approverId, title, context);

      // INTEGRATION FIX: Send rich email notification via ApprovalNotificationService
      if (this.notificationService) {
        try {
          // Build IJmlApproval object for notification service
          const workflowContext = context.workflowInstance.Context
            ? JSON.parse(context.workflowInstance.Context)
            : {};

          const approvalForNotification: Partial<IJmlApproval> = {
            Id: result.data.Id,
            ProcessID: context.workflowInstance.ProcessId,
            ProcessTitle: context.workflowInstance.Title || `Process ${context.workflowInstance.ProcessId}`,
            ProcessType: workflowContext.processType || 'Unknown',
            ApprovalLevel: 1,
            ApprovalSequence: 1,
            ApprovalType: ApprovalType.Sequential,
            Status: ApprovalStatus.Pending,
            ApproverId: typeof approverId === 'number' ? approverId : 0,
            Approver: { Id: typeof approverId === 'number' ? approverId : 0, Title: '' },
            RequestedDate: new Date(),
            DueDate: dueDate || new Date(Date.now() + 7 * 24 * 60 * 60 * 1000), // Default 7 days if not set
            ReasonRequired: false
          };

          await this.notificationService.sendNewApprovalNotification(approvalForNotification);
          logger.info('ApprovalActionHandler', `Sent email notification for approval ${result.data.Id}`);
        } catch (notifyError) {
          // Don't fail the approval creation if notification fails
          logger.warn('ApprovalActionHandler', 'Failed to send email notification for approval', notifyError);
        }
      }

      return {
        success: true,
        nextAction: 'wait',
        waitForItemType: 'approval',
        waitForItemIds: [result.data.Id],
        createdItemIds: [result.data.Id],
        outputVariables: {
          approvalId: result.data.Id,
          approverId
        }
      };
    } catch (error) {
      logger.error('ApprovalActionHandler', 'Error creating approval', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to create approval'
      };
    }
  }

  /**
   * Check approval status
   */
  public async checkApprovalStatus(approvalId: number): Promise<{
    status: ApprovalStatus;
    approved: boolean;
    rejected: boolean;
    comments?: string;
  }> {
    try {
      const approval = await this.sp.web.lists.getByTitle('JML_Approvals').items
        .getById(approvalId)
        .select('Status', 'Comments', 'ApproverComments')();

      return {
        status: approval.Status as ApprovalStatus,
        approved: approval.Status === ApprovalStatus.Approved,
        rejected: approval.Status === ApprovalStatus.Rejected,
        comments: approval.ApproverComments
      };
    } catch (error) {
      logger.error('ApprovalActionHandler', `Error checking approval ${approvalId}`, error);
      throw error;
    }
  }

  /**
   * Complete approval from workflow context
   * INTEGRATION FIX: Now sends completion notification to requester and handles multi-level progression
   */
  public async completeApproval(
    approvalId: number,
    approved: boolean,
    comments?: string,
    userId?: number
  ): Promise<IApprovalProgressionResult | void> {
    try {
      // Get the approval details first (needed for notifications and multi-level handling)
      const approval = await this.sp.web.lists.getByTitle('JML_Approvals').items
        .getById(approvalId)
        .select(
          'Id', 'Title', 'ProcessId', 'Status', 'ApproverId', 'RequestedById',
          'ApprovalLevel', 'TotalLevels', 'NextApproverId', 'WorkflowInstanceId',
          'WorkflowStepId', 'ApprovalTemplateId'
        )();

      // Update the approval status
      await this.sp.web.lists.getByTitle('JML_Approvals').items
        .getById(approvalId)
        .update({
          Status: approved ? ApprovalStatus.Approved : ApprovalStatus.Rejected,
          ApproverComments: comments,
          ResponseDate: new Date().toISOString(),
          RespondedById: userId,
          ActualCompletionDate: new Date().toISOString()
        });

      logger.info('ApprovalActionHandler', `Approval ${approvalId} ${approved ? 'approved' : 'rejected'}`);

      // INTEGRATION FIX: Send completion notification to requester
      if (this.notificationService && approval.RequestedById) {
        try {
          const approvalForNotification: Partial<IJmlApproval> = {
            Id: approvalId,
            ProcessID: approval.ProcessId,
            ProcessTitle: approval.Title || `Process ${approval.ProcessId}`,
            ProcessType: 'Unknown',
            ApprovalLevel: approval.ApprovalLevel || 1,
            ApprovalSequence: 1,
            ApprovalType: ApprovalType.Sequential,
            Status: approved ? ApprovalStatus.Approved : ApprovalStatus.Rejected,
            ApproverId: approval.ApproverId,
            Approver: { Id: approval.ApproverId, Title: '' },
            RequestedDate: new Date(),
            DueDate: new Date(),
            CompletedDate: new Date(),
            Notes: comments,
            ReasonRequired: false
          };

          await this.notificationService.sendCompletionNotification(approvalForNotification, approval.RequestedById);
          logger.info('ApprovalActionHandler', `Sent completion notification for approval ${approvalId}`);
        } catch (notifyError) {
          logger.warn('ApprovalActionHandler', 'Failed to send completion notification', notifyError);
        }
      }

      // INTEGRATION FIX: Handle multi-level approval progression
      if (approved && approval.TotalLevels && approval.ApprovalLevel < approval.TotalLevels && approval.NextApproverId) {
        const progressionResult = await this.progressToNextApprovalLevel(approval);
        return progressionResult;
      }

      // Return progression result for single-level or final level
      return {
        approvalId,
        currentLevel: approval.ApprovalLevel || 1,
        totalLevels: approval.TotalLevels || 1,
        isComplete: true,
        nextApproverId: undefined,
        createdNextLevelApprovalId: undefined
      };
    } catch (error) {
      logger.error('ApprovalActionHandler', `Error completing approval ${approvalId}`, error);
      throw error;
    }
  }

  /**
   * Progress to next level in multi-level approval
   * INTEGRATION FIX: Creates next level approval and sends notification
   */
  private async progressToNextApprovalLevel(
    currentApproval: Record<string, unknown>
  ): Promise<IApprovalProgressionResult> {
    const nextLevel = ((currentApproval.ApprovalLevel as number) || 1) + 1;
    const totalLevels = (currentApproval.TotalLevels as number) || 1;
    const nextApproverId = currentApproval.NextApproverId as number;

    logger.info('ApprovalActionHandler', `Progressing to approval level ${nextLevel} of ${totalLevels}`);

    // Create the next level approval
    const nextApprovalData = {
      Title: `[Level ${nextLevel}] ${currentApproval.Title}`,
      ProcessId: currentApproval.ProcessId,
      ApproverId: nextApproverId,
      Status: ApprovalStatus.Pending,
      RequestedDate: new Date().toISOString(),
      WorkflowInstanceId: currentApproval.WorkflowInstanceId,
      WorkflowStepId: currentApproval.WorkflowStepId,
      ApprovalTemplateId: currentApproval.ApprovalTemplateId,
      ApprovalLevel: nextLevel,
      TotalLevels: totalLevels,
      PreviousApprovalId: currentApproval.Id,
      Comments: `Multi-level approval (${nextLevel} of ${totalLevels})`
    };

    const result = await this.sp.web.lists.getByTitle('JML_Approvals').items.add(nextApprovalData);
    const newApprovalId = result.data.Id;

    logger.info('ApprovalActionHandler', `Created next level approval: ${newApprovalId} (Level ${nextLevel})`);

    // Send notification to next approver
    if (this.notificationService) {
      try {
        const approvalForNotification: Partial<IJmlApproval> = {
          Id: newApprovalId,
          ProcessID: currentApproval.ProcessId as number,
          ProcessTitle: currentApproval.Title as string || `Process ${currentApproval.ProcessId}`,
          ProcessType: 'Unknown',
          ApprovalLevel: nextLevel,
          ApprovalSequence: 1,
          ApprovalType: ApprovalType.Sequential,
          Status: ApprovalStatus.Pending,
          ApproverId: nextApproverId,
          Approver: { Id: nextApproverId, Title: '' },
          RequestedDate: new Date(),
          DueDate: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000),
          ReasonRequired: false
        };

        await this.notificationService.sendNewApprovalNotification(approvalForNotification);
        logger.info('ApprovalActionHandler', `Sent notification for level ${nextLevel} approval`);
      } catch (notifyError) {
        logger.warn('ApprovalActionHandler', 'Failed to send next level approval notification', notifyError);
      }
    }

    // Create in-app notification
    await this.sp.web.lists.getByTitle('JML_Notifications').items.add({
      Title: 'Multi-Level Approval Required',
      Message: `Level ${nextLevel} of ${totalLevels}: ${currentApproval.Title}`,
      Type: 'Approval',
      Priority: 'High',
      IsRead: false,
      RecipientId: nextApproverId,
      RelatedItemType: 'Approval',
      RelatedItemId: currentApproval.ProcessId,
      WorkflowInstanceId: currentApproval.WorkflowInstanceId
    });

    return {
      approvalId: currentApproval.Id as number,
      currentLevel: nextLevel,
      totalLevels,
      isComplete: false,
      nextApproverId,
      createdNextLevelApprovalId: newApprovalId
    };
  }

  /**
   * Handle multi-level approval
   */
  public async createMultiLevelApproval(
    config: IStepConfig,
    context: IActionContext,
    approverIds: number[]
  ): Promise<IActionResult> {
    try {
      if (approverIds.length === 0) {
        return { success: false, error: 'No approvers specified for multi-level approval' };
      }

      const createdApprovalIds: number[] = [];
      const title = `Approval Required: ${context.currentStep.name}`;

      // Create approval for first level
      // Subsequent levels will be triggered when previous level approves
      const firstApprovalData = {
        Title: title,
        ProcessId: context.workflowInstance.ProcessId,
        ApproverId: approverIds[0],
        Status: ApprovalStatus.Pending,
        RequestedDate: new Date().toISOString(),
        WorkflowInstanceId: context.workflowInstance.Id,
        WorkflowStepId: context.currentStep.id,
        ApprovalLevel: 1,
        TotalLevels: approverIds.length,
        NextApproverId: approverIds.length > 1 ? approverIds[1] : undefined,
        Comments: `Multi-level approval (1 of ${approverIds.length})`
      };

      const result = await this.sp.web.lists.getByTitle('JML_Approvals').items.add(firstApprovalData);
      createdApprovalIds.push(result.data.Id);

      // Store pending approvers for subsequent levels
      const pendingApprovers = approverIds.slice(1);

      logger.info('ApprovalActionHandler', `Created multi-level approval with ${approverIds.length} levels`);

      return {
        success: true,
        nextAction: 'wait',
        waitForItemType: 'approval',
        waitForItemIds: createdApprovalIds,
        createdItemIds: createdApprovalIds,
        outputVariables: {
          approvalId: createdApprovalIds[0],
          currentLevel: 1,
          totalLevels: approverIds.length,
          pendingApprovers
        }
      };
    } catch (error) {
      logger.error('ApprovalActionHandler', 'Error creating multi-level approval', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to create multi-level approval'
      };
    }
  }

  /**
   * Escalate approval
   * INTEGRATION FIX: Now sends escalation notification to escalation target
   */
  public async escalateApproval(approvalId: number, escalateToId: number): Promise<void> {
    try {
      const approval = await this.sp.web.lists.getByTitle('JML_Approvals').items
        .getById(approvalId)
        .select('Id', 'Title', 'ApproverId', 'ProcessId', 'ApprovalLevel', 'DueDate', 'RequestedDate', 'EscalationLevel')();

      // Update current approval
      await this.sp.web.lists.getByTitle('JML_Approvals').items
        .getById(approvalId)
        .update({
          Status: ApprovalStatus.Escalated,
          EscalatedToId: escalateToId,
          EscalatedDate: new Date().toISOString()
        });

      // Create new approval for escalation target
      const newEscalationLevel = (approval.EscalationLevel || 0) + 1;
      const escalatedApproval = await this.sp.web.lists.getByTitle('JML_Approvals').items.add({
        Title: `[ESCALATED] ${approval.Title}`,
        ApproverId: escalateToId,
        ProcessId: approval.ProcessId,
        Status: ApprovalStatus.Pending,
        RequestedDate: new Date().toISOString(),
        EscalatedFromId: approvalId,
        EscalationLevel: newEscalationLevel,
        DueDate: new Date(Date.now() + 2 * 24 * 60 * 60 * 1000).toISOString(), // 2 days for escalated approvals
        Comments: 'Escalated from original approver'
      });

      logger.info('ApprovalActionHandler', `Escalated approval ${approvalId} to user ${escalateToId}`);

      // INTEGRATION FIX: Send escalation notification via ApprovalNotificationService
      if (this.notificationService) {
        try {
          const approvalForNotification: Partial<IJmlApproval> = {
            Id: escalatedApproval.data.Id,
            ProcessID: approval.ProcessId,
            ProcessTitle: `[ESCALATED] ${approval.Title || `Process ${approval.ProcessId}`}`,
            ProcessType: 'Unknown',
            ApprovalLevel: approval.ApprovalLevel || 1,
            ApprovalSequence: 1,
            ApprovalType: ApprovalType.Sequential,
            Status: ApprovalStatus.Pending,
            ApproverId: escalateToId,
            Approver: { Id: escalateToId, Title: '' },
            OriginalApprover: { Id: approval.ApproverId, Title: '' },
            RequestedDate: approval.RequestedDate ? new Date(approval.RequestedDate) : new Date(),
            DueDate: approval.DueDate ? new Date(approval.DueDate) : new Date(),
            ReasonRequired: false
          };

          await this.notificationService.sendEscalationNotification(approvalForNotification, escalateToId);
          logger.info('ApprovalActionHandler', `Sent escalation notification to user ${escalateToId}`);
        } catch (notifyError) {
          logger.warn('ApprovalActionHandler', 'Failed to send escalation notification', notifyError);
        }
      }

      // Create in-app notification for escalation target
      await this.sp.web.lists.getByTitle('JML_Notifications').items.add({
        Title: 'Escalated Approval Requires Your Attention',
        Message: `An approval has been escalated to you: ${approval.Title}`,
        Type: 'Approval',
        Priority: 'Urgent',
        IsRead: false,
        RecipientId: escalateToId,
        RelatedItemType: 'Approval',
        RelatedItemId: approval.ProcessId
      });
    } catch (error) {
      logger.error('ApprovalActionHandler', `Error escalating approval ${approvalId}`, error);
      throw error;
    }
  }

  /**
   * Notify approver of pending approval
   */
  private async notifyApprover(approverId: number | string, title: string, context: IActionContext): Promise<void> {
    try {
      // Handle both SharePoint ID (number) and Entra ID (string)
      const notificationData: Record<string, unknown> = {
        Title: 'Approval Required',
        Message: `You have a pending approval: ${title}`,
        Type: 'Approval',
        Priority: 'High',
        IsRead: false,
        RelatedItemType: 'Approval',
        RelatedItemId: context.workflowInstance.ProcessId,
        WorkflowInstanceId: context.workflowInstance.Id
      };

      // Set recipient based on ID type
      if (typeof approverId === 'number') {
        notificationData.RecipientId = approverId;
      } else {
        // For Entra ID (string), store as external identifier
        notificationData.RecipientEntraId = approverId;
      }

      await this.sp.web.lists.getByTitle('JML_Notifications').items.add(notificationData);
    } catch (error) {
      // Non-critical - just log warning
      logger.warn('ApprovalActionHandler', 'Failed to create approval notification', error);
    }
  }
}
