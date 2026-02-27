// Approval Service
// Manages approval workflows for JML processes

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';
import {
  IJmlApproval,
  IJmlApprovalChain,
  IJmlApprovalHistory,
  IJmlApprovalDelegation,
  IJmlApprovalTemplate,
  ApprovalStatus,
  ApprovalType,
  IApprovalRequest,
  IApprovalDecision,
  IApprovalDelegationRequest,
  IApprovalSummary,
  EscalationAction,
  ProcessStatus,
  Priority
} from '../models';
import { WorkflowInstanceStatus } from '../models/IWorkflow';
import { WorkflowInstanceService } from './workflow/WorkflowInstanceService';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';
// INTEGRATION FIX: Add ApprovalNotificationService for email/Teams notifications
import { ApprovalNotificationService } from './ApprovalNotificationService';
import {
  retryWithDLQ,
  PROCESS_SYNC_RETRY_OPTIONS,
  workflowSyncDLQ
} from '../utils/retryUtils';

/**
 * Multi-level approval progression result
 */
export interface IApprovalProgressionResult {
  chainId: number;
  processId: number;
  previousLevel: number;
  currentLevel: number;
  isChainComplete: boolean;
  finalStatus?: ApprovalStatus;
  nextApproverIds?: number[];
  notificationsSent: number;
}

export class ApprovalService {
  private sp: SPFI;
  private readonly APPROVALS_LIST = 'PM_Approvals';
  private readonly APPROVAL_CHAINS_LIST = 'PM_ApprovalChains';
  private readonly APPROVAL_HISTORY_LIST = 'PM_ApprovalHistory';
  private readonly DELEGATIONS_LIST = 'PM_ApprovalDelegations';
  private readonly TEMPLATES_LIST = 'PM_ApprovalTemplates';
  private currentUserId: number = 0;
  private workflowInstanceService: WorkflowInstanceService;
  // INTEGRATION FIX: Add ApprovalNotificationService for email/Teams delivery
  private approvalNotificationService: ApprovalNotificationService | null = null;

  constructor(sp: SPFI, siteUrl?: string) {
    this.sp = sp;
    this.workflowInstanceService = new WorkflowInstanceService(sp);
    // INTEGRATION FIX: Initialize notification service if siteUrl provided
    if (siteUrl) {
      this.approvalNotificationService = new ApprovalNotificationService(sp, siteUrl);
    }
  }

  /**
   * INTEGRATION FIX: Initialize notification service (can be called after construction)
   */
  public initializeNotificationService(siteUrl: string): void {
    this.approvalNotificationService = new ApprovalNotificationService(this.sp, siteUrl);
  }

  /**
   * Initialize service with current user
   */
  public async initialize(): Promise<void> {
    const user = await this.sp.web.currentUser();
    this.currentUserId = user.Id;
  }

  /**
   * Initiate approval process for a process
   */
  public async initiateApproval(request: IApprovalRequest): Promise<IJmlApprovalChain> {
    try {
      let template: IJmlApprovalTemplate | undefined;

      if (request.templateId) {
        template = await this.getTemplateById(request.templateId);
      }

      // Create approval chain
      const chain = await this.sp.web.lists
        .getByTitle(this.APPROVAL_CHAINS_LIST)
        .items.add({
          ProcessID: request.processId,
          ChainName: template?.Title || 'Custom Approval Chain',
          ApprovalType: template?.ApprovalType || ApprovalType.Sequential,
          Levels: JSON.stringify(request.customChain || template?.Levels || []),
          RequireComments: template?.RequireComments || false,
          AllowDelegation: template?.AllowDelegation !== false,
          AutoEscalationDays: template?.AutoEscalationDays || 3,
          EscalationAction: template?.EscalationAction || EscalationAction.Notify,
          CurrentLevel: 1,
          OverallStatus: ApprovalStatus.Pending,
          IsActive: true,
          StartDate: new Date().toISOString()
        });

      // Create first level approvals
      const levels = request.customChain || template?.Levels || [];
      if (levels.length > 0) {
        const firstLevel = levels[0];
        await this.createLevelApprovals(chain.data.Id, request.processId, firstLevel, 1);

        // GAP FIX: Send initial notifications to first level approvers
        // For Sequential, only notify the first approver; for others, notify all
        const isSequential = firstLevel.ApprovalType === ApprovalType.Sequential;
        const approversToNotify = isSequential
          ? [firstLevel.ApproverIds[0]]  // Only first approver for Sequential
          : firstLevel.ApproverIds;      // All approvers for Parallel/FirstApprover

        await this.notifyNextLevelApprovers(
          request.processId,
          chain.data.Id,
          1,
          approversToNotify
        );
      }

      return await this.getChainById(chain.data.Id);
    } catch (error) {
      logger.error('ApprovalService', 'Failed to initiate approval:', error);
      throw error;
    }
  }

  /**
   * Create approvals for a specific level
   * GAP FIX: For Sequential type, only first approver is Pending, others are Queued
   */
  private async createLevelApprovals(
    chainId: number,
    processId: number,
    level: any,
    levelNumber: number
  ): Promise<void> {
    const dueDate = new Date();
    dueDate.setDate(dueDate.getDate() + (level.DueDays || 3));
    const isSequential = level.ApprovalType === ApprovalType.Sequential;

    for (let i = 0; i < level.ApproverIds.length; i++) {
      const approverId = level.ApproverIds[i];

      // Check for active delegations
      const delegatedTo = await this.getActiveDelegation(approverId);

      // GAP FIX: For Sequential, only first approver (i=0) starts as Pending
      // Others are Queued until it's their turn
      // For Parallel/FirstApprover, all start as Pending
      const initialStatus = isSequential && i > 0
        ? ApprovalStatus.Queued
        : ApprovalStatus.Pending;

      await this.sp.web.lists
        .getByTitle(this.APPROVALS_LIST)
        .items.add({
          ProcessID: processId,
          ApprovalChainId: chainId,
          ApprovalLevel: levelNumber,
          ApprovalSequence: i + 1,
          ApprovalType: level.ApprovalType,
          Status: initialStatus,
          ApproverId: delegatedTo || approverId,
          OriginalApproverId: delegatedTo ? approverId : undefined,
          RequestedDate: new Date().toISOString(),
          DueDate: dueDate.toISOString(),
          ReasonRequired: level.ReasonRequired || false,
          IsOverdue: false,
          EscalationLevel: 0
        });
    }

    logger.info('ApprovalService',
      `Created ${level.ApproverIds.length} approvals for level ${levelNumber} (${level.ApprovalType})` +
      (isSequential ? ' - sequential mode, only first approver active' : '')
    );
  }

  /**
   * Get active delegation for user
   */
  private async getActiveDelegation(userId: number): Promise<number | undefined> {
    try {
      // Validate input
      const validUserId = ValidationUtils.validateInteger(userId, 'userId', 1);

      const now = new Date();

      // Build secure filter
      const userFilter = ValidationUtils.buildFilter('DelegatedById', 'eq', validUserId);
      const startDateFilter = ValidationUtils.buildFilter('StartDate', 'le', now.toISOString());
      const endDateFilter = ValidationUtils.buildFilter('EndDate', 'ge', now.toISOString());
      const filter = `${userFilter} and IsActive eq 1 and ${startDateFilter} and ${endDateFilter}`;

      const delegations = await this.sp.web.lists
        .getByTitle(this.DELEGATIONS_LIST)
        .items.filter(filter)
        .top(1)();

      return delegations.length > 0 ? delegations[0].DelegatedToId : undefined;
    } catch (error) {
      logger.error('ApprovalService', 'Failed to get delegation:', error);
      return undefined;
    }
  }

  /**
   * Submit approval decision
   */
  public async submitDecision(decision: IApprovalDecision): Promise<void> {
    try {
      const approval = await this.getApprovalById(decision.approvalId);

      // Update approval
      await this.sp.web.lists
        .getByTitle(this.APPROVALS_LIST)
        .items.getById(decision.approvalId)
        .update({
          Status: decision.decision,
          Decision: decision.decision,
          Comments: decision.comments,
          CompletedDate: new Date().toISOString(),
          ResponseTime: this.calculateResponseTime(approval.RequestedDate)
        });

      // Log history
      await this.logHistory({
        ApprovalId: decision.approvalId,
        ProcessID: approval.ProcessID,
        Action: decision.decision === ApprovalStatus.Approved ? 'Approved' : 'Rejected',
        PerformedById: this.currentUserId,
        ActionDate: new Date().toISOString(),
        Comments: decision.comments,
        PreviousStatus: approval.Status,
        NewStatus: decision.decision
      });

      // Check if level is complete
      await this.checkLevelCompletion(approval.ProcessID, approval.ApprovalLevel);

      // Resume workflow if this approval is linked to a workflow instance
      if (approval.WorkflowInstanceId) {
        await this.resumeWorkflowAfterApproval(
          approval.WorkflowInstanceId,
          approval.WorkflowStepId,
          decision.decision === ApprovalStatus.Approved,
          decision.comments
        );
      }
    } catch (error) {
      logger.error('ApprovalService', 'Failed to submit decision:', error);
      throw error;
    }
  }

  /**
   * Resume workflow after approval decision
   */
  private async resumeWorkflowAfterApproval(
    instanceId: number,
    stepId: string | undefined,
    approved: boolean,
    comments?: string
  ): Promise<void> {
    try {
      // Get workflow instance
      const instance = await this.workflowInstanceService.getById(instanceId);

      if (!instance) {
        logger.warn('ApprovalService', `Workflow instance ${instanceId} not found for approval resumption`);
        return;
      }

      // Only resume if workflow is waiting for approval
      if (instance.Status !== WorkflowInstanceStatus.WaitingForApproval) {
        logger.info('ApprovalService', `Workflow ${instanceId} is not waiting for approval (status: ${instance.Status})`);
        return;
      }

      // Complete the approval step
      if (stepId) {
        await this.workflowInstanceService.completeStep(instanceId, stepId, {
          approved,
          approverComments: comments,
          approvedDate: new Date().toISOString()
        });
      }

      // Update workflow instance to resume execution
      if (approved) {
        // Approved - resume workflow to continue to next step
        await this.workflowInstanceService.update(instanceId, {
          Status: WorkflowInstanceStatus.Running
        });
        await this.workflowInstanceService.addLog(
          instanceId,
          stepId,
          'Approval Step',
          'Approval Completed',
          'Info' as any,
          `Approval was approved. Workflow resuming.`,
          { approved: true, comments }
        );
      } else {
        // Rejected - set to Running so WorkflowResumeService can handle rejection branches
        // CRITICAL FIX: Don't immediately fail - let the workflow engine evaluate rejection transitions
        await this.workflowInstanceService.update(instanceId, {
          Status: WorkflowInstanceStatus.Running
          // Note: Don't set ErrorMessage - rejection is a valid workflow path, not an error
        });

        // Store rejection outcome in step result for workflow transition evaluation
        if (stepId) {
          await this.workflowInstanceService.completeStep(instanceId, stepId, {
            approved: false,
            rejected: true, // Explicit flag for rejection branch evaluation
            rejectionReason: comments,
            rejectedDate: new Date().toISOString()
          });
        }

        await this.workflowInstanceService.addLog(
          instanceId,
          stepId,
          'Approval Step',
          'Approval Rejected',
          'Warning' as any,
          `Approval was rejected: ${comments || 'No reason provided'}. Workflow will evaluate rejection branches.`,
          { approved: false, rejected: true, comments }
        );
      }

      logger.info('ApprovalService', `Workflow ${instanceId} ${approved ? 'resumed after approval' : 'set to running for rejection branch evaluation'}`);
    } catch (error) {
      logger.error('ApprovalService', `Failed to resume workflow ${instanceId} after approval`, error);
      // Don't throw - approval was successful, workflow resumption is secondary
    }
  }

  /**
   * Calculate response time in hours
   */
  private calculateResponseTime(requestedDate: Date | string): number {
    const now = new Date();
    const requested = typeof requestedDate === 'string' ? new Date(requestedDate) : requestedDate;
    return Math.round((now.getTime() - requested.getTime()) / (1000 * 60 * 60));
  }

  /**
   * Check if approval level is complete and advance to next level
   * Enhanced with proper multi-level progression and process sync
   */
  private async checkLevelCompletion(processId: number, level: number): Promise<IApprovalProgressionResult | null> {
    try {
      const chain = await this.getChainByProcessId(processId);
      if (!chain) {
        return null;
      }

      const levelApprovals = await this.getLevelApprovals(processId, level);
      const levels = JSON.parse(chain.Levels as any);
      const currentLevel = levels[level - 1];

      let isLevelComplete = false;
      let isApproved = false;

      if (currentLevel.ApprovalType === ApprovalType.Sequential) {
        // GAP FIX: Sequential - must approve in order, one at a time
        // Check if there are any rejected approvals
        const anyRejected = levelApprovals.some(a => a.Status === ApprovalStatus.Rejected);

        if (anyRejected) {
          // Any rejection fails the level immediately
          isLevelComplete = true;
          isApproved = false;
        } else {
          // Check if there are queued approvers waiting for their turn
          const queuedApprovals = levelApprovals
            .filter(a => a.Status === ApprovalStatus.Queued)
            .sort((a, b) => a.ApprovalSequence - b.ApprovalSequence);

          if (queuedApprovals.length > 0) {
            // There are still queued approvers - check if we need to activate the next one
            const lastApprovedSequence = Math.max(
              ...levelApprovals
                .filter(a => a.Status === ApprovalStatus.Approved)
                .map(a => a.ApprovalSequence),
              0
            );

            // Find the next queued approval that should be activated
            const nextInQueue = queuedApprovals.find(
              a => a.ApprovalSequence === lastApprovedSequence + 1
            );

            if (nextInQueue) {
              // Activate the next approver in sequence
              await this.activateSequentialApprover(
                nextInQueue.Id,
                processId,
                chain.Id,
                level
              );
            }

            // Level not yet complete - still have queued approvers
            isLevelComplete = false;
            isApproved = false;
          } else {
            // No queued approvers - check if all are approved
            const allApproved = levelApprovals.every(
              a => a.Status === ApprovalStatus.Approved
            );
            isLevelComplete = allApproved;
            isApproved = allApproved;
          }
        }
      } else if (currentLevel.ApprovalType === ApprovalType.Parallel) {
        // Parallel - all must respond (approve/reject), notified simultaneously
        const allCompleted = levelApprovals.every(
          a => a.Status === ApprovalStatus.Approved || a.Status === ApprovalStatus.Rejected
        );
        const anyRejected = levelApprovals.some(a => a.Status === ApprovalStatus.Rejected);

        isLevelComplete = allCompleted;
        isApproved = allCompleted && !anyRejected;
      } else if (currentLevel.ApprovalType === ApprovalType.FirstApprover) {
        // First to respond wins
        const anyCompleted = levelApprovals.some(
          a => a.Status === ApprovalStatus.Approved || a.Status === ApprovalStatus.Rejected
        );

        isLevelComplete = anyCompleted;
        isApproved = levelApprovals.some(a => a.Status === ApprovalStatus.Approved);
      }

      const result: IApprovalProgressionResult = {
        chainId: chain.Id,
        processId,
        previousLevel: level,
        currentLevel: level,
        isChainComplete: false,
        notificationsSent: 0
      };

      if (isLevelComplete) {
        if (isApproved && level < levels.length) {
          // Move to next level
          const nextLevel = level + 1;
          result.currentLevel = nextLevel;
          const nextLevelConfig = levels[level]; // levels is 0-indexed, level is 1-indexed

          await this.sp.web.lists
            .getByTitle(this.APPROVAL_CHAINS_LIST)
            .items.getById(chain.Id)
            .update({
              CurrentLevel: nextLevel
            });

          // Create next level approvals
          await this.createLevelApprovals(chain.Id, processId, nextLevelConfig, nextLevel);

          // GAP FIX: For Sequential, only notify the first approver (others are Queued)
          // For Parallel/FirstApprover, notify all approvers
          const isNextLevelSequential = nextLevelConfig.ApprovalType === ApprovalType.Sequential;
          const approversToNotify = isNextLevelSequential
            ? [nextLevelConfig.ApproverIds[0]] // Only first approver for Sequential
            : nextLevelConfig.ApproverIds;     // All approvers for Parallel/FirstApprover

          // Get next level approver IDs for result
          result.nextApproverIds = approversToNotify;

          // Send notifications to active approvers
          result.notificationsSent = await this.notifyNextLevelApprovers(
            processId,
            chain.Id,
            nextLevel,
            approversToNotify
          );

          logger.info('ApprovalService',
            `Approval chain ${chain.Id} progressed from level ${level} to ${nextLevel}` +
            (isNextLevelSequential ? ' (sequential - first approver notified)' : ''));
        } else {
          // Chain complete - update chain status
          const finalStatus = isApproved ? ApprovalStatus.Approved : ApprovalStatus.Rejected;
          result.isChainComplete = true;
          result.finalStatus = finalStatus;

          await this.sp.web.lists
            .getByTitle(this.APPROVAL_CHAINS_LIST)
            .items.getById(chain.Id)
            .update({
              OverallStatus: finalStatus,
              CompletedDate: new Date().toISOString(),
              IsActive: false
            });

          // Sync process status based on approval outcome
          await this.syncProcessStatusFromApproval(processId, finalStatus);

          // Send completion notification
          await this.notifyApprovalChainComplete(processId, chain.Id, finalStatus);

          logger.info('ApprovalService',
            `Approval chain ${chain.Id} completed with status: ${finalStatus}`);
        }
      }

      return result;
    } catch (error) {
      logger.error('ApprovalService', 'Failed to check level completion:', error);
      return null;
    }
  }

  /**
   * Sync process status when approval chain completes
   * Uses retry with DLQ for reliability
   */
  private async syncProcessStatusFromApproval(
    processId: number,
    approvalStatus: ApprovalStatus
  ): Promise<void> {
    const processStatusMap: Record<ApprovalStatus, ProcessStatus | null> = {
      [ApprovalStatus.Approved]: ProcessStatus.InProgress, // Continue processing
      [ApprovalStatus.Rejected]: ProcessStatus.OnHold, // Put on hold for review
      [ApprovalStatus.Pending]: null,
      [ApprovalStatus.Delegated]: null,
      [ApprovalStatus.Escalated]: ProcessStatus.PendingApproval, // Still pending
      [ApprovalStatus.Cancelled]: ProcessStatus.Cancelled,
      [ApprovalStatus.Skipped]: ProcessStatus.InProgress,
      [ApprovalStatus.Queued]: null, // GAP FIX: Queued approvals don't affect process status
      [ApprovalStatus.Expired]: ProcessStatus.OnHold // GAP FIX: Expired approvals put process on hold
    };

    const newProcessStatus = processStatusMap[approvalStatus];
    if (!newProcessStatus) {
      return;
    }

    const syncPayload = {
      processId,
      approvalStatus,
      newProcessStatus,
      timestamp: new Date().toISOString()
    };

    const result = await retryWithDLQ<void>(
      async () => {
        await this.sp.web.lists
          .getByTitle('PM_Processes')
          .items
          .getById(processId)
          .update({
            ProcessStatus: newProcessStatus,
            ApprovalStatus: approvalStatus,
            ...(approvalStatus === ApprovalStatus.Approved ? {
              LastApprovalDate: new Date().toISOString()
            } : {}),
            ...(approvalStatus === ApprovalStatus.Rejected ? {
              RejectionDate: new Date().toISOString()
            } : {})
          });
      },
      'approval-to-process-sync',
      syncPayload,
      PROCESS_SYNC_RETRY_OPTIONS,
      workflowSyncDLQ,
      {
        source: 'ApprovalService',
        operation: 'syncProcessStatusFromApproval'
      }
    );

    if (!result.success) {
      logger.error('ApprovalService',
        `Failed to sync process ${processId} from approval status ${approvalStatus}. DLQ ID: ${result.deadLetterItemId}`);
    } else {
      logger.info('ApprovalService',
        `Process ${processId} synced to ${newProcessStatus} from approval ${approvalStatus}`);
    }
  }

  /**
   * Notify next level approvers that their approval is required
   * INTEGRATION FIX: Now also sends email/Teams notifications via ApprovalNotificationService
   */
  private async notifyNextLevelApprovers(
    processId: number,
    chainId: number,
    levelNumber: number,
    approverIds: number[]
  ): Promise<number> {
    let notificationsSent = 0;

    try {
      // Get process details for notification context
      const process = await this.sp.web.lists
        .getByTitle('PM_Processes')
        .items
        .getById(processId)
        .select('Id', 'Title', 'ProcessType', 'EmployeeName')() as {
          Id: number;
          Title: string;
          ProcessType: string;
          EmployeeName: string;
        };

      for (const approverId of approverIds) {
        try {
          // Check for delegation
          const actualApproverId = await this.getActiveDelegation(approverId) || approverId;

          // Create in-app notification
          await this.sp.web.lists.getByTitle('PM_Notifications').items.add({
            Title: 'Approval Required',
            Message: `Your approval is required for ${process.ProcessType} process: ${process.EmployeeName}. This is level ${levelNumber} of the approval chain.`,
            RecipientId: actualApproverId,
            NotificationType: 'ApprovalRequired',
            Priority: Priority.High,
            LinkUrl: `/sites/JML/SitePages/ApprovalCenter.aspx?processId=${processId}`,
            ProcessId: processId.toString(),
            IsRead: false,
            SentDate: new Date(),
            ExpirationDate: this.calculateApprovalDueDate(3) // 3 days default
          });

          // INTEGRATION FIX: Also send email/Teams notification via ApprovalNotificationService
          if (this.approvalNotificationService) {
            try {
              const dueDate = this.calculateApprovalDueDate(3);
              const approvalForNotification: Partial<IJmlApproval> = {
                Id: chainId, // Use chain ID for tracking
                ProcessID: processId,
                ProcessTitle: process.Title || `${process.ProcessType}: ${process.EmployeeName}`,
                ProcessType: process.ProcessType,
                ApprovalLevel: levelNumber,
                ApprovalSequence: 1,
                ApprovalType: ApprovalType.Sequential,
                Status: ApprovalStatus.Pending,
                ApproverId: actualApproverId,
                Approver: { Id: actualApproverId, Title: '' },
                RequestedDate: new Date(),
                DueDate: dueDate,
                ReasonRequired: false
              };

              await this.approvalNotificationService.sendNewApprovalNotification(approvalForNotification);
              logger.info('ApprovalService',
                `Sent email notification to approver ${actualApproverId} for level ${levelNumber}`);
            } catch (emailError) {
              // Don't fail if email notification fails - in-app notification was successful
              logger.warn('ApprovalService',
                `Failed to send email notification to approver ${actualApproverId}`, emailError);
            }
          }

          notificationsSent++;
        } catch (notifyError) {
          logger.warn('ApprovalService',
            `Failed to notify approver ${approverId} for process ${processId}`, notifyError);
        }
      }
    } catch (error) {
      logger.error('ApprovalService',
        `Failed to notify next level approvers for chain ${chainId}`, error);
    }

    return notificationsSent;
  }

  /**
   * Notify stakeholders that approval chain is complete
   * INTEGRATION FIX: Now also sends email/Teams notifications via ApprovalNotificationService
   */
  private async notifyApprovalChainComplete(
    processId: number,
    chainId: number,
    finalStatus: ApprovalStatus
  ): Promise<void> {
    try {
      // Get process details
      const process = await this.sp.web.lists
        .getByTitle('PM_Processes')
        .items
        .getById(processId)
        .select('Id', 'Title', 'ProcessType', 'EmployeeName', 'ManagerId', 'CreatedById')() as {
          Id: number;
          Title: string;
          ProcessType: string;
          EmployeeName: string;
          ManagerId?: number;
          CreatedById: number;
        };

      const isApproved = finalStatus === ApprovalStatus.Approved;
      const statusText = isApproved ? 'approved' : 'rejected';
      const recipientIds = new Set<number>();

      // Notify manager
      if (process.ManagerId) {
        recipientIds.add(process.ManagerId);
      }

      // Notify process creator
      if (process.CreatedById) {
        recipientIds.add(process.CreatedById);
      }

      for (const recipientId of Array.from(recipientIds)) {
        try {
          // Create in-app notification
          await this.sp.web.lists.getByTitle('PM_Notifications').items.add({
            Title: `Approval ${isApproved ? 'Completed' : 'Rejected'}`,
            Message: `The approval chain for ${process.ProcessType} process (${process.EmployeeName}) has been ${statusText}.`,
            RecipientId: recipientId,
            NotificationType: isApproved ? 'ApprovalComplete' : 'ApprovalRejected',
            Priority: isApproved ? Priority.Medium : Priority.High,
            LinkUrl: `/sites/JML/SitePages/ProcessDetails.aspx?processId=${processId}`,
            ProcessId: processId.toString(),
            IsRead: false,
            SentDate: new Date()
          });

          // INTEGRATION FIX: Send email/Teams notification via ApprovalNotificationService
          if (this.approvalNotificationService) {
            try {
              const approvalForNotification: Partial<IJmlApproval> = {
                Id: chainId,
                ProcessID: processId,
                ProcessTitle: process.Title || `${process.ProcessType}: ${process.EmployeeName}`,
                ProcessType: process.ProcessType,
                ApprovalLevel: 1,
                ApprovalSequence: 1,
                ApprovalType: ApprovalType.Sequential,
                Status: finalStatus,
                ApproverId: recipientId, // Use recipient as approver for notification purposes
                Approver: { Id: recipientId, Title: '' },
                RequestedDate: new Date(),
                DueDate: new Date(),
                CompletedDate: new Date(),
                ReasonRequired: false
              };

              await this.approvalNotificationService.sendCompletionNotification(
                approvalForNotification,
                recipientId
              );
              logger.info('ApprovalService',
                `Sent chain completion email to recipient ${recipientId}`);
            } catch (emailError) {
              logger.warn('ApprovalService',
                `Failed to send chain completion email to ${recipientId}`, emailError);
            }
          }
        } catch (notifyError) {
          logger.warn('ApprovalService',
            `Failed to notify recipient ${recipientId} of chain completion`, notifyError);
        }
      }
    } catch (error) {
      logger.error('ApprovalService',
        `Failed to send chain completion notifications for process ${processId}`, error);
    }
  }

  /**
   * Calculate approval due date
   */
  private calculateApprovalDueDate(days: number): Date {
    const dueDate = new Date();
    dueDate.setDate(dueDate.getDate() + days);
    return dueDate;
  }

  /**
   * Get the current approval status for a process
   */
  public async getProcessApprovalStatus(processId: number): Promise<{
    hasActiveChain: boolean;
    chainId?: number;
    currentLevel: number;
    totalLevels: number;
    overallStatus: ApprovalStatus;
    pendingApprovers: Array<{ id: number; name: string; level: number }>;
    completedLevels: number[];
  }> {
    const chain = await this.getChainByProcessId(processId);

    if (!chain) {
      return {
        hasActiveChain: false,
        currentLevel: 0,
        totalLevels: 0,
        overallStatus: ApprovalStatus.Pending,
        pendingApprovers: [],
        completedLevels: []
      };
    }

    const levels = typeof chain.Levels === 'string'
      ? JSON.parse(chain.Levels)
      : chain.Levels;

    // Get pending approvals
    const pendingApprovals = await this.sp.web.lists
      .getByTitle(this.APPROVALS_LIST)
      .items
      .filter(`ProcessID eq ${processId} and Status eq '${ApprovalStatus.Pending}'`)
      .select('Id', 'ApproverId', 'ApprovalLevel', 'Approver/Title')
      .expand('Approver')() as Array<{
        Id: number;
        ApproverId: number;
        ApprovalLevel: number;
        Approver?: { Title: string };
      }>;

    // Determine completed levels
    const completedLevels: number[] = [];
    for (let i = 1; i < chain.CurrentLevel; i++) {
      completedLevels.push(i);
    }

    return {
      hasActiveChain: chain.IsActive,
      chainId: chain.Id,
      currentLevel: chain.CurrentLevel,
      totalLevels: levels.length,
      overallStatus: chain.OverallStatus,
      pendingApprovers: pendingApprovals.map(a => ({
        id: a.ApproverId,
        name: a.Approver?.Title || 'Unknown',
        level: a.ApprovalLevel
      })),
      completedLevels
    };
  }

  /**
   * Get approvals for a specific level
   */
  private async getLevelApprovals(processId: number, level: number): Promise<IJmlApproval[]> {
    const items = await this.sp.web.lists
      .getByTitle(this.APPROVALS_LIST)
      .items.filter(`ProcessID eq ${processId} and ApprovalLevel eq ${level}`)
      .select(
        '*',
        'Approver/Title',
        'Approver/EMail',
        'OriginalApprover/Title',
        'OriginalApprover/EMail',
        'ModifiedBy/Title',
        'ModifiedBy/EMail'
      )
      .expand('Approver', 'OriginalApprover', 'ModifiedBy')();

    const approvals: IJmlApproval[] = [];
    for (let i = 0; i < items.length; i++) {
      approvals.push(this.mapToApproval(items[i]));
    }

    return approvals;
  }

  /**
   * Delegate approval to another user
   * GAP FIX: Now notifies the delegate via in-app and email/Teams notifications
   */
  public async delegateApproval(
    approvalId: number,
    delegateToId: number,
    reason: string
  ): Promise<void> {
    try {
      const approval = await this.getApprovalById(approvalId);

      await this.sp.web.lists
        .getByTitle(this.APPROVALS_LIST)
        .items.getById(approvalId)
        .update({
          OriginalApproverId: approval.ApproverId,
          ApproverId: delegateToId,
          DelegatedById: this.currentUserId,
          Status: ApprovalStatus.Delegated
        });

      await this.logHistory({
        ApprovalId: approvalId,
        ProcessID: approval.ProcessID,
        Action: 'Delegated',
        PerformedById: this.currentUserId,
        ActionDate: new Date().toISOString(),
        Comments: reason,
        PreviousStatus: approval.Status,
        NewStatus: ApprovalStatus.Delegated,
        DelegatedToId: delegateToId,
        DelegationReason: reason
      });

      // GAP FIX: Send notification to the delegate
      await this.sendDelegationNotifications(approval, delegateToId, reason);
    } catch (error) {
      logger.error('ApprovalService', 'Failed to delegate approval:', error);
      throw error;
    }
  }

  /**
   * Send delegation notifications to the delegate
   * GAP FIX: Notifies delegate that they have a new approval to review
   */
  private async sendDelegationNotifications(
    approval: IJmlApproval,
    delegateToId: number,
    reason: string
  ): Promise<void> {
    try {
      // Get process details for notification context
      const process = await this.sp.web.lists
        .getByTitle('PM_Processes')
        .items
        .getById(approval.ProcessID)
        .select('Id', 'Title', 'ProcessType', 'EmployeeName')() as {
          Id: number;
          Title: string;
          ProcessType: string;
          EmployeeName: string;
        };

      // Get original approver name for context
      let originalApproverName = 'Another user';
      try {
        const originalApprover = await this.sp.web.getUserById(approval.ApproverId)();
        originalApproverName = originalApprover.Title || originalApproverName;
      } catch {
        // Use default if user lookup fails
      }

      // Send in-app notification to the delegate
      await this.sp.web.lists.getByTitle('PM_Notifications').items.add({
        Title: 'Approval Delegated to You',
        Message: `${originalApproverName} has delegated an approval for ${process.ProcessType} process (${process.EmployeeName}) to you.${reason ? ` Reason: ${reason}` : ''} Please review and respond.`,
        RecipientId: delegateToId,
        NotificationType: 'ApprovalDelegated',
        Priority: Priority.High,
        LinkUrl: `/sites/JML/SitePages/ApprovalCenter.aspx?processId=${approval.ProcessID}`,
        ProcessId: approval.ProcessID.toString(),
        IsRead: false,
        SentDate: new Date(),
        ExpirationDate: approval.DueDate // Use original due date
      });

      // Also notify original approver that delegation was successful
      await this.sp.web.lists.getByTitle('PM_Notifications').items.add({
        Title: 'Approval Successfully Delegated',
        Message: `Your approval for ${process.ProcessType} process (${process.EmployeeName}) has been delegated successfully.`,
        RecipientId: approval.ApproverId,
        NotificationType: 'ApprovalDelegated',
        Priority: Priority.Medium,
        LinkUrl: `/sites/JML/SitePages/ProcessDetails.aspx?processId=${approval.ProcessID}`,
        ProcessId: approval.ProcessID.toString(),
        IsRead: false,
        SentDate: new Date()
      });

      // Send email/Teams notification via ApprovalNotificationService
      if (this.approvalNotificationService) {
        try {
          // Create approval object for notification service
          const delegatedApproval: Partial<IJmlApproval> = {
            ...approval,
            ApproverId: delegateToId,
            Approver: { Id: delegateToId, Title: '' },
            OriginalApproverId: approval.ApproverId,
            Status: ApprovalStatus.Delegated
          };

          await this.approvalNotificationService.sendDelegationNotification(
            delegatedApproval,
            delegateToId,
            reason
          );
          logger.info('ApprovalService',
            `Sent delegation email notification to delegate ${delegateToId}`);
        } catch (emailError) {
          // Don't fail if email notification fails - in-app notification was successful
          logger.warn('ApprovalService',
            `Failed to send delegation email notification`, emailError);
        }
      }

      logger.info('ApprovalService',
        `Sent delegation notifications for approval ${approval.Id} to delegate ${delegateToId}`);
    } catch (error) {
      // Log but don't fail - the delegation itself was successful
      logger.warn('ApprovalService',
        `Failed to send delegation notifications for approval ${approval.Id}`, error);
    }
  }

  /**
   * Create delegation rule
   */
  public async createDelegation(request: IApprovalDelegationRequest): Promise<IJmlApprovalDelegation> {
    try {
      const result = await this.sp.web.lists
        .getByTitle(this.DELEGATIONS_LIST)
        .items.add({
          DelegatedById: this.currentUserId,
          DelegatedToId: request.delegateToId,
          StartDate: request.startDate.toISOString(),
          EndDate: request.endDate.toISOString(),
          IsActive: true,
          Reason: request.reason,
          ProcessTypes: request.processTypes ? JSON.stringify(request.processTypes) : undefined,
          AutoDelegate: request.autoDelegate
        });

      return await this.getDelegationById(result.data.Id);
    } catch (error) {
      logger.error('ApprovalService', 'Failed to create delegation:', error);
      throw error;
    }
  }

  /**
   * Get my pending approvals
   */
  public async getMyPendingApprovals(): Promise<IJmlApproval[]> {
    try {
      // Build secure filter
      const approverFilter = ValidationUtils.buildFilter('ApproverId', 'eq', this.currentUserId);
      const statusFilter = ValidationUtils.buildFilter('Status', 'eq', ApprovalStatus.Pending);
      const filter = `${approverFilter} and ${statusFilter}`;

      const items = await this.sp.web.lists
        .getByTitle(this.APPROVALS_LIST)
        .items.filter(filter)
        .select(
          '*',
          'Approver/Title',
          'Approver/EMail',
          'OriginalApprover/Title',
          'OriginalApprover/EMail',
          'ModifiedBy/Title',
          'ModifiedBy/EMail'
        )
        .expand('Approver', 'OriginalApprover', 'ModifiedBy')
        .orderBy('DueDate', true)();

      const approvals: IJmlApproval[] = [];
      for (let i = 0; i < items.length; i++) {
        approvals.push(this.mapToApproval(items[i]));
      }

      return approvals;
    } catch (error) {
      logger.error('ApprovalService', 'Failed to get pending approvals:', error);
      return [];
    }
  }

  /**
   * Cancel all pending approvals for a process
   * GAP FIX: Prevents orphaned approvals when a process is cancelled
   */
  public async cancelPendingApprovalsForProcess(
    processId: number,
    reason?: string
  ): Promise<{ cancelled: number; errors: number }> {
    const result = { cancelled: 0, errors: 0 };

    try {
      // Get all pending approvals for this process
      const pendingApprovals = await this.sp.web.lists
        .getByTitle(this.APPROVALS_LIST)
        .items
        .filter(`ProcessID eq ${processId} and (Status eq '${ApprovalStatus.Pending}' or Status eq '${ApprovalStatus.Escalated}' or Status eq '${ApprovalStatus.Delegated}')`)
        .select('Id', 'Title', 'Status', 'ApproverId')();

      if (pendingApprovals.length === 0) {
        logger.info('ApprovalService', `No pending approvals to cancel for process ${processId}`);
        return result;
      }

      // Cancel each pending approval
      for (const approval of pendingApprovals) {
        try {
          await this.sp.web.lists
            .getByTitle(this.APPROVALS_LIST)
            .items.getById(approval.Id)
            .update({
              Status: ApprovalStatus.Cancelled,
              Comments: reason || 'Cancelled due to process cancellation',
              CompletedDate: new Date().toISOString()
            });

          // Log history
          await this.logHistory({
            ApprovalId: approval.Id,
            ProcessID: processId,
            Action: 'Cancelled',
            PerformedById: -1, // System
            ActionDate: new Date().toISOString(),
            PreviousStatus: approval.Status,
            NewStatus: ApprovalStatus.Cancelled,
            Comments: reason || 'Process cancelled'
          });

          // Notify the approver that their approval was cancelled
          try {
            await this.sp.web.lists.getByTitle('PM_Notifications').items.add({
              Title: 'Approval Cancelled',
              Message: `The approval request "${approval.Title}" has been cancelled because the process was cancelled.${reason ? ` Reason: ${reason}` : ''}`,
              RecipientId: approval.ApproverId,
              NotificationType: 'ApprovalCancelled',
              Priority: Priority.Medium,
              ProcessId: processId.toString(),
              IsRead: false,
              SentDate: new Date()
            });
          } catch (notifyError) {
            logger.warn('ApprovalService', `Failed to notify approver of cancellation for approval ${approval.Id}`, notifyError);
          }

          result.cancelled++;
        } catch (approvalError) {
          logger.error('ApprovalService', `Failed to cancel approval ${approval.Id}`, approvalError);
          result.errors++;
        }
      }

      // Also cancel any active approval chains for this process
      try {
        const activeChains = await this.sp.web.lists
          .getByTitle(this.APPROVAL_CHAINS_LIST)
          .items
          .filter(`ProcessID eq ${processId} and IsActive eq 1`)
          .select('Id')();

        for (const chain of activeChains) {
          await this.sp.web.lists
            .getByTitle(this.APPROVAL_CHAINS_LIST)
            .items.getById(chain.Id)
            .update({
              IsActive: false,
              OverallStatus: ApprovalStatus.Cancelled,
              CompletedDate: new Date().toISOString()
            });
        }

        if (activeChains.length > 0) {
          logger.info('ApprovalService', `Cancelled ${activeChains.length} approval chains for process ${processId}`);
        }
      } catch (chainError) {
        logger.warn('ApprovalService', `Failed to cancel approval chains for process ${processId}`, chainError);
      }

      logger.info('ApprovalService',
        `Cancelled ${result.cancelled} pending approvals for process ${processId}` +
        (result.errors > 0 ? ` (${result.errors} errors)` : '')
      );

      return result;
    } catch (error) {
      logger.error('ApprovalService', `Failed to cancel pending approvals for process ${processId}`, error);
      return result;
    }
  }

  /**
   * Get approval history for process
   */
  public async getApprovalHistory(processId: number): Promise<IJmlApprovalHistory[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.APPROVAL_HISTORY_LIST)
        .items.filter(`ProcessID eq ${processId}`)
        .select(
          '*',
          'PerformedBy/Title',
          'PerformedBy/EMail',
          'DelegatedTo/Title',
          'DelegatedTo/EMail',
          'EscalatedTo/Title',
          'EscalatedTo/EMail'
        )
        .expand('PerformedBy', 'DelegatedTo', 'EscalatedTo')
        .orderBy('ActionDate', false)();

      const history: IJmlApprovalHistory[] = [];
      for (let i = 0; i < items.length; i++) {
        history.push(this.mapToHistory(items[i]));
      }

      return history;
    } catch (error) {
      logger.error('ApprovalService', 'Failed to get approval history:', error);
      return [];
    }
  }

  /**
   * Get approval summary
   */
  public async getApprovalSummary(): Promise<IApprovalSummary> {
    try {
      const allApprovals: Array<{ Status: string; ApproverId: number; IsOverdue: boolean; ResponseTime: number }> =
        await this.sp.web.lists
          .getByTitle(this.APPROVALS_LIST)
          .items.select('Status', 'ApproverId', 'IsOverdue', 'ResponseTime')();

      const myPending = allApprovals.filter(
        a => a.ApproverId === this.currentUserId && a.Status === ApprovalStatus.Pending
      );

      const summary: IApprovalSummary = {
        totalPending: allApprovals.filter(a => a.Status === ApprovalStatus.Pending).length,
        totalApproved: allApprovals.filter(a => a.Status === ApprovalStatus.Approved).length,
        totalRejected: allApprovals.filter(a => a.Status === ApprovalStatus.Rejected).length,
        overdueCount: allApprovals.filter(a => a.IsOverdue).length,
        avgResponseTime: this.calculateAvgResponseTime(allApprovals),
        myPendingCount: myPending.length,
        delegatedCount: allApprovals.filter(a => a.Status === ApprovalStatus.Delegated).length
      };

      return summary;
    } catch (error) {
      logger.error('ApprovalService', 'Failed to get approval summary:', error);
      throw error;
    }
  }

  /**
   * Calculate average response time
   */
  private calculateAvgResponseTime(approvals: any[]): number {
    const completed = approvals.filter(a => a.ResponseTime > 0);
    if (completed.length === 0) {
      return 0;
    }

    let total = 0;
    for (let i = 0; i < completed.length; i++) {
      total += completed[i].ResponseTime;
    }

    return Math.round(total / completed.length);
  }

  /**
   * Check for overdue approvals and escalate
   */
  public async processEscalations(): Promise<void> {
    try {
      // Build secure filter
      const statusFilter = ValidationUtils.buildFilter('Status', 'eq', ApprovalStatus.Pending);
      const filter = `${statusFilter} and IsOverdue eq 0`;

      const pendingApprovals = await this.sp.web.lists
        .getByTitle(this.APPROVALS_LIST)
        .items.filter(filter)
        .select('*', 'ApprovalChainId')();

      for (let i = 0; i < pendingApprovals.length; i++) {
        const approval = pendingApprovals[i];
        const dueDate = new Date(approval.DueDate);
        const now = new Date();

        if (now > dueDate) {
          await this.escalateApproval(approval.Id);
        }
      }
    } catch (error) {
      logger.error('ApprovalService', 'Failed to process escalations:', error);
    }
  }

  /**
   * GAP FIX: Process approvals that have exceeded maximum allowed time
   * Default max age is 30 days after creation (configurable per chain)
   */
  public async processExpiredApprovals(maxAgeDays: number = 30): Promise<{
    expired: number;
    errors: number;
  }> {
    const result = { expired: 0, errors: 0 };

    try {
      // Get all pending/escalated/delegated approvals that haven't been completed

      const cutoffDate = new Date();
      cutoffDate.setDate(cutoffDate.getDate() - maxAgeDays);

      const approvals = await this.sp.web.lists
        .getByTitle(this.APPROVALS_LIST)
        .items
        .filter(
          `(Status eq '${ApprovalStatus.Pending}' or ` +
          `Status eq '${ApprovalStatus.Escalated}' or ` +
          `Status eq '${ApprovalStatus.Delegated}') and ` +
          `RequestedDate le datetime'${cutoffDate.toISOString()}'`
        )
        .select('Id', 'Title', 'ProcessID', 'ApproverId', 'ApprovalChainId', 'RequestedDate', 'Status')();

      for (const approval of approvals) {
        try {
          await this.expireApproval(approval);
          result.expired++;
        } catch (error) {
          logger.error('ApprovalService',
            `Failed to expire approval ${approval.Id}`, error);
          result.errors++;
        }
      }

      if (result.expired > 0) {
        logger.info('ApprovalService',
          `Expired ${result.expired} approvals older than ${maxAgeDays} days` +
          (result.errors > 0 ? ` (${result.errors} errors)` : '')
        );
      }

      return result;
    } catch (error) {
      logger.error('ApprovalService', 'Failed to process expired approvals:', error);
      return result;
    }
  }

  /**
   * GAP FIX: Mark an individual approval as expired
   */
  private async expireApproval(approval: {
    Id: number;
    Title: string;
    ProcessID: number;
    ApproverId: number;
    ApprovalChainId: number;
    Status: string;
  }): Promise<void> {
    // Update approval to Expired status
    await this.sp.web.lists
      .getByTitle(this.APPROVALS_LIST)
      .items.getById(approval.Id)
      .update({
        Status: ApprovalStatus.Expired,
        Comments: 'Automatically expired due to exceeding maximum approval time',
        CompletedDate: new Date().toISOString()
      });

    // Log history
    await this.logHistory({
      ApprovalId: approval.Id,
      ProcessID: approval.ProcessID,
      Action: 'Expired',
      PerformedById: -1, // System
      ActionDate: new Date().toISOString(),
      PreviousStatus: approval.Status,
      NewStatus: ApprovalStatus.Expired,
      Comments: 'Approval expired due to exceeding maximum allowed time'
    });

    // Notify the approver
    try {
      await this.sp.web.lists.getByTitle('PM_Notifications').items.add({
        Title: 'Approval Expired',
        Message: `Your approval request "${approval.Title}" has expired because it was not completed within the allowed time period.`,
        RecipientId: approval.ApproverId,
        NotificationType: 'ApprovalExpired',
        Priority: Priority.High,
        ProcessId: approval.ProcessID.toString(),
        IsRead: false,
        SentDate: new Date()
      });
    } catch (notifyError) {
      logger.warn('ApprovalService',
        `Failed to notify approver of expiration for approval ${approval.Id}`, notifyError);
    }

    // Check if this affects the approval chain
    // If all approvals at the level are either expired or rejected, fail the level
    await this.checkLevelExpiration(approval.ProcessID, approval.ApprovalChainId);
  }

  /**
   * GAP FIX: Check if expired approvals have caused a level to fail
   */
  private async checkLevelExpiration(processId: number, chainId: number): Promise<void> {
    try {
      const chain = await this.getChainById(chainId);
      if (!chain || !chain.IsActive) {
        return;
      }

      // Get all approvals at the current level
      const levelApprovals = await this.getLevelApprovals(processId, chain.CurrentLevel);

      // Check if all approvals are in terminal states (no pending/escalated/delegated)
      const allTerminal = levelApprovals.every(
        a => a.Status === ApprovalStatus.Approved ||
             a.Status === ApprovalStatus.Rejected ||
             a.Status === ApprovalStatus.Expired ||
             a.Status === ApprovalStatus.Cancelled
      );

      if (!allTerminal) {
        // Still have active approvals - don't fail the chain yet
        return;
      }

      // Check if any were approved
      const anyApproved = levelApprovals.some(a => a.Status === ApprovalStatus.Approved);

      if (!anyApproved) {
        // No approvals succeeded and all are terminal - fail the chain
        await this.sp.web.lists
          .getByTitle(this.APPROVAL_CHAINS_LIST)
          .items.getById(chainId)
          .update({
            OverallStatus: ApprovalStatus.Expired,
            CompletedDate: new Date().toISOString(),
            IsActive: false
          });

        // Notify stakeholders
        await this.notifyApprovalChainComplete(processId, chainId, ApprovalStatus.Rejected);

        logger.info('ApprovalService',
          `Approval chain ${chainId} failed due to all approvals expiring at level ${chain.CurrentLevel}`);
      }
    } catch (error) {
      logger.error('ApprovalService',
        `Failed to check level expiration for chain ${chainId}`, error);
    }
  }

  /**
   * Escalate overdue approval
   * GAP FIX: Now implements EscalationAction and sends notifications
   */
  private async escalateApproval(approvalId: number): Promise<void> {
    try {
      const approval = await this.getApprovalById(approvalId);
      const chain = await this.getChainByProcessId(approval.ProcessID);

      if (!chain) {
        return;
      }

      // Get the escalation action from the chain configuration
      const escalationAction = chain.EscalationAction || EscalationAction.Notify;
      let newApproverId: number | null = null;
      let autoApproved = false;

      // GAP FIX: Implement EscalationAction logic
      switch (escalationAction) {
        case EscalationAction.AutoApprove:
          // Auto-approve the overdue approval
          await this.sp.web.lists
            .getByTitle(this.APPROVALS_LIST)
            .items.getById(approvalId)
            .update({
              IsOverdue: true,
              EscalationLevel: (approval.EscalationLevel || 0) + 1,
              EscalationDate: new Date().toISOString(),
              Status: ApprovalStatus.Approved,
              Decision: ApprovalStatus.Approved,
              Comments: 'Auto-approved due to escalation policy',
              CompletedDate: new Date().toISOString()
            });
          autoApproved = true;

          // Check level completion after auto-approval
          await this.checkLevelCompletion(approval.ProcessID, approval.ApprovalLevel);
          break;

        case EscalationAction.AssignToManager:
          // Get the approver's manager and reassign
          newApproverId = await this.getApproverManager(approval.ApproverId);
          if (newApproverId) {
            await this.sp.web.lists
              .getByTitle(this.APPROVALS_LIST)
              .items.getById(approvalId)
              .update({
                IsOverdue: true,
                EscalationLevel: (approval.EscalationLevel || 0) + 1,
                EscalationDate: new Date().toISOString(),
                OriginalApproverId: approval.ApproverId,
                ApproverId: newApproverId,
                Status: ApprovalStatus.Pending // Reset to pending for new approver
              });
          } else {
            // No manager found - fall back to Notify behavior
            await this.updateApprovalAsEscalated(approvalId, approval.EscalationLevel);
          }
          break;

        case EscalationAction.AssignToAlternate:
          // Get alternate approver from chain config
          newApproverId = await this.getAlternateApprover(chain, approval.ApprovalLevel);
          if (newApproverId) {
            await this.sp.web.lists
              .getByTitle(this.APPROVALS_LIST)
              .items.getById(approvalId)
              .update({
                IsOverdue: true,
                EscalationLevel: (approval.EscalationLevel || 0) + 1,
                EscalationDate: new Date().toISOString(),
                OriginalApproverId: approval.ApproverId,
                ApproverId: newApproverId,
                Status: ApprovalStatus.Pending // Reset to pending for new approver
              });
          } else {
            // No alternate found - fall back to Notify behavior
            await this.updateApprovalAsEscalated(approvalId, approval.EscalationLevel);
          }
          break;

        case EscalationAction.Notify:
        default:
          // Just mark as escalated and notify
          await this.updateApprovalAsEscalated(approvalId, approval.EscalationLevel);
          break;
      }

      // Log history
      await this.logHistory({
        ApprovalId: approvalId,
        ProcessID: approval.ProcessID,
        Action: autoApproved ? 'AutoApproved' : 'Escalated',
        PerformedById: -1, // System
        ActionDate: new Date().toISOString(),
        PreviousStatus: approval.Status,
        NewStatus: autoApproved ? ApprovalStatus.Approved : ApprovalStatus.Escalated,
        EscalationReason: 'Approval overdue',
        EscalationAction: escalationAction,
        EscalatedToId: newApproverId
      });

      // GAP FIX: Send escalation notifications
      await this.sendEscalationNotifications(
        approval,
        escalationAction,
        newApproverId,
        autoApproved
      );

      logger.info('ApprovalService',
        `Escalated approval ${approvalId} with action: ${escalationAction}` +
        (newApproverId ? `, reassigned to: ${newApproverId}` : '') +
        (autoApproved ? ' (auto-approved)' : '')
      );
    } catch (error) {
      logger.error('ApprovalService', 'Failed to escalate approval:', error);
    }
  }

  /**
   * Update approval as escalated (without reassignment)
   */
  private async updateApprovalAsEscalated(approvalId: number, currentEscalationLevel: number): Promise<void> {
    await this.sp.web.lists
      .getByTitle(this.APPROVALS_LIST)
      .items.getById(approvalId)
      .update({
        IsOverdue: true,
        EscalationLevel: (currentEscalationLevel || 0) + 1,
        EscalationDate: new Date().toISOString(),
        Status: ApprovalStatus.Escalated
      });
  }

  /**
   * Activate the next sequential approver
   * GAP FIX: Changes Queued approval to Pending and notifies the approver
   */
  private async activateSequentialApprover(
    approvalId: number,
    processId: number,
    chainId: number,
    level: number
  ): Promise<void> {
    try {
      // Get the approval details
      const approval = await this.getApprovalById(approvalId);

      // Update status from Queued to Pending
      await this.sp.web.lists
        .getByTitle(this.APPROVALS_LIST)
        .items.getById(approvalId)
        .update({
          Status: ApprovalStatus.Pending,
          RequestedDate: new Date().toISOString() // Update request date to now
        });

      // Log the activation
      await this.logHistory({
        ApprovalId: approvalId,
        ProcessID: processId,
        Action: 'SequentialActivation',
        PerformedById: -1, // System
        ActionDate: new Date().toISOString(),
        PreviousStatus: ApprovalStatus.Queued,
        NewStatus: ApprovalStatus.Pending,
        Comments: `Activated as next in sequential approval (sequence ${approval.ApprovalSequence})`
      });

      // Notify the approver
      await this.notifyNextLevelApprovers(processId, chainId, level, [approval.ApproverId]);

      logger.info('ApprovalService',
        `Activated sequential approver ${approval.ApproverId} (sequence ${approval.ApprovalSequence}) for process ${processId}`
      );
    } catch (error) {
      logger.error('ApprovalService',
        `Failed to activate sequential approver ${approvalId}`, error);
      throw error;
    }
  }

  /**
   * Get the manager of an approver for escalation
   */
  private async getApproverManager(approverId: number): Promise<number | null> {
    try {
      // Try to get manager from user profile or PM_Employees list
      const employees = await this.sp.web.lists.getByTitle('PM_Employees')
        .items
        .filter(`UserId eq ${approverId}`)
        .select('ManagerId')
        .top(1)();

      if (employees.length > 0 && employees[0].ManagerId) {
        return employees[0].ManagerId;
      }

      // Fallback: Try to get manager from SharePoint user profile
      // This would require Graph API access, return null for now
      return null;
    } catch (error) {
      logger.warn('ApprovalService', `Failed to get manager for user ${approverId}`, error);
      return null;
    }
  }

  /**
   * Get alternate approver from chain configuration
   */
  private async getAlternateApprover(chain: IJmlApprovalChain, level: number): Promise<number | null> {
    try {
      const levels = JSON.parse(chain.Levels as any);
      const currentLevel = levels[level - 1];

      // Check for alternate approvers in level config
      if (currentLevel?.AlternateApproverIds && currentLevel.AlternateApproverIds.length > 0) {
        return currentLevel.AlternateApproverIds[0];
      }

      // Note: If chain-level alternate is needed, it would be stored in Levels JSON
      // For now, return null if no level-specific alternate is found
      return null;
    } catch (error) {
      logger.warn('ApprovalService', `Failed to get alternate approver for chain ${chain.Id}`, error);
      return null;
    }
  }

  /**
   * Send escalation notifications to relevant parties
   * GAP FIX: Implements proper notification sending on escalation
   */
  private async sendEscalationNotifications(
    approval: IJmlApproval,
    escalationAction: EscalationAction,
    newApproverId: number | null,
    autoApproved: boolean
  ): Promise<void> {
    try {
      // Get process details for notification context
      const process = await this.sp.web.lists
        .getByTitle('PM_Processes')
        .items
        .getById(approval.ProcessID)
        .select('Id', 'Title', 'ProcessType', 'EmployeeName')() as {
          Id: number;
          Title: string;
          ProcessType: string;
          EmployeeName: string;
        };

      // Notify original approver that their approval was escalated
      await this.sp.web.lists.getByTitle('PM_Notifications').items.add({
        Title: autoApproved ? 'Approval Auto-Approved' : 'Approval Escalated',
        Message: autoApproved
          ? `Your approval for ${process.ProcessType} process (${process.EmployeeName}) was auto-approved due to escalation policy.`
          : `Your approval for ${process.ProcessType} process (${process.EmployeeName}) has been escalated due to overdue status.`,
        RecipientId: approval.ApproverId,
        NotificationType: 'ApprovalEscalated',
        Priority: Priority.High,
        LinkUrl: `/sites/JML/SitePages/ApprovalCenter.aspx?processId=${approval.ProcessID}`,
        ProcessId: approval.ProcessID.toString(),
        IsRead: false,
        SentDate: new Date()
      });

      // If reassigned, notify the new approver
      if (newApproverId && !autoApproved) {
        await this.sp.web.lists.getByTitle('PM_Notifications').items.add({
          Title: 'Escalated Approval Requires Your Attention',
          Message: `An approval for ${process.ProcessType} process (${process.EmployeeName}) has been escalated to you. Please review urgently.`,
          RecipientId: newApproverId,
          NotificationType: 'ApprovalRequired',
          Priority: Priority.Critical,
          LinkUrl: `/sites/JML/SitePages/ApprovalCenter.aspx?processId=${approval.ProcessID}`,
          ProcessId: approval.ProcessID.toString(),
          IsRead: false,
          SentDate: new Date()
        });
      }

      // Send email/Teams notifications via ApprovalNotificationService
      if (this.approvalNotificationService) {
        try {
          // Send escalation notification - use newApproverId if reassigned, otherwise original approver
          const escalatedToUserId = newApproverId || approval.ApproverId;
          await this.approvalNotificationService.sendEscalationNotification(approval, escalatedToUserId);
          logger.info('ApprovalService',
            `Sent escalation email notification for approval ${approval.Id}`);
        } catch (emailError) {
          logger.warn('ApprovalService',
            `Failed to send escalation email for approval ${approval.Id}`, emailError);
        }
      }
    } catch (error) {
      logger.error('ApprovalService',
        `Failed to send escalation notifications for approval ${approval.Id}`, error);
    }
  }

  /**
   * Log approval history
   */
  private async logHistory(history: any): Promise<void> {
    await this.sp.web.lists
      .getByTitle(this.APPROVAL_HISTORY_LIST)
      .items.add(history);
  }

  /**
   * Get approval by ID
   */
  public async getApprovalById(approvalId: number): Promise<IJmlApproval> {
    const item = await this.sp.web.lists
      .getByTitle(this.APPROVALS_LIST)
      .items.getById(approvalId)
      .select(
        '*',
        'Approver/Title',
        'Approver/EMail',
        'OriginalApprover/Title',
        'OriginalApprover/EMail',
        'ModifiedBy/Title',
        'ModifiedBy/EMail'
      )
      .expand('Approver', 'OriginalApprover', 'ModifiedBy')();

    return this.mapToApproval(item);
  }

  /**
   * Get chain by ID
   */
  private async getChainById(chainId: number): Promise<IJmlApprovalChain> {
    const item = await this.sp.web.lists
      .getByTitle(this.APPROVAL_CHAINS_LIST)
      .items.getById(chainId)
      .select('*', 'CreatedBy/Title', 'CreatedBy/EMail', 'ModifiedBy/Title', 'ModifiedBy/EMail')
      .expand('CreatedBy', 'ModifiedBy')();

    return this.mapToChain(item);
  }

  /**
   * Get chain by process ID
   */
  private async getChainByProcessId(processId: number): Promise<IJmlApprovalChain | undefined> {
    const items = await this.sp.web.lists
      .getByTitle(this.APPROVAL_CHAINS_LIST)
      .items.filter(`ProcessID eq ${processId} and IsActive eq 1`)
      .select('*', 'CreatedBy/Title', 'CreatedBy/EMail', 'ModifiedBy/Title', 'ModifiedBy/EMail')
      .expand('CreatedBy', 'ModifiedBy')
      .top(1)();

    return items.length > 0 ? this.mapToChain(items[0]) : undefined;
  }

  /**
   * Get template by ID
   */
  private async getTemplateById(templateId: number): Promise<IJmlApprovalTemplate> {
    const item = await this.sp.web.lists
      .getByTitle(this.TEMPLATES_LIST)
      .items.getById(templateId)
      .select('*', 'CreatedBy/Title', 'CreatedBy/EMail', 'ModifiedBy/Title', 'ModifiedBy/EMail')
      .expand('CreatedBy', 'ModifiedBy')();

    return this.mapToTemplate(item);
  }

  /**
   * Get delegation by ID
   */
  private async getDelegationById(delegationId: number): Promise<IJmlApprovalDelegation> {
    const item = await this.sp.web.lists
      .getByTitle(this.DELEGATIONS_LIST)
      .items.getById(delegationId)
      .select('*', 'DelegatedBy/Title', 'DelegatedBy/EMail', 'DelegatedTo/Title', 'DelegatedTo/EMail')
      .expand('DelegatedBy', 'DelegatedTo')();

    return this.mapToDelegation(item);
  }

  // Mapping functions

  private mapToApproval(item: any): IJmlApproval {
    return {
      Id: item.Id,
      ProcessID: item.ProcessID,
      ProcessTitle: item.ProcessTitle || '',
      ProcessType: item.ProcessType || '',
      ApprovalLevel: item.ApprovalLevel,
      ApprovalSequence: item.ApprovalSequence,
      ApprovalType: item.ApprovalType as ApprovalType,
      Status: item.Status as ApprovalStatus,
      ApproverId: item.ApproverId,
      Approver: {
        Id: item.ApproverId,
        Title: item.Approver?.Title || '',
        EMail: item.Approver?.EMail || ''
      },
      OriginalApproverId: item.OriginalApproverId,
      OriginalApprover: item.OriginalApprover
        ? {
            Id: item.OriginalApproverId,
            Title: item.OriginalApprover.Title,
            EMail: item.OriginalApprover.EMail
          }
        : undefined,
      DelegatedById: item.DelegatedById,
      RequestedDate: new Date(item.RequestedDate),
      DueDate: new Date(item.DueDate),
      CompletedDate: item.ActualCompletionDate ? new Date(item.ActualCompletionDate) : undefined,
      ResponseTime: item.ResponseTime,
      Decision: item.Decision as ApprovalStatus,
      Comments: item.Notes,
      ReasonRequired: item.ReasonRequired || false,
      IsOverdue: item.IsOverdue || false,
      EscalationLevel: item.EscalationLevel || 0,
      EscalationDate: item.EscalationDate ? new Date(item.EscalationDate) : undefined,
      EscalationAction: item.EscalationAction as EscalationAction,
      // Workflow Integration
      WorkflowInstanceId: item.WorkflowInstanceId,
      WorkflowStepId: item.WorkflowStepId,
      ApprovalTemplateId: item.ApprovalTemplateId,
      Created: new Date(item.Created),
      Modified: new Date(item.Modified),
      ModifiedBy: {
        Id: item.EditorId,
        Title: item.ModifiedBy?.Title || '',
        EMail: item.ModifiedBy?.EMail || ''
      }
    };
  }

  private mapToChain(item: any): IJmlApprovalChain {
    return {
      Id: item.Id,
      ProcessID: item.ProcessID,
      ChainName: item.ChainName,
      ApprovalType: item.ApprovalType as ApprovalType,
      IsActive: item.IsActive,
      Levels: JSON.parse(item.Levels || '[]'),
      RequireComments: item.RequireComments || false,
      AllowDelegation: item.AllowDelegation !== false,
      AutoEscalationDays: item.AutoEscalationDays || 3,
      EscalationAction: item.EscalationAction as EscalationAction,
      CurrentLevel: item.CurrentLevel,
      OverallStatus: item.OverallStatus as ApprovalStatus,
      StartDate: item.StartDate ? new Date(item.StartDate) : undefined,
      CompletedDate: item.ActualCompletionDate ? new Date(item.ActualCompletionDate) : undefined,
      Created: new Date(item.Created),
      CreatedBy: {
        Id: item.AuthorId,
        Title: item.CreatedBy?.Title || '',
        EMail: item.CreatedBy?.EMail || ''
      },
      Modified: new Date(item.Modified),
      ModifiedBy: {
        Id: item.EditorId,
        Title: item.ModifiedBy?.Title || '',
        EMail: item.ModifiedBy?.EMail || ''
      }
    };
  }

  private mapToHistory(item: any): IJmlApprovalHistory {
    return {
      Id: item.Id,
      ApprovalId: item.ApprovalId,
      ProcessID: item.ProcessID,
      Action: item.Action,
      PerformedBy: {
        Id: item.PerformedById,
        Title: item.PerformedBy?.Title || 'System',
        EMail: item.PerformedBy?.EMail || ''
      },
      PerformedById: item.PerformedById,
      ActionDate: new Date(item.ActionDate),
      Comments: item.Notes,
      PreviousStatus: item.PreviousStatus as ApprovalStatus,
      NewStatus: item.NewStatus as ApprovalStatus,
      DelegatedTo: item.DelegatedTo
        ? {
            Id: item.DelegatedToId,
            Title: item.DelegatedTo.Title,
            EMail: item.DelegatedTo.EMail
          }
        : undefined,
      DelegatedToId: item.DelegatedToId,
      DelegationReason: item.DelegationReason,
      EscalatedTo: item.EscalatedTo
        ? {
            Id: item.EscalatedToId,
            Title: item.EscalatedTo.Title,
            EMail: item.EscalatedTo.EMail
          }
        : undefined,
      EscalatedToId: item.EscalatedToId,
      EscalationReason: item.EscalationReason,
      Created: new Date(item.Created)
    };
  }

  private mapToDelegation(item: any): IJmlApprovalDelegation {
    return {
      Id: item.Id,
      DelegatedById: item.DelegatedById,
      DelegatedBy: {
        Id: item.DelegatedById,
        Title: item.DelegatedBy?.Title || '',
        EMail: item.DelegatedBy?.EMail || ''
      },
      DelegatedToId: item.DelegatedToId,
      DelegatedTo: {
        Id: item.DelegatedToId,
        Title: item.DelegatedTo?.Title || '',
        EMail: item.DelegatedTo?.EMail || ''
      },
      StartDate: new Date(item.StartDate),
      EndDate: new Date(item.EndDate),
      IsActive: item.IsActive,
      Reason: item.Reason,
      ProcessTypes: item.ProcessTypes ? JSON.parse(item.ProcessTypes) : undefined,
      AutoDelegate: item.AutoDelegate || false,
      Created: new Date(item.Created),
      Modified: new Date(item.Modified)
    };
  }

  private mapToTemplate(item: any): IJmlApprovalTemplate {
    return {
      Id: item.Id,
      Title: item.Title,
      Description: item.Description,
      ProcessTypes: item.ProcessTypes ? JSON.parse(item.ProcessTypes) : [],
      ApprovalType: item.ApprovalType as ApprovalType,
      Levels: JSON.parse(item.Levels || '[]'),
      RequireComments: item.RequireComments || false,
      AllowDelegation: item.AllowDelegation !== false,
      AutoEscalationDays: item.AutoEscalationDays || 3,
      EscalationAction: item.EscalationAction as EscalationAction,
      IsActive: item.IsActive !== false,
      Created: new Date(item.Created),
      CreatedBy: {
        Id: item.AuthorId,
        Title: item.CreatedBy?.Title || '',
        EMail: item.CreatedBy?.EMail || ''
      },
      Modified: new Date(item.Modified),
      ModifiedBy: {
        Id: item.EditorId,
        Title: item.ModifiedBy?.Title || '',
        EMail: item.ModifiedBy?.EMail || ''
      }
    };
  }
}
