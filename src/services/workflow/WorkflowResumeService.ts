// @ts-nocheck
/* eslint-disable */
/**
 * WorkflowResumeService
 * Critical service for detecting completed tasks/approvals and resuming waiting workflows
 * Bridges the gap between external completions (UI) and workflow engine resumption
 *
 * This service solves the CRITICAL gap where workflows get stuck in 'WaitingForTask'
 * or 'WaitingForApproval' states because there's no mechanism to resume them when
 * tasks/approvals complete.
 */

import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import {
  IWorkflowInstance,
  IWorkflowStepStatus,
  WorkflowInstanceStatus,
  StepStatus
} from '../../models/IWorkflow';
import { TaskStatus } from '../../models/ICommon';
import { ApprovalStatus } from '../../models/IJmlApproval';
import { logger } from '../LoggingService';
import { WorkflowEngineService } from './WorkflowEngineService';
import { WorkflowInstanceService } from './WorkflowInstanceService';

/**
 * Result of a workflow resume operation
 */
export interface IWorkflowResumeResult {
  success: boolean;
  instanceId: number;
  resumedFromStatus: WorkflowInstanceStatus;
  newStatus?: WorkflowInstanceStatus;
  resumedStepId?: string;
  message?: string;
  error?: string;
}

/**
 * Result of polling for waiting workflows
 */
export interface IPollingResult {
  polledAt: Date;
  waitingWorkflowsFound: number;
  workflowsResumed: number;
  workflowsFailed: number;
  results: IWorkflowResumeResult[];
}

/**
 * Completed item that can trigger workflow resume
 */
export interface ICompletedItem {
  itemType: 'task' | 'approval';
  itemId: number;
  workflowInstanceId: number;
  workflowStepId: string;
  completedDate: Date;
  completionData?: Record<string, unknown>;
}

/**
 * Configuration for the polling mechanism
 */
export interface IPollingConfig {
  enabled: boolean;
  intervalMs: number;
  maxConcurrentResumes: number;
  retryFailedAfterMs: number;
}

export class WorkflowResumeService {
  private sp: SPFI;
  private context: WebPartContext;
  private workflowEngine: WorkflowEngineService;
  private instanceService: WorkflowInstanceService;
  private pollingConfig: IPollingConfig;
  private pollingTimer: ReturnType<typeof setInterval> | null = null;
  private isPolling: boolean = false;

  constructor(sp: SPFI, context: WebPartContext) {
    this.sp = sp;
    this.context = context;
    this.workflowEngine = new WorkflowEngineService(sp, context);
    this.instanceService = new WorkflowInstanceService(sp);

    // Default polling configuration
    // CRITICAL FIX: Enable polling by default as a safety net for missed events
    this.pollingConfig = {
      enabled: true, // Changed from false - ensures workflows are never permanently stuck
      intervalMs: 30000, // 30 seconds
      maxConcurrentResumes: 5,
      retryFailedAfterMs: 300000 // 5 minutes
    };
  }

  // ============================================================================
  // EVENT-DRIVEN RESUME (PREFERRED METHOD)
  // ============================================================================

  /**
   * Resume workflow when a task is completed
   * Call this from JmlMyTasks component when user completes a task
   */
  public async onTaskCompleted(
    taskId: number,
    completionData?: Record<string, unknown>
  ): Promise<IWorkflowResumeResult | null> {
    try {
      // Get the task with workflow reference
      const task = await this.sp.web.lists.getByTitle('JML_TaskAssignments').items
        .getById(taskId)
        .select(
          'Id', 'Title', 'Status', 'WorkflowInstanceId', 'WorkflowStepId',
          'ActualCompletionDate', 'CompletionNotes'
        )();

      // Check if task has workflow reference
      if (!task.WorkflowInstanceId || !task.WorkflowStepId) {
        logger.info('WorkflowResumeService', `Task ${taskId} has no workflow reference - no resume needed`);
        return null;
      }

      // Verify task is actually completed
      if (task.Status !== TaskStatus.Completed && task.Status !== TaskStatus.Skipped) {
        logger.warn('WorkflowResumeService', `Task ${taskId} is not completed (Status: ${task.Status}) - cannot resume workflow`);
        return null;
      }

      logger.info('WorkflowResumeService', `Task ${taskId} completed - checking workflow ${task.WorkflowInstanceId}`);

      // Check if workflow is waiting for this task
      const shouldResume = await this.shouldResumeForTask(
        task.WorkflowInstanceId,
        task.WorkflowStepId,
        taskId
      );

      if (!shouldResume) {
        logger.info('WorkflowResumeService', `Workflow ${task.WorkflowInstanceId} not ready to resume yet`);
        return null;
      }

      // Resume the workflow
      const result = await this.resumeWorkflowFromTask(
        task.WorkflowInstanceId,
        task.WorkflowStepId,
        {
          completedTaskId: taskId,
          completedTaskTitle: task.Title,
          completedAt: task.ActualCompletionDate || new Date().toISOString(),
          ...completionData
        }
      );

      return result;
    } catch (error) {
      logger.error('WorkflowResumeService', `Error processing task ${taskId} completion`, error);
      return {
        success: false,
        instanceId: 0,
        resumedFromStatus: WorkflowInstanceStatus.WaitingForTask,
        error: error instanceof Error ? error.message : 'Failed to process task completion'
      };
    }
  }

  /**
   * Resume workflow when an approval is completed
   * Call this from JmlApprovalCenter component when approver makes a decision
   */
  public async onApprovalCompleted(
    approvalId: number,
    approved: boolean,
    comments?: string,
    approverUserId?: number
  ): Promise<IWorkflowResumeResult | null> {
    try {
      // Get the approval with workflow reference
      const approval = await this.sp.web.lists.getByTitle('JML_Approvals').items
        .getById(approvalId)
        .select(
          'Id', 'Title', 'Status', 'WorkflowInstanceId', 'WorkflowStepId',
          'ResponseDate', 'ApproverComments', 'ApprovalLevel', 'TotalLevels'
        )();

      // Check if approval has workflow reference
      if (!approval.WorkflowInstanceId || !approval.WorkflowStepId) {
        logger.info('WorkflowResumeService', `Approval ${approvalId} has no workflow reference - no resume needed`);
        return null;
      }

      // Verify approval is actually completed
      const completedStatuses = [ApprovalStatus.Approved, ApprovalStatus.Rejected];
      if (!completedStatuses.includes(approval.Status)) {
        logger.warn('WorkflowResumeService', `Approval ${approvalId} is not completed (Status: ${approval.Status}) - cannot resume workflow`);
        return null;
      }

      logger.info('WorkflowResumeService', `Approval ${approvalId} ${approved ? 'approved' : 'rejected'} - checking workflow ${approval.WorkflowInstanceId}`);

      // For multi-level approvals, check if all levels are complete
      if (approval.TotalLevels && approval.TotalLevels > 1) {
        const allLevelsComplete = await this.checkMultiLevelApprovalComplete(
          approval.WorkflowInstanceId,
          approval.WorkflowStepId
        );

        if (!allLevelsComplete) {
          logger.info('WorkflowResumeService', `Multi-level approval not complete - waiting for more approvals`);
          return null;
        }
      }

      // Check if workflow is waiting for this approval
      const shouldResume = await this.shouldResumeForApproval(
        approval.WorkflowInstanceId,
        approval.WorkflowStepId
      );

      if (!shouldResume) {
        logger.info('WorkflowResumeService', `Workflow ${approval.WorkflowInstanceId} not ready to resume yet`);
        return null;
      }

      // Resume the workflow
      const result = await this.resumeWorkflowFromApproval(
        approval.WorkflowInstanceId,
        approval.WorkflowStepId,
        {
          approvalId,
          approved,
          approverComments: comments,
          approverUserId,
          respondedAt: approval.ResponseDate || new Date().toISOString(),
          // Include whether it was rejected to allow workflow to branch
          isRejected: !approved
        }
      );

      return result;
    } catch (error) {
      logger.error('WorkflowResumeService', `Error processing approval ${approvalId} completion`, error);
      return {
        success: false,
        instanceId: 0,
        resumedFromStatus: WorkflowInstanceStatus.WaitingForApproval,
        error: error instanceof Error ? error.message : 'Failed to process approval completion'
      };
    }
  }

  // ============================================================================
  // RESUME CONDITION CHECKS
  // ============================================================================

  /**
   * Check if workflow should be resumed after task completion
   * Handles 'wait for all' vs 'wait for any' conditions
   */
  private async shouldResumeForTask(
    workflowInstanceId: number,
    workflowStepId: string,
    completedTaskId: number
  ): Promise<boolean> {
    try {
      // Get workflow instance
      const instance = await this.instanceService.getById(workflowInstanceId);

      // Only resume if workflow is waiting for tasks
      if (instance.Status !== WorkflowInstanceStatus.WaitingForTask) {
        return false;
      }

      // Get the step status to check wait configuration
      const stepStatus = await this.instanceService.getStepStatus(workflowInstanceId, workflowStepId);

      if (!stepStatus) {
        // No step status means we should resume
        return true;
      }

      // Parse step result to get wait configuration
      const stepResult = stepStatus.Result ? JSON.parse(stepStatus.Result) : {};
      const waitCondition = stepResult.waitCondition || 'all';

      // Get all tasks for this workflow step
      const tasks = await this.sp.web.lists.getByTitle('JML_TaskAssignments').items
        .filter(`WorkflowInstanceId eq ${workflowInstanceId} and WorkflowStepId eq '${workflowStepId}'`)
        .select('Id', 'Status')();

      const completedTasks = tasks.filter(t =>
        t.Status === TaskStatus.Completed || t.Status === TaskStatus.Skipped
      );

      // Check based on wait condition
      if (waitCondition === 'any') {
        // Any single task completion triggers resume
        return completedTasks.length > 0;
      } else {
        // All tasks must be complete
        return completedTasks.length === tasks.length;
      }
    } catch (error) {
      logger.error('WorkflowResumeService', `Error checking task resume condition`, error);
      return false;
    }
  }

  /**
   * Check if workflow should be resumed after approval completion
   */
  private async shouldResumeForApproval(
    workflowInstanceId: number,
    workflowStepId: string
  ): Promise<boolean> {
    try {
      // Get workflow instance
      const instance = await this.instanceService.getById(workflowInstanceId);

      // Only resume if workflow is waiting for approval
      if (instance.Status !== WorkflowInstanceStatus.WaitingForApproval) {
        return false;
      }

      // Verify this is the current step being waited on
      if (instance.CurrentStepId !== workflowStepId) {
        logger.info('WorkflowResumeService', `Approval step ${workflowStepId} is not current step ${instance.CurrentStepId}`);
        return false;
      }

      return true;
    } catch (error) {
      logger.error('WorkflowResumeService', `Error checking approval resume condition`, error);
      return false;
    }
  }

  /**
   * Check if all levels of a multi-level approval are complete
   */
  private async checkMultiLevelApprovalComplete(
    workflowInstanceId: number,
    workflowStepId: string
  ): Promise<boolean> {
    try {
      // Get all approvals for this workflow step
      const approvals = await this.sp.web.lists.getByTitle('JML_Approvals').items
        .filter(`WorkflowInstanceId eq ${workflowInstanceId} and WorkflowStepId eq '${workflowStepId}'`)
        .select('Id', 'Status', 'ApprovalLevel', 'TotalLevels')
        .orderBy('ApprovalLevel', true)();

      if (approvals.length === 0) {
        return false;
      }

      const totalLevels = approvals[0].TotalLevels || 1;

      // Check each level is approved
      const completedLevels = approvals.filter(a =>
        a.Status === ApprovalStatus.Approved
      );

      // For rejection, we resume immediately (workflow handles branching)
      const hasRejection = approvals.some(a => a.Status === ApprovalStatus.Rejected);
      if (hasRejection) {
        return true;
      }

      return completedLevels.length >= totalLevels;
    } catch (error) {
      logger.error('WorkflowResumeService', `Error checking multi-level approval`, error);
      return false;
    }
  }

  // ============================================================================
  // WORKFLOW RESUME EXECUTION
  // ============================================================================

  /**
   * Resume workflow from task completion
   */
  private async resumeWorkflowFromTask(
    instanceId: number,
    stepId: string,
    completionData: Record<string, unknown>
  ): Promise<IWorkflowResumeResult> {
    try {
      const instance = await this.instanceService.getById(instanceId);
      const previousStatus = instance.Status;

      logger.info('WorkflowResumeService', `Resuming workflow ${instanceId} from step ${stepId} after task completion`);

      // Complete the waiting step and continue workflow
      const result = await this.workflowEngine.completeWaitingStep(instanceId, stepId, {
        taskCompleted: true,
        ...completionData
      });

      return {
        success: result.success,
        instanceId,
        resumedFromStatus: previousStatus,
        newStatus: result.status,
        resumedStepId: stepId,
        message: result.message || `Workflow resumed after task completion`
      };
    } catch (error) {
      logger.error('WorkflowResumeService', `Error resuming workflow ${instanceId}`, error);
      return {
        success: false,
        instanceId,
        resumedFromStatus: WorkflowInstanceStatus.WaitingForTask,
        error: error instanceof Error ? error.message : 'Failed to resume workflow'
      };
    }
  }

  /**
   * Resume workflow from approval completion
   */
  private async resumeWorkflowFromApproval(
    instanceId: number,
    stepId: string,
    completionData: Record<string, unknown>
  ): Promise<IWorkflowResumeResult> {
    try {
      const instance = await this.instanceService.getById(instanceId);
      const previousStatus = instance.Status;

      logger.info('WorkflowResumeService', `Resuming workflow ${instanceId} from step ${stepId} after approval`);

      // Complete the waiting step and continue workflow
      const result = await this.workflowEngine.completeWaitingStep(instanceId, stepId, {
        approvalCompleted: true,
        ...completionData
      });

      return {
        success: result.success,
        instanceId,
        resumedFromStatus: previousStatus,
        newStatus: result.status,
        resumedStepId: stepId,
        message: result.message || `Workflow resumed after approval`
      };
    } catch (error) {
      logger.error('WorkflowResumeService', `Error resuming workflow ${instanceId}`, error);
      return {
        success: false,
        instanceId,
        resumedFromStatus: WorkflowInstanceStatus.WaitingForApproval,
        error: error instanceof Error ? error.message : 'Failed to resume workflow'
      };
    }
  }

  // ============================================================================
  // POLLING MECHANISM (FALLBACK/BACKUP)
  // ============================================================================

  /**
   * Start polling for waiting workflows
   * Use as a backup if event-driven approach misses completions
   */
  public startPolling(config?: Partial<IPollingConfig>): void {
    if (config) {
      this.pollingConfig = { ...this.pollingConfig, ...config };
    }

    if (this.pollingTimer) {
      this.stopPolling();
    }

    this.pollingConfig.enabled = true;
    this.pollingTimer = setInterval(
      () => this.pollAndResumeWorkflows(),
      this.pollingConfig.intervalMs
    );

    logger.info('WorkflowResumeService', `Polling started with interval ${this.pollingConfig.intervalMs}ms`);
  }

  /**
   * Stop polling
   */
  public stopPolling(): void {
    if (this.pollingTimer) {
      clearInterval(this.pollingTimer);
      this.pollingTimer = null;
    }
    this.pollingConfig.enabled = false;
    logger.info('WorkflowResumeService', 'Polling stopped');
  }

  /**
   * Poll for waiting workflows and attempt to resume them
   */
  public async pollAndResumeWorkflows(): Promise<IPollingResult> {
    if (this.isPolling) {
      logger.warn('WorkflowResumeService', 'Polling already in progress - skipping');
      return {
        polledAt: new Date(),
        waitingWorkflowsFound: 0,
        workflowsResumed: 0,
        workflowsFailed: 0,
        results: []
      };
    }

    this.isPolling = true;
    const results: IWorkflowResumeResult[] = [];

    try {
      // Find workflows waiting for tasks
      const waitingForTasks = await this.findWorkflowsWaitingForTasks();

      // Find workflows waiting for approvals
      const waitingForApprovals = await this.findWorkflowsWaitingForApprovals();

      const totalWaiting = waitingForTasks.length + waitingForApprovals.length;

      logger.info('WorkflowResumeService', `Polling found ${totalWaiting} waiting workflows (${waitingForTasks.length} tasks, ${waitingForApprovals.length} approvals)`);

      // Process task completions
      for (const item of waitingForTasks.slice(0, this.pollingConfig.maxConcurrentResumes)) {
        const result = await this.resumeWorkflowFromTask(
          item.workflowInstanceId,
          item.workflowStepId,
          { polledCompletion: true, ...item.completionData }
        );
        results.push(result);
      }

      // Process approval completions
      for (const item of waitingForApprovals.slice(0, this.pollingConfig.maxConcurrentResumes)) {
        const result = await this.resumeWorkflowFromApproval(
          item.workflowInstanceId,
          item.workflowStepId,
          { polledCompletion: true, ...item.completionData }
        );
        results.push(result);
      }

      const succeeded = results.filter(r => r.success).length;
      const failed = results.filter(r => !r.success).length;

      return {
        polledAt: new Date(),
        waitingWorkflowsFound: totalWaiting,
        workflowsResumed: succeeded,
        workflowsFailed: failed,
        results
      };
    } catch (error) {
      logger.error('WorkflowResumeService', 'Error during polling', error);
      return {
        polledAt: new Date(),
        waitingWorkflowsFound: 0,
        workflowsResumed: 0,
        workflowsFailed: 0,
        results: []
      };
    } finally {
      this.isPolling = false;
    }
  }

  /**
   * Find workflows waiting for tasks that have been completed
   */
  private async findWorkflowsWaitingForTasks(): Promise<ICompletedItem[]> {
    const completedItems: ICompletedItem[] = [];

    try {
      // Get workflows in WaitingForTask status
      const waitingInstances = await this.sp.web.lists.getByTitle('JML_WorkflowInstances').items
        .filter(`Status eq '${WorkflowInstanceStatus.WaitingForTask}'`)
        .select('Id', 'CurrentStepId', 'ProcessId')
        .top(50)();

      for (const instance of waitingInstances) {
        // Get completed tasks for this workflow
        const completedTasks = await this.sp.web.lists.getByTitle('JML_TaskAssignments').items
          .filter(`WorkflowInstanceId eq ${instance.Id} and (Status eq '${TaskStatus.Completed}' or Status eq '${TaskStatus.Skipped}')`)
          .select('Id', 'Title', 'Status', 'WorkflowStepId', 'ActualCompletionDate')
          .top(10)();

        if (completedTasks.length > 0) {
          // Check if we should resume based on the step's wait condition
          const shouldResume = await this.shouldResumeForTask(
            instance.Id,
            instance.CurrentStepId,
            completedTasks[0].Id
          );

          if (shouldResume) {
            completedItems.push({
              itemType: 'task',
              itemId: completedTasks[0].Id,
              workflowInstanceId: instance.Id,
              workflowStepId: instance.CurrentStepId,
              completedDate: completedTasks[0].ActualCompletionDate || new Date(),
              completionData: {
                completedTaskIds: completedTasks.map(t => t.Id),
                taskCount: completedTasks.length
              }
            });
          }
        }
      }
    } catch (error) {
      logger.error('WorkflowResumeService', 'Error finding workflows waiting for tasks', error);
    }

    return completedItems;
  }

  /**
   * Find workflows waiting for approvals that have been completed
   */
  private async findWorkflowsWaitingForApprovals(): Promise<ICompletedItem[]> {
    const completedItems: ICompletedItem[] = [];

    try {
      // Get workflows in WaitingForApproval status
      const waitingInstances = await this.sp.web.lists.getByTitle('JML_WorkflowInstances').items
        .filter(`Status eq '${WorkflowInstanceStatus.WaitingForApproval}'`)
        .select('Id', 'CurrentStepId', 'ProcessId')
        .top(50)();

      for (const instance of waitingInstances) {
        // Get completed approvals for this workflow
        const completedApprovals = await this.sp.web.lists.getByTitle('JML_Approvals').items
          .filter(`WorkflowInstanceId eq ${instance.Id} and WorkflowStepId eq '${instance.CurrentStepId}' and (Status eq '${ApprovalStatus.Approved}' or Status eq '${ApprovalStatus.Rejected}')`)
          .select('Id', 'Title', 'Status', 'WorkflowStepId', 'ResponseDate', 'ApproverComments')
          .top(10)();

        if (completedApprovals.length > 0) {
          // Check if all levels are complete for multi-level approvals
          const allComplete = await this.checkMultiLevelApprovalComplete(
            instance.Id,
            instance.CurrentStepId
          );

          if (allComplete) {
            const approval = completedApprovals[0];
            completedItems.push({
              itemType: 'approval',
              itemId: approval.Id,
              workflowInstanceId: instance.Id,
              workflowStepId: instance.CurrentStepId,
              completedDate: approval.ResponseDate || new Date(),
              completionData: {
                approved: approval.Status === ApprovalStatus.Approved,
                approverComments: approval.ApproverComments,
                approvalCount: completedApprovals.length
              }
            });
          }
        }
      }
    } catch (error) {
      logger.error('WorkflowResumeService', 'Error finding workflows waiting for approvals', error);
    }

    return completedItems;
  }

  // ============================================================================
  // STATUS & DIAGNOSTICS
  // ============================================================================

  /**
   * Get count of workflows waiting for resume
   */
  public async getWaitingWorkflowCounts(): Promise<{
    waitingForTasks: number;
    waitingForApprovals: number;
    waitingForInput: number;
    total: number;
  }> {
    try {
      const [taskCount] = await this.sp.web.lists.getByTitle('JML_WorkflowInstances').items
        .filter(`Status eq '${WorkflowInstanceStatus.WaitingForTask}'`)
        .select('Id')();

      const [approvalCount] = await this.sp.web.lists.getByTitle('JML_WorkflowInstances').items
        .filter(`Status eq '${WorkflowInstanceStatus.WaitingForApproval}'`)
        .select('Id')();

      const [inputCount] = await this.sp.web.lists.getByTitle('JML_WorkflowInstances').items
        .filter(`Status eq '${WorkflowInstanceStatus.WaitingForInput}'`)
        .select('Id')();

      const waitingForTasks = Array.isArray(taskCount) ? taskCount.length : (taskCount ? 1 : 0);
      const waitingForApprovals = Array.isArray(approvalCount) ? approvalCount.length : (approvalCount ? 1 : 0);
      const waitingForInput = Array.isArray(inputCount) ? inputCount.length : (inputCount ? 1 : 0);

      return {
        waitingForTasks,
        waitingForApprovals,
        waitingForInput,
        total: waitingForTasks + waitingForApprovals + waitingForInput
      };
    } catch (error) {
      logger.error('WorkflowResumeService', 'Error getting waiting workflow counts', error);
      return { waitingForTasks: 0, waitingForApprovals: 0, waitingForInput: 0, total: 0 };
    }
  }

  /**
   * Get polling configuration status
   */
  public getPollingStatus(): {
    enabled: boolean;
    isActive: boolean;
    config: IPollingConfig;
  } {
    return {
      enabled: this.pollingConfig.enabled,
      isActive: this.pollingTimer !== null,
      config: { ...this.pollingConfig }
    };
  }

  /**
   * Force check and resume all stuck workflows
   * Useful for admin recovery operations
   */
  public async forceResumeAllStuckWorkflows(): Promise<IPollingResult> {
    logger.info('WorkflowResumeService', 'Force resuming all stuck workflows...');
    return await this.pollAndResumeWorkflows();
  }

  /**
   * Get detailed status of a specific waiting workflow
   */
  public async getWorkflowWaitStatus(instanceId: number): Promise<{
    instance: IWorkflowInstance | null;
    stepStatus: IWorkflowStepStatus | null;
    waitingFor: 'tasks' | 'approvals' | 'input' | 'none';
    pendingItems: Array<{ type: string; id: number; status: string }>;
    canResume: boolean;
    blockedReason?: string;
  }> {
    try {
      const instance = await this.instanceService.getById(instanceId);
      const stepStatus = instance.CurrentStepId
        ? await this.instanceService.getStepStatus(instanceId, instance.CurrentStepId)
        : null;

      let waitingFor: 'tasks' | 'approvals' | 'input' | 'none' = 'none';
      const pendingItems: Array<{ type: string; id: number; status: string }> = [];
      let canResume = false;
      let blockedReason: string | undefined;

      switch (instance.Status) {
        case WorkflowInstanceStatus.WaitingForTask:
          waitingFor = 'tasks';
          // Get pending tasks
          const tasks = await this.sp.web.lists.getByTitle('JML_TaskAssignments').items
            .filter(`WorkflowInstanceId eq ${instanceId}`)
            .select('Id', 'Title', 'Status')();

          for (const task of tasks) {
            pendingItems.push({ type: 'task', id: task.Id, status: task.Status });
          }

          const completedTasks = tasks.filter(t =>
            t.Status === TaskStatus.Completed || t.Status === TaskStatus.Skipped
          );
          canResume = completedTasks.length === tasks.length || completedTasks.length > 0;
          if (!canResume) {
            blockedReason = `Waiting for ${tasks.length - completedTasks.length} task(s) to complete`;
          }
          break;

        case WorkflowInstanceStatus.WaitingForApproval:
          waitingFor = 'approvals';
          // Get pending approvals
          const approvals = await this.sp.web.lists.getByTitle('JML_Approvals').items
            .filter(`WorkflowInstanceId eq ${instanceId}`)
            .select('Id', 'Title', 'Status')();

          for (const approval of approvals) {
            pendingItems.push({ type: 'approval', id: approval.Id, status: approval.Status });
          }

          const completedApprovals = approvals.filter(a =>
            a.Status === ApprovalStatus.Approved || a.Status === ApprovalStatus.Rejected
          );
          canResume = completedApprovals.length > 0;
          if (!canResume) {
            blockedReason = `Waiting for approval decision`;
          }
          break;

        case WorkflowInstanceStatus.WaitingForInput:
          waitingFor = 'input';
          blockedReason = 'Waiting for user input';
          break;

        default:
          waitingFor = 'none';
          canResume = instance.Status === WorkflowInstanceStatus.Running;
      }

      return {
        instance,
        stepStatus,
        waitingFor,
        pendingItems,
        canResume,
        blockedReason
      };
    } catch (error) {
      logger.error('WorkflowResumeService', `Error getting workflow ${instanceId} wait status`, error);
      return {
        instance: null,
        stepStatus: null,
        waitingFor: 'none',
        pendingItems: [],
        canResume: false,
        blockedReason: 'Error retrieving workflow status'
      };
    }
  }
}
