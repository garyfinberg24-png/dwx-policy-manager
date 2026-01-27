// @ts-nocheck
/* eslint-disable */
/**
 * WorkflowAdvancedService
 * Phase 7: Advanced Workflow Features
 *
 * Integrates:
 * 1. Task Dependency Management - Dependency validation and cascade operations
 * 2. Multi-Level Approval Chains - Enhanced approval handling with routing
 * 3. Scheduled Notification Processing - Deferred and scheduled notifications
 * 4. Parallel Step Synchronization - Parallel branch execution and join handling
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import {
  IWorkflowInstance,
  IWorkflowStep,
  IActionContext,
  IActionResult,
  WorkflowInstanceStatus,
  StepStatus,
  StepType
} from '../../models/IWorkflow';
import { TaskStatus, ProcessStatus, Priority, NotificationType } from '../../models/ICommon';
import { ApprovalStatus } from '../../models/IJmlApproval';
import { logger } from '../LoggingService';
import { retryWithDLQ, workflowSyncDLQ, PROCESS_SYNC_RETRY_OPTIONS } from '../../utils/retryUtils';

// ============================================================================
// INTERFACES
// ============================================================================

/**
 * Task dependency definition
 */
export interface ITaskDependency {
  taskId: number;
  dependsOnTaskIds: number[];
  dependencyType: 'all' | 'any';  // All deps must complete, or any one
}

/**
 * Task dependency validation result
 */
export interface IDependencyValidationResult {
  taskId: number;
  canStart: boolean;
  blockedBy: number[];
  blockedByTitles: string[];
  missingDependencies: number[];
  reason?: string;
}

/**
 * Cascade unblock result
 */
export interface ICascadeUnblockResult {
  completedTaskId: number;
  unblockedTaskIds: number[];
  unblockedCount: number;
  failedToUnblock: number[];
  notificationsSent: number;
}

/**
 * Approval chain level
 */
export interface IApprovalLevel {
  level: number;
  approverIds: number[];
  approverEmails: string[];
  approvalType: 'any' | 'all';  // Any one approver, or all must approve
  escalationDays?: number;
  escalationApproverIds?: number[];
}

/**
 * Multi-level approval status
 */
export interface IMultiLevelApprovalStatus {
  chainId: number;
  processId: number;
  currentLevel: number;
  totalLevels: number;
  levelStatuses: IApprovalLevelStatus[];
  overallStatus: ApprovalStatus;
  canProgress: boolean;
  nextApprovers?: number[];
}

/**
 * Individual level status
 */
export interface IApprovalLevelStatus {
  level: number;
  status: ApprovalStatus;
  approvedBy?: number[];
  rejectedBy?: number[];
  pendingApprovers: number[];
  completedDate?: Date;
}

/**
 * Scheduled notification
 */
export interface IScheduledNotification {
  id?: number;
  processId: number;
  workflowInstanceId: number;
  recipientIds: number[];
  recipientEmails: string[];
  subject: string;
  body: string;
  notificationType: NotificationType;
  scheduledDate: Date;
  status: 'Pending' | 'Sent' | 'Failed' | 'Cancelled';
  retryCount: number;
  maxRetries: number;
  priority: Priority;
  createdDate: Date;
  sentDate?: Date;
  errorMessage?: string;
}

/**
 * Scheduled notification result
 */
export interface IScheduledNotificationResult {
  notificationId: number;
  success: boolean;
  status: 'Sent' | 'Failed' | 'Rescheduled';
  error?: string;
  recipientCount: number;
}

/**
 * Parallel branch status
 */
export interface IParallelBranchStatus {
  branchId: string;
  stepId: string;
  stepName: string;
  status: StepStatus;
  startedDate?: Date;
  completedDate?: Date;
  outputVariables?: Record<string, unknown>;
  error?: string;
}

/**
 * Parallel execution context
 */
export interface IParallelExecutionContext {
  workflowInstanceId: number;
  parallelStepId: string;
  branches: IParallelBranchStatus[];
  joinType: 'all' | 'any' | 'first';  // Wait for all, any, or first success
  joinStepId: string;
  startedDate: Date;
  status: 'Running' | 'Completed' | 'Failed' | 'Partial';
}

/**
 * Parallel sync result
 */
export interface IParallelSyncResult {
  parallelStepId: string;
  allBranchesComplete: boolean;
  completedBranches: string[];
  pendingBranches: string[];
  failedBranches: string[];
  canProceedToJoin: boolean;
  mergedVariables: Record<string, unknown>;
}

/**
 * Approval escalation result
 */
export interface IApprovalEscalationResult {
  chainId: number;
  processId: number;
  level: number;
  daysOverdue: number;
  escalatedToIds: number[];
  originalApproverIds: number[];
  escalatedAt: Date;
}

/**
 * User pending approval info
 */
export interface IUserPendingApproval {
  approvalId: number;
  chainId: number;
  chainName: string;
  processId: number;
  level: number;
  totalLevels: number;
  isEscalation: boolean;
  escalationDays?: number;
  createdDate: Date;
  daysWaiting: number;
}

// ============================================================================
// TASK DEPENDENCY SERVICE
// ============================================================================

export class TaskDependencyService {
  private sp: SPFI;
  private readonly TASK_ASSIGNMENTS_LIST = 'JML_TaskAssignments';
  private readonly TASK_DEPENDENCIES_LIST = 'JML_TaskDependencies';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Validate if a task can start based on its dependencies
   */
  public async validateTaskDependencies(taskId: number): Promise<IDependencyValidationResult> {
    try {
      // Get task's dependencies
      const dependencies = await this.getTaskDependencies(taskId);

      if (dependencies.length === 0) {
        return {
          taskId,
          canStart: true,
          blockedBy: [],
          blockedByTitles: [],
          missingDependencies: []
        };
      }

      // Check status of all dependent tasks
      const dependencyIds = dependencies.map(d => d.dependsOnTaskId);
      const dependentTasks = await this.getTasksByIds(dependencyIds);

      const blockedBy: number[] = [];
      const blockedByTitles: string[] = [];
      const missingDependencies: number[] = [];

      // Check each dependency
      for (const depId of dependencyIds) {
        const depTask = dependentTasks.find(t => t.Id === depId);

        if (!depTask) {
          missingDependencies.push(depId);
          continue;
        }

        const isComplete = depTask.Status === TaskStatus.Completed ||
                          depTask.Status === TaskStatus.Skipped;

        if (!isComplete) {
          blockedBy.push(depId);
          blockedByTitles.push(depTask.Title);
        }
      }

      // Determine if task can start based on dependency type
      const dependencyDef = dependencies[0]; // Assume same type for all deps
      const dependencyType = dependencyDef.dependencyType || 'all';

      let canStart = false;
      let reason: string | undefined;

      if (dependencyType === 'all') {
        canStart = blockedBy.length === 0 && missingDependencies.length === 0;
        if (!canStart) {
          reason = `Waiting for ${blockedBy.length} task(s) to complete: ${blockedByTitles.join(', ')}`;
        }
      } else if (dependencyType === 'any') {
        const completedCount = dependencyIds.length - blockedBy.length - missingDependencies.length;
        canStart = completedCount > 0;
        if (!canStart) {
          reason = `Waiting for at least one dependency to complete`;
        }
      }

      return {
        taskId,
        canStart,
        blockedBy,
        blockedByTitles,
        missingDependencies,
        reason
      };
    } catch (error) {
      logger.error('TaskDependencyService', `Error validating dependencies for task ${taskId}`, error);
      return {
        taskId,
        canStart: false,
        blockedBy: [],
        blockedByTitles: [],
        missingDependencies: [],
        reason: `Error validating dependencies: ${error instanceof Error ? error.message : 'Unknown error'}`
      };
    }
  }

  /**
   * Handle task completion - cascade unblock dependent tasks
   */
  public async onTaskCompleted(taskId: number): Promise<ICascadeUnblockResult> {
    const result: ICascadeUnblockResult = {
      completedTaskId: taskId,
      unblockedTaskIds: [],
      unblockedCount: 0,
      failedToUnblock: [],
      notificationsSent: 0
    };

    try {
      // Find tasks that depend on this completed task
      const dependentTasks = await this.findDependentTasks(taskId);

      for (const depTask of dependentTasks) {
        try {
          // Re-validate the dependent task's dependencies
          const validation = await this.validateTaskDependencies(depTask.Id);

          if (validation.canStart && depTask.Status === TaskStatus.Blocked) {
            // Unblock the task
            await this.sp.web.lists.getByTitle(this.TASK_ASSIGNMENTS_LIST)
              .items.getById(depTask.Id)
              .update({
                Status: TaskStatus.NotStarted,
                BlockedReason: null,
                UnblockedDate: new Date().toISOString()
              });

            result.unblockedTaskIds.push(depTask.Id);
            result.unblockedCount++;

            // Send notification to assignee
            await this.notifyTaskUnblocked(depTask);
            result.notificationsSent++;

            logger.info('TaskDependencyService',
              `Task ${depTask.Id} unblocked after completion of task ${taskId}`);
          }
        } catch (error) {
          result.failedToUnblock.push(depTask.Id);
          logger.error('TaskDependencyService',
            `Failed to unblock task ${depTask.Id}`, error);
        }
      }

      return result;
    } catch (error) {
      logger.error('TaskDependencyService',
        `Error processing task completion cascade for ${taskId}`, error);
      return result;
    }
  }

  /**
   * Create task dependency
   */
  public async createDependency(
    taskId: number,
    dependsOnTaskId: number,
    dependencyType: 'all' | 'any' = 'all'
  ): Promise<number> {
    try {
      const result = await this.sp.web.lists.getByTitle(this.TASK_DEPENDENCIES_LIST)
        .items.add({
          TaskId: taskId,
          DependsOnTaskId: dependsOnTaskId,
          DependencyType: dependencyType,
          CreatedDate: new Date().toISOString()
        });

      logger.info('TaskDependencyService',
        `Created dependency: Task ${taskId} depends on ${dependsOnTaskId}`);

      return result.data.Id;
    } catch (error) {
      logger.error('TaskDependencyService', 'Error creating dependency', error);
      throw error;
    }
  }

  /**
   * Remove task dependency
   */
  public async removeDependency(taskId: number, dependsOnTaskId: number): Promise<void> {
    try {
      const deps = await this.sp.web.lists.getByTitle(this.TASK_DEPENDENCIES_LIST)
        .items
        .filter(`TaskId eq ${taskId} and DependsOnTaskId eq ${dependsOnTaskId}`)
        .select('Id')();

      for (const dep of deps) {
        await this.sp.web.lists.getByTitle(this.TASK_DEPENDENCIES_LIST)
          .items.getById(dep.Id).delete();
      }

      logger.info('TaskDependencyService',
        `Removed dependency: Task ${taskId} no longer depends on ${dependsOnTaskId}`);
    } catch (error) {
      logger.error('TaskDependencyService', 'Error removing dependency', error);
      throw error;
    }
  }

  /**
   * Get all dependencies for a task
   */
  private async getTaskDependencies(taskId: number): Promise<Array<{ dependsOnTaskId: number; dependencyType: string }>> {
    try {
      const deps = await this.sp.web.lists.getByTitle(this.TASK_DEPENDENCIES_LIST)
        .items
        .filter(`TaskId eq ${taskId}`)
        .select('DependsOnTaskId', 'DependencyType')();

      return deps.map(d => ({
        dependsOnTaskId: d.DependsOnTaskId,
        dependencyType: d.DependencyType || 'all'
      }));
    } catch (error) {
      // List might not exist yet, return empty
      return [];
    }
  }

  /**
   * Find tasks that depend on the given task
   */
  private async findDependentTasks(taskId: number): Promise<Array<{ Id: number; Title: string; Status: string; AssignedToId: number }>> {
    try {
      const deps = await this.sp.web.lists.getByTitle(this.TASK_DEPENDENCIES_LIST)
        .items
        .filter(`DependsOnTaskId eq ${taskId}`)
        .select('TaskId')();

      if (deps.length === 0) return [];

      const taskIds = deps.map(d => d.TaskId);
      return this.getTasksByIds(taskIds);
    } catch (error) {
      return [];
    }
  }

  /**
   * Get tasks by their IDs
   */
  private async getTasksByIds(taskIds: number[]): Promise<Array<{ Id: number; Title: string; Status: string; AssignedToId: number }>> {
    if (taskIds.length === 0) return [];

    try {
      const filter = taskIds.map(id => `Id eq ${id}`).join(' or ');
      return await this.sp.web.lists.getByTitle(this.TASK_ASSIGNMENTS_LIST)
        .items
        .filter(filter)
        .select('Id', 'Title', 'Status', 'AssignedToId')();
    } catch (error) {
      return [];
    }
  }

  /**
   * Notify task assignee that their task is unblocked
   */
  private async notifyTaskUnblocked(task: { Id: number; Title: string; AssignedToId: number }): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('JML_Notifications').items.add({
        Title: 'Task Ready to Start',
        RecipientId: task.AssignedToId,
        Message: `Your task "${task.Title}" is now unblocked and ready to start.`,
        NotificationType: 'InApp',
        Status: 'Pending',
        RelatedItemId: task.Id,
        RelatedItemType: 'Task',
        CreatedDate: new Date().toISOString()
      });
    } catch (error) {
      logger.error('TaskDependencyService', 'Error sending unblock notification', error);
    }
  }
}

// ============================================================================
// MULTI-LEVEL APPROVAL SERVICE
// ============================================================================

export class MultiLevelApprovalService {
  private sp: SPFI;
  private readonly APPROVAL_CHAINS_LIST = 'JML_ApprovalChains';
  private readonly APPROVALS_LIST = 'JML_Approvals';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Get detailed multi-level approval status
   */
  public async getApprovalStatus(chainId: number): Promise<IMultiLevelApprovalStatus | null> {
    try {
      const chain = await this.sp.web.lists.getByTitle(this.APPROVAL_CHAINS_LIST)
        .items.getById(chainId)
        .select('*')();

      if (!chain) return null;

      const levels: IApprovalLevel[] = JSON.parse(chain.Levels || '[]');
      const levelStatuses: IApprovalLevelStatus[] = [];

      // Get all approvals for this chain
      const approvals = await this.sp.web.lists.getByTitle(this.APPROVALS_LIST)
        .items
        .filter(`ApprovalChainId eq ${chainId}`)
        .select('Id', 'Level', 'ApproverId', 'Status', 'ApprovedDate', 'Comments')();

      // Build status for each level
      for (let i = 0; i < levels.length; i++) {
        const level = levels[i];
        const levelApprovals = approvals.filter(a => a.Level === level.level);

        const approvedBy = levelApprovals
          .filter(a => a.Status === ApprovalStatus.Approved)
          .map(a => a.ApproverId);

        const rejectedBy = levelApprovals
          .filter(a => a.Status === ApprovalStatus.Rejected)
          .map(a => a.ApproverId);

        const pendingApprovers = level.approverIds.filter(id =>
          !approvedBy.includes(id) && !rejectedBy.includes(id)
        );

        let status = ApprovalStatus.Pending;
        let completedDate: Date | undefined;

        // Determine level status
        if (rejectedBy.length > 0) {
          status = ApprovalStatus.Rejected;
          const rejectedApproval = levelApprovals.find(a => a.Status === ApprovalStatus.Rejected);
          if (rejectedApproval?.ApprovedDate) {
            completedDate = new Date(rejectedApproval.ApprovedDate);
          }
        } else if (level.approvalType === 'all' && approvedBy.length === level.approverIds.length) {
          status = ApprovalStatus.Approved;
          const lastApproval = levelApprovals
            .filter(a => a.Status === ApprovalStatus.Approved)
            .sort((a, b) => new Date(b.ApprovedDate).getTime() - new Date(a.ApprovedDate).getTime())[0];
          if (lastApproval?.ApprovedDate) {
            completedDate = new Date(lastApproval.ApprovedDate);
          }
        } else if (level.approvalType === 'any' && approvedBy.length > 0) {
          status = ApprovalStatus.Approved;
          const firstApproval = levelApprovals
            .filter(a => a.Status === ApprovalStatus.Approved)
            .sort((a, b) => new Date(a.ApprovedDate).getTime() - new Date(b.ApprovedDate).getTime())[0];
          if (firstApproval?.ApprovedDate) {
            completedDate = new Date(firstApproval.ApprovedDate);
          }
        }

        levelStatuses.push({
          level: level.level,
          status,
          approvedBy: approvedBy.length > 0 ? approvedBy : undefined,
          rejectedBy: rejectedBy.length > 0 ? rejectedBy : undefined,
          pendingApprovers,
          completedDate
        });
      }

      // Determine overall status and progression
      const currentLevelStatus = levelStatuses.find(l => l.level === chain.CurrentLevel);
      const anyRejected = levelStatuses.some(l => l.status === ApprovalStatus.Rejected);
      const allApproved = levelStatuses.every(l => l.status === ApprovalStatus.Approved);

      let overallStatus = ApprovalStatus.Pending;
      let canProgress = false;
      let nextApprovers: number[] | undefined;

      if (anyRejected) {
        overallStatus = ApprovalStatus.Rejected;
      } else if (allApproved) {
        overallStatus = ApprovalStatus.Approved;
      } else if (currentLevelStatus?.status === ApprovalStatus.Approved) {
        canProgress = true;
        const nextLevel = levels.find(l => l.level === chain.CurrentLevel + 1);
        if (nextLevel) {
          nextApprovers = nextLevel.approverIds;
        }
      }

      return {
        chainId,
        processId: chain.ProcessID,
        currentLevel: chain.CurrentLevel,
        totalLevels: levels.length,
        levelStatuses,
        overallStatus,
        canProgress,
        nextApprovers
      };
    } catch (error) {
      logger.error('MultiLevelApprovalService', `Error getting approval status for chain ${chainId}`, error);
      return null;
    }
  }

  /**
   * Progress to next approval level
   */
  public async progressToNextLevel(chainId: number): Promise<boolean> {
    try {
      const status = await this.getApprovalStatus(chainId);
      if (!status || !status.canProgress) {
        return false;
      }

      const nextLevel = status.currentLevel + 1;
      if (nextLevel > status.totalLevels) {
        // Chain complete
        await this.sp.web.lists.getByTitle(this.APPROVAL_CHAINS_LIST)
          .items.getById(chainId)
          .update({
            OverallStatus: ApprovalStatus.Approved,
            CompletedDate: new Date().toISOString()
          });

        await this.notifyChainComplete(chainId, true);
        return true;
      }

      // Progress to next level
      await this.sp.web.lists.getByTitle(this.APPROVAL_CHAINS_LIST)
        .items.getById(chainId)
        .update({
          CurrentLevel: nextLevel
        });

      // Notify next level approvers
      if (status.nextApprovers && status.nextApprovers.length > 0) {
        await this.notifyLevelApprovers(chainId, nextLevel, status.nextApprovers);
      }

      logger.info('MultiLevelApprovalService',
        `Progressed chain ${chainId} to level ${nextLevel}`);

      return true;
    } catch (error) {
      logger.error('MultiLevelApprovalService',
        `Error progressing chain ${chainId}`, error);
      return false;
    }
  }

  /**
   * Handle approval decision
   */
  public async processApprovalDecision(
    chainId: number,
    approverId: number,
    approved: boolean,
    comments?: string
  ): Promise<boolean> {
    try {
      const status = await this.getApprovalStatus(chainId);
      if (!status) return false;

      // Record the decision
      await this.sp.web.lists.getByTitle(this.APPROVALS_LIST)
        .items.add({
          ApprovalChainId: chainId,
          ProcessId: status.processId,
          Level: status.currentLevel,
          ApproverId: approverId,
          Status: approved ? ApprovalStatus.Approved : ApprovalStatus.Rejected,
          ApprovedDate: new Date().toISOString(),
          Comments: comments
        });

      // If rejected, fail the entire chain
      if (!approved) {
        await this.sp.web.lists.getByTitle(this.APPROVAL_CHAINS_LIST)
          .items.getById(chainId)
          .update({
            OverallStatus: ApprovalStatus.Rejected,
            CompletedDate: new Date().toISOString()
          });

        await this.notifyChainComplete(chainId, false);
        return true;
      }

      // Check if current level is now complete
      const updatedStatus = await this.getApprovalStatus(chainId);
      if (updatedStatus?.canProgress) {
        await this.progressToNextLevel(chainId);
      }

      return true;
    } catch (error) {
      logger.error('MultiLevelApprovalService',
        `Error processing approval decision`, error);
      return false;
    }
  }

  /**
   * Notify approvers at a specific level
   */
  private async notifyLevelApprovers(
    chainId: number,
    level: number,
    approverIds: number[]
  ): Promise<void> {
    try {
      for (const approverId of approverIds) {
        await this.sp.web.lists.getByTitle('JML_Notifications').items.add({
          Title: 'Approval Required',
          RecipientId: approverId,
          Message: `Your approval is required (Level ${level}). Please review and approve or reject.`,
          NotificationType: 'InApp',
          Status: 'Pending',
          RelatedItemId: chainId,
          RelatedItemType: 'ApprovalChain',
          Priority: 'High',
          CreatedDate: new Date().toISOString()
        });
      }

      logger.info('MultiLevelApprovalService',
        `Notified ${approverIds.length} approvers for level ${level}`);
    } catch (error) {
      logger.error('MultiLevelApprovalService', 'Error notifying approvers', error);
    }
  }

  /**
   * Check for and process overdue approvals (auto-escalation)
   */
  public async processOverdueApprovals(): Promise<IApprovalEscalationResult[]> {
    const results: IApprovalEscalationResult[] = [];

    try {
      // Get all pending approval chains
      const pendingChains = await this.sp.web.lists.getByTitle(this.APPROVAL_CHAINS_LIST)
        .items
        .filter(`OverallStatus eq 'Pending'`)
        .select('Id', 'Levels', 'CurrentLevel', 'ProcessID', 'ChainName', 'LevelStartDate')();

      for (const chain of pendingChains) {
        const escalationResult = await this.checkAndEscalateChain(chain);
        if (escalationResult) {
          results.push(escalationResult);
        }
      }

      if (results.length > 0) {
        logger.info('MultiLevelApprovalService',
          `Processed ${results.length} approval escalations`);
      }

      return results;
    } catch (error) {
      logger.error('MultiLevelApprovalService', 'Error processing overdue approvals', error);
      return results;
    }
  }

  /**
   * Check and escalate a single approval chain if overdue
   */
  private async checkAndEscalateChain(
    chain: Record<string, unknown>
  ): Promise<IApprovalEscalationResult | null> {
    try {
      const levels: IApprovalLevel[] = JSON.parse(chain.Levels as string || '[]');
      const currentLevel = chain.CurrentLevel as number;
      const levelStartDate = chain.LevelStartDate ? new Date(chain.LevelStartDate as string) : null;

      if (!levelStartDate) return null;

      const currentLevelConfig = levels.find(l => l.level === currentLevel);
      if (!currentLevelConfig || !currentLevelConfig.escalationDays) return null;

      // Calculate if escalation is due
      const now = new Date();
      const daysSinceStart = Math.floor(
        (now.getTime() - levelStartDate.getTime()) / (1000 * 60 * 60 * 24)
      );

      if (daysSinceStart < currentLevelConfig.escalationDays) return null;

      // Check if already escalated
      const existingEscalations = await this.sp.web.lists.getByTitle(this.APPROVALS_LIST)
        .items
        .filter(`ApprovalChainId eq ${chain.Id} and Level eq ${currentLevel} and IsEscalation eq 1`)();

      if (existingEscalations.length > 0) return null; // Already escalated

      // Perform escalation
      const escalationApproverIds = currentLevelConfig.escalationApproverIds || [];

      if (escalationApproverIds.length === 0) {
        logger.warn('MultiLevelApprovalService',
          `Chain ${chain.Id} is overdue but has no escalation approvers configured`);
        return null;
      }

      // Add escalation approvers to the pending list
      for (const approverId of escalationApproverIds) {
        await this.sp.web.lists.getByTitle(this.APPROVALS_LIST)
          .items.add({
            ApprovalChainId: chain.Id,
            ProcessId: chain.ProcessID,
            Level: currentLevel,
            ApproverId: approverId,
            Status: ApprovalStatus.Pending,
            IsEscalation: true,
            EscalatedDate: new Date().toISOString()
          });
      }

      // Notify escalation approvers
      await this.notifyEscalationApprovers(
        chain.Id as number,
        chain.ChainName as string,
        currentLevel,
        escalationApproverIds,
        daysSinceStart
      );

      // Notify original approvers that escalation occurred
      await this.notifyOriginalApproversOfEscalation(
        chain.Id as number,
        currentLevel,
        currentLevelConfig.approverIds
      );

      logger.info('MultiLevelApprovalService',
        `Escalated chain ${chain.Id} level ${currentLevel} after ${daysSinceStart} days`);

      return {
        chainId: chain.Id as number,
        processId: chain.ProcessID as number,
        level: currentLevel,
        daysOverdue: daysSinceStart - currentLevelConfig.escalationDays,
        escalatedToIds: escalationApproverIds,
        originalApproverIds: currentLevelConfig.approverIds,
        escalatedAt: new Date()
      };
    } catch (error) {
      logger.error('MultiLevelApprovalService',
        `Error checking escalation for chain ${chain.Id}`, error);
      return null;
    }
  }

  /**
   * Notify escalation approvers
   */
  private async notifyEscalationApprovers(
    chainId: number,
    chainName: string,
    level: number,
    approverIds: number[],
    daysWaiting: number
  ): Promise<void> {
    try {
      for (const approverId of approverIds) {
        await this.sp.web.lists.getByTitle('JML_Notifications').items.add({
          Title: 'Escalated Approval Required',
          RecipientId: approverId,
          Message: `An approval request "${chainName}" (Level ${level}) has been escalated to you after ${daysWaiting} days without response. Please review and take action urgently.`,
          NotificationType: 'InApp',
          Status: 'Pending',
          RelatedItemId: chainId,
          RelatedItemType: 'ApprovalChain',
          Priority: 'Urgent',
          CreatedDate: new Date().toISOString()
        });
      }
    } catch (error) {
      logger.error('MultiLevelApprovalService', 'Error notifying escalation approvers', error);
    }
  }

  /**
   * Notify original approvers that their request was escalated
   */
  private async notifyOriginalApproversOfEscalation(
    chainId: number,
    level: number,
    originalApproverIds: number[]
  ): Promise<void> {
    try {
      for (const approverId of originalApproverIds) {
        await this.sp.web.lists.getByTitle('JML_Notifications').items.add({
          Title: 'Approval Request Escalated',
          RecipientId: approverId,
          Message: `An approval request assigned to you (Level ${level}) has been escalated due to timeout. The request has been routed to escalation contacts.`,
          NotificationType: 'InApp',
          Status: 'Pending',
          RelatedItemId: chainId,
          RelatedItemType: 'ApprovalChain',
          Priority: 'Normal',
          CreatedDate: new Date().toISOString()
        });
      }
    } catch (error) {
      logger.error('MultiLevelApprovalService', 'Error notifying original approvers', error);
    }
  }

  /**
   * Create a new approval chain with levels
   */
  public async createApprovalChain(
    processId: number,
    chainName: string,
    levels: IApprovalLevel[]
  ): Promise<number> {
    try {
      const result = await this.sp.web.lists.getByTitle(this.APPROVAL_CHAINS_LIST)
        .items.add({
          ProcessID: processId,
          ChainName: chainName,
          Levels: JSON.stringify(levels),
          CurrentLevel: 1,
          OverallStatus: ApprovalStatus.Pending,
          LevelStartDate: new Date().toISOString(),
          CreatedDate: new Date().toISOString()
        });

      // Notify first level approvers
      const firstLevel = levels.find(l => l.level === 1);
      if (firstLevel) {
        await this.notifyLevelApprovers(result.data.Id, 1, firstLevel.approverIds);
      }

      logger.info('MultiLevelApprovalService',
        `Created approval chain ${result.data.Id} with ${levels.length} levels`);

      return result.data.Id;
    } catch (error) {
      logger.error('MultiLevelApprovalService', 'Error creating approval chain', error);
      throw error;
    }
  }

  /**
   * Get pending approvals for a user
   */
  public async getPendingApprovalsForUser(userId: number): Promise<IUserPendingApproval[]> {
    try {
      const pendingApprovals = await this.sp.web.lists.getByTitle(this.APPROVALS_LIST)
        .items
        .filter(`ApproverId eq ${userId} and Status eq 'Pending'`)
        .select('Id', 'ApprovalChainId', 'Level', 'IsEscalation', 'EscalatedDate', 'Created')();

      const results: IUserPendingApproval[] = [];

      for (const approval of pendingApprovals) {
        const chain = await this.sp.web.lists.getByTitle(this.APPROVAL_CHAINS_LIST)
          .items.getById(approval.ApprovalChainId)
          .select('ProcessID', 'ChainName', 'Levels')();

        const levels: IApprovalLevel[] = JSON.parse(chain.Levels || '[]');
        const levelConfig = levels.find(l => l.level === approval.Level);

        results.push({
          approvalId: approval.Id,
          chainId: approval.ApprovalChainId,
          chainName: chain.ChainName,
          processId: chain.ProcessID,
          level: approval.Level,
          totalLevels: levels.length,
          isEscalation: approval.IsEscalation || false,
          escalationDays: levelConfig?.escalationDays,
          createdDate: new Date(approval.Created),
          daysWaiting: Math.floor(
            (new Date().getTime() - new Date(approval.Created).getTime()) / (1000 * 60 * 60 * 24)
          )
        });
      }

      return results;
    } catch (error) {
      logger.error('MultiLevelApprovalService', 'Error getting pending approvals for user', error);
      return [];
    }
  }

  /**
   * Notify that approval chain is complete
   */
  private async notifyChainComplete(chainId: number, approved: boolean): Promise<void> {
    try {
      const chain = await this.sp.web.lists.getByTitle(this.APPROVAL_CHAINS_LIST)
        .items.getById(chainId)
        .select('ProcessID', 'ChainName')();

      // Get process owner/initiator
      const process = await this.sp.web.lists.getByTitle('JML_Processes')
        .items.getById(chain.ProcessID)
        .select('CreatedById')();

      await this.sp.web.lists.getByTitle('JML_Notifications').items.add({
        Title: approved ? 'Approval Complete' : 'Approval Rejected',
        RecipientId: process.CreatedById,
        Message: approved
          ? `The approval chain "${chain.ChainName}" has been fully approved.`
          : `The approval chain "${chain.ChainName}" has been rejected.`,
        NotificationType: 'InApp',
        Status: 'Pending',
        RelatedItemId: chainId,
        RelatedItemType: 'ApprovalChain',
        Priority: approved ? 'Normal' : 'High',
        CreatedDate: new Date().toISOString()
      });
    } catch (error) {
      logger.error('MultiLevelApprovalService', 'Error notifying chain complete', error);
    }
  }
}

// ============================================================================
// SCHEDULED NOTIFICATION SERVICE
// ============================================================================

export class ScheduledNotificationService {
  private sp: SPFI;
  private readonly SCHEDULED_NOTIFICATIONS_LIST = 'JML_ScheduledNotifications';
  private readonly NOTIFICATIONS_LIST = 'JML_Notifications';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Schedule a notification for future delivery
   */
  public async scheduleNotification(
    notification: Omit<IScheduledNotification, 'id' | 'status' | 'retryCount' | 'createdDate'>
  ): Promise<number> {
    try {
      const result = await this.sp.web.lists.getByTitle(this.SCHEDULED_NOTIFICATIONS_LIST)
        .items.add({
          ProcessId: notification.processId,
          WorkflowInstanceId: notification.workflowInstanceId,
          RecipientIds: JSON.stringify(notification.recipientIds),
          RecipientEmails: JSON.stringify(notification.recipientEmails),
          Subject: notification.subject,
          Body: notification.body,
          NotificationType: notification.notificationType,
          ScheduledDate: notification.scheduledDate.toISOString(),
          Status: 'Pending',
          RetryCount: 0,
          MaxRetries: notification.maxRetries,
          Priority: notification.priority,
          CreatedDate: new Date().toISOString()
        });

      logger.info('ScheduledNotificationService',
        `Scheduled notification ${result.data.Id} for ${notification.scheduledDate.toISOString()}`);

      return result.data.Id;
    } catch (error) {
      logger.error('ScheduledNotificationService', 'Error scheduling notification', error);
      throw error;
    }
  }

  /**
   * Process all due scheduled notifications
   */
  public async processDueNotifications(): Promise<IScheduledNotificationResult[]> {
    const results: IScheduledNotificationResult[] = [];

    try {
      const now = new Date().toISOString();

      // Get pending notifications that are due
      const dueNotifications = await this.sp.web.lists.getByTitle(this.SCHEDULED_NOTIFICATIONS_LIST)
        .items
        .filter(`Status eq 'Pending' and ScheduledDate le '${now}'`)
        .orderBy('Priority', false)
        .orderBy('ScheduledDate', true)
        .top(50)();

      for (const notif of dueNotifications) {
        const result = await this.processScheduledNotification(notif);
        results.push(result);
      }

      logger.info('ScheduledNotificationService',
        `Processed ${results.length} scheduled notifications`);

      return results;
    } catch (error) {
      logger.error('ScheduledNotificationService', 'Error processing due notifications', error);
      return results;
    }
  }

  /**
   * Process a single scheduled notification
   */
  private async processScheduledNotification(
    notif: Record<string, unknown>
  ): Promise<IScheduledNotificationResult> {
    const notificationId = notif.Id as number;
    const recipientIds: number[] = JSON.parse(notif.RecipientIds as string || '[]');

    try {
      // Create actual notifications for each recipient
      for (const recipientId of recipientIds) {
        await this.sp.web.lists.getByTitle(this.NOTIFICATIONS_LIST)
          .items.add({
            Title: notif.Subject,
            RecipientId: recipientId,
            Message: notif.Body,
            NotificationType: notif.NotificationType,
            Status: 'Pending',
            RelatedItemId: notif.ProcessId,
            RelatedItemType: 'Process',
            Priority: notif.Priority,
            CreatedDate: new Date().toISOString()
          });
      }

      // Mark as sent
      await this.sp.web.lists.getByTitle(this.SCHEDULED_NOTIFICATIONS_LIST)
        .items.getById(notificationId)
        .update({
          Status: 'Sent',
          SentDate: new Date().toISOString()
        });

      logger.info('ScheduledNotificationService',
        `Successfully sent scheduled notification ${notificationId}`);

      return {
        notificationId,
        success: true,
        status: 'Sent',
        recipientCount: recipientIds.length
      };
    } catch (error) {
      const retryCount = (notif.RetryCount as number || 0) + 1;
      const maxRetries = notif.MaxRetries as number || 3;

      if (retryCount >= maxRetries) {
        // Max retries reached, mark as failed
        await this.sp.web.lists.getByTitle(this.SCHEDULED_NOTIFICATIONS_LIST)
          .items.getById(notificationId)
          .update({
            Status: 'Failed',
            RetryCount: retryCount,
            ErrorMessage: error instanceof Error ? error.message : 'Unknown error'
          });

        return {
          notificationId,
          success: false,
          status: 'Failed',
          error: error instanceof Error ? error.message : 'Unknown error',
          recipientCount: recipientIds.length
        };
      } else {
        // Reschedule with backoff
        const backoffMinutes = Math.pow(2, retryCount) * 5; // 10, 20, 40 minutes
        const nextScheduled = new Date();
        nextScheduled.setMinutes(nextScheduled.getMinutes() + backoffMinutes);

        await this.sp.web.lists.getByTitle(this.SCHEDULED_NOTIFICATIONS_LIST)
          .items.getById(notificationId)
          .update({
            RetryCount: retryCount,
            ScheduledDate: nextScheduled.toISOString(),
            ErrorMessage: error instanceof Error ? error.message : 'Unknown error'
          });

        return {
          notificationId,
          success: false,
          status: 'Rescheduled',
          error: error instanceof Error ? error.message : 'Unknown error',
          recipientCount: recipientIds.length
        };
      }
    }
  }

  /**
   * Cancel a scheduled notification
   */
  public async cancelNotification(notificationId: number): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(this.SCHEDULED_NOTIFICATIONS_LIST)
        .items.getById(notificationId)
        .update({
          Status: 'Cancelled'
        });

      logger.info('ScheduledNotificationService',
        `Cancelled scheduled notification ${notificationId}`);
    } catch (error) {
      logger.error('ScheduledNotificationService', 'Error cancelling notification', error);
      throw error;
    }
  }

  /**
   * Get pending scheduled notifications for a process
   */
  public async getPendingNotifications(processId: number): Promise<IScheduledNotification[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.SCHEDULED_NOTIFICATIONS_LIST)
        .items
        .filter(`ProcessId eq ${processId} and Status eq 'Pending'`)
        .orderBy('ScheduledDate', true)();

      return items.map(item => ({
        id: item.Id,
        processId: item.ProcessId,
        workflowInstanceId: item.WorkflowInstanceId,
        recipientIds: JSON.parse(item.RecipientIds || '[]'),
        recipientEmails: JSON.parse(item.RecipientEmails || '[]'),
        subject: item.Subject,
        body: item.Body,
        notificationType: item.NotificationType,
        scheduledDate: new Date(item.ScheduledDate),
        status: item.Status,
        retryCount: item.RetryCount,
        maxRetries: item.MaxRetries,
        priority: item.Priority,
        createdDate: new Date(item.CreatedDate)
      }));
    } catch (error) {
      logger.error('ScheduledNotificationService', 'Error getting pending notifications', error);
      return [];
    }
  }
}

// ============================================================================
// PARALLEL STEP SERVICE
// ============================================================================

export class ParallelStepService {
  private sp: SPFI;
  private readonly PARALLEL_CONTEXTS_LIST = 'JML_ParallelExecutions';
  private readonly WORKFLOW_INSTANCES_LIST = 'JML_WorkflowInstances';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Initialize parallel execution
   */
  public async initializeParallelExecution(
    workflowInstanceId: number,
    parallelStepId: string,
    branchStepIds: string[],
    joinType: 'all' | 'any' | 'first',
    joinStepId: string
  ): Promise<number> {
    try {
      const branches: IParallelBranchStatus[] = branchStepIds.map((stepId, index) => ({
        branchId: `branch_${index}`,
        stepId,
        stepName: `Branch ${index + 1}`,
        status: StepStatus.Pending
      }));

      const context: IParallelExecutionContext = {
        workflowInstanceId,
        parallelStepId,
        branches,
        joinType,
        joinStepId,
        startedDate: new Date(),
        status: 'Running'
      };

      const result = await this.sp.web.lists.getByTitle(this.PARALLEL_CONTEXTS_LIST)
        .items.add({
          WorkflowInstanceId: workflowInstanceId,
          ParallelStepId: parallelStepId,
          Branches: JSON.stringify(branches),
          JoinType: joinType,
          JoinStepId: joinStepId,
          StartedDate: new Date().toISOString(),
          Status: 'Running'
        });

      logger.info('ParallelStepService',
        `Initialized parallel execution ${result.data.Id} with ${branchStepIds.length} branches`);

      return result.data.Id;
    } catch (error) {
      logger.error('ParallelStepService', 'Error initializing parallel execution', error);
      throw error;
    }
  }

  /**
   * Update branch status
   */
  public async updateBranchStatus(
    parallelExecutionId: number,
    branchId: string,
    status: StepStatus,
    outputVariables?: Record<string, unknown>,
    error?: string
  ): Promise<void> {
    try {
      const context = await this.sp.web.lists.getByTitle(this.PARALLEL_CONTEXTS_LIST)
        .items.getById(parallelExecutionId)
        .select('Branches', 'JoinType', 'Status')();

      const branches: IParallelBranchStatus[] = JSON.parse(context.Branches || '[]');
      const branch = branches.find(b => b.branchId === branchId);

      if (branch) {
        branch.status = status;
        branch.outputVariables = outputVariables;
        branch.error = error;

        if (status === StepStatus.Completed || status === StepStatus.Failed) {
          branch.completedDate = new Date();
        }
      }

      // Check if parallel execution is complete
      const syncResult = this.evaluateParallelCompletion(branches, context.JoinType);

      await this.sp.web.lists.getByTitle(this.PARALLEL_CONTEXTS_LIST)
        .items.getById(parallelExecutionId)
        .update({
          Branches: JSON.stringify(branches),
          Status: syncResult.allBranchesComplete ? 'Completed' :
                  syncResult.failedBranches.length > 0 ? 'Partial' : 'Running',
          CompletedDate: syncResult.allBranchesComplete ? new Date().toISOString() : null,
          MergedVariables: JSON.stringify(syncResult.mergedVariables)
        });

      logger.info('ParallelStepService',
        `Updated branch ${branchId} to ${status}`);
    } catch (error) {
      logger.error('ParallelStepService', 'Error updating branch status', error);
      throw error;
    }
  }

  /**
   * Check if parallel execution can proceed to join step
   */
  public async checkParallelSync(parallelExecutionId: number): Promise<IParallelSyncResult> {
    try {
      const context = await this.sp.web.lists.getByTitle(this.PARALLEL_CONTEXTS_LIST)
        .items.getById(parallelExecutionId)
        .select('ParallelStepId', 'Branches', 'JoinType')();

      const branches: IParallelBranchStatus[] = JSON.parse(context.Branches || '[]');

      return this.evaluateParallelCompletion(branches, context.JoinType, context.ParallelStepId);
    } catch (error) {
      logger.error('ParallelStepService', 'Error checking parallel sync', error);
      return {
        parallelStepId: '',
        allBranchesComplete: false,
        completedBranches: [],
        pendingBranches: [],
        failedBranches: [],
        canProceedToJoin: false,
        mergedVariables: {}
      };
    }
  }

  /**
   * Evaluate if parallel execution is complete
   */
  private evaluateParallelCompletion(
    branches: IParallelBranchStatus[],
    joinType: string,
    parallelStepId: string = ''
  ): IParallelSyncResult {
    const completedBranches = branches
      .filter(b => b.status === StepStatus.Completed)
      .map(b => b.branchId);

    const failedBranches = branches
      .filter(b => b.status === StepStatus.Failed)
      .map(b => b.branchId);

    const pendingBranches = branches
      .filter(b => b.status === StepStatus.Pending || b.status === StepStatus.InProgress)
      .map(b => b.branchId);

    const allBranchesComplete = pendingBranches.length === 0;

    let canProceedToJoin = false;

    switch (joinType) {
      case 'all':
        canProceedToJoin = allBranchesComplete && failedBranches.length === 0;
        break;
      case 'any':
        canProceedToJoin = completedBranches.length > 0;
        break;
      case 'first':
        canProceedToJoin = completedBranches.length > 0 || failedBranches.length > 0;
        break;
    }

    // Merge output variables from completed branches
    const mergedVariables: Record<string, unknown> = {};
    branches
      .filter(b => b.status === StepStatus.Completed && b.outputVariables)
      .forEach(b => {
        Object.assign(mergedVariables, {
          [`${b.branchId}_output`]: b.outputVariables
        });
      });

    return {
      parallelStepId,
      allBranchesComplete,
      completedBranches,
      pendingBranches,
      failedBranches,
      canProceedToJoin,
      mergedVariables
    };
  }

  /**
   * Get parallel execution context
   */
  public async getParallelContext(workflowInstanceId: number, parallelStepId: string): Promise<IParallelExecutionContext | null> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.PARALLEL_CONTEXTS_LIST)
        .items
        .filter(`WorkflowInstanceId eq ${workflowInstanceId} and ParallelStepId eq '${parallelStepId}'`)
        .top(1)();

      if (items.length === 0) return null;

      const item = items[0];
      return {
        workflowInstanceId: item.WorkflowInstanceId,
        parallelStepId: item.ParallelStepId,
        branches: JSON.parse(item.Branches || '[]'),
        joinType: item.JoinType,
        joinStepId: item.JoinStepId,
        startedDate: new Date(item.StartedDate),
        status: item.Status
      };
    } catch (error) {
      logger.error('ParallelStepService', 'Error getting parallel context', error);
      return null;
    }
  }
}

// ============================================================================
// UNIFIED WORKFLOW ADVANCED SERVICE
// ============================================================================

/**
 * Unified service that coordinates all Phase 7 advanced features
 */
export class WorkflowAdvancedService {
  private sp: SPFI;

  public readonly taskDependency: TaskDependencyService;
  public readonly multiLevelApproval: MultiLevelApprovalService;
  public readonly scheduledNotification: ScheduledNotificationService;
  public readonly parallelStep: ParallelStepService;

  constructor(sp: SPFI) {
    this.sp = sp;
    this.taskDependency = new TaskDependencyService(sp);
    this.multiLevelApproval = new MultiLevelApprovalService(sp);
    this.scheduledNotification = new ScheduledNotificationService(sp);
    this.parallelStep = new ParallelStepService(sp);
  }

  /**
   * Process all pending workflow operations
   * Call this periodically (e.g., from a timer job or Azure Function)
   */
  public async processAllPending(): Promise<{
    scheduledNotificationsSent: number;
    approvalsEscalated: number;
    errors: string[];
  }> {
    const errors: string[] = [];
    let scheduledNotificationsSent = 0;
    let approvalsEscalated = 0;

    try {
      // Process scheduled notifications
      const notifResults = await this.scheduledNotification.processDueNotifications();
      scheduledNotificationsSent = notifResults.filter(r => r.success).length;

      // Log any notification errors
      notifResults
        .filter(r => !r.success && r.error)
        .forEach(r => errors.push(`Notification ${r.notificationId}: ${r.error}`));

      // Process overdue approvals (auto-escalation)
      const escalationResults = await this.multiLevelApproval.processOverdueApprovals();
      approvalsEscalated = escalationResults.length;

    } catch (error) {
      errors.push(`General processing error: ${error instanceof Error ? error.message : 'Unknown'}`);
    }

    logger.info('WorkflowAdvancedService',
      `Processed pending operations: ${scheduledNotificationsSent} notifications sent, ${approvalsEscalated} approvals escalated, ${errors.length} errors`);

    return {
      scheduledNotificationsSent,
      approvalsEscalated,
      errors
    };
  }

  /**
   * Handle task completion with dependency cascade
   */
  public async handleTaskCompletion(taskId: number): Promise<ICascadeUnblockResult> {
    return this.taskDependency.onTaskCompleted(taskId);
  }

  /**
   * Handle approval decision with multi-level progression
   */
  public async handleApprovalDecision(
    chainId: number,
    approverId: number,
    approved: boolean,
    comments?: string
  ): Promise<boolean> {
    return this.multiLevelApproval.processApprovalDecision(
      chainId,
      approverId,
      approved,
      comments
    );
  }

  /**
   * Check parallel step readiness
   */
  public async checkParallelStepReadiness(
    workflowInstanceId: number,
    parallelStepId: string
  ): Promise<IParallelSyncResult | null> {
    const context = await this.parallelStep.getParallelContext(
      workflowInstanceId,
      parallelStepId
    );

    if (!context) return null;

    // Get the stored item ID
    const items = await this.sp.web.lists.getByTitle('JML_ParallelExecutions')
      .items
      .filter(`WorkflowInstanceId eq ${workflowInstanceId} and ParallelStepId eq '${parallelStepId}'`)
      .select('Id')
      .top(1)();

    if (items.length === 0) return null;

    return this.parallelStep.checkParallelSync(items[0].Id);
  }
}

// Export all services and interfaces
export default WorkflowAdvancedService;
