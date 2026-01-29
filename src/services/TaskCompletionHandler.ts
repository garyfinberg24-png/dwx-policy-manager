// @ts-nocheck
/**
 * TaskCompletionHandler
 * Handles task completion events and coordinates with workflow engine
 *
 * This service is called whenever a task is marked as complete,
 * and ensures that:
 * 1. Process progress is updated
 * 2. Workflow is notified and potentially resumed
 * 3. Process status is synchronized
 * 4. Notifications are sent as needed
 */

import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import { logger } from './LoggingService';
import { ProcessOrchestrationService, ITaskCompletionResult } from './ProcessOrchestrationService';
import { ProcessService } from './ProcessService';
import { TaskDependencyService } from './TaskDependencyService';
import { TaskNotificationService } from './TaskNotificationService';
import { IJmlTaskAssignment } from '../models/IJmlTaskAssignment';
import { TaskStatus } from '../models/ICommon';
import { TaskNotificationType } from '../models/IJmlTaskEscalation';

// ============================================================================
// INTERFACES
// ============================================================================

/**
 * Options for completing a task
 */
export interface ITaskCompletionOptions {
  taskAssignmentId: number;
  completedByUserId: number;
  completedByUserName?: string;
  actualHours?: number;
  notes?: string;
  result?: Record<string, unknown>;
  skipWorkflowUpdate?: boolean;
}

/**
 * Extended completion result
 */
export interface IExtendedCompletionResult extends ITaskCompletionResult {
  taskTitle?: string;
  processId?: number;
  processProgress?: number;
  notificationsSent?: boolean;
  // INTEGRATION FIX: Task dependency unblocking results
  unblockedTaskIds?: number[];
  unblockedTaskCount?: number;
}

/**
 * Bulk completion options
 */
export interface IBulkCompletionOptions {
  taskAssignmentIds: number[];
  completedByUserId: number;
  completedByUserName?: string;
  notes?: string;
}

/**
 * Bulk completion result
 */
export interface IBulkCompletionResult {
  success: boolean;
  totalTasks: number;
  completedTasks: number;
  failedTasks: number;
  processesUpdated: number;
  workflowsResumed: number;
  errors: string[];
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Extract numeric process ID from ProcessID field which can be string or lookup object
 */
function extractProcessId(processId: string | { Id: number; Title: string } | undefined): number | undefined {
  if (!processId) return undefined;
  if (typeof processId === 'string') {
    const id = parseInt(processId, 10);
    return isNaN(id) ? undefined : id;
  }
  if (typeof processId === 'object' && 'Id' in processId) {
    return processId.Id;
  }
  return undefined;
}

/**
 * Convert ProcessID to string for notifications (handles both string and object types)
 */
function processIdToString(processId: string | { Id: number; Title: string } | undefined): string | undefined {
  if (!processId) return undefined;
  if (typeof processId === 'string') return processId;
  if (typeof processId === 'object' && 'Id' in processId) return processId.Id.toString();
  return undefined;
}

// ============================================================================
// TASK COMPLETION HANDLER
// ============================================================================

export class TaskCompletionHandler {
  private sp: SPFI;
  private context: WebPartContext;
  private orchestrationService: ProcessOrchestrationService;
  private processService: ProcessService;
  // INTEGRATION FIX: Task dependency and notification services for auto-unblocking
  private taskDependencyService: TaskDependencyService;
  private taskNotificationService: TaskNotificationService;

  constructor(sp: SPFI, context: WebPartContext) {
    this.sp = sp;
    this.context = context;
    this.orchestrationService = new ProcessOrchestrationService(sp, context);
    this.processService = new ProcessService(sp, context);
    // INTEGRATION FIX: Initialize dependency and notification services
    this.taskDependencyService = new TaskDependencyService(sp);
    this.taskNotificationService = new TaskNotificationService(sp, context);
  }

  // ============================================================================
  // SINGLE TASK COMPLETION
  // ============================================================================

  /**
   * Complete a single task and handle all related updates
   */
  public async completeTask(options: ITaskCompletionOptions): Promise<IExtendedCompletionResult> {
    try {
      logger.info('TaskCompletionHandler', `Completing task ${options.taskAssignmentId}`);

      // Step 1: Get the task assignment
      const task = await this.getTaskAssignment(options.taskAssignmentId);
      if (!task) {
        return {
          success: false,
          processUpdated: false,
          workflowResumed: false,
          processCompleted: false,
          error: 'Task assignment not found'
        };
      }

      // Check if already completed
      if (task.Status === TaskStatus.Completed) {
        return {
          success: true,
          processUpdated: false,
          workflowResumed: false,
          processCompleted: false,
          taskTitle: task.Title,
          processId: extractProcessId(task.ProcessID),
          error: 'Task already completed'
        };
      }

      // Step 2: Update the task status
      // Note: CompletionNotes is used instead of a CompletedById field
      await this.updateTaskStatus(options.taskAssignmentId, {
        Status: TaskStatus.Completed,
        ActualCompletionDate: new Date(),
        ActualHours: options.actualHours,
        PercentComplete: 100,
        CompletionNotes: options.completedByUserName
          ? `Completed by user ID ${options.completedByUserId} (${options.completedByUserName})`
          : `Completed by user ID ${options.completedByUserId}`,
        Notes: options.notes
          ? (task.Notes ? `${task.Notes}\n${options.notes}` : options.notes)
          : task.Notes
      });

      logger.info('TaskCompletionHandler', `Task ${options.taskAssignmentId} marked as completed`);

      // Step 3: Notify orchestration service (unless skipped)
      let orchestrationResult: ITaskCompletionResult = {
        success: true,
        processUpdated: false,
        workflowResumed: false,
        processCompleted: false
      };

      if (!options.skipWorkflowUpdate) {
        orchestrationResult = await this.orchestrationService.handleTaskCompletion(
          options.taskAssignmentId,
          options.completedByUserId,
          options.result
        );
      } else {
        // Still update process progress even if skipping workflow
        const processId = extractProcessId(task.ProcessID);
        if (processId) {
          await this.processService.recalculateProgress(processId);
          orchestrationResult.processUpdated = true;
        }
      }

      // Step 4: INTEGRATION FIX - Unblock dependent tasks
      // This is the critical fix for task dependency auto-unblocking
      let unblockedTaskIds: number[] = [];
      try {
        await this.taskDependencyService.onTaskCompleted(options.taskAssignmentId);

        // Get the list of tasks that were unblocked
        const dependentTasks = await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
          .filter(`DependsOnTaskId eq ${options.taskAssignmentId} and IsBlocked eq false`)
          .select('Id', 'Title', 'AssignedToId')();

        unblockedTaskIds = dependentTasks.map(t => t.Id);

        if (unblockedTaskIds.length > 0) {
          logger.info('TaskCompletionHandler',
            `Unblocked ${unblockedTaskIds.length} dependent tasks after completing task ${options.taskAssignmentId}`);

          // Send notifications to assignees of unblocked tasks
          for (const depTask of dependentTasks) {
            if (depTask.AssignedToId) {
              try {
                await this.sendTaskUnblockedNotification(
                  depTask.Id,
                  depTask.Title,
                  depTask.AssignedToId,
                  task.Title,
                  extractProcessId(task.ProcessID)
                );
              } catch (notifyErr) {
                logger.warn('TaskCompletionHandler',
                  `Failed to send unblock notification for task ${depTask.Id}`, notifyErr);
              }
            }
          }
        }
      } catch (depError) {
        // Don't fail task completion if dependency unblocking fails
        logger.warn('TaskCompletionHandler',
          `Error unblocking dependent tasks for ${options.taskAssignmentId}`, depError);
      }

      // Step 5: Send completion notification
      let notificationsSent = false;
      try {
        await this.sendCompletionNotification(task, options.completedByUserName || 'Unknown');
        notificationsSent = true;
      } catch (notifyError) {
        logger.warn('TaskCompletionHandler', 'Failed to send completion notification', notifyError);
      }

      // Step 6: Create audit log
      await this.createAuditLog(task, options);

      return {
        ...orchestrationResult,
        taskTitle: task.Title,
        processId: extractProcessId(task.ProcessID),
        processProgress: await this.getProcessProgress(extractProcessId(task.ProcessID)),
        notificationsSent,
        // INTEGRATION FIX: Include unblocked task info in result
        unblockedTaskIds,
        unblockedTaskCount: unblockedTaskIds.length
      };

    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : 'Unknown error';
      logger.error('TaskCompletionHandler', `Error completing task ${options.taskAssignmentId}`, error);
      return {
        success: false,
        processUpdated: false,
        workflowResumed: false,
        processCompleted: false,
        error: errorMsg
      };
    }
  }

  /**
   * Complete a task and check if approval is required
   */
  public async completeWithApproval(options: ITaskCompletionOptions): Promise<IExtendedCompletionResult> {
    const task = await this.getTaskAssignment(options.taskAssignmentId);

    if (task?.RequiresApproval) {
      // Set to pending approval instead of completed
      await this.updateTaskStatus(options.taskAssignmentId, {
        Status: 'Pending Approval' as TaskStatus,
        PercentComplete: 100,
        Notes: options.notes
          ? (task.Notes ? `${task.Notes}\n${options.notes}` : options.notes)
          : task.Notes
      });

      // Create approval request
      await this.createApprovalRequest(task, options.completedByUserId);

      return {
        success: true,
        processUpdated: false,
        workflowResumed: false,
        processCompleted: false,
        taskTitle: task.Title,
        error: 'Task requires approval'
      };
    }

    // No approval required - proceed with normal completion
    return this.completeTask(options);
  }

  // ============================================================================
  // BULK COMPLETION
  // ============================================================================

  /**
   * Complete multiple tasks at once
   */
  public async completeTasks(options: IBulkCompletionOptions): Promise<IBulkCompletionResult> {
    const result: IBulkCompletionResult = {
      success: true,
      totalTasks: options.taskAssignmentIds.length,
      completedTasks: 0,
      failedTasks: 0,
      processesUpdated: 0,
      workflowsResumed: 0,
      errors: []
    };

    const processesAffected = new Set<number>();
    const workflowsResumed = new Set<number>();

    for (const taskId of options.taskAssignmentIds) {
      try {
        const completionResult = await this.completeTask({
          taskAssignmentId: taskId,
          completedByUserId: options.completedByUserId,
          completedByUserName: options.completedByUserName,
          notes: options.notes,
          skipWorkflowUpdate: true // We'll update workflows in batch at the end
        });

        if (completionResult.success) {
          result.completedTasks++;
          if (completionResult.processId) {
            processesAffected.add(completionResult.processId);
          }
        } else {
          result.failedTasks++;
          if (completionResult.error) {
            result.errors.push(`Task ${taskId}: ${completionResult.error}`);
          }
        }
      } catch (error) {
        result.failedTasks++;
        result.errors.push(`Task ${taskId}: ${error instanceof Error ? error.message : 'Unknown error'}`);
      }
    }

    // Update all affected processes and workflows
    for (const processId of Array.from(processesAffected)) {
      try {
        // Recalculate progress
        await this.processService.recalculateProgress(processId);
        result.processesUpdated++;

        // Check if workflow needs to be resumed
        const workflowResult = await this.orchestrationService.handleTaskCompletion(
          options.taskAssignmentIds[0], // Use first task as trigger
          options.completedByUserId
        );

        if (workflowResult.workflowResumed) {
          result.workflowsResumed++;
        }
      } catch (error) {
        logger.warn('TaskCompletionHandler', `Error updating process ${processId}`, error);
      }
    }

    result.success = result.failedTasks === 0;
    return result;
  }

  // ============================================================================
  // TASK STATUS UPDATES
  // ============================================================================

  /**
   * Skip a task
   * INTEGRATION FIX: Now also triggers dependency unblocking for dependent tasks
   */
  public async skipTask(
    taskAssignmentId: number,
    skippedByUserId: number,
    reason?: string
  ): Promise<IExtendedCompletionResult> {
    try {
      const task = await this.getTaskAssignment(taskAssignmentId);
      if (!task) {
        return {
          success: false,
          processUpdated: false,
          workflowResumed: false,
          processCompleted: false,
          error: 'Task assignment not found'
        };
      }

      await this.updateTaskStatus(taskAssignmentId, {
        Status: TaskStatus.Skipped,
        Notes: reason
          ? (task.Notes ? `${task.Notes}\nSkipped: ${reason}` : `Skipped: ${reason}`)
          : task.Notes
      });

      // INTEGRATION FIX: Unblock dependent tasks (skipped tasks also unblock dependents)
      let unblockedTaskIds: number[] = [];
      try {
        await this.taskDependencyService.onTaskCompleted(taskAssignmentId);

        // Get the list of tasks that were unblocked
        const dependentTasks = await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
          .filter(`DependsOnTaskId eq ${taskAssignmentId} and IsBlocked eq false`)
          .select('Id', 'Title', 'AssignedToId')();

        unblockedTaskIds = dependentTasks.map(t => t.Id);

        if (unblockedTaskIds.length > 0) {
          logger.info('TaskCompletionHandler',
            `Unblocked ${unblockedTaskIds.length} dependent tasks after skipping task ${taskAssignmentId}`);

          // Send notifications to assignees of unblocked tasks
          for (const depTask of dependentTasks) {
            if (depTask.AssignedToId) {
              try {
                await this.sendTaskUnblockedNotification(
                  depTask.Id,
                  depTask.Title,
                  depTask.AssignedToId,
                  task.Title,
                  extractProcessId(task.ProcessID)
                );
              } catch (notifyErr) {
                logger.warn('TaskCompletionHandler',
                  `Failed to send unblock notification for task ${depTask.Id}`, notifyErr);
              }
            }
          }
        }
      } catch (depError) {
        logger.warn('TaskCompletionHandler',
          `Error unblocking dependent tasks for skipped task ${taskAssignmentId}`, depError);
      }

      // Handle like completion for workflow purposes
      const orchestrationResult = await this.orchestrationService.handleTaskCompletion(
        taskAssignmentId,
        skippedByUserId,
        { skipped: true, reason }
      );

      return {
        ...orchestrationResult,
        unblockedTaskIds,
        unblockedTaskCount: unblockedTaskIds.length
      };

    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : 'Unknown error';
      logger.error('TaskCompletionHandler', `Error skipping task ${taskAssignmentId}`, error);
      return {
        success: false,
        processUpdated: false,
        workflowResumed: false,
        processCompleted: false,
        error: errorMsg
      };
    }
  }

  /**
   * Block a task
   */
  public async blockTask(
    taskAssignmentId: number,
    blockedByUserId: number,
    reason: string
  ): Promise<boolean> {
    try {
      const task = await this.getTaskAssignment(taskAssignmentId);
      if (!task) return false;

      await this.updateTaskStatus(taskAssignmentId, {
        Status: TaskStatus.Blocked,
        IsBlocked: true,
        Notes: task.Notes
          ? `${task.Notes}\nBlocked: ${reason}`
          : `Blocked: ${reason}`
      });

      // Create notification for manager
      const processId = extractProcessId(task.ProcessID);
      if (processId) {
        const process = await this.processService.getById(processId);
        if (process.ManagerId) {
          await this.sendBlockedNotification(task, process.ManagerId, reason);
        }
      }

      return true;
    } catch (error) {
      logger.error('TaskCompletionHandler', `Error blocking task ${taskAssignmentId}`, error);
      return false;
    }
  }

  /**
   * Unblock a task
   */
  public async unblockTask(
    taskAssignmentId: number,
    unblockedByUserId: number,
    notes?: string
  ): Promise<boolean> {
    try {
      const task = await this.getTaskAssignment(taskAssignmentId);
      if (!task) return false;

      await this.updateTaskStatus(taskAssignmentId, {
        Status: TaskStatus.InProgress,
        IsBlocked: false,
        Notes: notes
          ? (task.Notes ? `${task.Notes}\nUnblocked: ${notes}` : `Unblocked: ${notes}`)
          : task.Notes
      });

      return true;
    } catch (error) {
      logger.error('TaskCompletionHandler', `Error unblocking task ${taskAssignmentId}`, error);
      return false;
    }
  }

  // ============================================================================
  // PRIVATE HELPERS
  // ============================================================================

  /**
   * Get task assignment by ID
   */
  private async getTaskAssignment(taskId: number): Promise<IJmlTaskAssignment | null> {
    try {
      const item = await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
        .getById(taskId)
        .select(
          'Id', 'Title', 'ProcessID', 'TaskID', 'Status', 'Priority',
          'DueDate', 'Notes', 'RequiresApproval', 'AssignedToId',
          'PercentComplete', 'IsBlocked'
        )();
      return item as IJmlTaskAssignment;
    } catch {
      return null;
    }
  }

  /**
   * Update task status
   */
  private async updateTaskStatus(
    taskId: number,
    updates: Partial<IJmlTaskAssignment>
  ): Promise<void> {
    await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
      .getById(taskId)
      .update(updates);
  }

  /**
   * Get process progress
   */
  private async getProcessProgress(processId: number | undefined): Promise<number | undefined> {
    if (!processId) return undefined;

    try {
      const process = await this.processService.getById(processId);
      return process.ProgressPercentage;
    } catch {
      return undefined;
    }
  }

  /**
   * Send task completion notification
   */
  private async sendCompletionNotification(
    task: IJmlTaskAssignment,
    completedByName: string
  ): Promise<void> {
    const processId = extractProcessId(task.ProcessID);
    if (!processId) return;

    const process = await this.processService.getById(processId);

    // Notify process owner or manager
    const recipientId = process.ProcessOwnerId || process.ManagerId;
    if (!recipientId) return;

    await this.sp.web.lists.getByTitle('PM_Notifications').items.add({
      Title: `Task Completed: ${task.Title}`,
      NotificationType: 'TaskCompleted',
      MessageBody: `Task "${task.Title}" has been completed by ${completedByName}.`,
      Priority: 'Normal',
      RecipientId: recipientId,
      ProcessId: processId.toString(),
      TaskId: task.Id?.toString(),
      Status: 'Pending'
    });
  }

  /**
   * Send blocked task notification
   */
  private async sendBlockedNotification(
    task: IJmlTaskAssignment,
    managerId: number,
    reason: string
  ): Promise<void> {
    await this.sp.web.lists.getByTitle('PM_Notifications').items.add({
      Title: `Task Blocked: ${task.Title}`,
      NotificationType: 'TaskBlocked',
      MessageBody: `Task "${task.Title}" has been blocked. Reason: ${reason}`,
      Priority: 'High',
      RecipientId: managerId,
      ProcessId: processIdToString(task.ProcessID),
      TaskId: task.Id?.toString(),
      Status: 'Pending'
    });
  }

  /**
   * INTEGRATION FIX: Send notification when a task is unblocked
   * Notifies the assignee that their previously blocked task is now available
   */
  private async sendTaskUnblockedNotification(
    taskId: number,
    taskTitle: string,
    assigneeId: number,
    completedBlockingTaskTitle: string,
    processId: number | undefined
  ): Promise<void> {
    // Create in-app notification in PM_Notifications list
    await this.sp.web.lists.getByTitle('PM_Notifications').items.add({
      Title: 'Task Now Available',
      NotificationType: 'TaskUnblocked',
      MessageBody: `Your task "${taskTitle}" is now available. The prerequisite task "${completedBlockingTaskTitle}" has been completed.`,
      Priority: 'Normal',
      RecipientId: assigneeId,
      ProcessId: processId?.toString(),
      TaskId: taskId.toString(),
      LinkUrl: `/sites/JML/SitePages/MyTasks.aspx?taskId=${taskId}`,
      Status: 'Pending',
      IsRead: false,
      SentDate: new Date()
    });

    // Also send via TaskNotificationService for email/Teams delivery
    try {
      const taskForNotification: Partial<IJmlTaskAssignment> = {
        Id: taskId,
        Title: taskTitle,
        AssignedToId: assigneeId
      };

      await this.taskNotificationService.sendNotification({
        TaskAssignmentId: taskId,
        NotificationType: TaskNotificationType.Reminder, // Use Reminder type for unblock notifications
        ScheduledFor: new Date(),
        Priority: 'Normal',
        Recipients: [assigneeId],
        Message: `Your task "${taskTitle}" is now available! The prerequisite task "${completedBlockingTaskTitle}" has been completed. You can now start working on this task.`,
        IsProcessed: false
      });

      logger.info('TaskCompletionHandler',
        `Sent task unblocked notification to user ${assigneeId} for task ${taskId}`);
    } catch (error) {
      // Don't fail if email notification fails - in-app notification was already created
      logger.warn('TaskCompletionHandler',
        `Failed to send email/Teams notification for unblocked task ${taskId}`, error);
    }
  }

  /**
   * Create approval request
   */
  private async createApprovalRequest(
    task: IJmlTaskAssignment,
    requestedByUserId: number
  ): Promise<void> {
    const processId = extractProcessId(task.ProcessID);
    if (!processId) return;

    const process = await this.processService.getById(processId);

    await this.sp.web.lists.getByTitle('PM_Approvals').items.add({
      Title: `Approval Required: ${task.Title}`,
      ApprovalType: 'TaskCompletion',
      Status: 'Pending',
      RequestedById: requestedByUserId,
      ApproverId: process.ManagerId || process.ProcessOwnerId,
      ProcessId: processId.toString(),
      TaskId: task.Id?.toString(),
      Comments: `Task completion requires approval for: ${task.Title}`
    });
  }

  /**
   * Create audit log entry
   */
  private async createAuditLog(
    task: IJmlTaskAssignment,
    options: ITaskCompletionOptions
  ): Promise<void> {
    try {
      const processId = extractProcessId(task.ProcessID);

      await this.sp.web.lists.getByTitle('PM_AuditLogs').items.add({
        Title: `Task Completed: ${task.Title}`,
        EventType: 'TaskCompleted',
        EntityType: 'TaskAssignment',
        EntityId: task.Id,
        ProcessId: processId || undefined,
        Action: 'Complete',
        Description: `Task "${task.Title}" completed by user ${options.completedByUserId}`,
        AdditionalData: JSON.stringify({
          completedByUserId: options.completedByUserId,
          completedByUserName: options.completedByUserName,
          actualHours: options.actualHours,
          notes: options.notes
        })
      });
    } catch (error) {
      logger.warn('TaskCompletionHandler', 'Failed to create audit log', error);
    }
  }
}

export default TaskCompletionHandler;
