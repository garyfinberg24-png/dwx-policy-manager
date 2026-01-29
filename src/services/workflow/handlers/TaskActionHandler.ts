// @ts-nocheck
/**
 * TaskActionHandler
 * Handles task-related workflow actions
 * Creates, assigns, and tracks tasks within workflow execution
 *
 * INTEGRATION: Now supports callback for ProcessOrchestrationService integration
 * to ensure task completions properly update process status and workflow state
 *
 * INTEGRATION FIX: Now integrated with TaskNotificationService for:
 * - New task assignment email notifications
 * - Task completion notifications
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import {
  IStepConfig,
  IActionContext,
  IActionResult
} from '../../../models/IWorkflow';
import { TaskStatus, Priority as PriorityEnum } from '../../../models/ICommon';
import { IJmlTaskAssignment } from '../../../models/IJmlTaskAssignment';
import { logger } from '../../LoggingService';
import { TaskNotificationService } from '../../TaskNotificationService';

/**
 * Task dependency validation result
 */
export interface ITaskDependencyValidation {
  canStart: boolean;
  blockedBy: number[];
  blockedByTitles: string[];
  reason?: string;
}

/**
 * Task completion cascade result
 */
export interface ITaskCompletionCascade {
  unblockedTaskIds: number[];
  unblockedCount: number;
  failedToUnblock: number[];
}

/**
 * Callback type for process orchestration integration
 * Called when a task is completed to update process and workflow state
 */
export type TaskCompletionCallback = (
  taskAssignmentId: number,
  completedByUserId: number,
  result?: Record<string, unknown>
) => Promise<{
  success: boolean;
  processUpdated: boolean;
  workflowResumed: boolean;
  processCompleted: boolean;
  error?: string;
}>;

export class TaskActionHandler {
  private sp: SPFI;

  // INTEGRATION FIX: Callback for ProcessOrchestrationService
  private processOrchestrationCallback?: TaskCompletionCallback;

  // INTEGRATION FIX: TaskNotificationService for email notifications
  private taskNotificationService: TaskNotificationService | null = null;

  constructor(sp: SPFI, context?: WebPartContext) {
    this.sp = sp;
    if (context) {
      this.taskNotificationService = new TaskNotificationService(sp, context);
    }
  }

  /**
   * Initialize notification service (can be called after construction)
   * INTEGRATION FIX: Enable email notifications for task assignments
   */
  public initializeNotificationService(context: WebPartContext): void {
    this.taskNotificationService = new TaskNotificationService(this.sp, context);
    logger.info('TaskActionHandler', 'Task notification service initialized');
  }

  /**
   * Register a callback for process orchestration integration
   * This callback is invoked when tasks are completed to update process/workflow state
   * INTEGRATION FIX: Connect TaskActionHandler to ProcessOrchestrationService
   */
  public onTaskCompletion(callback: TaskCompletionCallback): void {
    this.processOrchestrationCallback = callback;
    logger.info('TaskActionHandler', 'Process orchestration callback registered');
  }

  /**
   * Check if process orchestration callback is registered
   */
  public hasProcessOrchestrationCallback(): boolean {
    return !!this.processOrchestrationCallback;
  }

  /**
   * Create a single task
   */
  public async createTask(config: IStepConfig, context: IActionContext): Promise<IActionResult> {
    try {
      const title = config.taskTitle || `Task for ${context.currentStep.name}`;
      const assigneeId = config.assigneeId;

      if (!assigneeId) {
        return { success: false, error: 'Task assignee not specified' };
      }

      // Calculate due date
      let dueDate: Date | undefined;
      if (config.dueDaysFromNow) {
        dueDate = new Date();
        dueDate.setDate(dueDate.getDate() + config.dueDaysFromNow);
      }

      // Create task assignment
      const taskData = {
        Title: title,
        ProcessId: context.workflowInstance.ProcessId,
        AssignedToId: assigneeId,
        Status: TaskStatus.NotStarted,
        DueDate: dueDate?.toISOString(),
        Priority: 'Normal',
        WorkflowInstanceId: context.workflowInstance.Id,
        WorkflowStepId: context.currentStep.id,
        TaskTemplateId: config.taskTemplateId,
        Comments: `Auto-created by workflow step: ${context.currentStep.name}`
      };

      const result = await this.sp.web.lists.getByTitle('PM_TaskAssignments').items.add(taskData);

      logger.info('TaskActionHandler', `Created task: ${result.data.Id} - ${title}`);

      // INTEGRATION FIX: Send email notification to assignee
      if (this.taskNotificationService) {
        try {
          const taskForNotification: Partial<IJmlTaskAssignment> = {
            Id: result.data.Id,
            Title: title,
            AssignedToId: assigneeId as number,
            Status: TaskStatus.NotStarted,
            Priority: PriorityEnum.Medium,
            DueDate: dueDate,
            Notes: taskData.Comments
          };

          await this.taskNotificationService.sendTaskAssignmentNotification(
            taskForNotification as IJmlTaskAssignment,
            context.workflowInstance.Title || `Process ${context.workflowInstance.ProcessId}`
          );
          logger.info('TaskActionHandler', `Sent assignment notification for task ${result.data.Id}`);
        } catch (notifyError) {
          // Don't fail task creation if notification fails
          logger.warn('TaskActionHandler', 'Failed to send task assignment notification', notifyError);
        }
      }

      return {
        success: true,
        nextAction: 'wait',
        waitForItemType: 'task',
        waitForItemIds: [result.data.Id],
        createdItemIds: [result.data.Id],
        outputVariables: {
          createdTaskId: result.data.Id,
          createdTaskTitle: title
        }
      };
    } catch (error) {
      logger.error('TaskActionHandler', 'Error creating task', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to create task'
      };
    }
  }

  /**
   * Assign tasks from a template OR by role
   * If taskTemplateId is provided, uses template-based assignment
   * If assigneeRole is provided without template, creates a workflow-driven task
   */
  public async assignTasksFromTemplate(config: IStepConfig, context: IActionContext): Promise<IActionResult> {
    try {
      const templateId = config.taskTemplateId;

      // If no template, try role-based task creation
      if (!templateId) {
        return await this.assignTasksByRole(config, context);
      }

      // Get template tasks
      const templateTasks = await this.sp.web.lists.getByTitle('PM_TemplateTaskMappings').items
        .filter(`ChecklistTemplateId eq ${templateId}`)
        .select('Id', 'TaskId', 'Order', 'DefaultDaysOffset', 'AssigneeRole', 'Category')
        .orderBy('Order', true)();

      if (templateTasks.length === 0) {
        // FIXED: Empty template should be a failure unless explicitly allowed
        const allowEmptyTemplate = config.allowEmptyTemplate === true;
        if (!allowEmptyTemplate) {
          logger.error('TaskActionHandler', `Template ${templateId} has no tasks configured`);
          return {
            success: false,
            error: `Task template ${templateId} has no tasks configured. Please configure tasks in the template or set allowEmptyTemplate: true if this is intentional.`
          };
        }
        logger.warn('TaskActionHandler', `Template ${templateId} has no tasks but allowEmptyTemplate is true`);
        return {
          success: true,
          nextAction: 'continue',
          outputVariables: { tasksCreated: 0, templateEmpty: true }
        };
      }

      const createdTaskIds: number[] = [];
      const processStartDate = new Date(context.workflowInstance.StartedDate || new Date());

      for (const templateTask of templateTasks) {
        // Get task definition
        const taskDef = await this.sp.web.lists.getByTitle('PM_Tasks').items
          .getById(templateTask.TaskId)
          .select('Id', 'Title', 'Description', 'Category', 'EstimatedDuration', 'Priority')();

        // Determine assignee
        let assigneeId = config.assigneeId;
        if (!assigneeId && templateTask.AssigneeRole) {
          assigneeId = this.resolveRoleToUserId(templateTask.AssigneeRole, context);
        }

        // Calculate due date
        const dueDate = new Date(processStartDate);
        dueDate.setDate(dueDate.getDate() + (templateTask.DefaultDaysOffset || 7));

        // Create task assignment
        const assignmentData = {
          Title: taskDef.Title,
          ProcessId: context.workflowInstance.ProcessId,
          TaskId: taskDef.Id,
          AssignedToId: assigneeId,
          Status: TaskStatus.NotStarted,
          DueDate: dueDate.toISOString(),
          Priority: taskDef.Priority || 'Normal',
          Category: taskDef.Category,
          WorkflowInstanceId: context.workflowInstance.Id,
          WorkflowStepId: context.currentStep.id,
          Order: templateTask.Order
        };

        const result = await this.sp.web.lists.getByTitle('PM_TaskAssignments').items.add(assignmentData);
        createdTaskIds.push(result.data.Id);
      }

      logger.info('TaskActionHandler', `Assigned ${createdTaskIds.length} tasks from template ${templateId}`);

      return {
        success: true,
        nextAction: 'wait',
        waitForItemType: 'task',
        waitForItemIds: createdTaskIds,
        createdItemIds: createdTaskIds,
        outputVariables: {
          tasksCreated: createdTaskIds.length,
          createdTaskIds
        }
      };
    } catch (error) {
      logger.error('TaskActionHandler', 'Error assigning tasks from template', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to assign tasks'
      };
    }
  }

  /**
   * Assign tasks based on role (no template required)
   * Used when workflow steps define tasks by assigneeRole
   */
  private async assignTasksByRole(config: IStepConfig, context: IActionContext): Promise<IActionResult> {
    try {
      const assigneeRole = config.assigneeRole;
      const stepName = context.currentStep.name;
      const stepDescription = context.currentStep.description;

      if (!assigneeRole) {
        logger.warn('TaskActionHandler', 'No assigneeRole or templateId specified, skipping task creation');
        return {
          success: true,
          nextAction: 'continue',
          outputVariables: { tasksCreated: 0, skipped: true }
        };
      }

      // Resolve role to user ID
      const assigneeId = await this.resolveRoleToUserIdAsync(assigneeRole, context);

      if (!assigneeId) {
        // FIXED: Unresolved role should be a failure - this is a configuration issue
        const allowUnresolvedRole = config.allowUnresolvedRole === true;
        if (!allowUnresolvedRole) {
          logger.error('TaskActionHandler', `Could not resolve role "${assigneeRole}" to a user - check PM_RoleAssignments`);
          return {
            success: false,
            error: `Could not find a user assigned to role "${assigneeRole}". Please configure this role in the PM_RoleAssignments list or ensure the process has the required role field populated.`
          };
        }
        logger.warn('TaskActionHandler', `Role "${assigneeRole}" not resolved but allowUnresolvedRole is true - skipping task creation`);
        return {
          success: true,
          nextAction: 'continue',
          outputVariables: { tasksCreated: 0, roleNotResolved: assigneeRole, skippedDueToConfig: true }
        };
      }

      // Calculate due date
      let dueDate: Date = new Date();
      if (config.dueDaysFromNow) {
        dueDate.setDate(dueDate.getDate() + config.dueDaysFromNow);
      } else {
        dueDate.setDate(dueDate.getDate() + 7); // Default 7 days
      }

      // Build task title from step name
      const employeeName = context.process.employeeName || context.variables.employeeName || 'Employee';
      const taskTitle = `${stepName} - ${employeeName}`;

      // Create task assignment
      const taskData = {
        Title: taskTitle,
        Description: stepDescription || `Workflow task: ${stepName}`,
        ProcessId: context.workflowInstance.ProcessId,
        AssignedToId: assigneeId,
        Status: TaskStatus.NotStarted,
        DueDate: dueDate.toISOString(),
        Priority: 'Normal',
        Category: assigneeRole,
        WorkflowInstanceId: context.workflowInstance.Id,
        WorkflowStepId: context.currentStep.id,
        Comments: `Auto-created by workflow step: ${stepName}`
      };

      const result = await this.sp.web.lists.getByTitle('PM_TaskAssignments').items.add(taskData);

      logger.info('TaskActionHandler', `Created role-based task: ${result.data.Id} - ${taskTitle} for ${assigneeRole}`);

      return {
        success: true,
        nextAction: 'continue', // Don't wait by default, use WaitForTasks step if needed
        createdItemIds: [result.data.Id],
        outputVariables: {
          tasksCreated: 1,
          createdTaskId: result.data.Id,
          createdTaskTitle: taskTitle,
          assigneeRole
        }
      };
    } catch (error) {
      logger.error('TaskActionHandler', 'Error assigning tasks by role', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to assign tasks by role'
      };
    }
  }

  /**
   * Wait for tasks to complete
   */
  public async waitForTasks(config: IStepConfig, context: IActionContext): Promise<IActionResult> {
    try {
      const waitCondition = config.waitCondition || 'all';

      // Get tasks created by this workflow instance
      const tasks = await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
        .filter(`WorkflowInstanceId eq ${context.workflowInstance.Id}`)
        .select('Id', 'Status', 'WorkflowStepId')();

      // Filter to relevant steps if specified
      let relevantTasks = tasks;
      if (config.waitForTaskIds && config.waitForTaskIds.length > 0) {
        relevantTasks = tasks.filter(t =>
          config.waitForTaskIds!.includes(t.WorkflowStepId)
        );
      }

      const completedTasks = relevantTasks.filter(t =>
        t.Status === TaskStatus.Completed || t.Status === TaskStatus.Skipped
      );

      const allComplete = completedTasks.length === relevantTasks.length;
      const anyComplete = completedTasks.length > 0;

      if ((waitCondition === 'all' && allComplete) ||
          (waitCondition === 'any' && anyComplete)) {
        return {
          success: true,
          nextAction: 'continue',
          outputVariables: {
            completedTaskCount: completedTasks.length,
            totalTaskCount: relevantTasks.length
          }
        };
      }

      // Still waiting
      return {
        success: true,
        nextAction: 'wait',
        waitForItemType: 'task',
        waitForItemIds: relevantTasks.filter(t => t.Status !== TaskStatus.Completed).map(t => t.Id)
      };
    } catch (error) {
      logger.error('TaskActionHandler', 'Error waiting for tasks', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to check task status'
      };
    }
  }

  /**
   * Update task status
   */
  public async updateTaskStatus(
    config: Record<string, unknown>,
    context: IActionContext
  ): Promise<IActionResult> {
    try {
      const taskId = config.taskId as number;
      const newStatus = config.status as string;

      if (!taskId) {
        return { success: false, error: 'Task ID not specified' };
      }

      await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
        .getById(taskId)
        .update({
          Status: newStatus,
          Modified: new Date().toISOString()
        });

      logger.info('TaskActionHandler', `Updated task ${taskId} status to ${newStatus}`);

      return {
        success: true,
        nextAction: 'continue',
        outputVariables: {
          updatedTaskId: taskId,
          newStatus
        }
      };
    } catch (error) {
      logger.error('TaskActionHandler', 'Error updating task status', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to update task status'
      };
    }
  }

  /**
   * Resolve role name to user ID (sync version for template-based)
   */
  private resolveRoleToUserId(role: string, context: IActionContext): number | undefined {
    // Map common roles to context fields
    const normalizedRole = role.toLowerCase().trim();
    switch (normalizedRole) {
      case 'manager':
        return context.process.managerId as number;
      case 'hr':
      case 'hr admin':
        return context.process.hrAdminId as number;
      case 'it':
      case 'it admin':
        return context.process.itAdminId as number;
      case 'processowner':
      case 'process owner':
        return context.process.processOwnerId as number;
      case 'facilities':
        return context.process.facilitiesId as number;
      default:
        return undefined;
    }
  }

  /**
   * Resolve role name to user ID (async version with SharePoint lookup)
   * First checks context, then falls back to PM_RoleAssignments list
   * Supports department-specific assignments with priority ordering
   */
  private async resolveRoleToUserIdAsync(role: string, context: IActionContext): Promise<number | undefined> {
    // First try sync resolution from context
    const contextUserId = this.resolveRoleToUserId(role, context);
    if (contextUserId) {
      return contextUserId;
    }

    // Fall back to configuration list lookup
    try {
      const department = context.process.department as string;

      // Look up role assignment from configuration
      // Priority: Department-specific (lower Priority) > Global (higher Priority)
      let filter = `Role eq '${role}' and IsActive eq 1`;
      if (department) {
        filter = `Role eq '${role}' and IsActive eq 1 and (Department eq '${department}' or Department eq null or Department eq '')`;
      }

      const roleAssignments = await this.sp.web.lists.getByTitle('PM_RoleAssignments').items
        .filter(filter)
        .select('Id', 'Role', 'AssignedUserId', 'AssignedUserEmail', 'AssignedUser/Id', 'Department', 'IsActive', 'Priority')
        .expand('AssignedUser')
        .orderBy('Priority', true) // Lower priority number = higher precedence
        .top(5)(); // Get top 5 to find best match

      if (roleAssignments.length > 0) {
        // Prefer department-specific match
        const deptMatch = department
          ? roleAssignments.find(r => r.Department?.toLowerCase() === department.toLowerCase())
          : undefined;

        const bestMatch = deptMatch || roleAssignments[0];

        // Return AssignedUserId if set, otherwise try to get from AssignedUser lookup
        if (bestMatch.AssignedUserId) {
          return bestMatch.AssignedUserId;
        }
        if (bestMatch.AssignedUser?.Id) {
          return bestMatch.AssignedUser.Id;
        }

        // Log warning if email is set but no user ID
        if (bestMatch.AssignedUserEmail) {
          logger.warn('TaskActionHandler', `Role "${role}" has email (${bestMatch.AssignedUserEmail}) but no AssignedUserId. Please configure the list properly.`);
        }
      }

      // Try alternative role names
      const roleAliases: Record<string, string[]> = {
        'hr admin': ['hr', 'human resources', 'hr administrator'],
        'it admin': ['it', 'information technology', 'it administrator', 'it support'],
        'facilities': ['facility', 'facilities manager', 'office manager'],
        'payroll': ['payroll admin', 'payroll administrator'],
        'finance admin': ['finance', 'finance administrator'],
        'compliance': ['compliance officer'],
        'it security': ['security', 'security admin']
      };

      const normalizedRole = role.toLowerCase().trim();
      for (const [key, aliases] of Object.entries(roleAliases)) {
        if (normalizedRole === key || aliases.includes(normalizedRole)) {
          // Try to find by primary role name
          const altAssignments = await this.sp.web.lists.getByTitle('PM_RoleAssignments').items
            .filter(`Role eq '${key}' and IsActive eq 1`)
            .select('AssignedUserId', 'AssignedUser/Id')
            .expand('AssignedUser')
            .orderBy('Priority', true)
            .top(1)();

          if (altAssignments.length > 0) {
            return altAssignments[0].AssignedUserId || altAssignments[0].AssignedUser?.Id;
          }
        }
      }

      logger.warn('TaskActionHandler', `No role assignment found for: ${role}`);
      return undefined;
    } catch (error) {
      // List might not exist - just log and return undefined
      logger.warn('TaskActionHandler', `Could not look up role ${role}`, error);
      return undefined;
    }
  }

  // ============================================================================
  // TASK DEPENDENCY VALIDATION
  // ============================================================================

  /**
   * Validate if a task can be started based on its dependencies
   * Returns detailed information about blocking tasks
   */
  public async validateTaskDependencies(taskId: number): Promise<ITaskDependencyValidation> {
    try {
      // Get the task with dependency info
      const task = await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
        .getById(taskId)
        .select('Id', 'Title', 'Status', 'IsDependentTask', 'DependsOnTaskId', 'IsBlocked', 'ProcessId')() as {
          Id: number;
          Title: string;
          Status: string;
          IsDependentTask: boolean;
          DependsOnTaskId?: number;
          IsBlocked: boolean;
          ProcessId: string;
        };

      // If task is not dependent, it can start
      if (!task.IsDependentTask || !task.DependsOnTaskId) {
        return {
          canStart: true,
          blockedBy: [],
          blockedByTitles: []
        };
      }

      // Get the blocking task(s)
      const blockingTask = await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
        .getById(task.DependsOnTaskId)
        .select('Id', 'Title', 'Status')() as { Id: number; Title: string; Status: string };

      // Check if blocking task is completed
      const isBlockingTaskComplete =
        blockingTask.Status === TaskStatus.Completed ||
        blockingTask.Status === TaskStatus.Skipped;

      if (isBlockingTaskComplete) {
        return {
          canStart: true,
          blockedBy: [],
          blockedByTitles: []
        };
      }

      // Task is blocked
      return {
        canStart: false,
        blockedBy: [blockingTask.Id],
        blockedByTitles: [blockingTask.Title],
        reason: `Waiting for "${blockingTask.Title}" (ID: ${blockingTask.Id}) to complete`
      };
    } catch (error) {
      logger.error('TaskActionHandler', `Error validating dependencies for task ${taskId}`, error);
      // On error, allow task to proceed but log warning
      return {
        canStart: true,
        blockedBy: [],
        blockedByTitles: [],
        reason: 'Could not validate dependencies - allowing task to proceed'
      };
    }
  }

  /**
   * Validate all dependencies for tasks in a process
   * Returns a map of taskId -> validation result
   */
  public async validateAllProcessDependencies(
    processId: number
  ): Promise<Map<number, ITaskDependencyValidation>> {
    const results = new Map<number, ITaskDependencyValidation>();

    try {
      // Get all tasks for the process with dependency info
      const tasks = await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
        .filter(`ProcessId eq '${processId}' and IsDeleted ne true`)
        .select('Id', 'Title', 'Status', 'IsDependentTask', 'DependsOnTaskId', 'IsBlocked')() as Array<{
          Id: number;
          Title: string;
          Status: string;
          IsDependentTask: boolean;
          DependsOnTaskId?: number;
          IsBlocked: boolean;
        }>;

      // Build task map for efficient lookup
      const taskMap = new Map<number, typeof tasks[0]>();
      for (const task of tasks) {
        taskMap.set(task.Id, task);
      }

      // Validate each task
      for (const task of tasks) {
        if (!task.IsDependentTask || !task.DependsOnTaskId) {
          results.set(task.Id, {
            canStart: true,
            blockedBy: [],
            blockedByTitles: []
          });
          continue;
        }

        const blockingTask = taskMap.get(task.DependsOnTaskId);
        if (!blockingTask) {
          // Blocking task not found in this process - allow
          results.set(task.Id, {
            canStart: true,
            blockedBy: [],
            blockedByTitles: [],
            reason: 'Dependency task not found - allowing task to proceed'
          });
          continue;
        }

        const isComplete =
          blockingTask.Status === TaskStatus.Completed ||
          blockingTask.Status === TaskStatus.Skipped;

        results.set(task.Id, {
          canStart: isComplete,
          blockedBy: isComplete ? [] : [blockingTask.Id],
          blockedByTitles: isComplete ? [] : [blockingTask.Title],
          reason: isComplete ? undefined : `Waiting for "${blockingTask.Title}" to complete`
        });
      }
    } catch (error) {
      logger.error('TaskActionHandler', `Error validating process ${processId} dependencies`, error);
    }

    return results;
  }

  /**
   * Called when a task is completed to unblock dependent tasks
   * Cascades through the dependency chain
   */
  public async onTaskCompleted(completedTaskId: number): Promise<ITaskCompletionCascade> {
    const unblockedTaskIds: number[] = [];
    const failedToUnblock: number[] = [];

    try {
      // Get the completed task to know its process
      const completedTask = await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
        .getById(completedTaskId)
        .select('Id', 'Title', 'ProcessId')() as { Id: number; Title: string; ProcessId: string };

      // Find all tasks that depend on this completed task
      const dependentTasks = await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
        .filter(`DependsOnTaskId eq ${completedTaskId} and IsDeleted ne true`)
        .select('Id', 'Title', 'Status', 'IsBlocked')() as Array<{
          Id: number;
          Title: string;
          Status: string;
          IsBlocked: boolean;
        }>;

      if (dependentTasks.length === 0) {
        logger.info('TaskActionHandler', `No dependent tasks found for completed task ${completedTaskId}`);
        return { unblockedTaskIds: [], unblockedCount: 0, failedToUnblock: [] };
      }

      // Unblock each dependent task
      for (const dependentTask of dependentTasks) {
        // Only unblock if task is actually blocked and not already completed
        if (dependentTask.IsBlocked &&
            dependentTask.Status !== TaskStatus.Completed &&
            dependentTask.Status !== TaskStatus.Skipped) {
          try {
            await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
              .getById(dependentTask.Id)
              .update({
                IsBlocked: false,
                BlockedReason: null,
                // If task was NotStarted and is now unblocked, keep it NotStarted
                // The assignee will need to start it
              });

            unblockedTaskIds.push(dependentTask.Id);
            logger.info('TaskActionHandler', `Unblocked task ${dependentTask.Id} "${dependentTask.Title}" after ${completedTaskId} completed`);

            // Send notification to assignee that task is now available
            await this.notifyTaskUnblocked(dependentTask.Id, completedTask.Title);
          } catch (updateError) {
            logger.error('TaskActionHandler', `Failed to unblock task ${dependentTask.Id}`, updateError);
            failedToUnblock.push(dependentTask.Id);
          }
        }
      }

      logger.info('TaskActionHandler', `Task ${completedTaskId} completion cascade: unblocked ${unblockedTaskIds.length} tasks`);

      return {
        unblockedTaskIds,
        unblockedCount: unblockedTaskIds.length,
        failedToUnblock
      };
    } catch (error) {
      logger.error('TaskActionHandler', `Error in task completion cascade for ${completedTaskId}`, error);
      return { unblockedTaskIds: [], unblockedCount: 0, failedToUnblock: [] };
    }
  }

  /**
   * Notify task assignee that their blocked task is now available
   */
  private async notifyTaskUnblocked(taskId: number, completedTaskTitle: string): Promise<void> {
    try {
      // Get task details including assignee
      const task = await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
        .getById(taskId)
        .select('Id', 'Title', 'AssignedToId', 'ProcessId')() as {
          Id: number;
          Title: string;
          AssignedToId: number;
          ProcessId: string;
        };

      if (!task.AssignedToId) {
        return;
      }

      // Create notification
      await this.sp.web.lists.getByTitle('PM_Notifications').items.add({
        Title: 'Task Now Available',
        Message: `Your task "${task.Title}" is now available. The prerequisite task "${completedTaskTitle}" has been completed.`,
        RecipientId: task.AssignedToId,
        NotificationType: 'TaskUnblocked',
        Priority: PriorityEnum.Medium,
        LinkUrl: `/sites/JML/SitePages/MyTasks.aspx?taskId=${taskId}`,
        ProcessId: task.ProcessId,
        IsRead: false,
        SentDate: new Date()
      });
    } catch (error) {
      // Don't fail the cascade for notification errors
      logger.warn('TaskActionHandler', `Failed to send task unblocked notification for ${taskId}`, error);
    }
  }

  /**
   * Attempt to start a task - validates dependencies first
   */
  public async attemptStartTask(taskId: number): Promise<IActionResult> {
    try {
      // First validate dependencies
      const validation = await this.validateTaskDependencies(taskId);

      if (!validation.canStart) {
        return {
          success: false,
          error: validation.reason || `Task is blocked by ${validation.blockedByTitles.join(', ')}`
        };
      }

      // Dependencies met - update task status to InProgress
      await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
        .getById(taskId)
        .update({
          Status: TaskStatus.InProgress,
          StartDate: new Date().toISOString(),
          IsBlocked: false,
          BlockedReason: null
        });

      logger.info('TaskActionHandler', `Task ${taskId} started successfully`);

      return {
        success: true,
        nextAction: 'continue',
        outputVariables: {
          taskId,
          startedAt: new Date().toISOString()
        }
      };
    } catch (error) {
      logger.error('TaskActionHandler', `Error starting task ${taskId}`, error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to start task'
      };
    }
  }

  /**
   * Complete a task and trigger dependency cascade
   * INTEGRATION FIX: Now calls ProcessOrchestrationService callback if registered
   * INTEGRATION FIX: Now sends completion notification via TaskNotificationService
   */
  public async completeTask(
    taskId: number,
    completionData?: { comments?: string; actualHours?: number; completedByUserId?: number }
  ): Promise<IActionResult> {
    try {
      // INTEGRATION FIX: Fetch task details for notification before completing
      let taskDetails: Partial<IJmlTaskAssignment> | undefined;
      let processTitle: string | undefined;

      if (this.taskNotificationService) {
        try {
          const fetchedTask = await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
            .getById(taskId)
            .select('Id', 'Title', 'AssignedToId', 'ProcessId', 'Priority', 'DueDate')();

          taskDetails = {
            Id: fetchedTask.Id,
            Title: fetchedTask.Title,
            AssignedToId: fetchedTask.AssignedToId,
            Priority: fetchedTask.Priority,
            DueDate: fetchedTask.DueDate ? new Date(fetchedTask.DueDate) : undefined
          };

          // Try to get process title for notification context
          if (fetchedTask.ProcessId) {
            try {
              const process = await this.sp.web.lists.getByTitle('PM_Processes').items
                .getById(parseInt(fetchedTask.ProcessId, 10))
                .select('Title')();
              processTitle = process.Title;
            } catch {
              processTitle = `Process ${fetchedTask.ProcessId}`;
            }
          }
        } catch (fetchError) {
          logger.warn('TaskActionHandler', `Could not fetch task details for notification: ${taskId}`, fetchError);
        }
      }

      // Update task to completed
      await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
        .getById(taskId)
        .update({
          Status: TaskStatus.Completed,
          CompletedDate: new Date().toISOString(),
          PercentComplete: 100,
          ...(completionData?.comments ? { Comments: completionData.comments } : {}),
          ...(completionData?.actualHours ? { ActualHours: completionData.actualHours } : {}),
          ...(completionData?.completedByUserId ? { CompletedById: completionData.completedByUserId } : {})
        });

      logger.info('TaskActionHandler', `Task ${taskId} marked as completed`);

      // INTEGRATION FIX: Send completion notification
      if (this.taskNotificationService && taskDetails) {
        try {
          // Mark the task as completed for the notification
          taskDetails.Status = TaskStatus.Completed;

          await this.taskNotificationService.sendTaskCompletionNotification(
            taskDetails as IJmlTaskAssignment,
            completionData?.completedByUserId || 0,
            false, // Don't notify assignee (they completed it)
            undefined, // No additional recipients for now
            undefined // Use default delivery options
          );
          logger.info('TaskActionHandler', `Sent completion notification for task ${taskId}`);
        } catch (notifyError) {
          // Don't fail task completion if notification fails
          logger.warn('TaskActionHandler', 'Failed to send task completion notification', notifyError);
        }
      }

      // Trigger dependency cascade (unblock dependent tasks)
      const cascadeResult = await this.onTaskCompleted(taskId);

      // INTEGRATION FIX: Notify ProcessOrchestrationService of task completion
      // This ensures process status is updated and workflow step is resumed
      let orchestrationResult: {
        success: boolean;
        processUpdated: boolean;
        workflowResumed: boolean;
        processCompleted: boolean;
        error?: string;
      } | undefined;

      if (this.processOrchestrationCallback) {
        try {
          orchestrationResult = await this.processOrchestrationCallback(
            taskId,
            completionData?.completedByUserId || 0,
            {
              comments: completionData?.comments,
              actualHours: completionData?.actualHours,
              unblockedTasks: cascadeResult.unblockedTaskIds,
              completedAt: new Date().toISOString()
            }
          );

          if (orchestrationResult.success) {
            logger.info(
              'TaskActionHandler',
              `Process orchestration updated for task ${taskId}: ` +
              `processUpdated=${orchestrationResult.processUpdated}, ` +
              `workflowResumed=${orchestrationResult.workflowResumed}, ` +
              `processCompleted=${orchestrationResult.processCompleted}`
            );
          } else {
            logger.warn(
              'TaskActionHandler',
              `Process orchestration failed for task ${taskId}: ${orchestrationResult.error}`
            );
          }
        } catch (orchError) {
          // Don't fail task completion if orchestration fails - log and continue
          logger.error('TaskActionHandler', `Error in process orchestration for task ${taskId}`, orchError);
        }
      } else {
        logger.debug('TaskActionHandler', `No process orchestration callback registered for task ${taskId}`);
      }

      return {
        success: true,
        nextAction: 'continue',
        outputVariables: {
          taskId,
          completedAt: new Date().toISOString(),
          unblockedTasks: cascadeResult.unblockedTaskIds,
          unblockedCount: cascadeResult.unblockedCount,
          // Include orchestration results if available
          ...(orchestrationResult ? {
            processUpdated: orchestrationResult.processUpdated,
            workflowResumed: orchestrationResult.workflowResumed,
            processCompleted: orchestrationResult.processCompleted
          } : {})
        }
      };
    } catch (error) {
      logger.error('TaskActionHandler', `Error completing task ${taskId}`, error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to complete task'
      };
    }
  }

  /**
   * Recalculate blocked status for all tasks in a process
   * Useful for fixing sync issues or after bulk operations
   */
  public async recalculateProcessBlockedStatus(processId: number): Promise<{
    processed: number;
    blocked: number;
    unblocked: number;
    errors: number;
  }> {
    let processed = 0;
    let blocked = 0;
    let unblocked = 0;
    let errors = 0;

    try {
      // Get all tasks for the process
      const tasks = await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
        .filter(`ProcessId eq '${processId}' and IsDeleted ne true`)
        .select('Id', 'Title', 'Status', 'IsDependentTask', 'DependsOnTaskId', 'IsBlocked')() as Array<{
          Id: number;
          Title: string;
          Status: string;
          IsDependentTask: boolean;
          DependsOnTaskId?: number;
          IsBlocked: boolean;
        }>;

      // Build task map
      const taskMap = new Map<number, typeof tasks[0]>();
      for (const task of tasks) {
        taskMap.set(task.Id, task);
      }

      // Process each task
      for (const task of tasks) {
        processed++;

        // Skip completed tasks
        if (task.Status === TaskStatus.Completed || task.Status === TaskStatus.Skipped) {
          continue;
        }

        // If no dependency, should not be blocked
        if (!task.IsDependentTask || !task.DependsOnTaskId) {
          if (task.IsBlocked) {
            try {
              await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
                .getById(task.Id)
                .update({ IsBlocked: false, BlockedReason: null });
              unblocked++;
            } catch {
              errors++;
            }
          }
          continue;
        }

        // Has dependency - check if blocking task is complete
        const blockingTask = taskMap.get(task.DependsOnTaskId);
        const shouldBeBlocked = blockingTask &&
          blockingTask.Status !== TaskStatus.Completed &&
          blockingTask.Status !== TaskStatus.Skipped;

        if (shouldBeBlocked && !task.IsBlocked) {
          // Should be blocked but isn't
          try {
            await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
              .getById(task.Id)
              .update({
                IsBlocked: true,
                BlockedReason: `Waiting for "${blockingTask!.Title}" to complete`
              });
            blocked++;
          } catch {
            errors++;
          }
        } else if (!shouldBeBlocked && task.IsBlocked) {
          // Shouldn't be blocked but is
          try {
            await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
              .getById(task.Id)
              .update({ IsBlocked: false, BlockedReason: null });
            unblocked++;
          } catch {
            errors++;
          }
        }
      }

      logger.info('TaskActionHandler',
        `Recalculated blocked status for process ${processId}: processed=${processed}, blocked=${blocked}, unblocked=${unblocked}, errors=${errors}`);
    } catch (error) {
      logger.error('TaskActionHandler', `Error recalculating blocked status for process ${processId}`, error);
    }

    return { processed, blocked, unblocked, errors };
  }
}
