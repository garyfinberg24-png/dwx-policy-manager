// @ts-nocheck
/**
 * TaskAssignmentService
 * Central service for task assignment operations
 * Handles task completion with workflow integration
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { IJmlTaskAssignment, IUser } from '../models';
import { TaskStatus, Priority } from '../models/ICommon';
import { ApprovalStatus } from '../models/IJmlApproval';
import { logger } from './LoggingService';

/**
 * Task completion result
 */
export interface ITaskCompletionResult {
  success: boolean;
  taskId: number;
  workflowResumed?: boolean;
  workflowInstanceId?: number;
  error?: string;
}

/**
 * Concurrency error types
 */
export enum ConcurrencyErrorType {
  None = 'None',
  VersionConflict = 'VersionConflict',
  ItemDeleted = 'ItemDeleted',
  Unknown = 'Unknown'
}

/**
 * Result of a concurrent update operation
 */
export interface IConcurrentUpdateResult {
  success: boolean;
  taskId: number;
  error?: string;
  errorType?: ConcurrencyErrorType;
  serverVersion?: string;
  localVersion?: string;
  conflictData?: {
    serverModified: Date;
    serverModifiedBy?: string;
  };
}

/**
 * Task update data
 */
export interface ITaskUpdateData {
  Status?: TaskStatus;
  Priority?: Priority;
  PercentComplete?: number;
  Notes?: string;
  CompletionNotes?: string;
  ActualHours?: number;
  DueDate?: Date;
  AssignedToId?: number;
}

export class TaskAssignmentService {
  private sp: SPFI;
  private context?: WebPartContext;
  private tasksListTitle = 'PM_TaskAssignments';

  constructor(sp: SPFI, context?: WebPartContext) {
    this.sp = sp;
    this.context = context;
  }

  /**
   * Get task by ID
   */
  public async getById(taskId: number): Promise<IJmlTaskAssignment | null> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.tasksListTitle)
        .items.getById(taskId)
        .select(
          'Id', 'Title', 'Status', 'Priority', 'DueDate', 'AssignedToId',
          'AssignedDate', 'ProcessIDId', 'TaskIDId', 'Modified', 'Created',
          'PercentComplete', 'Notes', 'CompletionNotes', 'ActualHours',
          'WorkflowInstanceId', 'WorkflowStepId', 'RequiresApproval', 'ApprovalStatus',
          'IsBlocked', 'BlockedReason', 'ReminderSent', 'EscalationSent',
          'AssignedTo/Id', 'AssignedTo/Title', 'AssignedTo/EMail'
        )
        .expand('AssignedTo')();

      return this.mapToTaskAssignment(item);
    } catch (error) {
      logger.error('TaskAssignmentService', `Error getting task ${taskId}`, error);
      return null;
    }
  }

  /**
   * Get all tasks for a process
   */
  public async getByProcessId(processId: number): Promise<IJmlTaskAssignment[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.tasksListTitle)
        .items.filter(`ProcessIDId eq ${processId}`)
        .select(
          'Id', 'Title', 'Status', 'Priority', 'DueDate', 'AssignedToId',
          'ProcessIDId', 'TaskIDId', 'PercentComplete', 'WorkflowInstanceId',
          'AssignedTo/Id', 'AssignedTo/Title', 'AssignedTo/EMail'
        )
        .expand('AssignedTo')
        .orderBy('Created', true)();

      return items.map(item => this.mapToTaskAssignment(item));
    } catch (error) {
      logger.error('TaskAssignmentService', `Error getting tasks for process ${processId}`, error);
      return [];
    }
  }

  /**
   * Get tasks for current user
   */
  public async getMyTasks(userId: number): Promise<IJmlTaskAssignment[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.tasksListTitle)
        .items.filter(`AssignedToId eq ${userId} and Status ne '${TaskStatus.Completed}' and Status ne '${TaskStatus.Cancelled}'`)
        .select(
          'Id', 'Title', 'Status', 'Priority', 'DueDate', 'AssignedToId',
          'ProcessIDId', 'TaskIDId', 'PercentComplete', 'WorkflowInstanceId',
          'AssignedTo/Id', 'AssignedTo/Title', 'AssignedTo/EMail'
        )
        .expand('AssignedTo')
        .orderBy('DueDate', true)();

      return items.map(item => this.mapToTaskAssignment(item));
    } catch (error) {
      logger.error('TaskAssignmentService', `Error getting tasks for user ${userId}`, error);
      return [];
    }
  }

  /**
   * Update task
   */
  public async updateTask(taskId: number, updates: ITaskUpdateData): Promise<boolean> {
    try {
      const updateData: Record<string, unknown> = {
        Modified: new Date().toISOString()
      };

      if (updates.Status !== undefined) updateData.Status = updates.Status;
      if (updates.Priority !== undefined) updateData.Priority = updates.Priority;
      if (updates.PercentComplete !== undefined) updateData.PercentComplete = updates.PercentComplete;
      if (updates.Notes !== undefined) updateData.Notes = updates.Notes;
      if (updates.CompletionNotes !== undefined) updateData.CompletionNotes = updates.CompletionNotes;
      if (updates.ActualHours !== undefined) updateData.ActualHours = updates.ActualHours;
      if (updates.DueDate !== undefined) updateData.DueDate = updates.DueDate.toISOString();
      if (updates.AssignedToId !== undefined) updateData.AssignedToId = updates.AssignedToId;

      await this.sp.web.lists
        .getByTitle(this.tasksListTitle)
        .items.getById(taskId)
        .update(updateData);

      logger.info('TaskAssignmentService', `Updated task ${taskId}`);
      return true;
    } catch (error) {
      logger.error('TaskAssignmentService', `Error updating task ${taskId}`, error);
      return false;
    }
  }

  /**
   * Update task with optimistic concurrency control
   * Uses SharePoint's ETag mechanism to detect conflicts
   * @param taskId Task ID to update
   * @param updates Update data
   * @param expectedVersion Optional: The odata.etag value from when item was last read
   * @returns Concurrency result with conflict details if update fails
   */
  public async updateTaskWithConcurrency(
    taskId: number,
    updates: ITaskUpdateData,
    expectedVersion?: string
  ): Promise<IConcurrentUpdateResult> {
    try {
      // First get current item with its ETag
      const item = this.sp.web.lists
        .getByTitle(this.tasksListTitle)
        .items.getById(taskId);

      // Get current server state to compare versions
      const currentItem = await item.select('Id', 'Modified', 'Editor/Title', 'odata.etag')
        .expand('Editor')();

      const serverEtag = (currentItem as any)['odata.etag'] || currentItem['@odata.etag'];

      // If caller provided expected version, check for conflict
      if (expectedVersion && serverEtag && expectedVersion !== serverEtag) {
        logger.warn('TaskAssignmentService',
          `Concurrency conflict on task ${taskId}: expected ${expectedVersion}, server has ${serverEtag}`);

        return {
          success: false,
          taskId,
          error: 'The task was modified by another user. Please refresh and try again.',
          errorType: ConcurrencyErrorType.VersionConflict,
          serverVersion: serverEtag,
          localVersion: expectedVersion,
          conflictData: {
            serverModified: new Date(currentItem.Modified),
            serverModifiedBy: (currentItem as any).Editor?.Title
          }
        };
      }

      // Build update data
      const updateData: Record<string, unknown> = {};
      if (updates.Status !== undefined) updateData.Status = updates.Status;
      if (updates.Priority !== undefined) updateData.Priority = updates.Priority;
      if (updates.PercentComplete !== undefined) updateData.PercentComplete = updates.PercentComplete;
      if (updates.Notes !== undefined) updateData.Notes = updates.Notes;
      if (updates.CompletionNotes !== undefined) updateData.CompletionNotes = updates.CompletionNotes;
      if (updates.ActualHours !== undefined) updateData.ActualHours = updates.ActualHours;
      if (updates.DueDate !== undefined) updateData.DueDate = updates.DueDate.toISOString();
      if (updates.AssignedToId !== undefined) updateData.AssignedToId = updates.AssignedToId;

      // Update with ETag header for SharePoint concurrency
      const headers: HeadersInit = {};
      if (serverEtag) {
        headers['If-Match'] = serverEtag;
      }

      await item.update(updateData, serverEtag);

      logger.info('TaskAssignmentService', `Updated task ${taskId} with concurrency check`);

      return {
        success: true,
        taskId,
        errorType: ConcurrencyErrorType.None,
        serverVersion: serverEtag
      };
    } catch (error: unknown) {
      // Check for 412 Precondition Failed (concurrency conflict)
      const errorObj = error as { status?: number; message?: string };
      if (errorObj.status === 412) {
        logger.warn('TaskAssignmentService',
          `Concurrency conflict detected on task ${taskId} (412 response)`);

        return {
          success: false,
          taskId,
          error: 'The task was modified by another user. Please refresh and try again.',
          errorType: ConcurrencyErrorType.VersionConflict
        };
      }

      // Check for 404 Not Found (item deleted)
      if (errorObj.status === 404) {
        return {
          success: false,
          taskId,
          error: 'The task no longer exists.',
          errorType: ConcurrencyErrorType.ItemDeleted
        };
      }

      logger.error('TaskAssignmentService',
        `Error updating task ${taskId} with concurrency`, error);

      return {
        success: false,
        taskId,
        error: errorObj.message || 'Failed to update task',
        errorType: ConcurrencyErrorType.Unknown
      };
    }
  }

  /**
   * Complete task with optimistic concurrency control
   * Ensures the task hasn't been modified by another user before completing
   */
  public async completeTaskWithConcurrency(
    taskId: number,
    expectedVersion: string,
    completionNotes?: string,
    actualHours?: number
  ): Promise<IConcurrentUpdateResult & { workflowResumed?: boolean }> {
    try {
      // Get the task first to check for workflow linkage
      const task = await this.getById(taskId);
      if (!task) {
        return {
          success: false,
          taskId,
          error: 'Task not found',
          errorType: ConcurrencyErrorType.ItemDeleted
        };
      }

      // Check if already completed
      if (task.Status === TaskStatus.Completed) {
        return {
          success: false,
          taskId,
          error: 'Task is already completed',
          errorType: ConcurrencyErrorType.None
        };
      }

      // Check if task requires approval
      if (task.RequiresApproval && task.ApprovalStatus !== 'Approved') {
        return {
          success: false,
          taskId,
          error: 'Task requires approval before completion',
          errorType: ConcurrencyErrorType.None
        };
      }

      // Use concurrency-safe update
      const updateResult = await this.updateTaskWithConcurrency(
        taskId,
        {
          Status: TaskStatus.Completed,
          PercentComplete: 100,
          CompletionNotes: completionNotes,
          ActualHours: actualHours || task.ActualHours
        },
        expectedVersion
      );

      if (!updateResult.success) {
        return updateResult;
      }

      // Update the completion date separately (can't use ITaskUpdateData)
      await this.sp.web.lists
        .getByTitle(this.tasksListTitle)
        .items.getById(taskId)
        .update({
          ActualCompletionDate: new Date().toISOString()
        });

      logger.info('TaskAssignmentService', `Completed task ${taskId} with concurrency check`);

      // Trigger workflow resume if applicable
      let workflowResumed = false;
      if (task.WorkflowInstanceId) {
        workflowResumed = await this.triggerWorkflowResume(
          task.WorkflowInstanceId,
          task.WorkflowStepId,
          taskId
        );
      }

      return {
        ...updateResult,
        success: true,
        workflowResumed
      };
    } catch (error) {
      logger.error('TaskAssignmentService',
        `Error completing task ${taskId} with concurrency`, error);

      return {
        success: false,
        taskId,
        error: error instanceof Error ? error.message : 'Failed to complete task',
        errorType: ConcurrencyErrorType.Unknown
      };
    }
  }

  /**
   * Get task with ETag for concurrency checks
   */
  public async getByIdWithVersion(taskId: number): Promise<(IJmlTaskAssignment & { etag?: string }) | null> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.tasksListTitle)
        .items.getById(taskId)
        .select(
          'Id', 'Title', 'Status', 'Priority', 'DueDate', 'AssignedToId',
          'AssignedDate', 'ProcessIDId', 'TaskIDId', 'Modified', 'Created',
          'PercentComplete', 'Notes', 'CompletionNotes', 'ActualHours',
          'WorkflowInstanceId', 'WorkflowStepId', 'RequiresApproval', 'ApprovalStatus',
          'IsBlocked', 'BlockedReason', 'ReminderSent', 'EscalationSent',
          'AssignedTo/Id', 'AssignedTo/Title', 'AssignedTo/EMail'
        )
        .expand('AssignedTo')();

      const mapped = this.mapToTaskAssignment(item);
      const etag = (item as any)['odata.etag'] || (item as any)['@odata.etag'];

      return { ...mapped, etag };
    } catch (error) {
      logger.error('TaskAssignmentService', `Error getting task ${taskId} with version`, error);
      return null;
    }
  }

  /**
   * Complete task - THE KEY METHOD that integrates with workflow and training
   * When a task is completed:
   * 1. Updates task status to Completed
   * 2. If task has WorkflowInstanceId, triggers workflow resume
   * 3. If task is linked to training, updates training enrollment
   */
  public async completeTask(
    taskId: number,
    completionNotes?: string,
    actualHours?: number,
    trainingScore?: number
  ): Promise<ITaskCompletionResult> {
    try {
      // Get the task first to check for workflow linkage
      const task = await this.getById(taskId);
      if (!task) {
        return { success: false, taskId, error: 'Task not found' };
      }

      // Check if already completed
      if (task.Status === TaskStatus.Completed) {
        return { success: false, taskId, error: 'Task is already completed' };
      }

      // Check if task requires approval
      if (task.RequiresApproval && task.ApprovalStatus !== 'Approved') {
        return { success: false, taskId, error: 'Task requires approval before completion' };
      }

      // Update task to completed
      await this.sp.web.lists
        .getByTitle(this.tasksListTitle)
        .items.getById(taskId)
        .update({
          Status: TaskStatus.Completed,
          PercentComplete: 100,
          ActualCompletionDate: new Date().toISOString(),
          CompletionNotes: completionNotes || '',
          ActualHours: actualHours || task.ActualHours
        });

      logger.info('TaskAssignmentService', `Completed task ${taskId}: ${task.Title}`);

      // Check if this task is linked to a workflow
      const result: ITaskCompletionResult = {
        success: true,
        taskId,
        workflowInstanceId: task.WorkflowInstanceId
      };

      // Handle training task completion - update linked training enrollment
      await this.handleTrainingTaskCompletion(taskId, task.AssignedToId, trainingScore);

      if (task.WorkflowInstanceId) {
        // Trigger workflow resume
        result.workflowResumed = await this.triggerWorkflowResume(
          task.WorkflowInstanceId,
          task.WorkflowStepId,
          taskId
        );
      }

      return result;
    } catch (error) {
      logger.error('TaskAssignmentService', `Error completing task ${taskId}`, error);
      return {
        success: false,
        taskId,
        error: error instanceof Error ? error.message : 'Failed to complete task'
      };
    }
  }

  /**
   * Handle training task completion - update linked training enrollment
   * This bridges JML tasks to the Training Skills Builder system
   */
  private async handleTrainingTaskCompletion(
    taskId: number,
    userId: number,
    score?: number
  ): Promise<void> {
    try {
      // Check if there's a training enrollment linked to this task
      const enrollments = await this.sp.web.lists
        .getByTitle('PM_TrainingEnrollments')
        .items.filter(`TaskAssignmentId eq ${taskId}`)
        .select('Id', 'CourseId', 'Status', 'UserId', 'Score', 'Progress')
        .top(1)();

      if (enrollments.length === 0) {
        // No linked training enrollment - this is not a training task
        return;
      }

      const enrollment = enrollments[0];

      // Update enrollment to completed
      await this.sp.web.lists
        .getByTitle('PM_TrainingEnrollments')
        .items.getById(enrollment.Id)
        .update({
          Status: 'Completed',
          Progress: 100,
          CompletedDate: new Date().toISOString(),
          Score: score || enrollment.Score || 100
        });

      logger.info('TaskAssignmentService',
        `Updated training enrollment ${enrollment.Id} to Completed for task ${taskId}`);

      // Update user skill if the course has associated skills
      await this.updateUserSkillFromTraining(userId, enrollment.CourseId, score);

    } catch (error) {
      // Log but don't fail the task completion if training update fails
      logger.warn('TaskAssignmentService',
        `Failed to update training enrollment for task ${taskId}`, error);
    }
  }

  /**
   * Update user skill record when training is completed
   * Links training completion to skills profile in jmlTrainingSkillsBuilder
   */
  private async updateUserSkillFromTraining(
    userId: number,
    courseId: number,
    score?: number
  ): Promise<void> {
    try {
      // Get the course to find associated skills
      const courses = await this.sp.web.lists
        .getByTitle('PM_TrainingCourses')
        .items.filter(`Id eq ${courseId}`)
        .select('Id', 'RelatedSkillIds', 'Tags')
        .top(1)();

      if (courses.length === 0) return;

      const course = courses[0];
      const relatedSkillIds = course.RelatedSkillIds
        ? String(course.RelatedSkillIds).split(';').filter((s: string) => s).map(Number)
        : [];

      if (relatedSkillIds.length === 0) {
        logger.info('TaskAssignmentService', `Course ${courseId} has no related skills`);
        return;
      }

      // Get user's email for skill record
      const user = await this.sp.web.siteUsers.getById(userId)();

      // Update user skills for each related skill
      for (const skillId of relatedSkillIds) {
        await this.upsertUserSkill(userId, user.Email, skillId, score);
      }

      logger.info('TaskAssignmentService',
        `Updated ${relatedSkillIds.length} skills for user ${userId} from course ${courseId}`);

    } catch (error) {
      logger.warn('TaskAssignmentService',
        `Failed to update user skills from training`, error);
    }
  }

  /**
   * Create or update user skill record
   */
  private async upsertUserSkill(
    userId: number,
    userEmail: string,
    skillId: number,
    score?: number
  ): Promise<void> {
    try {
      // Check if user already has this skill
      const existingSkills = await this.sp.web.lists
        .getByTitle('PM_UserSkills')
        .items.filter(`UserId eq ${userId} and SkillId eq ${skillId}`)
        .select('Id', 'SelfRating', 'VerifiedRating')
        .top(1)();

      // Calculate proficiency level from score (1-5 scale)
      const proficiencyLevel = score ? Math.min(5, Math.max(1, Math.round((score / 100) * 5))) : 3;

      if (existingSkills.length > 0) {
        // Update existing skill - increment verified rating based on training
        const existing = existingSkills[0];
        const newVerifiedRating = Math.max(
          existing.VerifiedRating || 0,
          proficiencyLevel
        );

        await this.sp.web.lists
          .getByTitle('PM_UserSkills')
          .items.getById(existing.Id)
          .update({
            VerifiedRating: newVerifiedRating,
            LastAssessedDate: new Date().toISOString(),
            SkillSource: 'Training',
            Evidence: `Training completed with score: ${score || 'N/A'}%`
          });
      } else {
        // Create new skill record
        await this.sp.web.lists
          .getByTitle('PM_UserSkills')
          .items.add({
            Title: `Skill ${skillId} - User ${userId}`,
            UserId: userId,
            UserEmail: userEmail,
            SkillId: skillId,
            SelfRating: proficiencyLevel,
            VerifiedRating: proficiencyLevel,
            LastAssessedDate: new Date().toISOString(),
            SkillSource: 'Training',
            Evidence: `Initial skill from training completion. Score: ${score || 'N/A'}%`
          });
      }
    } catch (error) {
      logger.warn('TaskAssignmentService',
        `Failed to upsert user skill ${skillId} for user ${userId}`, error);
    }
  }

  /**
   * Bulk complete tasks
   */
  public async completeTasks(
    taskIds: number[],
    completionNotes?: string
  ): Promise<ITaskCompletionResult[]> {
    const results: ITaskCompletionResult[] = [];

    for (const taskId of taskIds) {
      const result = await this.completeTask(taskId, completionNotes);
      results.push(result);
    }

    return results;
  }

  /**
   * Trigger workflow resume after task completion
   * This is the critical integration point between tasks and workflows
   */
  private async triggerWorkflowResume(
    workflowInstanceId: number,
    workflowStepId?: string,
    completedTaskId?: number
  ): Promise<boolean> {
    try {
      logger.info('TaskAssignmentService',
        `Triggering workflow resume for instance ${workflowInstanceId} after task ${completedTaskId} completion`);

      // Import workflow engine dynamically to avoid circular dependencies
      const { WorkflowEngineService } = await import('./workflow/WorkflowEngineService');

      if (!this.context) {
        logger.warn('TaskAssignmentService',
          'WebPartContext not available, cannot resume workflow. Workflow will be resumed by scheduled processor.');
        return false;
      }

      const workflowEngine = new WorkflowEngineService(this.sp, this.context);

      // Resume the workflow with task completion data
      const result = await workflowEngine.resumeWorkflow(workflowInstanceId, {
        completedTaskId,
        completedStepId: workflowStepId,
        completedAt: new Date().toISOString(),
        trigger: 'taskCompletion'
      });

      logger.info('TaskAssignmentService',
        `Workflow resume result: ${result.success ? 'Success' : 'Failed'} - Status: ${result.status}`);

      return result.success;
    } catch (error) {
      logger.error('TaskAssignmentService',
        `Error resuming workflow ${workflowInstanceId}`, error);
      // Don't fail the task completion if workflow resume fails
      // The scheduled processor will pick it up
      return false;
    }
  }

  /**
   * Start task (change from Not Started to In Progress)
   */
  public async startTask(taskId: number): Promise<boolean> {
    try {
      const task = await this.getById(taskId);
      if (!task) {
        return false;
      }

      if (task.Status !== TaskStatus.NotStarted) {
        logger.warn('TaskAssignmentService', `Task ${taskId} is not in Not Started status`);
        return false;
      }

      // Check if blocked
      if (task.IsBlocked) {
        logger.warn('TaskAssignmentService', `Task ${taskId} is blocked: ${task.BlockedReason}`);
        return false;
      }

      await this.sp.web.lists
        .getByTitle(this.tasksListTitle)
        .items.getById(taskId)
        .update({
          Status: TaskStatus.InProgress,
          StartDate: new Date().toISOString()
        });

      logger.info('TaskAssignmentService', `Started task ${taskId}`);
      return true;
    } catch (error) {
      logger.error('TaskAssignmentService', `Error starting task ${taskId}`, error);
      return false;
    }
  }

  /**
   * Block task
   */
  public async blockTask(taskId: number, reason: string): Promise<boolean> {
    try {
      await this.sp.web.lists
        .getByTitle(this.tasksListTitle)
        .items.getById(taskId)
        .update({
          Status: TaskStatus.Blocked,
          IsBlocked: true,
          BlockedReason: reason
        });

      logger.info('TaskAssignmentService', `Blocked task ${taskId}: ${reason}`);
      return true;
    } catch (error) {
      logger.error('TaskAssignmentService', `Error blocking task ${taskId}`, error);
      return false;
    }
  }

  /**
   * Unblock task
   */
  public async unblockTask(taskId: number): Promise<boolean> {
    try {
      const task = await this.getById(taskId);
      if (!task) return false;

      // Restore to previous status or In Progress
      const newStatus = task.StartDate ? TaskStatus.InProgress : TaskStatus.NotStarted;

      await this.sp.web.lists
        .getByTitle(this.tasksListTitle)
        .items.getById(taskId)
        .update({
          Status: newStatus,
          IsBlocked: false,
          BlockedReason: ''
        });

      logger.info('TaskAssignmentService', `Unblocked task ${taskId}`);
      return true;
    } catch (error) {
      logger.error('TaskAssignmentService', `Error unblocking task ${taskId}`, error);
      return false;
    }
  }

  /**
   * Reassign task
   */
  public async reassignTask(taskId: number, newAssigneeId: number): Promise<boolean> {
    try {
      const task = await this.getById(taskId);
      if (!task) return false;

      await this.sp.web.lists
        .getByTitle(this.tasksListTitle)
        .items.getById(taskId)
        .update({
          AssignedToId: newAssigneeId,
          AssignedDate: new Date().toISOString()
        });

      logger.info('TaskAssignmentService', `Reassigned task ${taskId} to user ${newAssigneeId}`);
      return true;
    } catch (error) {
      logger.error('TaskAssignmentService', `Error reassigning task ${taskId}`, error);
      return false;
    }
  }

  /**
   * Get tasks awaiting workflow resume
   * These are completed tasks with WorkflowInstanceId where workflow may be waiting
   */
  public async getTasksAwaitingWorkflowResume(): Promise<IJmlTaskAssignment[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.tasksListTitle)
        .items.filter(`Status eq '${TaskStatus.Completed}' and WorkflowInstanceId ne null`)
        .select(
          'Id', 'Title', 'Status', 'WorkflowInstanceId', 'WorkflowStepId',
          'ProcessIDId', 'ActualCompletionDate'
        )
        .orderBy('ActualCompletionDate', false)
        .top(100)();

      return items.map(item => this.mapToTaskAssignment(item));
    } catch (error) {
      logger.error('TaskAssignmentService', 'Error getting tasks awaiting workflow resume', error);
      return [];
    }
  }

  /**
   * Map SharePoint item to IJmlTaskAssignment
   */
  private mapToTaskAssignment(item: Record<string, unknown>): IJmlTaskAssignment {
    const assignedTo = item.AssignedTo as Record<string, unknown> | undefined;

    return {
      Id: item.Id as number,
      Title: item.Title as string,
      Status: item.Status as TaskStatus,
      Priority: item.Priority as Priority,
      DueDate: item.DueDate ? new Date(item.DueDate as string) : undefined,
      AssignedToId: item.AssignedToId as number,
      AssignedTo: assignedTo ? {
        Id: assignedTo.Id as number,
        Title: assignedTo.Title as string,
        EMail: assignedTo.EMail as string
      } as IUser : undefined,
      AssignedDate: item.AssignedDate ? new Date(item.AssignedDate as string) : undefined,
      StartDate: item.StartDate ? new Date(item.StartDate as string) : undefined,
      ActualCompletionDate: item.ActualCompletionDate ? new Date(item.ActualCompletionDate as string) : undefined,
      ProcessIDId: item.ProcessIDId as number,
      TaskIDId: item.TaskIDId as number,
      PercentComplete: item.PercentComplete as number,
      Notes: item.Notes as string,
      CompletionNotes: item.CompletionNotes as string,
      ActualHours: item.ActualHours as number,
      WorkflowInstanceId: item.WorkflowInstanceId as number,
      WorkflowStepId: item.WorkflowStepId as string,
      RequiresApproval: item.RequiresApproval as boolean,
      ApprovalStatus: item.ApprovalStatus as ApprovalStatus,
      IsBlocked: item.IsBlocked as boolean,
      BlockedReason: item.BlockedReason as string,
      ReminderSent: item.ReminderSent as boolean,
      EscalationSent: item.EscalationSent as boolean,
      Modified: item.Modified ? new Date(item.Modified as string) : undefined,
      Created: item.Created ? new Date(item.Created as string) : undefined
    };
  }
}
