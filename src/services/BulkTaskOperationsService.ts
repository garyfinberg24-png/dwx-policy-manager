// @ts-nocheck
// BulkTaskOperationsService - Handles batch operations on multiple tasks
// Provides bulk complete, reassign, update status, and delete functionality
// ENHANCED: Integrates with WorkflowResumeService for workflow automation

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/batching';
import { IJmlTaskAssignment, TaskStatus } from '../models';
import { logger } from './LoggingService';
import { WorkflowResumeService } from './workflow/WorkflowResumeService';

export interface IBulkOperationResult {
  totalItems: number;
  successCount: number;
  failureCount: number;
  errors: IBulkOperationError[];
  rollbackData?: IBulkRollbackData[];
}

export interface IBulkOperationError {
  taskId: number;
  taskTitle: string;
  error: string;
}

export interface IBulkRollbackData {
  taskId: number;
  originalValues: Partial<IJmlTaskAssignment>;
}

export interface IBulkUpdateOptions {
  status?: TaskStatus;
  assignedToId?: number;
  dueDate?: Date;
  priority?: string;
  notes?: string;
}

export class BulkTaskOperationsService {
  private sp: SPFI;
  private workflowResumeService: WorkflowResumeService | undefined;

  constructor(sp: SPFI, workflowResumeService?: WorkflowResumeService) {
    this.sp = sp;
    this.workflowResumeService = workflowResumeService;
  }

  /**
   * Set workflow resume service for workflow integration
   * Can be called after construction if service wasn't available initially
   */
  public setWorkflowResumeService(service: WorkflowResumeService): void {
    this.workflowResumeService = service;
  }

  /**
   * Bulk complete tasks
   */
  public async bulkCompleteTasks(taskIds: number[]): Promise<IBulkOperationResult> {
    logger.info('BulkTaskOperationsService', `Bulk completing ${taskIds.length} tasks`);

    const result: IBulkOperationResult = {
      totalItems: taskIds.length,
      successCount: 0,
      failureCount: 0,
      errors: [],
      rollbackData: []
    };

    try {
      const list = this.sp.web.lists.getByTitle('PM_TaskAssignments');
      const [batch] = this.sp.web.batched();
      const batchedList = this.sp.web.lists.getByTitle('PM_TaskAssignments').using(batch);

      // Store original values for rollback
      const rollbackData: IBulkRollbackData[] = [];

      for (const taskId of taskIds) {
        try {
          // Get original task for rollback
          const originalTask = await list.items.getById(taskId)
            .select('Id', 'Title', 'Status', 'PercentComplete', 'ActualCompletionDate')();

          rollbackData.push({
            taskId,
            originalValues: {
              Status: originalTask.Status,
              PercentComplete: originalTask.PercentComplete,
              ActualCompletionDate: originalTask.ActualCompletionDate
            }
          });

          // Update task in batch
          batchedList.items.getById(taskId).update({
            Status: TaskStatus.Completed,
            PercentComplete: 100,
            ActualCompletionDate: new Date().toISOString()
          });

          result.successCount++;
        } catch (error) {
          result.failureCount++;
          result.errors.push({
            taskId,
            taskTitle: `Task ${taskId}`,
            error: error.message || 'Unknown error'
          });
          logger.error('BulkTaskOperationsService', `Error completing task ${taskId}`, error);
        }
      }

      // Execute batch
      await batch;
      result.rollbackData = rollbackData;

      logger.info('BulkTaskOperationsService',
        `Bulk complete finished: ${result.successCount} succeeded, ${result.failureCount} failed`);

      // WORKFLOW INTEGRATION: Trigger workflow resume for completed tasks
      // This ensures workflows waiting for these tasks are properly advanced
      if (this.workflowResumeService && result.successCount > 0) {
        await this.triggerWorkflowResumeForTasks(taskIds);
      }

      return result;
    } catch (error) {
      logger.error('BulkTaskOperationsService', 'Bulk complete operation failed', error);
      throw error;
    }
  }

  /**
   * Trigger workflow resume for completed tasks
   * Calls WorkflowResumeService to check if any waiting workflows should be advanced
   */
  private async triggerWorkflowResumeForTasks(taskIds: number[]): Promise<void> {
    if (!this.workflowResumeService) {
      return;
    }

    try {
      logger.info('BulkTaskOperationsService',
        `Triggering workflow resume check for ${taskIds.length} completed tasks`);

      for (const taskId of taskIds) {
        try {
          // WorkflowResumeService.onTaskCompleted will handle fetching task details
          // and checking if workflow should be resumed
          await this.workflowResumeService.onTaskCompleted(taskId, {
            completedVia: 'bulk-operation',
            completedAt: new Date().toISOString()
          });
        } catch (taskError) {
          // Log but don't fail the entire operation for individual task workflow resume failures
          logger.warn('BulkTaskOperationsService',
            `Failed to trigger workflow resume for task ${taskId}`, taskError);
        }
      }

      logger.info('BulkTaskOperationsService',
        `Workflow resume checks completed for bulk task completion`);
    } catch (error) {
      // Log but don't throw - the task completion was successful, workflow resume is secondary
      logger.error('BulkTaskOperationsService',
        'Failed to trigger workflow resume for bulk completed tasks', error);
    }
  }

  /**
   * Bulk reassign tasks to a new user
   */
  public async bulkReassignTasks(taskIds: number[], newAssigneeId: number): Promise<IBulkOperationResult> {
    logger.info('BulkTaskOperationsService',
      `Bulk reassigning ${taskIds.length} tasks to user ${newAssigneeId}`);

    const result: IBulkOperationResult = {
      totalItems: taskIds.length,
      successCount: 0,
      failureCount: 0,
      errors: [],
      rollbackData: []
    };

    try {
      const list = this.sp.web.lists.getByTitle('PM_TaskAssignments');
      const [batch] = this.sp.web.batched();
      const batchedList = this.sp.web.lists.getByTitle('PM_TaskAssignments').using(batch);

      // Store original values for rollback
      const rollbackData: IBulkRollbackData[] = [];

      for (const taskId of taskIds) {
        try {
          // Get original task for rollback
          const originalTask = await list.items.getById(taskId)
            .select('Id', 'Title', 'AssignedToId', 'AssignedDate')();

          rollbackData.push({
            taskId,
            originalValues: {
              AssignedToId: originalTask.AssignedToId,
              AssignedDate: originalTask.AssignedDate
            }
          });

          // Update task in batch
          batchedList.items.getById(taskId).update({
            AssignedToId: newAssigneeId,
            AssignedDate: new Date().toISOString()
          });

          result.successCount++;
        } catch (error) {
          result.failureCount++;
          result.errors.push({
            taskId,
            taskTitle: `Task ${taskId}`,
            error: error.message || 'Unknown error'
          });
          logger.error('BulkTaskOperationsService', `Error reassigning task ${taskId}`, error);
        }
      }

      // Execute batch
      await batch;
      result.rollbackData = rollbackData;

      logger.info('BulkTaskOperationsService',
        `Bulk reassign finished: ${result.successCount} succeeded, ${result.failureCount} failed`);

      return result;
    } catch (error) {
      logger.error('BulkTaskOperationsService', 'Bulk reassign operation failed', error);
      throw error;
    }
  }

  /**
   * Bulk update task status
   */
  public async bulkUpdateStatus(taskIds: number[], newStatus: TaskStatus): Promise<IBulkOperationResult> {
    logger.info('BulkTaskOperationsService',
      `Bulk updating status of ${taskIds.length} tasks to ${newStatus}`);

    const result: IBulkOperationResult = {
      totalItems: taskIds.length,
      successCount: 0,
      failureCount: 0,
      errors: [],
      rollbackData: []
    };

    try {
      const list = this.sp.web.lists.getByTitle('PM_TaskAssignments');
      const [batch] = this.sp.web.batched();
      const batchedList = this.sp.web.lists.getByTitle('PM_TaskAssignments').using(batch);

      // Store original values for rollback
      const rollbackData: IBulkRollbackData[] = [];

      for (const taskId of taskIds) {
        try {
          // Get original task for rollback
          const originalTask = await list.items.getById(taskId)
            .select('Id', 'Title', 'Status')();

          rollbackData.push({
            taskId,
            originalValues: {
              Status: originalTask.Status
            }
          });

          // Update task in batch
          batchedList.items.getById(taskId).update({
            Status: newStatus
          });

          result.successCount++;
        } catch (error) {
          result.failureCount++;
          result.errors.push({
            taskId,
            taskTitle: `Task ${taskId}`,
            error: error.message || 'Unknown error'
          });
          logger.error('BulkTaskOperationsService', `Error updating status for task ${taskId}`, error);
        }
      }

      // Execute batch
      await batch;
      result.rollbackData = rollbackData;

      logger.info('BulkTaskOperationsService',
        `Bulk status update finished: ${result.successCount} succeeded, ${result.failureCount} failed`);

      return result;
    } catch (error) {
      logger.error('BulkTaskOperationsService', 'Bulk status update operation failed', error);
      throw error;
    }
  }

  /**
   * Bulk update multiple fields
   */
  public async bulkUpdateTasks(
    taskIds: number[],
    updates: IBulkUpdateOptions
  ): Promise<IBulkOperationResult> {
    logger.info('BulkTaskOperationsService',
      `Bulk updating ${taskIds.length} tasks with custom fields`);

    const result: IBulkOperationResult = {
      totalItems: taskIds.length,
      successCount: 0,
      failureCount: 0,
      errors: [],
      rollbackData: []
    };

    try {
      const list = this.sp.web.lists.getByTitle('PM_TaskAssignments');
      const [batch] = this.sp.web.batched();
      const batchedList = this.sp.web.lists.getByTitle('PM_TaskAssignments').using(batch);

      // Build update object
      const updateObj: any = {};
      if (updates.status) updateObj.Status = updates.status;
      if (updates.assignedToId) updateObj.AssignedToId = updates.assignedToId;
      if (updates.dueDate) updateObj.DueDate = updates.dueDate.toISOString();
      if (updates.priority) updateObj.Priority = updates.priority;
      if (updates.notes) updateObj.Notes = updates.notes;

      // Store original values for rollback
      const rollbackData: IBulkRollbackData[] = [];
      const fieldsToGet = ['Id', 'Title'];
      if (updates.status) fieldsToGet.push('Status');
      if (updates.assignedToId) fieldsToGet.push('AssignedToId');
      if (updates.dueDate) fieldsToGet.push('DueDate');
      if (updates.priority) fieldsToGet.push('Priority');
      if (updates.notes) fieldsToGet.push('Notes');

      for (const taskId of taskIds) {
        try {
          // Get original task for rollback
          const originalTask = await list.items.getById(taskId)
            .select(...fieldsToGet)();

          const originalValues: any = {};
          if (updates.status) originalValues.Status = originalTask.Status;
          if (updates.assignedToId) originalValues.AssignedToId = originalTask.AssignedToId;
          if (updates.dueDate) originalValues.DueDate = originalTask.DueDate;
          if (updates.priority) originalValues.Priority = originalTask.Priority;
          if (updates.notes) originalValues.Notes = originalTask.Notes;

          rollbackData.push({
            taskId,
            originalValues
          });

          // Update task in batch
          batchedList.items.getById(taskId).update(updateObj);

          result.successCount++;
        } catch (error) {
          result.failureCount++;
          result.errors.push({
            taskId,
            taskTitle: `Task ${taskId}`,
            error: error.message || 'Unknown error'
          });
          logger.error('BulkTaskOperationsService', `Error updating task ${taskId}`, error);
        }
      }

      // Execute batch
      await batch;
      result.rollbackData = rollbackData;

      logger.info('BulkTaskOperationsService',
        `Bulk update finished: ${result.successCount} succeeded, ${result.failureCount} failed`);

      return result;
    } catch (error) {
      logger.error('BulkTaskOperationsService', 'Bulk update operation failed', error);
      throw error;
    }
  }

  /**
   * Bulk delete tasks (soft delete by setting IsDeleted flag)
   */
  public async bulkDeleteTasks(taskIds: number[], hardDelete: boolean = false): Promise<IBulkOperationResult> {
    logger.info('BulkTaskOperationsService',
      `Bulk ${hardDelete ? 'hard' : 'soft'} deleting ${taskIds.length} tasks`);

    const result: IBulkOperationResult = {
      totalItems: taskIds.length,
      successCount: 0,
      failureCount: 0,
      errors: [],
      rollbackData: []
    };

    try {
      const list = this.sp.web.lists.getByTitle('PM_TaskAssignments');

      if (hardDelete) {
        // Hard delete - actually remove items
        const [batch] = this.sp.web.batched();
        const batchedList = this.sp.web.lists.getByTitle('PM_TaskAssignments').using(batch);

        for (const taskId of taskIds) {
          try {
            batchedList.items.getById(taskId).delete();
            result.successCount++;
          } catch (error) {
            result.failureCount++;
            result.errors.push({
              taskId,
              taskTitle: `Task ${taskId}`,
              error: error.message || 'Unknown error'
            });
            logger.error('BulkTaskOperationsService', `Error hard deleting task ${taskId}`, error);
          }
        }

        await batch;
      } else {
        // Soft delete - set IsDeleted flag
        const [batch] = this.sp.web.batched();
        const batchedList = this.sp.web.lists.getByTitle('PM_TaskAssignments').using(batch);
        const rollbackData: IBulkRollbackData[] = [];

        for (const taskId of taskIds) {
          try {
            // Get original task for rollback
            const originalTask = await list.items.getById(taskId)
              .select('Id', 'Title', 'IsDeleted')();

            rollbackData.push({
              taskId,
              originalValues: {
                IsDeleted: originalTask.IsDeleted || false
              }
            });

            // Update task in batch
            batchedList.items.getById(taskId).update({
              IsDeleted: true
            });

            result.successCount++;
          } catch (error) {
            result.failureCount++;
            result.errors.push({
              taskId,
              taskTitle: `Task ${taskId}`,
              error: error.message || 'Unknown error'
            });
            logger.error('BulkTaskOperationsService', `Error soft deleting task ${taskId}`, error);
          }
        }

        await batch;
        result.rollbackData = rollbackData;
      }

      logger.info('BulkTaskOperationsService',
        `Bulk delete finished: ${result.successCount} succeeded, ${result.failureCount} failed`);

      return result;
    } catch (error) {
      logger.error('BulkTaskOperationsService', 'Bulk delete operation failed', error);
      throw error;
    }
  }

  /**
   * Rollback a bulk operation using stored rollback data
   */
  public async rollbackOperation(rollbackData: IBulkRollbackData[]): Promise<IBulkOperationResult> {
    logger.info('BulkTaskOperationsService', `Rolling back ${rollbackData.length} tasks`);

    const result: IBulkOperationResult = {
      totalItems: rollbackData.length,
      successCount: 0,
      failureCount: 0,
      errors: []
    };

    try {
      const list = this.sp.web.lists.getByTitle('PM_TaskAssignments');
      const [batch] = this.sp.web.batched();
      const batchedList = this.sp.web.lists.getByTitle('PM_TaskAssignments').using(batch);

      for (const item of rollbackData) {
        try {
          // Restore original values
          batchedList.items.getById(item.taskId).update(item.originalValues);
          result.successCount++;
        } catch (error) {
          result.failureCount++;
          result.errors.push({
            taskId: item.taskId,
            taskTitle: `Task ${item.taskId}`,
            error: error.message || 'Unknown error'
          });
          logger.error('BulkTaskOperationsService', `Error rolling back task ${item.taskId}`, error);
        }
      }

      // Execute batch
      await batch;

      logger.info('BulkTaskOperationsService',
        `Rollback finished: ${result.successCount} succeeded, ${result.failureCount} failed`);

      return result;
    } catch (error) {
      logger.error('BulkTaskOperationsService', 'Rollback operation failed', error);
      throw error;
    }
  }

  /**
   * Validate if bulk operation can be performed
   */
  public async validateBulkOperation(
    taskIds: number[],
    operationType: 'complete' | 'reassign' | 'delete'
  ): Promise<{ valid: boolean; errors: string[] }> {
    const errors: string[] = [];

    try {
      const list = this.sp.web.lists.getByTitle('PM_TaskAssignments');

      for (const taskId of taskIds) {
        const task = await list.items.getById(taskId)
          .select('Id', 'Status', 'IsDeleted', 'RequiresApproval', 'ApprovalStatus')();

        // Check if already deleted
        if (task.IsDeleted) {
          errors.push(`Task ${taskId} is already deleted`);
        }

        // Operation-specific validation
        if (operationType === 'complete') {
          if (task.Status === TaskStatus.Completed) {
            errors.push(`Task ${taskId} is already completed`);
          }
          if (task.RequiresApproval && task.ApprovalStatus !== 'Approved') {
            errors.push(`Task ${taskId} requires approval before completion`);
          }
        }
      }

      return {
        valid: errors.length === 0,
        errors
      };
    } catch (error) {
      logger.error('BulkTaskOperationsService', 'Validation failed', error);
      return {
        valid: false,
        errors: ['Validation failed: ' + error.message]
      };
    }
  }
}
