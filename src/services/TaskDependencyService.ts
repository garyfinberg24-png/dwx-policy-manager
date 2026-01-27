// @ts-nocheck
// TaskDependencyService - Manages task dependencies, blocking logic, and critical paths

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { IJmlTaskAssignment, TaskStatus } from '../models';
import { logger } from './LoggingService';

export interface ITaskDependencyInfo {
  taskId: number;
  taskTitle: string;
  dependsOnTasks: ITaskDependencyLink[];
  blockedByTasks: ITaskDependencyLink[];
  blockingTasks: ITaskDependencyLink[]; // Tasks that depend on this one
  isBlocked: boolean;
  canStart: boolean;
  criticalPath: boolean;
}

export interface ITaskDependencyLink {
  taskId: number;
  taskTitle: string;
  status: TaskStatus;
  dueDate: Date;
  percentComplete?: number;
}

export interface IDependencyValidationResult {
  valid: boolean;
  error?: string;
  circularPath?: number[];
}

export class TaskDependencyService {
  private sp: SPFI;
  private dependencyColumnsExist: boolean | null = null;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Check if dependency columns exist in the list
   */
  private async checkDependencyColumnsExist(): Promise<boolean> {
    if (this.dependencyColumnsExist !== null) {
      return this.dependencyColumnsExist;
    }

    try {
      // Try to get field to check if it exists
      await this.sp.web.lists.getByTitle('JML_TaskAssignments').fields
        .getByInternalNameOrTitle('DependsOnTaskId')
        .select('InternalName')();
      this.dependencyColumnsExist = true;
      return true;
    } catch {
      logger.warn('TaskDependencyService', 'DependsOnTaskId column does not exist. Task dependencies feature will be unavailable.');
      this.dependencyColumnsExist = false;
      return false;
    }
  }

  /**
   * Get dependency information for a specific task
   */
  public async getTaskDependencyInfo(taskId: number): Promise<ITaskDependencyInfo> {
    try {
      // Check if dependency columns exist
      const columnsExist = await this.checkDependencyColumnsExist();

      // If columns don't exist, return a default response
      if (!columnsExist) {
        try {
          const basicTask = await this.sp.web.lists.getByTitle('JML_TaskAssignments').items
            .getById(taskId)
            .select('Id', 'Title', 'Status', 'DueDate', 'PercentComplete')();

          return {
            taskId: basicTask.Id,
            taskTitle: basicTask.Title,
            dependsOnTasks: [],
            blockedByTasks: [],
            blockingTasks: [],
            isBlocked: false,
            canStart: true,
            criticalPath: false
          };
        } catch {
          // Return minimal response if even basic query fails
          return {
            taskId: taskId,
            taskTitle: 'Unknown',
            dependsOnTasks: [],
            blockedByTasks: [],
            blockingTasks: [],
            isBlocked: false,
            canStart: true,
            criticalPath: false
          };
        }
      }

      const task = await this.sp.web.lists.getByTitle('JML_TaskAssignments').items
        .getById(taskId)
        .select(
          'Id', 'Title', 'Status', 'DueDate', 'PercentComplete',
          'DependsOnTaskId', 'IsBlocked', 'BlockedReason'
        )();

      // Get tasks this task depends on
      const dependsOnTasks: ITaskDependencyLink[] = [];
      if (task.DependsOnTaskId) {
        const dependsOnTask = await this.sp.web.lists.getByTitle('JML_TaskAssignments').items
          .getById(task.DependsOnTaskId)
          .select('Id', 'Title', 'Status', 'DueDate', 'PercentComplete')();

        dependsOnTasks.push({
          taskId: dependsOnTask.Id,
          taskTitle: dependsOnTask.Title,
          status: dependsOnTask.Status,
          dueDate: new Date(dependsOnTask.DueDate),
          percentComplete: dependsOnTask.PercentComplete
        });
      }

      // Get tasks that depend on this task
      const blockingTasks = await this.sp.web.lists.getByTitle('JML_TaskAssignments').items
        .filter(`DependsOnTaskId eq ${taskId}`)
        .select('Id', 'Title', 'Status', 'DueDate', 'PercentComplete')();

      const blockingTaskLinks: ITaskDependencyLink[] = blockingTasks.map(t => ({
        taskId: t.Id,
        taskTitle: t.Title,
        status: t.Status,
        dueDate: new Date(t.DueDate),
        percentComplete: t.PercentComplete
      }));

      // Determine if task is blocked and can start
      const canStart = this.canTaskStart(dependsOnTasks);
      const isBlocked = !canStart;

      return {
        taskId: task.Id,
        taskTitle: task.Title,
        dependsOnTasks,
        blockedByTasks: isBlocked ? dependsOnTasks.filter(t => t.status !== TaskStatus.Completed) : [],
        blockingTasks: blockingTaskLinks,
        isBlocked,
        canStart,
        criticalPath: false // Will be calculated by critical path algorithm
      };
    } catch (error) {
      logger.error('TaskDependencyService', `Error getting dependency info for task ${taskId}`, error);
      throw error;
    }
  }

  /**
   * Check if a task can start based on its dependencies
   */
  private canTaskStart(dependencies: ITaskDependencyLink[]): boolean {
    if (!dependencies || dependencies.length === 0) {
      return true;
    }

    // All dependencies must be completed
    return dependencies.every(dep => dep.status === TaskStatus.Completed);
  }

  /**
   * Add a dependency between tasks
   */
  public async addDependency(taskId: number, dependsOnTaskId: number): Promise<void> {
    try {
      // Validate for circular dependencies
      const validation = await this.validateDependency(taskId, dependsOnTaskId);
      if (!validation.valid) {
        throw new Error(validation.error);
      }

      // Update the task with the dependency
      await this.sp.web.lists.getByTitle('JML_TaskAssignments').items
        .getById(taskId)
        .update({
          DependsOnTaskId: dependsOnTaskId,
          IsDependentTask: true
        });

      // Check if task should be blocked
      await this.updateBlockedStatus(taskId);

      logger.info('TaskDependencyService', `Added dependency: Task ${taskId} depends on Task ${dependsOnTaskId}`);
    } catch (error) {
      logger.error('TaskDependencyService', `Error adding dependency`, error);
      throw error;
    }
  }

  /**
   * Remove a dependency
   */
  public async removeDependency(taskId: number): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('JML_TaskAssignments').items
        .getById(taskId)
        .update({
          DependsOnTaskId: null,
          IsDependentTask: false,
          IsBlocked: false,
          BlockedReason: null
        });

      logger.info('TaskDependencyService', `Removed dependency from Task ${taskId}`);
    } catch (error) {
      logger.error('TaskDependencyService', `Error removing dependency`, error);
      throw error;
    }
  }

  /**
   * Validate a dependency to prevent circular references
   */
  public async validateDependency(taskId: number, dependsOnTaskId: number): Promise<IDependencyValidationResult> {
    try {
      // Can't depend on itself
      if (taskId === dependsOnTaskId) {
        return {
          valid: false,
          error: 'A task cannot depend on itself'
        };
      }

      // Check if adding this dependency would create a circular reference
      const wouldCreateCircular = await this.wouldCreateCircularDependency(taskId, dependsOnTaskId);
      if (wouldCreateCircular) {
        return {
          valid: false,
          error: 'This dependency would create a circular reference'
        };
      }

      return { valid: true };
    } catch (error) {
      logger.error('TaskDependencyService', 'Error validating dependency', error);
      return {
        valid: false,
        error: error.message || 'Validation failed'
      };
    }
  }

  /**
   * Check if a dependency would create a circular reference
   */
  private async wouldCreateCircularDependency(taskId: number, dependsOnTaskId: number): Promise<boolean> {
    try {
      // Build dependency chain from the target task
      const visited = new Set<number>();
      const queue: number[] = [dependsOnTaskId];

      while (queue.length > 0) {
        const currentTaskId = queue.shift()!;

        // If we've reached the original task, we have a circular dependency
        if (currentTaskId === taskId) {
          return true;
        }

        // Prevent infinite loops
        if (visited.has(currentTaskId)) {
          continue;
        }
        visited.add(currentTaskId);

        // Get tasks that currentTask depends on
        const currentTask = await this.sp.web.lists.getByTitle('JML_TaskAssignments').items
          .getById(currentTaskId)
          .select('DependsOnTaskId')();

        if (currentTask.DependsOnTaskId) {
          queue.push(currentTask.DependsOnTaskId);
        }
      }

      return false;
    } catch (error) {
      logger.error('TaskDependencyService', 'Error checking circular dependency', error);
      return true; // Err on the side of caution
    }
  }

  /**
   * Update blocked status based on dependencies
   */
  public async updateBlockedStatus(taskId: number): Promise<void> {
    try {
      const task = await this.sp.web.lists.getByTitle('JML_TaskAssignments').items
        .getById(taskId)
        .select('Id', 'DependsOnTaskId', 'Status')();

      if (!task.DependsOnTaskId || task.Status === TaskStatus.Completed) {
        return; // No dependency or already completed
      }

      // Check if dependency is completed
      const dependsOnTask = await this.sp.web.lists.getByTitle('JML_TaskAssignments').items
        .getById(task.DependsOnTaskId)
        .select('Id', 'Title', 'Status')();

      const isBlocked = dependsOnTask.Status !== TaskStatus.Completed;
      const blockedReason = isBlocked ? `Waiting for "${dependsOnTask.Title}" to be completed` : null;

      await this.sp.web.lists.getByTitle('JML_TaskAssignments').items
        .getById(taskId)
        .update({
          IsBlocked: isBlocked,
          BlockedReason: blockedReason
        });
    } catch (error) {
      logger.error('TaskDependencyService', `Error updating blocked status for task ${taskId}`, error);
      throw error;
    }
  }

  /**
   * When a task is completed, unblock dependent tasks
   */
  public async onTaskCompleted(taskId: number): Promise<void> {
    try {
      // Find all tasks that depend on this task
      const dependentTasks = await this.sp.web.lists.getByTitle('JML_TaskAssignments').items
        .filter(`DependsOnTaskId eq ${taskId}`)
        .select('Id')();

      // Update blocked status for each dependent task
      for (const task of dependentTasks) {
        await this.updateBlockedStatus(task.Id);
      }

      logger.info('TaskDependencyService', `Unblocked ${dependentTasks.length} tasks after completing Task ${taskId}`);
    } catch (error) {
      logger.error('TaskDependencyService', `Error unblocking dependent tasks`, error);
      throw error;
    }
  }

  /**
   * Get all tasks for a process with dependency information
   */
  public async getProcessTasksWithDependencies(processId: number): Promise<ITaskDependencyInfo[]> {
    try {
      const tasks = await this.sp.web.lists.getByTitle('JML_TaskAssignments').items
        .filter(`ProcessID eq '${processId}'`)
        .select('Id', 'Title', 'Status', 'DueDate', 'PercentComplete', 'DependsOnTaskId', 'IsBlocked')
        .orderBy('DueDate', true)();

      const taskInfos: ITaskDependencyInfo[] = [];

      for (const task of tasks) {
        const info = await this.getTaskDependencyInfo(task.Id);
        taskInfos.push(info);
      }

      return taskInfos;
    } catch (error) {
      logger.error('TaskDependencyService', `Error getting process tasks with dependencies`, error);
      throw error;
    }
  }

  /**
   * Calculate critical path for a process
   */
  public async calculateCriticalPath(processId: number): Promise<number[]> {
    try {
      const tasks = await this.getProcessTasksWithDependencies(processId);

      // Build dependency graph
      const graph = new Map<number, number[]>();
      const inDegree = new Map<number, number>();
      const taskDueDates = new Map<number, Date>();

      for (const task of tasks) {
        graph.set(task.taskId, task.blockingTasks.map(t => t.taskId));
        inDegree.set(task.taskId, task.dependsOnTasks.length);
        // taskDueDates would need to be populated from task data
      }

      // Topological sort to find longest path (critical path)
      const criticalPath: number[] = [];
      const queue: number[] = [];

      // Find tasks with no dependencies
      inDegree.forEach((degree, taskId) => {
        if (degree === 0) {
          queue.push(taskId);
        }
      });

      // Process tasks in topological order
      while (queue.length > 0) {
        const taskId = queue.shift()!;
        criticalPath.push(taskId);

        const dependents = graph.get(taskId) || [];
        for (const depTaskId of dependents) {
          const degree = inDegree.get(depTaskId)! - 1;
          inDegree.set(depTaskId, degree);
          if (degree === 0) {
            queue.push(depTaskId);
          }
        }
      }

      return criticalPath;
    } catch (error) {
      logger.error('TaskDependencyService', `Error calculating critical path`, error);
      throw error;
    }
  }

  /**
   * Get available tasks that can be set as dependencies (prevent circular refs)
   */
  public async getAvailableDependencies(taskId: number, processId: number): Promise<IJmlTaskAssignment[]> {
    try {
      // Get all tasks in the process
      const allTasks = await this.sp.web.lists.getByTitle('JML_TaskAssignments').items
        .filter(`ProcessID eq '${processId}'`)
        .select('Id', 'Title', 'Status', 'DueDate', 'PercentComplete')
        .orderBy('DueDate', true)();

      // Filter out tasks that would create circular dependencies
      const availableTasks: IJmlTaskAssignment[] = [];

      for (const task of allTasks) {
        // Skip the task itself
        if (task.Id === taskId) {
          continue;
        }

        // Check if this would create a circular dependency
        const validation = await this.validateDependency(taskId, task.Id);
        if (validation.valid) {
          availableTasks.push(task as IJmlTaskAssignment);
        }
      }

      return availableTasks;
    } catch (error) {
      logger.error('TaskDependencyService', 'Error getting available dependencies', error);
      throw error;
    }
  }
}
