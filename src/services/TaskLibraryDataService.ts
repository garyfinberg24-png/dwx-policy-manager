// @ts-nocheck
// TaskLibraryDataService - SharePoint-backed service for managing PM_Tasks list
// This service provides CRUD operations for task library templates

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/batching';
import { IJmlTask } from '../models/IJmlTask';
import { logger } from './LoggingService';

export class TaskLibraryDataService {
  private sp: SPFI;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Get all tasks from the PM_Tasks list
   */
  public async getAllTasks(): Promise<IJmlTask[]> {
    try {
      // Note: ApprovalRole field may not exist in all environments
      // Query only the core fields that are guaranteed to exist
      // NOTE: Author/Editor may be text fields or Person lookups depending on list config
      // Do NOT use expand() as it will fail if they are text fields
      // Note: Dependencies field removed - does not exist in PM_Tasks list
      // If task dependencies are needed, add the column to the SharePoint list first
      const items = await this.sp.web.lists.getByTitle('PM_Tasks').items
        .select(
          'Id', 'Title', 'TaskCode', 'Category', 'Description', 'Instructions',
          'Department', 'SLAHours', 'RequiresApproval',
          'IsActive', 'Priority', 'Tags',
          'RelatedLinks', 'AutomationAvailable',
          'Created', 'Modified'
        )
        .orderBy('TaskCode', true)();

      // Map list columns to expected interface properties
      return items.map((item: any) => ({
        ...item,
        ApproverRole: item.ApprovalRole || '' // May be undefined if field doesn't exist
      })) as IJmlTask[];
    } catch (error) {
      logger.error('TaskLibraryDataService', 'Error getting all tasks', error);
      throw error;
    }
  }

  /**
   * Get active tasks only
   */
  public async getActiveTasks(): Promise<IJmlTask[]> {
    try {
      const items = await this.sp.web.lists.getByTitle('PM_Tasks').items
        .filter('IsActive eq true')
        .select(
          'Id', 'Title', 'TaskCode', 'Category', 'Description', 'Instructions',
          'Department', 'SLAHours',
          'RequiresApproval', 'Priority', 'IsActive'
        )
        .orderBy('Category', true)
        .orderBy('TaskCode', true)();

      return items as IJmlTask[];
    } catch (error) {
      logger.error('TaskLibraryDataService', 'Error getting active tasks', error);
      throw error;
    }
  }

  /**
   * Get a single task by ID
   */
  public async getTaskById(id: number): Promise<IJmlTask> {
    try {
      // Note: ApprovalRole field may not exist in all environments
      // Note: Dependencies field removed - does not exist in PM_Tasks list
      const item: any = await this.sp.web.lists.getByTitle('PM_Tasks').items
        .getById(id)
        .select(
          'Id', 'Title', 'TaskCode', 'Category', 'Description', 'Instructions',
          'Department', 'SLAHours', 'RequiresApproval',
          'IsActive', 'Priority', 'Tags',
          'RelatedLinks', 'AutomationAvailable'
        )();

      return {
        ...item,
        ApproverRole: item.ApprovalRole || '' // May be undefined if field doesn't exist
      } as IJmlTask;
    } catch (error) {
      logger.error('TaskLibraryDataService', `Error getting task ${id}`, error);
      throw error;
    }
  }

  /**
   * Create a new task
   */
  public async createTask(task: Partial<IJmlTask>): Promise<IJmlTask> {
    try {
      const result = await this.sp.web.lists.getByTitle('PM_Tasks').items.add({
        Title: task.Title,
        TaskCode: task.TaskCode,
        Category: task.Category,
        Description: task.Description,
        Instructions: task.Instructions,
        Department: task.Department,
        DefaultAssigneeRole: task.AssigneeRole,
        SLAHours: task.SLAHours,
        RequiresApproval: task.RequiresApproval,
        ApprovalRole: task.ApproverRole,
        IsActive: task.IsActive !== false,
        Priority: task.Priority,
        Tags: task.Tags,
        Dependencies: task.DependsOn,
        RelatedLinks: task.RelatedLinks,
        AutomationAvailable: task.AutomationAvailable
      });

      logger.info('TaskLibraryDataService', `Created task: ${task.TaskCode}`);

      // Return the created item
      return await this.getTaskById(result.data.Id);
    } catch (error) {
      logger.error('TaskLibraryDataService', 'Error creating task', error);
      throw error;
    }
  }

  /**
   * Update an existing task
   */
  public async updateTask(id: number, updates: Partial<IJmlTask>): Promise<void> {
    try {
      const updateData: any = {};

      // Only include fields that are provided - map interface names to SharePoint column names
      if (updates.Title !== undefined) updateData.Title = updates.Title;
      if (updates.TaskCode !== undefined) updateData.TaskCode = updates.TaskCode;
      if (updates.Category !== undefined) updateData.Category = updates.Category;
      if (updates.Description !== undefined) updateData.Description = updates.Description;
      if (updates.Instructions !== undefined) updateData.Instructions = updates.Instructions;
      if (updates.Department !== undefined) updateData.Department = updates.Department;
      if (updates.AssigneeRole !== undefined) updateData.DefaultAssigneeRole = updates.AssigneeRole;
      if (updates.SLAHours !== undefined) updateData.SLAHours = updates.SLAHours;
      if (updates.RequiresApproval !== undefined) updateData.RequiresApproval = updates.RequiresApproval;
      if (updates.ApproverRole !== undefined) updateData.ApprovalRole = updates.ApproverRole;
      if (updates.DependsOn !== undefined) updateData.Dependencies = updates.DependsOn;
      if (updates.IsActive !== undefined) updateData.IsActive = updates.IsActive;
      if (updates.Priority !== undefined) updateData.Priority = updates.Priority;
      if (updates.Tags !== undefined) updateData.Tags = updates.Tags;
      if (updates.RelatedLinks !== undefined) updateData.RelatedLinks = updates.RelatedLinks;
      if (updates.AutomationAvailable !== undefined) updateData.AutomationAvailable = updates.AutomationAvailable;

      await this.sp.web.lists.getByTitle('PM_Tasks').items
        .getById(id)
        .update(updateData);

      logger.info('TaskLibraryDataService', `Updated task ${id}`);
    } catch (error) {
      logger.error('TaskLibraryDataService', `Error updating task ${id}`, error);
      throw error;
    }
  }

  /**
   * Delete a task
   */
  public async deleteTask(id: number): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('PM_Tasks').items
        .getById(id)
        .delete();

      logger.info('TaskLibraryDataService', `Deleted task ${id}`);
    } catch (error) {
      logger.error('TaskLibraryDataService', `Error deleting task ${id}`, error);
      throw error;
    }
  }

  /**
   * Get tasks by category
   */
  public async getTasksByCategory(category: string): Promise<IJmlTask[]> {
    try {
      const items = await this.sp.web.lists.getByTitle('PM_Tasks').items
        .filter(`Category eq '${category}'`)
        .select(
          'Id', 'Title', 'TaskCode', 'Category', 'Description',
          'Department', 'Priority', 'IsActive', 'SLAHours'
        )
        .orderBy('TaskCode', true)();

      return items as IJmlTask[];
    } catch (error) {
      logger.error('TaskLibraryDataService', `Error getting tasks by category ${category}`, error);
      throw error;
    }
  }

  /**
   * Get tasks by department
   */
  public async getTasksByDepartment(department: string): Promise<IJmlTask[]> {
    try {
      const items = await this.sp.web.lists.getByTitle('PM_Tasks').items
        .filter(`Department eq '${department}'`)
        .select(
          'Id', 'Title', 'TaskCode', 'Category', 'Description',
          'Department', 'Priority', 'IsActive', 'SLAHours'
        )
        .orderBy('TaskCode', true)();

      return items as IJmlTask[];
    } catch (error) {
      logger.error('TaskLibraryDataService', `Error getting tasks by department ${department}`, error);
      throw error;
    }
  }

  /**
   * Search tasks by keyword
   */
  public async searchTasks(keyword: string): Promise<IJmlTask[]> {
    try {
      const lowerKeyword = keyword.toLowerCase();

      // Get all tasks and filter client-side for better search
      const allTasks = await this.getAllTasks();

      return allTasks.filter(task =>
        task.Title.toLowerCase().includes(lowerKeyword) ||
        task.TaskCode.toLowerCase().includes(lowerKeyword) ||
        task.Description?.toLowerCase().includes(lowerKeyword) ||
        task.Department.toLowerCase().includes(lowerKeyword)
      );
    } catch (error) {
      logger.error('TaskLibraryDataService', `Error searching tasks with keyword ${keyword}`, error);
      throw error;
    }
  }

  /**
   * Toggle task active status
   */
  public async toggleTaskActive(id: number, isActive: boolean): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('PM_Tasks').items
        .getById(id)
        .update({ IsActive: isActive });

      logger.info('TaskLibraryDataService', `Toggled task ${id} active status to ${isActive}`);
    } catch (error) {
      logger.error('TaskLibraryDataService', `Error toggling task ${id} active status`, error);
      throw error;
    }
  }

  /**
   * Get task usage statistics
   */
  public async getTaskStats(): Promise<{
    total: number;
    active: number;
    inactive: number;
    byCategory: { [key: string]: number };
    byDepartment: { [key: string]: number };
  }> {
    try {
      const allTasks = await this.getAllTasks();

      const stats = {
        total: allTasks.length,
        active: allTasks.filter(t => t.IsActive).length,
        inactive: allTasks.filter(t => !t.IsActive).length,
        byCategory: {} as { [key: string]: number },
        byDepartment: {} as { [key: string]: number }
      };

      // Count by category
      allTasks.forEach(task => {
        stats.byCategory[task.Category] = (stats.byCategory[task.Category] || 0) + 1;
        stats.byDepartment[task.Department] = (stats.byDepartment[task.Department] || 0) + 1;
      });

      return stats;
    } catch (error) {
      logger.error('TaskLibraryDataService', 'Error getting task stats', error);
      throw error;
    }
  }
}
