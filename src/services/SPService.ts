// @ts-nocheck
// SPService - SharePoint Data Access Layer
// Handles all SharePoint list operations using PnP JS

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/batching';
import '@pnp/sp/fields';
import '@pnp/sp/views';
import '@pnp/sp/site-users/web';

import {
  IJmlProcess,
  IJmlChecklistTemplate,
  IJmlTask,
  IJmlTaskAssignment,
  IJmlConfiguration,
  IJmlAuditLog,
  IJmlTemplateTaskMapping,
  TaskStatus
} from '../models';
import { ValidationUtils } from '../utils/ValidationUtils';
import { logger } from './LoggingService';

export class SPService {
  private sp: SPFI;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ===================================================================
  // PM_Processes Operations
  // ===================================================================

  /**
   * Get all processes
   */
  public async getProcesses(filter?: string, orderBy?: string, top?: number): Promise<IJmlProcess[]> {
    try {
      let query = this.sp.web.lists.getByTitle('PM_Processes').items
        .select(
          'Id', 'Title', 'ProcessType', 'ProcessStatus', 'EmployeeName', 'EmployeeEmail',
          'Department', 'JobTitle', 'Location', 'StartDate', 'TargetCompletionDate',
          'ActualCompletionDate', 'Priority', 'TotalTasks', 'CompletedTasks',
          'ProgressPercentage', 'Comments', 'OverdueTasks',
          'Manager/Id', 'Manager/Title', 'Manager/EMail',
          'ProcessOwner/Id', 'ProcessOwner/Title', 'ProcessOwner/EMail',
          'ChecklistTemplateID',
          'Created', 'Modified'
        )
        .expand('Manager', 'ProcessOwner');

      if (filter) {
        query = query.filter(filter);
      }

      if (orderBy) {
        query = query.orderBy(orderBy, false);
      }

      if (top) {
        query = query.top(top);
      }

      const items = await query();
      return items as IJmlProcess[];
    } catch (error) {
      logger.error('SPService', 'Error fetching processes', error);
      throw error;
    }
  }

  /**
   * Get process by ID
   */
  public async getProcessById(id: number): Promise<IJmlProcess> {
    try {
      const item = await this.sp.web.lists.getByTitle('PM_Processes').items
        .getById(id)
        .select(
          'Id', 'Title', 'ProcessType', 'ProcessStatus', 'EmployeeName', 'EmployeeEmail',
          'EmployeeID', 'Department', 'JobTitle', 'Location', 'StartDate',
          'TargetCompletionDate', 'ActualCompletionDate', 'Priority',
          'TotalTasks', 'CompletedTasks', 'ProgressPercentage', 'OverdueTasks',
          'Comments', 'BusinessUnit', 'CostCenter', 'ContractType',
          'Manager/Id', 'Manager/Title', 'Manager/EMail',
          'ProcessOwner/Id', 'ProcessOwner/Title', 'ProcessOwner/EMail',
          'ChecklistTemplateID',
          'Created', 'Modified'
        )
        .expand('Manager', 'ProcessOwner')();

      return item as IJmlProcess;
    } catch (error) {
      logger.error('SPService', `Error fetching process ${id}`, error);
      throw error;
    }
  }

  /**
   * Create new process
   */
  public async createProcess(process: Partial<IJmlProcess>): Promise<IJmlProcess> {
    try {
      const result = await this.sp.web.lists.getByTitle('PM_Processes').items.add({
        Title: process.Title,
        ProcessType: process.ProcessType,
        ProcessStatus: process.ProcessStatus || 'Draft',
        EmployeeName: process.EmployeeName,
        EmployeeEmail: process.EmployeeEmail,
        EmployeeID: process.EmployeeID,
        Department: process.Department,
        JobTitle: process.JobTitle,
        Location: process.Location,
        StartDate: process.StartDate,
        TargetCompletionDate: process.TargetCompletionDate,
        Priority: process.Priority,
        ManagerId: process.ManagerId,
        ProcessOwnerId: process.ProcessOwnerId,
        ChecklistTemplateID: process.ChecklistTemplateID,
        Comments: process.Comments,
        BusinessUnit: process.BusinessUnit,
        CostCenter: process.CostCenter,
        ContractType: process.ContractType,
        TotalTasks: process.TotalTasks || 0,
        CompletedTasks: process.CompletedTasks || 0,
        OverdueTasks: process.OverdueTasks || 0,
        ProgressPercentage: process.ProgressPercentage || 0
      });

      return await this.getProcessById(result.data.Id);
    } catch (error) {
      logger.error('SPService', 'Error creating process', error);
      throw error;
    }
  }

  /**
   * Update process
   */
  public async updateProcess(id: number, updates: Partial<IJmlProcess>): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('PM_Processes').items
        .getById(id)
        .update(updates);
    } catch (error) {
      logger.error('SPService', `Error updating process ${id}`, error);
      throw error;
    }
  }

  /**
   * Delete process
   */
  public async deleteProcess(id: number): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('PM_Processes').items
        .getById(id)
        .delete();
    } catch (error) {
      logger.error('SPService', `Error deleting process ${id}`, error);
      throw error;
    }
  }

  // ===================================================================
  // PM_ProcessChecklistTemplates Operations
  // ===================================================================
  // Note: This list is separate from PM_ChecklistTemplates which is used for task mapping relationships

  /**
   * Get all active templates for Process Wizard
   */
  public async getTemplates(processType?: string): Promise<IJmlChecklistTemplate[]> {
    try {
      // Build secure filter
      let filter = 'IsActive eq 1';
      if (processType) {
        const sanitizedType = ValidationUtils.sanitizeForOData(processType);
        const processTypeFilter = ValidationUtils.buildFilter('ProcessType', 'eq', sanitizedType);
        filter = `${filter} and ${processTypeFilter}`;
      }

      const items = await this.sp.web.lists.getByTitle('PM_ProcessChecklistTemplates').items
        .select('Id', 'Title', 'TemplateCode', 'ProcessType', 'Description',
                'JobRole', 'EstimatedDuration', 'TaskCount', 'IsActive', 'Version')
        .filter(filter)();

      return items as IJmlChecklistTemplate[];
    } catch (error) {
      logger.error('SPService', 'Error fetching templates', error);
      throw error;
    }
  }

  /**
   * Get template by ID
   */
  public async getTemplateById(id: number): Promise<IJmlChecklistTemplate> {
    try {
      const item = await this.sp.web.lists.getByTitle('PM_ProcessChecklistTemplates').items
        .getById(id)();
      return item as IJmlChecklistTemplate;
    } catch (error) {
      logger.error('SPService', `Error fetching template ${id}`, error);
      throw error;
    }
  }

  /**
   * Update template
   */
  public async updateTemplate(id: number, updates: Partial<IJmlChecklistTemplate>): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('PM_ProcessChecklistTemplates').items
        .getById(id)
        .update(updates);
    } catch (error) {
      logger.error('SPService', `Error updating template ${id}:`, error);
      throw error;
    }
  }

  /**
   * Create a new template
   */
  public async createTemplate(template: Partial<IJmlChecklistTemplate>): Promise<IJmlChecklistTemplate> {
    try {
      const result = await this.sp.web.lists.getByTitle('PM_ProcessChecklistTemplates').items
        .add(template);
      return result as unknown as IJmlChecklistTemplate;
    } catch (error) {
      logger.error('SPService', 'Error creating template:', error);
      throw error;
    }
  }

  // ===================================================================
  // PM_Tasks Operations
  // ===================================================================

  /**
   * Get all active tasks
   */
  public async getTasks(category?: string): Promise<IJmlTask[]> {
    try {
      // Build secure filter
      let filter = 'IsActive eq 1';
      if (category) {
        const sanitizedCategory = ValidationUtils.sanitizeForOData(category);
        const categoryFilter = ValidationUtils.buildFilter('Category', 'eq', sanitizedCategory);
        filter = `${filter} and ${categoryFilter}`;
      }

      const items = await this.sp.web.lists.getByTitle('PM_Tasks').items
        .select('Id', 'Title', 'TaskCode', 'Category', 'Description', 'Instructions',
                'Department', 'SLAHours', 'EstimatedHours', 'RequiresApproval',
                'Priority', 'IsActive', 'DefaultAssignee/Id', 'DefaultAssignee/Title')
        .expand('DefaultAssignee')
        .filter(filter)();

      return items as IJmlTask[];
    } catch (error) {
      logger.error('SPService', 'Error fetching tasks:', error);
      throw error;
    }
  }

  // ===================================================================
  // PM_TaskAssignments Operations
  // ===================================================================

  /**
   * Get task assignments for a process
   */
  public async getTaskAssignmentsByProcess(processId: number): Promise<IJmlTaskAssignment[]> {
    try {
      // Validate process ID
      const validProcessId = ValidationUtils.validateInteger(processId, 'processId', 1);

      // Note: ProcessID might be Text or Lookup depending on whether Script 09 has been run
      // For now, query as Text field until Script 09 is executed
      const filter = `ProcessID eq '${validProcessId}'`;

      const items = await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
        .select(
          'Id', 'Title', 'Status', 'Priority', 'DueDate', 'StartDate', 'ActualCompletionDate',
          'PercentComplete', 'ActualHours', 'Notes', 'RequiresApproval', 'BlockedReason',
          'ProcessID', 'TaskID', 'TaskCode',
          'AssignedTo/Id', 'AssignedTo/Title', 'AssignedTo/EMail',
          'CompletedBy/Id', 'CompletedBy/Title', 'CompletedBy/EMail'
        )
        .expand('AssignedTo', 'CompletedBy')
        .filter(filter)
        .orderBy('DueDate', true)();

      return items as IJmlTaskAssignment[];
    } catch (error) {
      logger.error('SPService', `Error fetching task assignments for process ${processId}:`, error);
      throw error;
    }
  }

  /**
   * Get my tasks (current user)
   */
  public async getMyTasks(userId: number): Promise<IJmlTaskAssignment[]> {
    try {
      // Validate user ID
      const validUserId = ValidationUtils.validateInteger(userId, 'userId', 1);

      // Build filter for Person field - AssignedToId is the lookup ID suffix
      const filter = `AssignedToId eq ${validUserId} and Status ne 'Completed' and Status ne 'Cancelled'`;
      console.log('[SPService.getMyTasks] Filter:', filter);

      // Query essential columns - simplified to avoid User field expansion issues
      // Note: Removed AssignedTo expansion which can cause "Invalid Request" in some SharePoint tenants
      // Note: Removed Department and Category which may not exist on all list configurations
      const items = await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
        .select(
          'Id', 'Title', 'Status', 'Priority', 'DueDate', 'PercentComplete',
          'Notes', 'BlockedReason', 'AssignedToId', 'ProcessID', 'AssignedDate'
        )
        .filter(filter)
        .orderBy('DueDate', true)
        .top(100)();

      console.log('[SPService.getMyTasks] Found', items.length, 'tasks');
      return items as IJmlTaskAssignment[];
    } catch (error) {
      console.error('[SPService.getMyTasks] ERROR:', error);
      logger.error('SPService', `Error fetching tasks for user ${userId}:`, error);
      // Re-throw so UI can display the error
      throw error;
    }
  }

  /**
   * Create task assignment
   */
  public async createTaskAssignment(assignment: Partial<IJmlTaskAssignment>): Promise<IJmlTaskAssignment> {
    try {
      const result = await this.sp.web.lists.getByTitle('PM_TaskAssignments').items.add({
        Title: assignment.Title,
        ProcessIDId: assignment.ProcessIDId,
        TaskIDId: assignment.TaskIDId,
        AssignedToId: assignment.AssignedToId,
        DueDate: assignment.DueDate,
        Status: assignment.Status || 'Not Started',
        Priority: assignment.Priority,
        RequiresApproval: assignment.RequiresApproval,
        Notes: assignment.Notes
      });

      const item = await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
        .getById(result.data.Id)();
      return item as IJmlTaskAssignment;
    } catch (error) {
      logger.error('SPService', 'Error creating task assignment:', error);
      throw error;
    }
  }

  /**
   * Update task assignment
   */
  public async updateTaskAssignment(id: number, updates: Partial<IJmlTaskAssignment>): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
        .getById(id)
        .update(updates);
    } catch (error) {
      logger.error('SPService', `Error updating task assignment ${id}:`, error);
      throw error;
    }
  }

  /**
   * Update task status
   */
  public async updateTaskStatus(id: number, status: TaskStatus): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
        .getById(id)
        .update({ Status: status });
    } catch (error) {
      logger.error('SPService', `Error updating task status ${id}:`, error);
      throw error;
    }
  }

  // ===================================================================
  // PM_Configuration Operations
  // ===================================================================

  /**
   * Get configuration value by key
   */
  public async getConfigValue(key: string): Promise<string | null> {
    try {
      // Sanitize config key
      if (!key || typeof key !== 'string') {
        return null;
      }
      const sanitizedKey = ValidationUtils.sanitizeForOData(key.substring(0, 100));

      // Build secure filter
      const filter = `ConfigKey eq '${sanitizedKey}' and IsActive eq 1`;

      const items = await this.sp.web.lists.getByTitle('PM_Configuration').items
        .select('ConfigValue')
        .filter(filter)
        .top(1)();

      return items.length > 0 ? items[0].ConfigValue : null;
    } catch (error) {
      logger.error('SPService', `Error fetching config ${key}:`, error);
      return null;
    }
  }

  /**
   * Set a configuration value by key (upsert)
   */
  public async setConfigValue(key: string, value: string, category?: string): Promise<void> {
    try {
      if (!key || typeof key !== 'string') return;
      const sanitizedKey = ValidationUtils.sanitizeForOData(key.substring(0, 100));
      const filter = `ConfigKey eq '${sanitizedKey}'`;

      const items = await this.sp.web.lists.getByTitle('PM_Configuration').items
        .select('Id', 'ConfigKey')
        .filter(filter)
        .top(1)();

      if (items.length > 0) {
        // Update existing
        await this.sp.web.lists.getByTitle('PM_Configuration').items.getById(items[0].Id).update({
          ConfigValue: value,
          IsActive: true
        });
      } else {
        // Create new
        await this.sp.web.lists.getByTitle('PM_Configuration').items.add({
          Title: key,
          ConfigKey: key,
          ConfigValue: value,
          Category: category || 'Integration',
          IsActive: true,
          IsSystemConfig: true
        });
      }
    } catch (error) {
      logger.error('SPService', `Error setting config ${key}:`, error);
      throw error;
    }
  }

  /**
   * Get all active configurations
   */
  public async getAllConfigurations(): Promise<IJmlConfiguration[]> {
    try {
      const items = await this.sp.web.lists.getByTitle('PM_Configuration').items
        .select('Id', 'Title', 'ConfigKey', 'ConfigValue', 'Category', 'Description', 'IsActive')
        .filter('IsActive eq 1')();

      return items as IJmlConfiguration[];
    } catch (error) {
      logger.error('SPService', 'Error fetching configurations:', error);
      throw error;
    }
  }

  // ===================================================================
  // PM_AuditLog Operations
  // ===================================================================

  /**
   * Create audit log entry
   */
  public async createAuditLog(log: Partial<IJmlAuditLog>): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('PM_AuditLog').items.add({
        Title: log.Title || log.EventType,
        EventType: log.EventType,
        EntityType: log.EntityType,
        EntityId: log.EntityId,
        UserId: log.UserId,
        Action: log.Action,
        FieldChanged: log.FieldChanged,
        OldValue: log.OldValue,
        NewValue: log.NewValue,
        ProcessId: log.ProcessId,
        TaskId: log.TaskId,
        Description: log.Description,
        Timestamp: new Date(),
        AdditionalData: log.AdditionalData
      });
    } catch (error) {
      logger.error('SPService', 'Error creating audit log:', error);
      // Don't throw - audit logging shouldn't break the app
    }
  }

  // ===================================================================
  // PM_TemplateTaskMapping Operations
  // ===================================================================

  /**
   * Get tasks for a template
   */
  public async getTemplateTaskMappings(templateId: number): Promise<IJmlTemplateTaskMapping[]> {
    try {
      // Validate template ID
      const validTemplateId = ValidationUtils.validateInteger(templateId, 'templateId', 1);

      // Build secure filter
      const templateFilter = ValidationUtils.buildFilter('TemplateID/Id', 'eq', validTemplateId);
      const filter = `${templateFilter} and IsActive eq 1`;

      const items = await this.sp.web.lists.getByTitle('PM_TemplateTaskMapping').items
        .select(
          'Id', 'Title', 'SequenceOrder', 'IsMandatory', 'OffsetDays', 'DependsOnTaskID',
          'CustomInstructions', 'OverrideSLAHours', 'IsActive',
          'TemplateID/Id', 'TemplateID/Title',
          'TaskID/Id', 'TaskID/Title', 'TaskID/TaskCode', 'TaskID/Category',
          'TaskID/EstimatedHours', 'TaskID/SLAHours',
          'OverrideAssignee/Id', 'OverrideAssignee/Title'
        )
        .expand('TemplateID', 'TaskID', 'OverrideAssignee')
        .filter(filter)
        .orderBy('SequenceOrder', true)();

      return items as IJmlTemplateTaskMapping[];
    } catch (error) {
      logger.error('SPService', `Error fetching template mappings for template ${templateId}:`, error);
      throw error;
    }
  }

  // ===================================================================
  // Utility Operations
  // ===================================================================

  /**
   * Get current user
   */
  public async getCurrentUser(): Promise<any> {
    try {
      const user = await this.sp.web.currentUser();
      return user;
    } catch (error) {
      logger.error('SPService', 'Error fetching current user:', error);
      throw error;
    }
  }

  /**
   * Search users
   */
  public async searchUsers(searchTerm: string): Promise<any[]> {
    try {
      // Sanitize search term for OData substringof
      if (!searchTerm || typeof searchTerm !== 'string') {
        return [];
      }
      const sanitizedTerm = ValidationUtils.sanitizeForOData(searchTerm.substring(0, 100));

      const users = await this.sp.web.siteUsers
        .filter(`substringof('${sanitizedTerm}', Title) or substringof('${sanitizedTerm}', Email')`)
        .top(10)();
      return users;
    } catch (error) {
      logger.error('SPService', 'Error searching users:', error);
      throw error;
    }
  }

  // ===================================================================
  // Notification Operations
  // ===================================================================

  /**
   * Create a notification for task assignment
   */
  public async createTaskAssignmentNotification(
    taskId: number,
    assigneeId: number,
    processTitle: string,
    taskTitle: string
  ): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('PM_Notifications').items.add({
        Title: `New task assigned: ${taskTitle}`,
        NotificationType: 'TaskAssigned',
        MessageBody: `You have been assigned a task "${taskTitle}" in process "${processTitle}".`,
        Priority: 'Normal',
        RecipientId: assigneeId,
        TaskId: taskId.toString(),
        Status: 'Pending'
      });
    } catch (error) {
      logger.error('SPService', 'Error creating task assignment notification:', error);
      // Don't throw - notifications are not critical
    }
  }

  /**
   * Create a notification for task completion
   */
  public async createTaskCompletionNotification(
    taskId: number,
    managerId: number,
    processTitle: string,
    taskTitle: string,
    completedBy: string
  ): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('PM_Notifications').items.add({
        Title: `Task completed: ${taskTitle}`,
        NotificationType: 'TaskCompleted',
        MessageBody: `${completedBy} completed task "${taskTitle}" in process "${processTitle}".`,
        Priority: 'Normal',
        RecipientId: managerId,
        TaskId: taskId.toString(),
        Status: 'Pending'
      });
    } catch (error) {
      logger.error('SPService', 'Error creating task completion notification:', error);
      // Don't throw - notifications are not critical
    }
  }

  /**
   * Create a notification for process completion
   */
  public async createProcessCompletionNotification(
    processId: number,
    managerId: number,
    processTitle: string
  ): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('PM_Notifications').items.add({
        Title: `Process completed: ${processTitle}`,
        NotificationType: 'ProcessCompleted',
        MessageBody: `The process "${processTitle}" has been completed.`,
        Priority: 'Normal',
        RecipientId: managerId,
        ProcessId: processId.toString(),
        Status: 'Pending'
      });
    } catch (error) {
      logger.error('SPService', 'Error creating process completion notification:', error);
      // Don't throw - notifications are not critical
    }
  }

  /**
   * Create a notification for due date approaching
   */
  public async createDueDateNotification(
    taskId: number,
    assigneeId: number,
    taskTitle: string,
    dueDate: Date
  ): Promise<void> {
    try {
      const daysUntilDue = Math.ceil((dueDate.getTime() - new Date().getTime()) / (1000 * 60 * 60 * 24));
      await this.sp.web.lists.getByTitle('PM_Notifications').items.add({
        Title: `Due date approaching: ${taskTitle}`,
        NotificationType: 'DueDateApproaching',
        MessageBody: `Task "${taskTitle}" is due in ${daysUntilDue} day(s).`,
        Priority: daysUntilDue <= 1 ? 'High' : 'Normal',
        RecipientId: assigneeId,
        TaskId: taskId.toString(),
        Status: 'Pending'
      });
    } catch (error) {
      logger.error('SPService', 'Error creating due date notification:', error);
      // Don't throw - notifications are not critical
    }
  }

  /**
   * Get notifications for user
   */
  public async getNotificationsForUser(userId: number, unreadOnly: boolean = false): Promise<any[]> {
    try {
      // Validate user ID
      const validUserId = ValidationUtils.validateInteger(userId, 'userId', 1);

      // Build secure filter
      const userFilter = ValidationUtils.buildFilter('RecipientId', 'eq', validUserId);
      let filter = userFilter;

      if (unreadOnly) {
        filter = `${filter} and Status eq 'Pending'`;
      }

      const items = await this.sp.web.lists.getByTitle('PM_Notifications').items
        .select('Id', 'Title', 'NotificationType', 'MessageBody', 'Priority', 'Status', 'Created', 'ProcessId', 'TaskId')
        .filter(filter)
        .orderBy('Created', false)
        .top(50)();

      return items;
    } catch (error) {
      logger.error('SPService', 'Error fetching notifications:', error);
      return [];
    }
  }

  /**
   * Mark notification as read (LEGACY - keeping for reference)
   */
  private async markNotificationAsReadLegacy(userId: number, unreadOnly: boolean): Promise<any[]> {
    try {
      let query = this.sp.web.lists.getByTitle('PM_Notifications').items
        .select('Id', 'Title', 'NotificationType', 'MessageBody', 'Priority', 'Status', 'Created', 'ProcessId', 'TaskId')
        .filter(`RecipientId eq ${userId}`)
        .orderBy('Created', false)
        .top(50);

      if (unreadOnly) {
        query = query.filter(`Status eq 'Pending'`);
      }

      return await query();
    } catch (error) {
      logger.error('SPService', 'Error getting notifications:', error);
      throw error;
    }
  }

  /**
   * Mark notification as read
   */
  public async markNotificationAsRead(notificationId: number): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('PM_Notifications').items
        .getById(notificationId)
        .update({
          Status: 'Sent'
        });
    } catch (error) {
      logger.error('SPService', 'Error marking notification as read:', error);
      throw error;
    }
  }

  // ============================================
  // Filter Preset Management Methods
  // ============================================

  /**
   * Create or update a filter preset
   */
  public async saveFilterPreset(preset: any): Promise<void> {
    try {
      // Validate and sanitize preset ID
      if (!preset.id || typeof preset.id !== 'string') {
        throw new Error('Invalid preset ID');
      }

      const presetData = {
        Title: preset.title,
        Description: preset.description || '',
        FilterData: JSON.stringify(preset.filters),
        UserId: preset.userId,
        IsDefault: preset.isDefault || false,
        IsShared: preset.isShared || false,
        UseCount: preset.useCount || 0
      };

      // Check if preset exists with secure filter
      const filter = ValidationUtils.buildFilter('PresetId', 'eq', preset.id.substring(0, 100));
      const existingPresets = await this.sp.web.lists.getByTitle('PM_FilterPresets').items
        .filter(filter)
        .top(1)();

      if (existingPresets.length > 0) {
        // Update existing preset
        await this.sp.web.lists.getByTitle('PM_FilterPresets').items
          .getById(existingPresets[0].Id)
          .update(presetData);
      } else {
        // Create new preset
        await this.sp.web.lists.getByTitle('PM_FilterPresets').items.add({
          ...presetData,
          PresetId: preset.id.substring(0, 100)
        });
      }
    } catch (error) {
      logger.error('SPService', 'Error saving filter preset:', error);
      throw error;
    }
  }

  /**
   * Get filter presets for a user
   */
  public async getFilterPresets(userId: number, includeShared: boolean = true): Promise<any[]> {
    try {
      // Check if list exists first
      const lists = await this.sp.web.lists();
      const filterPresetsList = lists.find(l => l.Title === 'PM_FilterPresets');

      if (!filterPresetsList) {
        logger.warn('SPService', 'PM_FilterPresets list does not exist. Returning empty array.');
        return [];
      }

      // Validate user ID
      const validUserId = ValidationUtils.validateInteger(userId, 'userId', 1);

      // Build secure filter
      const userFilter = ValidationUtils.buildFilter('UserId', 'eq', validUserId);
      let filter = userFilter;
      if (includeShared) {
        filter = `(${userFilter}) or (IsShared eq 1)`;
      }

      const items = await this.sp.web.lists.getByTitle('PM_FilterPresets').items
        .filter(filter)
        .orderBy('IsDefault', false)
        .orderBy('Modified', false)();

      return items.map(item => ({
        id: item.PresetId,
        title: item.Title,
        description: item.Description,
        filters: JSON.parse(item.FilterData || '{}'),
        userId: item.UserId,
        isDefault: item.IsDefault,
        isShared: item.IsShared,
        createdDate: new Date(item.Created),
        modifiedDate: new Date(item.Modified),
        useCount: item.UseCount || 0
      }));
    } catch (error) {
      logger.error('SPService', 'Error getting filter presets:', error);
      return [];
    }
  }

  /**
   * Delete a filter preset
   */
  public async deleteFilterPreset(presetId: string): Promise<void> {
    try {
      // Validate and sanitize preset ID
      if (!presetId || typeof presetId !== 'string') {
        throw new Error('Invalid preset ID');
      }

      // Build secure filter
      const filter = ValidationUtils.buildFilter('PresetId', 'eq', presetId.substring(0, 100));

      const items = await this.sp.web.lists.getByTitle('PM_FilterPresets').items
        .filter(filter)
        .top(1)();

      if (items.length > 0) {
        await this.sp.web.lists.getByTitle('PM_FilterPresets').items
          .getById(items[0].Id)
          .delete();
      }
    } catch (error) {
      logger.error('SPService', 'Error deleting filter preset:', error);
      throw error;
    }
  }

  /**
   * Set a preset as default
   */
  public async setDefaultPreset(userId: number, presetId: string): Promise<void> {
    try {
      // Validate inputs
      const validUserId = ValidationUtils.validateInteger(userId, 'userId', 1);
      if (!presetId || typeof presetId !== 'string') {
        throw new Error('Invalid preset ID');
      }

      // First, unset all defaults for this user with secure filter
      const userFilter = ValidationUtils.buildFilter('UserId', 'eq', validUserId);
      const allPresets = await this.sp.web.lists.getByTitle('PM_FilterPresets').items
        .filter(`${userFilter} and IsDefault eq 1`)();

      for (let i = 0; i < allPresets.length; i++) {
        await this.sp.web.lists.getByTitle('PM_FilterPresets').items
          .getById(allPresets[i].Id)
          .update({ IsDefault: false });
      }

      // Set the new default with secure filter
      const presetFilter = ValidationUtils.buildFilter('PresetId', 'eq', presetId.substring(0, 100));
      const targetPresets = await this.sp.web.lists.getByTitle('PM_FilterPresets').items
        .filter(presetFilter)
        .top(1)();

      if (targetPresets.length > 0) {
        await this.sp.web.lists.getByTitle('PM_FilterPresets').items
          .getById(targetPresets[0].Id)
          .update({ IsDefault: true });
      }
    } catch (error) {
      logger.error('SPService', 'Error setting default preset:', error);
      throw error;
    }
  }

  /**
   * Increment preset use count
   */
  public async incrementPresetUseCount(presetId: string): Promise<void> {
    try {
      // Validate and sanitize preset ID
      if (!presetId || typeof presetId !== 'string') {
        throw new Error('Invalid preset ID');
      }

      // Build secure filter
      const filter = ValidationUtils.buildFilter('PresetId', 'eq', presetId.substring(0, 100));

      const items = await this.sp.web.lists.getByTitle('PM_FilterPresets').items
        .filter(filter)
        .select('Id', 'UseCount')
        .top(1)();

      if (items.length > 0) {
        const currentCount = items[0].UseCount || 0;
        await this.sp.web.lists.getByTitle('PM_FilterPresets').items
          .getById(items[0].Id)
          .update({ UseCount: currentCount + 1 });
      }
    } catch (error) {
      logger.error('SPService', 'Error incrementing preset use count:', error);
    }
  }

  // ===================================================================
  // PM_Assets Operations
  // ===================================================================

  /**
   * Get all assets
   */
  public async getAssets(filter?: string, orderBy?: string, top?: number): Promise<any[]> {
    try {
      let query = this.sp.web.lists.getByTitle('PM_Assets').items
        .select(
          'Id', 'Title', 'AssetTag', 'AssetType', 'Manufacturer', 'Model',
          'SerialNumber', 'PurchaseDate', 'PurchaseCost', 'WarrantyExpiry',
          'AssetStatus', 'Department', 'Location', 'Condition',
          'Specifications', 'Notes', 'LastServiceDate', 'NextServiceDate',
          'Supplier', 'AssetValue', 'IsActive',
          'CurrentOwner/Id', 'CurrentOwner/Title', 'CurrentOwner/EMail',
          'Created', 'Modified'
        )
        .expand('CurrentOwner');

      if (filter) {
        query = query.filter(filter);
      }

      if (orderBy) {
        query = query.orderBy(orderBy, false);
      }

      if (top) {
        query = query.top(top);
      }

      const items = await query();
      return items;
    } catch (error) {
      logger.error('SPService', 'Error fetching assets:', error);
      throw error;
    }
  }

  /**
   * Get asset by ID
   */
  public async getAssetById(id: number): Promise<any> {
    try {
      const item = await this.sp.web.lists.getByTitle('PM_Assets').items
        .getById(id)
        .select(
          'Id', 'Title', 'AssetTag', 'AssetType', 'Manufacturer', 'Model',
          'SerialNumber', 'PurchaseDate', 'PurchaseCost', 'WarrantyExpiry',
          'AssetStatus', 'Department', 'Location', 'Condition',
          'Specifications', 'Notes', 'LastServiceDate', 'NextServiceDate',
          'Supplier', 'AssetValue', 'IsActive',
          'CurrentOwner/Id', 'CurrentOwner/Title', 'CurrentOwner/EMail',
          'Created', 'Modified'
        )
        .expand('CurrentOwner')();

      return item;
    } catch (error) {
      logger.error('SPService', `Error fetching asset ${id}:`, error);
      throw error;
    }
  }

  /**
   * Get asset by asset tag
   */
  public async getAssetByTag(assetTag: string): Promise<any> {
    try {
      const sanitizedTag = ValidationUtils.sanitizeForOData(assetTag);
      const filter = ValidationUtils.buildFilter('AssetTag', 'eq', sanitizedTag);

      const items = await this.sp.web.lists.getByTitle('PM_Assets').items
        .select(
          'Id', 'Title', 'AssetTag', 'AssetType', 'AssetStatus',
          'CurrentOwner/Id', 'CurrentOwner/Title', 'CurrentOwner/EMail'
        )
        .expand('CurrentOwner')
        .filter(filter)
        .top(1)();

      return items.length > 0 ? items[0] : null;
    } catch (error) {
      logger.error('SPService', `Error fetching asset by tag ${assetTag}:`, error);
      throw error;
    }
  }

  /**
   * Create new asset
   */
  public async createAsset(asset: any): Promise<any> {
    try {
      const result = await this.sp.web.lists.getByTitle('PM_Assets').items.add({
        Title: asset.Title,
        AssetTag: asset.AssetTag,
        AssetType: asset.AssetType,
        Manufacturer: asset.Manufacturer,
        Model: asset.Model,
        SerialNumber: asset.SerialNumber,
        PurchaseDate: asset.PurchaseDate,
        PurchaseCost: asset.PurchaseCost,
        WarrantyExpiry: asset.WarrantyExpiry,
        AssetStatus: asset.AssetStatus || 'Available',
        CurrentOwnerId: asset.CurrentOwnerId,
        Department: asset.Department,
        Location: asset.Location,
        Condition: asset.Condition || 'Good',
        Specifications: asset.Specifications,
        Notes: asset.Notes,
        LastServiceDate: asset.LastServiceDate,
        NextServiceDate: asset.NextServiceDate,
        Supplier: asset.Supplier,
        AssetValue: asset.AssetValue,
        IsActive: asset.IsActive !== undefined ? asset.IsActive : true
      });

      return await this.getAssetById(result.data.Id);
    } catch (error) {
      logger.error('SPService', 'Error creating asset:', error);
      throw error;
    }
  }

  /**
   * Update asset
   */
  public async updateAsset(id: number, updates: any): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('PM_Assets').items
        .getById(id)
        .update(updates);
    } catch (error) {
      logger.error('SPService', `Error updating asset ${id}:`, error);
      throw error;
    }
  }

  /**
   * Delete asset (soft delete - set IsActive to false)
   */
  public async deleteAsset(id: number): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('PM_Assets').items
        .getById(id)
        .update({ IsActive: false });
    } catch (error) {
      logger.error('SPService', `Error deleting asset ${id}:`, error);
      throw error;
    }
  }

  // ===================================================================
  // PM_AssetCheckouts Operations
  // ===================================================================

  /**
   * Get all asset checkouts
   */
  public async getAssetCheckouts(filter?: string, orderBy?: string, top?: number): Promise<any[]> {
    try {
      let query = this.sp.web.lists.getByTitle('PM_AssetCheckouts').items
        .select(
          'Id', 'Title', 'AssetTag', 'EmployeeName', 'EmployeeEmail',
          'CheckedOutDate', 'ExpectedReturnDate', 'ActualReturnDate',
          'CheckoutReason', 'ReturnCondition', 'CheckoutStatus',
          'Department', 'ProcessID', 'Notes',
          'CheckedOutBy/Id', 'CheckedOutBy/Title', 'CheckedOutBy/EMail',
          'ReturnedBy/Id', 'ReturnedBy/Title', 'ReturnedBy/EMail',
          'Created', 'Modified'
        )
        .expand('CheckedOutBy', 'ReturnedBy');

      if (filter) {
        query = query.filter(filter);
      }

      if (orderBy) {
        query = query.orderBy(orderBy, false);
      } else {
        query = query.orderBy('CheckedOutDate', false);
      }

      if (top) {
        query = query.top(top);
      }

      const items = await query();
      return items;
    } catch (error) {
      logger.error('SPService', 'Error fetching asset checkouts:', error);
      throw error;
    }
  }

  /**
   * Get checkout by ID
   */
  public async getAssetCheckoutById(id: number): Promise<any> {
    try {
      const item = await this.sp.web.lists.getByTitle('PM_AssetCheckouts').items
        .getById(id)
        .select(
          'Id', 'Title', 'AssetTag', 'EmployeeName', 'EmployeeEmail',
          'CheckedOutDate', 'ExpectedReturnDate', 'ActualReturnDate',
          'CheckoutReason', 'ReturnCondition', 'CheckoutStatus',
          'Department', 'ProcessID', 'Notes',
          'CheckedOutBy/Id', 'CheckedOutBy/Title', 'CheckedOutBy/EMail',
          'ReturnedBy/Id', 'ReturnedBy/Title', 'ReturnedBy/EMail',
          'Created', 'Modified'
        )
        .expand('CheckedOutBy', 'ReturnedBy')();

      return item;
    } catch (error) {
      logger.error('SPService', `Error fetching asset checkout ${id}:`, error);
      throw error;
    }
  }

  /**
   * Get active checkouts for an asset
   */
  public async getActiveCheckoutsForAsset(assetTag: string): Promise<any[]> {
    try {
      const sanitizedTag = ValidationUtils.sanitizeForOData(assetTag);
      const filter = `AssetTag eq '${sanitizedTag}' and CheckoutStatus eq 'Active'`;

      const items = await this.sp.web.lists.getByTitle('PM_AssetCheckouts').items
        .select(
          'Id', 'Title', 'EmployeeName', 'CheckedOutDate', 'ExpectedReturnDate',
          'CheckedOutBy/Title'
        )
        .expand('CheckedOutBy')
        .filter(filter)();

      return items;
    } catch (error) {
      logger.error('SPService', `Error fetching active checkouts for asset ${assetTag}:`, error);
      throw error;
    }
  }

  /**
   * Create asset checkout
   */
  public async createAssetCheckout(checkout: any): Promise<any> {
    try {
      const result = await this.sp.web.lists.getByTitle('PM_AssetCheckouts').items.add({
        Title: checkout.Title,
        AssetTag: checkout.AssetTag,
        EmployeeName: checkout.EmployeeName,
        EmployeeEmail: checkout.EmployeeEmail,
        CheckedOutById: checkout.CheckedOutById,
        CheckedOutDate: checkout.CheckedOutDate || new Date(),
        ExpectedReturnDate: checkout.ExpectedReturnDate,
        CheckoutReason: checkout.CheckoutReason,
        CheckoutStatus: checkout.CheckoutStatus || 'Active',
        Department: checkout.Department,
        ProcessID: checkout.ProcessID,
        Notes: checkout.Notes
      });

      // Update asset status to "Checked Out"
      if (checkout.AssetId) {
        await this.updateAsset(checkout.AssetId, {
          AssetStatus: 'Checked Out',
          CurrentOwnerId: checkout.CheckedOutById
        });
      }

      return await this.getAssetCheckoutById(result.data.Id);
    } catch (error) {
      logger.error('SPService', 'Error creating asset checkout:', error);
      throw error;
    }
  }

  /**
   * Return asset (check in)
   */
  public async returnAsset(checkoutId: number, returnData: any): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('PM_AssetCheckouts').items
        .getById(checkoutId)
        .update({
          ActualReturnDate: returnData.ActualReturnDate || new Date(),
          ReturnCondition: returnData.ReturnCondition,
          CheckoutStatus: 'Returned',
          ReturnedById: returnData.ReturnedById,
          Notes: returnData.Notes
        });

      // Update asset status back to "Available"
      if (returnData.AssetId) {
        await this.updateAsset(returnData.AssetId, {
          AssetStatus: 'Available',
          CurrentOwnerId: null,
          Condition: returnData.ReturnCondition
        });
      }
    } catch (error) {
      logger.error('SPService', `Error returning asset ${checkoutId}:`, error);
      throw error;
    }
  }

  /**
   * Update asset checkout
   */
  public async updateAssetCheckout(id: number, updates: any): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('PM_AssetCheckouts').items
        .getById(id)
        .update(updates);
    } catch (error) {
      logger.error('SPService', `Error updating asset checkout ${id}:`, error);
      throw error;
    }
  }

  // ===================================================================
  // PM_M365Licenses Operations
  // ===================================================================

  /**
   * Get all M365 licenses
   */
  public async getM365Licenses(filter?: string, orderBy?: string, top?: number): Promise<any[]> {
    try {
      let query = this.sp.web.lists.getByTitle('PM_M365Licenses').items
        .select(
          'Id', 'Title', 'LicenseType', 'EmployeeName', 'EmployeeEmail',
          'AssignmentDate', 'RemovalDate', 'CostPerMonth', 'AnnualCost',
          'Department', 'LicenseStatus', 'ProcessID', 'SubscriptionID',
          'Notes',
          'AssignedTo/Id', 'AssignedTo/Title', 'AssignedTo/EMail',
          'Created', 'Modified'
        )
        .expand('AssignedTo');

      if (filter) {
        query = query.filter(filter);
      }

      if (orderBy) {
        query = query.orderBy(orderBy, false);
      } else {
        query = query.orderBy('AssignmentDate', false);
      }

      if (top) {
        query = query.top(top);
      }

      const items = await query();
      return items;
    } catch (error) {
      logger.error('SPService', 'Error fetching M365 licenses:', error);
      throw error;
    }
  }

  /**
   * Get M365 license by ID
   */
  public async getM365LicenseById(id: number): Promise<any> {
    try {
      const item = await this.sp.web.lists.getByTitle('PM_M365Licenses').items
        .getById(id)
        .select(
          'Id', 'Title', 'LicenseType', 'EmployeeName', 'EmployeeEmail',
          'AssignmentDate', 'RemovalDate', 'CostPerMonth', 'AnnualCost',
          'Department', 'LicenseStatus', 'ProcessID', 'SubscriptionID',
          'Notes',
          'AssignedTo/Id', 'AssignedTo/Title', 'AssignedTo/EMail',
          'Created', 'Modified'
        )
        .expand('AssignedTo')();

      return item;
    } catch (error) {
      logger.error('SPService', `Error fetching M365 license ${id}:`, error);
      throw error;
    }
  }

  /**
   * Get active licenses for a user
   */
  public async getActiveLicensesForUser(userId: number): Promise<any[]> {
    try {
      const validUserId = ValidationUtils.validateInteger(userId, 'userId', 1);
      const filter = `AssignedToId eq ${validUserId} and LicenseStatus eq 'Active'`;

      const items = await this.sp.web.lists.getByTitle('PM_M365Licenses').items
        .select(
          'Id', 'Title', 'LicenseType', 'AssignmentDate', 'CostPerMonth',
          'AssignedTo/Title'
        )
        .expand('AssignedTo')
        .filter(filter)();

      return items;
    } catch (error) {
      logger.error('SPService', `Error fetching active licenses for user ${userId}:`, error);
      throw error;
    }
  }

  /**
   * Assign M365 license
   */
  public async assignM365License(license: any): Promise<any> {
    try {
      const result = await this.sp.web.lists.getByTitle('PM_M365Licenses').items.add({
        Title: license.Title,
        LicenseType: license.LicenseType,
        EmployeeName: license.EmployeeName,
        EmployeeEmail: license.EmployeeEmail,
        AssignedToId: license.AssignedToId,
        AssignmentDate: license.AssignmentDate || new Date(),
        CostPerMonth: license.CostPerMonth,
        AnnualCost: license.AnnualCost,
        Department: license.Department,
        LicenseStatus: license.LicenseStatus || 'Active',
        ProcessID: license.ProcessID,
        SubscriptionID: license.SubscriptionID,
        Notes: license.Notes
      });

      return await this.getM365LicenseById(result.data.Id);
    } catch (error) {
      logger.error('SPService', 'Error assigning M365 license:', error);
      throw error;
    }
  }

  /**
   * Remove M365 license
   */
  public async removeM365License(licenseId: number, removalNotes?: string): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('PM_M365Licenses').items
        .getById(licenseId)
        .update({
          RemovalDate: new Date(),
          LicenseStatus: 'Removed',
          Notes: removalNotes
        });
    } catch (error) {
      logger.error('SPService', `Error removing M365 license ${licenseId}:`, error);
      throw error;
    }
  }

  /**
   * Update M365 license
   */
  public async updateM365License(id: number, updates: any): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('PM_M365Licenses').items
        .getById(id)
        .update(updates);
    } catch (error) {
      logger.error('SPService', `Error updating M365 license ${id}:`, error);
      throw error;
    }
  }

  /**
   * Get license cost summary by type
   */
  public async getLicenseCostSummary(): Promise<any[]> {
    try {
      const items = await this.sp.web.lists.getByTitle('PM_M365Licenses').items
        .select('LicenseType', 'LicenseStatus', 'CostPerMonth', 'AnnualCost')
        .filter('LicenseStatus eq \'Active\'')();

      // Group by license type and calculate totals
      const summary = items.reduce((acc: any, item: any) => {
        const type = item.LicenseType || 'Unknown';
        if (!acc[type]) {
          acc[type] = {
            LicenseType: type,
            Count: 0,
            TotalMonthlyCost: 0,
            TotalAnnualCost: 0
          };
        }
        acc[type].Count++;
        acc[type].TotalMonthlyCost += item.CostPerMonth || 0;
        acc[type].TotalAnnualCost += item.AnnualCost || 0;
        return acc;
      }, {});

      return Object.values(summary);
    } catch (error) {
      logger.error('SPService', 'Error fetching license cost summary:', error);
      throw error;
    }
  }
}
