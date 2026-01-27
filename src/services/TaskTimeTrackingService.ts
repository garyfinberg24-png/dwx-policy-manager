// @ts-nocheck
// TaskTimeTrackingService - Handles time tracking for task assignments
// Provides start/stop timer, manual entry, and time summary functionality

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import {
  IJmlTaskTimeEntry,
  ITaskTimeEntryForm,
  ITaskTimeEntryView,
  ITaskTimeSummary,
  WorkType
} from '../models';
import { logger } from './LoggingService';

export class TaskTimeTrackingService {
  private sp: SPFI;
  private listTitle = 'JML_TaskTimeEntries';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Get all time entries for a task
   */
  public async getTaskTimeEntries(taskAssignmentId: number, currentUserId: number): Promise<ITaskTimeEntryView[]> {
    try {
      logger.info('TaskTimeTrackingService', `Fetching time entries for task ${taskAssignmentId}`);

      const items = await this.sp.web.lists
        .getByTitle(this.listTitle)
        .items.filter(`TaskAssignmentId eq ${taskAssignmentId}`)
        .select(
          'Id', 'TaskAssignmentId', 'UserId', 'StartTime', 'EndTime',
          'HoursLogged', 'IsActive', 'WorkType', 'ActivityDescription',
          'IsBillable', 'Notes', 'Created', 'Modified',
          'Author/Id', 'Author/Title', 'Author/EMail',
          'Editor/Id', 'Editor/Title', 'Editor/EMail'
        )
        .expand('Author', 'Editor')
        .orderBy('StartTime', false)();

      const entries: ITaskTimeEntryView[] = items.map((item: any) => ({
        Id: item.Id,
        Title: item.ActivityDescription || `Time Entry ${item.Id}`,
        TaskAssignmentId: item.TaskAssignmentId,
        UserId: item.UserId,
        User: {
          Id: item.UserId,
          Title: item.Author?.Title || 'Unknown',
          Email: item.Author?.EMail || ''
        },
        StartTime: new Date(item.StartTime),
        EndTime: item.EndTime ? new Date(item.EndTime) : undefined,
        HoursLogged: item.HoursLogged,
        IsActive: item.IsActive || false,
        WorkType: item.WorkType as WorkType,
        ActivityDescription: item.ActivityDescription,
        IsBillable: item.IsBillable || false,
        Notes: item.Notes,
        Created: new Date(item.Created),
        Modified: new Date(item.Modified),
        Author: {
          Id: item.Author?.Id,
          Title: item.Author?.Title || 'Unknown',
          Email: item.Author?.EMail || ''
        },
        Editor: {
          Id: item.Editor?.Id,
          Title: item.Editor?.Title || 'Unknown',
          Email: item.Editor?.EMail || ''
        },
        Duration: this.formatDuration(item.HoursLogged),
        CanEdit: item.UserId === currentUserId,
        CanDelete: item.UserId === currentUserId
      }));

      logger.info('TaskTimeTrackingService', `Retrieved ${entries.length} time entries`);
      return entries;
    } catch (error) {
      logger.error('TaskTimeTrackingService', 'Error fetching time entries', error);
      throw error;
    }
  }

  /**
   * Get active timer for current user and task
   */
  public async getActiveTimer(taskAssignmentId: number, userId: number): Promise<IJmlTaskTimeEntry | null> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.listTitle)
        .items.filter(`TaskAssignmentId eq ${taskAssignmentId} and UserId eq ${userId} and IsActive eq 1`)
        .select('Id', 'TaskAssignmentId', 'UserId', 'StartTime', 'WorkType', 'ActivityDescription', 'Notes')
        .top(1)();

      if (items.length === 0) return null;

      const item = items[0];
      return {
        Id: item.Id,
        Title: item.ActivityDescription || `Time Entry ${item.Id}`,
        TaskAssignmentId: item.TaskAssignmentId,
        UserId: item.UserId,
        StartTime: new Date(item.StartTime),
        HoursLogged: 0,
        IsActive: true,
        WorkType: item.WorkType as WorkType,
        ActivityDescription: item.ActivityDescription,
        Notes: item.Notes,
        Created: new Date(),
        Modified: new Date()
      };
    } catch (error) {
      logger.error('TaskTimeTrackingService', 'Error getting active timer', error);
      throw error;
    }
  }

  /**
   * Start a new timer
   */
  public async startTimer(
    taskAssignmentId: number,
    userId: number,
    workType: WorkType,
    activityDescription?: string
  ): Promise<IJmlTaskTimeEntry> {
    try {
      logger.info('TaskTimeTrackingService', `Starting timer for task ${taskAssignmentId}`);

      // Check for existing active timer
      const existingTimer = await this.getActiveTimer(taskAssignmentId, userId);
      if (existingTimer) {
        throw new Error('You already have an active timer for this task. Please stop it first.');
      }

      const newItem = await this.sp.web.lists
        .getByTitle(this.listTitle)
        .items.add({
          TaskAssignmentId: taskAssignmentId,
          UserId: userId,
          StartTime: new Date().toISOString(),
          HoursLogged: 0,
          IsActive: true,
          WorkType: workType,
          ActivityDescription: activityDescription || '',
          IsBillable: false
        });

      logger.info('TaskTimeTrackingService', `Timer started with ID ${newItem.data.Id}`);

      return {
        Id: newItem.data.Id,
        Title: activityDescription || `Time Entry ${newItem.data.Id}`,
        TaskAssignmentId: taskAssignmentId,
        UserId: userId,
        StartTime: new Date(),
        HoursLogged: 0,
        IsActive: true,
        WorkType: workType,
        ActivityDescription: activityDescription,
        Created: new Date(),
        Modified: new Date()
      };
    } catch (error) {
      logger.error('TaskTimeTrackingService', 'Error starting timer', error);
      throw error;
    }
  }

  /**
   * Stop active timer
   */
  public async stopTimer(timerId: number): Promise<IJmlTaskTimeEntry> {
    try {
      logger.info('TaskTimeTrackingService', `Stopping timer ${timerId}`);

      // Get current timer to calculate hours
      const item = await this.sp.web.lists
        .getByTitle(this.listTitle)
        .items.getById(timerId)
        .select('Id', 'TaskAssignmentId', 'UserId', 'StartTime', 'WorkType', 'ActivityDescription', 'Notes')();

      const startTime = new Date(item.StartTime);
      const endTime = new Date();
      const hoursLogged = (endTime.getTime() - startTime.getTime()) / (1000 * 60 * 60);

      // Update item
      await this.sp.web.lists
        .getByTitle(this.listTitle)
        .items.getById(timerId)
        .update({
          EndTime: endTime.toISOString(),
          HoursLogged: parseFloat(hoursLogged.toFixed(2)),
          IsActive: false
        });

      logger.info('TaskTimeTrackingService', `Timer stopped. Hours logged: ${hoursLogged.toFixed(2)}`);

      return {
        Id: item.Id,
        Title: item.ActivityDescription || `Time Entry ${item.Id}`,
        TaskAssignmentId: item.TaskAssignmentId,
        UserId: item.UserId,
        StartTime: startTime,
        EndTime: endTime,
        HoursLogged: parseFloat(hoursLogged.toFixed(2)),
        IsActive: false,
        WorkType: item.WorkType as WorkType,
        ActivityDescription: item.ActivityDescription,
        Notes: item.Notes,
        Created: new Date(),
        Modified: new Date()
      };
    } catch (error) {
      logger.error('TaskTimeTrackingService', 'Error stopping timer', error);
      throw error;
    }
  }

  /**
   * Add manual time entry
   */
  public async addTimeEntry(entry: ITaskTimeEntryForm, userId: number): Promise<IJmlTaskTimeEntry> {
    try {
      logger.info('TaskTimeTrackingService', `Adding manual time entry for task ${entry.TaskAssignmentId}`);

      // Calculate hours if not provided
      let hoursLogged = entry.HoursLogged || 0;
      if (!hoursLogged && entry.StartTime && entry.EndTime) {
        const startTime = new Date(entry.StartTime);
        const endTime = new Date(entry.EndTime);
        hoursLogged = (endTime.getTime() - startTime.getTime()) / (1000 * 60 * 60);
      }

      const newItem = await this.sp.web.lists
        .getByTitle(this.listTitle)
        .items.add({
          TaskAssignmentId: entry.TaskAssignmentId,
          UserId: userId,
          StartTime: entry.StartTime.toISOString(),
          EndTime: entry.EndTime ? entry.EndTime.toISOString() : null,
          HoursLogged: parseFloat(hoursLogged.toFixed(2)),
          IsActive: false,
          WorkType: entry.WorkType,
          ActivityDescription: entry.ActivityDescription || '',
          IsBillable: entry.IsBillable || false,
          Notes: entry.Notes || ''
        });

      logger.info('TaskTimeTrackingService', `Time entry added with ID ${newItem.data.Id}`);

      return {
        Id: newItem.data.Id,
        Title: entry.ActivityDescription || `Time Entry ${newItem.data.Id}`,
        TaskAssignmentId: entry.TaskAssignmentId,
        UserId: userId,
        StartTime: entry.StartTime,
        EndTime: entry.EndTime,
        HoursLogged: parseFloat(hoursLogged.toFixed(2)),
        IsActive: false,
        WorkType: entry.WorkType,
        ActivityDescription: entry.ActivityDescription,
        IsBillable: entry.IsBillable,
        Notes: entry.Notes,
        Created: new Date(),
        Modified: new Date()
      };
    } catch (error) {
      logger.error('TaskTimeTrackingService', 'Error adding time entry', error);
      throw error;
    }
  }

  /**
   * Update time entry
   */
  public async updateTimeEntry(entryId: number, updates: Partial<ITaskTimeEntryForm>): Promise<void> {
    try {
      logger.info('TaskTimeTrackingService', `Updating time entry ${entryId}`);

      const updateObj: any = {};
      if (updates.StartTime) updateObj.StartTime = updates.StartTime.toISOString();
      if (updates.EndTime) updateObj.EndTime = updates.EndTime.toISOString();
      if (updates.HoursLogged !== undefined) updateObj.HoursLogged = updates.HoursLogged;
      if (updates.WorkType) updateObj.WorkType = updates.WorkType;
      if (updates.ActivityDescription !== undefined) updateObj.ActivityDescription = updates.ActivityDescription;
      if (updates.IsBillable !== undefined) updateObj.IsBillable = updates.IsBillable;
      if (updates.Notes !== undefined) updateObj.Notes = updates.Notes;

      // Recalculate hours if start/end time changed
      if (updates.StartTime || updates.EndTime) {
        const item = await this.sp.web.lists
          .getByTitle(this.listTitle)
          .items.getById(entryId)
          .select('StartTime', 'EndTime')();

        const startTime = updates.StartTime ? new Date(updates.StartTime) : new Date(item.StartTime);
        const endTime = updates.EndTime ? new Date(updates.EndTime) : (item.EndTime ? new Date(item.EndTime) : null);

        if (endTime) {
          const hours = (endTime.getTime() - startTime.getTime()) / (1000 * 60 * 60);
          updateObj.HoursLogged = parseFloat(hours.toFixed(2));
        }
      }

      await this.sp.web.lists
        .getByTitle(this.listTitle)
        .items.getById(entryId)
        .update(updateObj);

      logger.info('TaskTimeTrackingService', `Time entry ${entryId} updated`);
    } catch (error) {
      logger.error('TaskTimeTrackingService', 'Error updating time entry', error);
      throw error;
    }
  }

  /**
   * Delete time entry
   */
  public async deleteTimeEntry(entryId: number): Promise<void> {
    try {
      logger.info('TaskTimeTrackingService', `Deleting time entry ${entryId}`);

      await this.sp.web.lists
        .getByTitle(this.listTitle)
        .items.getById(entryId)
        .delete();

      logger.info('TaskTimeTrackingService', `Time entry ${entryId} deleted`);
    } catch (error) {
      logger.error('TaskTimeTrackingService', 'Error deleting time entry', error);
      throw error;
    }
  }

  /**
   * Get time summary for a task
   */
  public async getTaskTimeSummary(taskAssignmentId: number): Promise<ITaskTimeSummary> {
    try {
      logger.info('TaskTimeTrackingService', `Calculating time summary for task ${taskAssignmentId}`);

      const entries = await this.sp.web.lists
        .getByTitle(this.listTitle)
        .items.filter(`TaskAssignmentId eq ${taskAssignmentId}`)
        .select('Id', 'HoursLogged', 'WorkType', 'IsActive', 'StartTime', 'Modified')
        .orderBy('Modified', false)();

      let totalHours = 0;
      const byWorkType: { [key in WorkType]?: number } = {};
      let activeEntry: IJmlTaskTimeEntry | undefined;
      let lastEntry: Date | undefined;

      entries.forEach((item: any) => {
        const hours = parseFloat(item.HoursLogged) || 0;
        totalHours += hours;

        // Track by work type
        const workType = item.WorkType as WorkType;
        if (!byWorkType[workType]) {
          byWorkType[workType] = 0;
        }
        byWorkType[workType]! += hours;

        // Track active entry
        if (item.IsActive) {
          activeEntry = {
            Id: item.Id,
            Title: item.ActivityDescription || `Time Entry ${item.Id}`,
            TaskAssignmentId: taskAssignmentId,
            UserId: 0, // Will be filled from expand if needed
            StartTime: new Date(item.StartTime),
            HoursLogged: 0,
            IsActive: true,
            WorkType: workType,
            Created: new Date(),
            Modified: new Date()
          };
        }

        // Track last entry date
        const modifiedDate = new Date(item.Modified);
        if (!lastEntry || modifiedDate > lastEntry) {
          lastEntry = modifiedDate;
        }
      });

      const summary: ITaskTimeSummary = {
        TaskAssignmentId: taskAssignmentId,
        TotalHours: parseFloat(totalHours.toFixed(2)),
        VarianceHours: 0, // Will be calculated with EstimatedHours
        PercentComplete: 0, // Will be calculated with EstimatedHours
        ByWorkType: byWorkType,
        LastEntry: lastEntry,
        ActiveEntry: activeEntry
      };

      logger.info('TaskTimeTrackingService', `Time summary calculated: ${summary.TotalHours} hours`);
      return summary;
    } catch (error) {
      logger.error('TaskTimeTrackingService', 'Error calculating time summary', error);
      throw error;
    }
  }

  /**
   * Format duration for display
   */
  private formatDuration(hours: number): string {
    if (hours < 1) {
      const minutes = Math.round(hours * 60);
      return `${minutes}m`;
    } else if (hours < 24) {
      const wholeHours = Math.floor(hours);
      const minutes = Math.round((hours - wholeHours) * 60);
      return minutes > 0 ? `${wholeHours}h ${minutes}m` : `${wholeHours}h`;
    } else {
      const days = Math.floor(hours / 24);
      const remainingHours = Math.floor(hours % 24);
      return remainingHours > 0 ? `${days}d ${remainingHours}h` : `${days}d`;
    }
  }

  /**
   * Get user's time entries across all tasks (for reporting)
   */
  public async getUserTimeEntries(
    userId: number,
    startDate?: Date,
    endDate?: Date
  ): Promise<ITaskTimeEntryView[]> {
    try {
      logger.info('TaskTimeTrackingService', `Fetching time entries for user ${userId}`);

      let filter = `UserId eq ${userId}`;
      if (startDate) {
        filter += ` and StartTime ge datetime'${startDate.toISOString()}'`;
      }
      if (endDate) {
        filter += ` and StartTime le datetime'${endDate.toISOString()}'`;
      }

      const items = await this.sp.web.lists
        .getByTitle(this.listTitle)
        .items.filter(filter)
        .select(
          'Id', 'TaskAssignmentId', 'UserId', 'StartTime', 'EndTime',
          'HoursLogged', 'IsActive', 'WorkType', 'ActivityDescription',
          'IsBillable', 'Notes', 'Created', 'Modified',
          'Author/Id', 'Author/Title', 'Author/EMail',
          'TaskAssignment/Id', 'TaskAssignment/Title'
        )
        .expand('Author', 'TaskAssignment')
        .orderBy('StartTime', false)();

      const entries: ITaskTimeEntryView[] = items.map((item: any) => ({
        Id: item.Id,
        Title: item.ActivityDescription || `Time Entry ${item.Id}`,
        TaskAssignmentId: item.TaskAssignmentId,
        TaskAssignment: item.TaskAssignment ? {
          Id: item.TaskAssignment.Id,
          Title: item.TaskAssignment.Title
        } : undefined,
        UserId: item.UserId,
        User: {
          Id: item.UserId,
          Title: item.Author?.Title || 'Unknown',
          Email: item.Author?.EMail || ''
        },
        StartTime: new Date(item.StartTime),
        EndTime: item.EndTime ? new Date(item.EndTime) : undefined,
        HoursLogged: item.HoursLogged,
        IsActive: item.IsActive || false,
        WorkType: item.WorkType as WorkType,
        ActivityDescription: item.ActivityDescription,
        IsBillable: item.IsBillable || false,
        Notes: item.Notes,
        Created: new Date(item.Created),
        Modified: new Date(item.Modified),
        Author: {
          Id: item.Author?.Id,
          Title: item.Author?.Title || 'Unknown',
          Email: item.Author?.EMail || ''
        },
        Duration: this.formatDuration(item.HoursLogged),
        CanEdit: true,
        CanDelete: true
      }));

      logger.info('TaskTimeTrackingService', `Retrieved ${entries.length} time entries for user`);
      return entries;
    } catch (error) {
      logger.error('TaskTimeTrackingService', 'Error fetching user time entries', error);
      throw error;
    }
  }
}
