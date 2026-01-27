// @ts-nocheck
// SavedViewsService - Manages saved view persistence and default views
// Uses localStorage for client-side storage

import {
  IJmlSavedView,
  IViewFilterConfig,
  IDefaultView,
  DefaultViewType,
  ISavedViewsState
} from '../models';
import { TaskStatus, Priority } from '../models/ICommon';
import { logger } from './LoggingService';

export class SavedViewsService {
  private storageKey = 'jml_saved_views';
  private activeViewKey = 'jml_active_view';
  private userId: number;

  constructor(userId: number) {
    this.userId = userId;
  }

  /**
   * Get all saved views for current user
   */
  public getSavedViews(): IJmlSavedView[] {
    try {
      const stored = localStorage.getItem(`${this.storageKey}_${this.userId}`);
      if (!stored) return [];

      const views: IJmlSavedView[] = JSON.parse(stored);

      // Convert date strings back to Date objects
      return views.map(view => ({
        ...view,
        createdDate: new Date(view.createdDate),
        lastModified: new Date(view.lastModified),
        filterConfig: {
          ...view.filterConfig,
          dueDateRange: view.filterConfig.dueDateRange ? {
            start: view.filterConfig.dueDateRange.start ? new Date(view.filterConfig.dueDateRange.start) : undefined,
            end: view.filterConfig.dueDateRange.end ? new Date(view.filterConfig.dueDateRange.end) : undefined
          } : undefined
        }
      }));
    } catch (error) {
      logger.error('SavedViewsService', 'Error loading saved views', error);
      return [];
    }
  }

  /**
   * Save a new view
   */
  public saveView(name: string, filterConfig: IViewFilterConfig): IJmlSavedView {
    try {
      const views = this.getSavedViews();

      const newView: IJmlSavedView = {
        id: `view_${Date.now()}_${Math.random().toString(36).substring(7)}`,
        name,
        isDefault: false,
        userId: this.userId,
        createdDate: new Date(),
        lastModified: new Date(),
        filterConfig
      };

      views.push(newView);
      this.persistViews(views);

      logger.info('SavedViewsService', `Saved view: ${name}`);
      return newView;
    } catch (error) {
      logger.error('SavedViewsService', 'Error saving view', error);
      throw error;
    }
  }

  /**
   * Update an existing view
   */
  public updateView(viewId: string, name: string, filterConfig: IViewFilterConfig): void {
    try {
      const views = this.getSavedViews();
      const viewIndex = views.findIndex(v => v.id === viewId);

      if (viewIndex === -1) {
        throw new Error(`View not found: ${viewId}`);
      }

      views[viewIndex] = {
        ...views[viewIndex],
        name,
        filterConfig,
        lastModified: new Date()
      };

      this.persistViews(views);
      logger.info('SavedViewsService', `Updated view: ${name}`);
    } catch (error) {
      logger.error('SavedViewsService', 'Error updating view', error);
      throw error;
    }
  }

  /**
   * Delete a saved view
   */
  public deleteView(viewId: string): void {
    try {
      const views = this.getSavedViews();
      const filteredViews = views.filter(v => v.id !== viewId);

      if (filteredViews.length === views.length) {
        throw new Error(`View not found: ${viewId}`);
      }

      this.persistViews(filteredViews);

      // Clear active view if it was deleted
      if (this.getActiveViewId() === viewId) {
        this.setActiveViewId(null);
      }

      logger.info('SavedViewsService', `Deleted view: ${viewId}`);
    } catch (error) {
      logger.error('SavedViewsService', 'Error deleting view', error);
      throw error;
    }
  }

  /**
   * Get active view ID
   */
  public getActiveViewId(): string | null {
    return localStorage.getItem(`${this.activeViewKey}_${this.userId}`);
  }

  /**
   * Set active view
   */
  public setActiveViewId(viewId: string | null): void {
    if (viewId) {
      localStorage.setItem(`${this.activeViewKey}_${this.userId}`, viewId);
    } else {
      localStorage.removeItem(`${this.activeViewKey}_${this.userId}`);
    }
  }

  /**
   * Get default/preset views
   */
  public getDefaultViews(): IDefaultView[] {
    const now = new Date();
    const startOfWeek = new Date(now);
    startOfWeek.setDate(now.getDate() - now.getDay());
    startOfWeek.setHours(0, 0, 0, 0);

    const endOfWeek = new Date(startOfWeek);
    endOfWeek.setDate(startOfWeek.getDate() + 6);
    endOfWeek.setHours(23, 59, 59, 999);

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const endOfToday = new Date();
    endOfToday.setHours(23, 59, 59, 999);

    return [
      {
        type: DefaultViewType.AllTasks,
        name: 'All Tasks',
        description: 'Show all tasks assigned to me',
        icon: 'TaskList',
        filterConfig: {
          statusFilters: [
            TaskStatus.NotStarted,
            TaskStatus.InProgress,
            TaskStatus.Blocked,
            TaskStatus.Waiting,
            TaskStatus.Completed
          ],
          priorityFilters: [],
          assignedToMe: true,
          sortBy: 'dueDate',
          sortDirection: 'asc',
          groupBy: 'none'
        }
      },
      {
        type: DefaultViewType.MyOpenTasks,
        name: 'My Open Tasks',
        description: 'Active tasks that need attention',
        icon: 'CheckboxComposite',
        filterConfig: {
          statusFilters: [
            TaskStatus.NotStarted,
            TaskStatus.InProgress
          ],
          priorityFilters: [],
          assignedToMe: true,
          sortBy: 'priority',
          sortDirection: 'desc',
          groupBy: 'status'
        }
      },
      {
        type: DefaultViewType.OverdueTasks,
        name: 'Overdue',
        description: 'Tasks past their due date',
        icon: 'Warning',
        filterConfig: {
          statusFilters: [
            TaskStatus.NotStarted,
            TaskStatus.InProgress,
            TaskStatus.Blocked,
            TaskStatus.Waiting
          ],
          priorityFilters: [],
          assignedToMe: true,
          overdue: true,
          sortBy: 'dueDate',
          sortDirection: 'asc',
          groupBy: 'priority'
        }
      },
      {
        type: DefaultViewType.HighPriority,
        name: 'High Priority',
        description: 'Critical and high priority tasks',
        icon: 'Important',
        filterConfig: {
          statusFilters: [
            TaskStatus.NotStarted,
            TaskStatus.InProgress
          ],
          priorityFilters: [Priority.Critical, Priority.High],
          assignedToMe: true,
          sortBy: 'dueDate',
          sortDirection: 'asc',
          groupBy: 'none'
        }
      },
      {
        type: DefaultViewType.DueThisWeek,
        name: 'Due This Week',
        description: 'Tasks due within this week',
        icon: 'Calendar',
        filterConfig: {
          statusFilters: [
            TaskStatus.NotStarted,
            TaskStatus.InProgress
          ],
          priorityFilters: [],
          assignedToMe: true,
          dueThisWeek: true,
          dueDateRange: {
            start: startOfWeek,
            end: endOfWeek
          },
          sortBy: 'dueDate',
          sortDirection: 'asc',
          groupBy: 'none'
        }
      },
      {
        type: DefaultViewType.DueToday,
        name: 'Due Today',
        description: 'Tasks due today',
        icon: 'CalendarDay',
        filterConfig: {
          statusFilters: [
            TaskStatus.NotStarted,
            TaskStatus.InProgress
          ],
          priorityFilters: [],
          assignedToMe: true,
          dueToday: true,
          dueDateRange: {
            start: today,
            end: endOfToday
          },
          sortBy: 'priority',
          sortDirection: 'desc',
          groupBy: 'none'
        }
      },
      {
        type: DefaultViewType.Completed,
        name: 'Completed',
        description: 'Recently completed tasks',
        icon: 'CompletedSolid',
        filterConfig: {
          statusFilters: [TaskStatus.Completed],
          priorityFilters: [],
          assignedToMe: true,
          sortBy: 'dueDate',
          sortDirection: 'desc',
          groupBy: 'none'
        }
      },
      {
        type: DefaultViewType.Blocked,
        name: 'Blocked',
        description: 'Tasks with blockers',
        icon: 'Blocked2',
        filterConfig: {
          statusFilters: [TaskStatus.Blocked],
          priorityFilters: [],
          assignedToMe: true,
          sortBy: 'dueDate',
          sortDirection: 'asc',
          groupBy: 'priority'
        }
      },
      {
        type: DefaultViewType.UnassignedTasks,
        name: 'Unassigned',
        description: 'Tasks without an assignee',
        icon: 'UnknownSolid',
        filterConfig: {
          statusFilters: [
            TaskStatus.NotStarted,
            TaskStatus.InProgress
          ],
          priorityFilters: [],
          unassigned: true,
          sortBy: 'priority',
          sortDirection: 'desc',
          groupBy: 'none'
        }
      }
    ];
  }

  /**
   * Get view by ID (from saved views or default views)
   */
  public getViewById(viewId: string): IJmlSavedView | IDefaultView | null {
    // Check saved views first
    const savedViews = this.getSavedViews();
    const savedView = savedViews.find(v => v.id === viewId);
    if (savedView) return savedView;

    // Check default views
    const defaultViews = this.getDefaultViews();
    const defaultView = defaultViews.find(v => v.type === viewId);
    if (defaultView) {
      // Convert IDefaultView to IJmlSavedView format
      return {
        id: defaultView.type,
        name: defaultView.name,
        isDefault: true,
        createdDate: new Date(),
        lastModified: new Date(),
        filterConfig: defaultView.filterConfig
      };
    }

    return null;
  }

  /**
   * Get complete state (saved + default views)
   */
  public getViewsState(): ISavedViewsState {
    return {
      views: this.getSavedViews(),
      activeViewId: this.getActiveViewId(),
      defaultViews: this.getDefaultViews()
    };
  }

  /**
   * Persist views to localStorage
   */
  private persistViews(views: IJmlSavedView[]): void {
    localStorage.setItem(`${this.storageKey}_${this.userId}`, JSON.stringify(views));
  }

  /**
   * Export views (for backup/migration)
   */
  public exportViews(): string {
    const state = this.getViewsState();
    return JSON.stringify(state, null, 2);
  }

  /**
   * Import views (for backup/migration)
   */
  public importViews(data: string): void {
    try {
      const parsed = JSON.parse(data);
      if (parsed.views && Array.isArray(parsed.views)) {
        this.persistViews(parsed.views);
        logger.info('SavedViewsService', `Imported ${parsed.views.length} views`);
      } else {
        throw new Error('Invalid import data format');
      }
    } catch (error) {
      logger.error('SavedViewsService', 'Error importing views', error);
      throw error;
    }
  }

  /**
   * Clear all saved views
   */
  public clearAll(): void {
    localStorage.removeItem(`${this.storageKey}_${this.userId}`);
    localStorage.removeItem(`${this.activeViewKey}_${this.userId}`);
    logger.info('SavedViewsService', 'Cleared all saved views');
  }
}
