// JML Saved Views Model
// Manages user-defined filter presets and view configurations

import { TaskStatus, Priority } from './ICommon';

export interface IJmlSavedView {
  id: string; // Unique identifier (GUID or timestamp-based)
  name: string; // User-defined name
  isDefault: boolean; // System-provided default view
  isActive?: boolean; // Currently active view
  userId?: number; // Owner of the view (for user-specific views)
  createdDate: Date;
  lastModified: Date;
  filterConfig: IViewFilterConfig;
}

export interface IViewFilterConfig {
  // Status filters
  statusFilters: TaskStatus[];

  // Priority filters
  priorityFilters: Priority[];

  // Assignment filters
  assignedToMe?: boolean;
  assignedToOthers?: boolean;
  unassigned?: boolean;

  // Date filters
  dueDateRange?: {
    start?: Date;
    end?: Date;
  };
  overdue?: boolean;
  dueThisWeek?: boolean;
  dueToday?: boolean;

  // Search/text filter
  searchText?: string;

  // Sorting
  sortBy?: 'dueDate' | 'priority' | 'status' | 'title' | 'created';
  sortDirection?: 'asc' | 'desc';

  // Grouping
  groupBy?: 'none' | 'status' | 'priority' | 'assignee' | 'dueDate';

  // Additional filters
  processTypes?: string[]; // Filter by process type (Joiner, Mover, Leaver)
  departments?: string[];
  hasComments?: boolean;
  hasAttachments?: boolean;
  hasDependencies?: boolean;
}

// Preset default views
export enum DefaultViewType {
  AllTasks = 'all-tasks',
  MyOpenTasks = 'my-open-tasks',
  OverdueTasks = 'overdue-tasks',
  HighPriority = 'high-priority',
  DueThisWeek = 'due-this-week',
  DueToday = 'due-today',
  Completed = 'completed',
  Blocked = 'blocked',
  UnassignedTasks = 'unassigned-tasks'
}

export interface IDefaultView {
  type: DefaultViewType;
  name: string;
  description: string;
  icon: string; // Fluent UI icon name
  filterConfig: IViewFilterConfig;
}

// View management operations
export interface ISavedViewsState {
  views: IJmlSavedView[];
  activeViewId: string | null;
  defaultViews: IDefaultView[];
}

// Form model for creating/editing views
export interface ISaveViewForm {
  name: string;
  captureCurrentFilters: boolean;
  makeDefault?: boolean;
}
