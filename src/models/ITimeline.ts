export interface ITimelineTask {
  id: number;
  title: string;
  description: string;
  startDate: Date;
  endDate: Date;

  // Progress Tracking
  progress: number; // 0-100
  status: TaskStatus;

  // Process & Context
  processType: ProcessType;
  processInstanceId?: number;
  assignedTo?: string[];
  owner: string;

  // Dependencies
  dependsOn: number[]; // Array of task IDs
  blockedBy: number[]; // Tasks that are blocking this one
  isOnCriticalPath: boolean;

  // Resource Allocation
  resources: IResourceAllocation[];
  estimatedHours: number;
  actualHours: number;

  // Milestone
  isMilestone: boolean;
  milestoneType?: MilestoneType;

  // Styling
  color: string;
  backgroundColor: string;

  // Metadata
  createdBy: string;
  createdDate: Date;
  modifiedBy: string;
  modifiedDate: Date;
}

export enum TaskStatus {
  NotStarted = 'Not Started',
  InProgress = 'In Progress',
  OnHold = 'On Hold',
  Completed = 'Completed',
  Cancelled = 'Cancelled',
  Overdue = 'Overdue',
  AtRisk = 'At Risk'
}

export enum ProcessType {
  Onboarding = 'Onboarding',
  Offboarding = 'Offboarding',
  Probation = 'Probation',
  Training = 'Training',
  Review = 'Review',
  Project = 'Project',
  Compliance = 'Compliance',
  Other = 'Other'
}

export enum MilestoneType {
  StartDate = 'Start Date',
  EndDate = 'End Date',
  Checkpoint = 'Checkpoint',
  Deliverable = 'Deliverable',
  Review = 'Review',
  Approval = 'Approval'
}

export interface IResourceAllocation {
  resourceId: string;
  resourceName: string;
  resourceEmail: string;
  allocatedHours: number;
  role: string;
  availability: number; // Percentage 0-100
}

export interface ITaskDependency {
  taskId: number;
  dependsOnTaskId: number;
  dependencyType: DependencyType;
  lag: number; // Days of lag/lead time (negative for lead)
}

export enum DependencyType {
  FinishToStart = 'Finish-to-Start', // Default - predecessor must finish before successor starts
  StartToStart = 'Start-to-Start',   // Both tasks start at the same time
  FinishToFinish = 'Finish-to-Finish', // Both tasks finish at the same time
  StartToFinish = 'Start-to-Finish'   // Successor can't finish until predecessor starts
}

export interface ICriticalPath {
  tasks: ITimelineTask[];
  totalDuration: number; // In days
  startDate: Date;
  endDate: Date;
  slack: number; // Float/slack time in days
}

export interface IJmlTimelineProps {
  sp: any;
  siteUrl: string;
  userEmail: string;
  userDisplayName: string;

  // View Settings
  viewMode: TimelineView;
  showCriticalPath: boolean;
  showResourceAllocation: boolean;

  // Date Range Settings
  monthsBack: number;
  monthsForward: number;
  timelineScale: string;

  // Filter Settings
  defaultProcessType: string;
  defaultStatusFilter: string;
  showCompletedTasks: boolean;

  // Display Options
  showWeekends: boolean;
  showTodayIndicator: boolean;
  showMilestonesOnly: boolean;
  groupByProcessType: boolean;
  showTaskDependencies: boolean;
  showResourceNames: boolean;
  compactView: boolean;

  // Performance Settings
  maxTasksToDisplay: number;
  autoRefreshMinutes: number;
}

export type TimelineView = 'timeline' | 'gantt' | 'resource' | 'critical-path';

export interface IJmlTimelineState {
  loading: boolean;
  error: string;
  tasks: ITimelineTask[];
  dependencies: ITaskDependency[];
  criticalPath?: ICriticalPath;
  selectedTask?: ITimelineTask;
  showTaskDetails: boolean;
  viewMode: TimelineView;

  // Filters
  filterProcessType?: ProcessType;
  filterStatus?: TaskStatus;
  filterAssignedTo?: string;

  // Date Range
  startDate: Date;
  endDate: Date;

  // Resource View
  resources: IResourceAllocation[];
  showResourceUtilization: boolean;
}

// For drag-and-drop timeline adjustments
export interface ITimelineDragContext {
  taskId: number;
  originalStartDate: Date;
  originalEndDate: Date;
  dragType: 'move' | 'resize-start' | 'resize-end';
}

// For Gantt chart rendering
export interface IGanttRow {
  task: ITimelineTask;
  level: number; // For hierarchy/indentation
  children?: IGanttRow[];
  collapsed: boolean;
}

export interface IGanttTimeScale {
  unit: 'day' | 'week' | 'month' | 'quarter';
  pixelsPerUnit: number;
  startDate: Date;
  endDate: Date;
}

// Resource utilization
export interface IResourceUtilization {
  resource: IResourceAllocation;
  allocatedHours: number;
  availableHours: number;
  utilizationPercent: number; // 0-100+, can exceed 100 if overallocated
  tasks: ITimelineTask[];
  isOverallocated: boolean;
}

// Timeline filters
export interface ITimelineFilter {
  processTypes: ProcessType[];
  statuses: TaskStatus[];
  assignedTo: string[];
  dateRange: {
    start: Date;
    end: Date;
  };
  showMilestonesOnly: boolean;
  showCriticalPathOnly: boolean;
}

// Export options
export interface ITimelineExportOptions {
  format: 'pdf' | 'png' | 'mpp' | 'xlsx';
  includeDetails: boolean;
  includeDependencies: boolean;
  includeResources: boolean;
}
