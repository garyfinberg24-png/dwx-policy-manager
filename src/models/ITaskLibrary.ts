// Task Library Data Models
// Comprehensive task library system for Process Wizard

import { ProcessType, Priority, TaskCategory } from './ICommon';

// ============================================================================
// ENUMS
// ============================================================================

/**
 * Task phases for timeline organization
 */
export enum TaskPhase {
  PreArrival = 'Pre-Arrival',
  Day1 = 'Day 1',
  Week1 = 'Week 1',
  Month1 = 'Month 1',
  Month2Plus = 'Month 2+',
  Ongoing = 'Ongoing'
}

/**
 * Task trigger types for automation
 */
export enum TriggerType {
  Manual = 'Manual',
  AutoTrigger = 'Auto-Trigger',
  Scheduled = 'Scheduled',
  Conditional = 'Conditional'
}

/**
 * Automation action types
 */
export enum AutomationAction {
  SendEmail = 'Send Email',
  SendTeamsNotification = 'Send Teams Notification',
  CreateCalendarEvent = 'Create Calendar Event',
  AssignTask = 'Assign Task',
  UpdateStatus = 'Update Status',
  SendReminder = 'Send Reminder',
  Escalate = 'Escalate',
  CompleteTask = 'Complete Task',
  RunWorkflow = 'Run Workflow'
}

/**
 * Template tiers for different organizational needs
 */
export enum TemplateTier {
  Enterprise = 'Enterprise',
  Professional = 'Professional',
  Essential = 'Essential'
}

// ============================================================================
// TASK LIBRARY ITEM
// ============================================================================

/**
 * Comprehensive task definition for the task library
 */
export interface ITaskLibraryItem {
  // Core Identification
  id: string;                    // Unique ID (e.g., "task-it-hardware-001")
  code: string;                  // Display code (e.g., "IT-HW-001")
  title: string;                 // Task title
  category: TaskCategory;        // Task category
  phase: TaskPhase;              // When in process this typically occurs

  // Description & Instructions
  description: string;           // Short description
  instructions?: string;         // Detailed instructions

  // Assignment
  defaultDepartment: string;     // Default owning department
  defaultRole?: string;          // Default assignee role

  // Timing
  estimatedHours: number;        // Estimated time to complete
  defaultDaysOffset: number;     // Days before/after start date (negative = before)
  slaHours?: number;             // Service Level Agreement

  // Priority & Requirements
  priority: Priority;
  requiresApproval: boolean;
  approverRole?: string;

  // Dependencies
  dependencies?: string[];       // Array of task IDs this depends on
  blocksOtherTasks?: boolean;    // If true, dependent tasks wait for this

  // Automation
  triggerType: TriggerType;
  automationRules?: IAutomationRule[];

  // Resources
  formUrl?: string;
  documentationUrl?: string;
  systemUrl?: string;
  attachments?: string[];

  // Tags & Metadata
  tags: string[];                // For filtering/searching
  icon: string;                  // Fluent UI icon name
  isCommon: boolean;             // Commonly used across processes
  isOptional: boolean;           // Can be skipped

  // Template Association
  includedInTiers: TemplateTier[]; // Which templates include this by default
  applicableProcessTypes: ProcessType[]; // Joiner, Mover, Leaver
}

// ============================================================================
// AUTOMATION
// ============================================================================

/**
 * Automation rule definition
 */
export interface IAutomationRule {
  id: string;
  action: AutomationAction;
  trigger: 'onAssign' | 'onStart' | 'onComplete' | 'onOverdue' | 'beforeDue';
  triggerOffset?: number;        // Days/hours before/after trigger

  // Conditional execution
  condition?: {
    field: string;
    operator: 'equals' | 'notEquals' | 'contains' | 'greaterThan' | 'lessThan';
    value: any;
  };

  // Action parameters
  parameters?: {
    recipient?: string;          // Email/Teams recipient
    template?: string;           // Message template ID
    subject?: string;            // Email subject
    message?: string;            // Message body
    escalateTo?: string;         // Escalation role/person
    workflowId?: string;         // Power Automate workflow ID
  };

  isEnabled: boolean;
  description: string;
}

// ============================================================================
// TASK CATEGORY DEFINITION
// ============================================================================

/**
 * Task category with metadata and icon
 */
export interface ITaskCategoryDefinition {
  category: TaskCategory;
  displayName: string;
  description: string;
  icon: string;                  // Fluent UI icon name
  color: string;                 // Hex color for UI
  department: string;            // Primary owning department
  taskCount?: number;            // Number of tasks in this category
}

// ============================================================================
// TEMPLATE DEFINITION
// ============================================================================

/**
 * Process template definition
 */
export interface IProcessTemplate {
  // Identification
  id: string;
  code: string;                  // Template code (e.g., "TPL-JOIN-ENT-001")
  name: string;
  tier: TemplateTier;
  processType: ProcessType;

  // Description
  description: string;
  shortDescription: string;      // For cards/lists

  // Metadata
  estimatedDays: number;
  taskCount: number;
  automationPercentage: number;  // 0-100

  // Tasks
  tasks: string[];               // Array of task IDs from library

  // Customization
  isCustomizable: boolean;
  isPublic: boolean;             // Available to all vs. org-specific

  // Analytics
  rating?: number;               // 1-5 stars
  usageCount?: number;

  // Tags & Industry
  tags: string[];
  industries?: string[];         // Healthcare, Tech, Finance, etc.
  companySize?: string[];        // <50, 50-500, 500+

  // Visual
  icon: string;
  color?: string;

  // Metadata
  createdBy?: string;
  createdDate?: Date;
  modifiedDate?: Date;
  version?: string;
}

// ============================================================================
// SELECTED TASK (User's customization)
// ============================================================================

/**
 * Task selected for a specific process (with customizations)
 */
export interface ISelectedTask {
  // Reference to library task
  libraryTaskId: string;

  // Customizations (override library defaults)
  title?: string;
  description?: string;
  instructions?: string;

  // Assignment overrides
  assignedTo?: string;           // User ID or role
  department?: string;

  // Timing overrides
  daysOffset?: number;
  estimatedHours?: number;
  scheduledDate?: Date;
  dueDate?: Date;

  // Priority override
  priority?: Priority;

  // Dependencies (custom for this process)
  dependencies?: string[];

  // Automation overrides
  triggerType?: TriggerType;
  automationRules?: IAutomationRule[];

  // Custom fields
  customFields?: Record<string, any>;
  notes?: string;

  // UI state
  isExpanded?: boolean;
  order?: number;                // Custom ordering
}

// ============================================================================
// PROCESS CONFIGURATION
// ============================================================================

/**
 * Complete process configuration from wizard
 */
export interface IProcessConfiguration {
  // Template source
  templateId?: string;
  templateTier?: TemplateTier;

  // Process details
  processType: ProcessType;
  employeeName: string;
  employeeEmail?: string;
  startDate: Date;
  department: string;
  jobTitle: string;
  managerId?: number;
  buddyId?: number;

  // Selected tasks
  tasks: ISelectedTask[];

  // Automation settings
  enableAutomation: boolean;
  notifyStakeholders: boolean;
  sendWelcomeEmail: boolean;
  scheduleKickoffMeeting: boolean;

  // Metadata
  createdBy?: string;
  createdDate?: Date;
  notes?: string;
}

// ============================================================================
// UI/UX HELPERS
// ============================================================================

/**
 * Task group for UI display (by phase, category, etc.)
 */
export interface ITaskGroup {
  id: string;
  title: string;
  description?: string;
  icon: string;
  tasks: ITaskLibraryItem[];
  isExpanded: boolean;
  order: number;
}

/**
 * Template card for UI display
 */
export interface ITemplateCard {
  template: IProcessTemplate;
  isRecommended: boolean;
  isPopular: boolean;
  matchScore?: number;           // 0-100 relevance score
}

/**
 * Automation preview for UI
 */
export interface IAutomationPreview {
  phase: TaskPhase;
  day: number;
  actions: {
    task: ITaskLibraryItem;
    automations: IAutomationRule[];
  }[];
}
