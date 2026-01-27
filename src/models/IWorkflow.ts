/**
 * JML Workflow Engine Models
 * Defines interfaces for workflow definitions, instances, and execution
 */

import { IBaseListItem, ProcessType, IUser } from './ICommon';

// Re-export ProcessType as WorkflowProcessType for backward compatibility
export { ProcessType as WorkflowProcessType };

// ============================================================================
// WORKFLOW STATUS ENUMS
// ============================================================================

export enum WorkflowStatus {
  Draft = 'Draft',
  Active = 'Active',
  Paused = 'Paused',
  Completed = 'Completed',
  Failed = 'Failed',
  Cancelled = 'Cancelled'
}

export enum WorkflowInstanceStatus {
  Pending = 'Pending',
  Running = 'Running',
  Paused = 'Paused',
  WaitingForInput = 'Waiting for Input',
  WaitingForApproval = 'Waiting for Approval',
  WaitingForTask = 'Waiting for Task',
  Completed = 'Completed',
  Failed = 'Failed',
  Cancelled = 'Cancelled'
}

export enum StepStatus {
  Pending = 'Pending',
  InProgress = 'In Progress',
  Completed = 'Completed',
  Skipped = 'Skipped',
  Failed = 'Failed',
  Cancelled = 'Cancelled'
}

export enum StepType {
  Start = 'Start',
  End = 'End',
  AssignTasks = 'AssignTasks',
  CreateTask = 'CreateTask',
  WaitForTasks = 'WaitForTasks',
  Approval = 'Approval',
  Condition = 'Condition',
  Action = 'Action',
  Notification = 'Notification',
  Wait = 'Wait',
  Parallel = 'Parallel',
  SetVariable = 'SetVariable',
  // New step types for enhanced workflow engine
  ForEach = 'ForEach',           // Loop/iterator for collections
  CallWorkflow = 'CallWorkflow', // Sub-workflow invocation
  Webhook = 'Webhook'            // External webhook call
}

// Alias for backward compatibility with workflow designer
export const WorkflowStepType = StepType;
export type WorkflowStepType = StepType;

export enum ActionType {
  CreateTask = 'CreateTask',
  AssignTasksFromTemplate = 'AssignTasksFromTemplate',
  UpdateTaskStatus = 'UpdateTaskStatus',
  UpdateTask = 'UpdateTask',
  CompleteTask = 'CompleteTask',
  CreateApproval = 'CreateApproval',
  SendNotification = 'SendNotification',
  SendEmail = 'SendEmail',
  UpdateListItem = 'UpdateListItem',
  CreateListItem = 'CreateListItem',
  SetVariable = 'SetVariable',
  CallWebhook = 'CallWebhook',
  SendTeamsMessage = 'SendTeamsMessage',
  Wait = 'Wait',
  Custom = 'Custom',
  // Azure AD Actions
  AddUserToGroup = 'AddUserToGroup',
  RemoveUserFromGroup = 'RemoveUserFromGroup',
  DisableUserAccount = 'DisableUserAccount',
  EnableUserAccount = 'EnableUserAccount',
  UpdateUserProfile = 'UpdateUserProfile',
  // Calendar Actions
  CreateCalendarEvent = 'CreateCalendarEvent',
  UpdateCalendarEvent = 'UpdateCalendarEvent',
  DeleteCalendarEvent = 'DeleteCalendarEvent',
  // Asset & Equipment Actions
  CreateEquipmentRequest = 'CreateEquipmentRequest',
  CreateAssetReturnRequest = 'CreateAssetReturnRequest',
  ReclaimLicense = 'ReclaimLicense'
}

// Alias for backward compatibility with workflow designer
export const WorkflowActionType = ActionType;
export type WorkflowActionType = ActionType;

/**
 * Workflow Action - Used by workflow designer for step actions
 */
export interface IWorkflowAction {
  id: string;
  type: ActionType;
  name: string;
  config: Record<string, unknown>;
  order: number;
}

/**
 * Workflow Condition - Used by workflow designer for branching
 */
export interface IWorkflowCondition {
  id: string;
  name: string;
  operator: 'AND' | 'OR';
  rules: IConditionRule[];
  nextStepId: string;
}

/**
 * Condition Rule - Single rule in a condition
 */
export interface IConditionRule {
  field: string;
  operator: ConditionOperator;
  value: string | number | boolean;
}

export enum ConditionOperator {
  Equals = 'eq',
  NotEquals = 'ne',
  Contains = 'contains',
  StartsWith = 'startsWith',
  EndsWith = 'endsWith',
  GreaterThan = 'gt',
  GreaterThanOrEqual = 'gte',
  LessThan = 'lt',
  LessThanOrEqual = 'lte',
  IsEmpty = 'isEmpty',
  IsNotEmpty = 'isNotEmpty',
  In = 'in',
  NotIn = 'notIn',
  DateBefore = 'dateBefore',
  DateAfter = 'dateAfter',
  DateEquals = 'dateEquals'
}

export enum TransitionType {
  Next = 'next',          // Go to next step in order
  Goto = 'goto',          // Jump to specific step
  Branch = 'branch',      // Conditional branching
  Parallel = 'parallel',  // Execute multiple paths
  End = 'end'             // End workflow
}

export enum LogLevel {
  Debug = 'Debug',
  Info = 'Info',
  Warning = 'Warning',
  Error = 'Error'
}

export enum ScheduledActionType {
  ExecuteStep = 'ExecuteStep',
  Reminder = 'Reminder',
  Escalation = 'Escalation',
  SLAWarning = 'SLAWarning',
  SLABreach = 'SLABreach'
}

export enum ScheduledItemStatus {
  Pending = 'Pending',
  Processing = 'Processing',
  Completed = 'Completed',
  Failed = 'Failed',
  Cancelled = 'Cancelled'
}

// ============================================================================
// WORKFLOW DEFINITION INTERFACES
// ============================================================================

/**
 * Workflow Definition - The template/blueprint for a workflow
 */
export interface IWorkflowDefinition extends IBaseListItem {
  // Identification
  WorkflowCode: string;         // Unique code (e.g., "WF-JOIN-STD")
  Description?: string;
  Version: string;              // Semantic version (e.g., "1.0.0")

  // Configuration
  ProcessType: ProcessType;     // Joiner | Mover | Leaver
  IsActive: boolean;
  IsDefault: boolean;           // Default workflow for this process type

  // Trigger Conditions (JSON)
  TriggerConditions?: string;   // JSON: ITriggerCondition[]

  // Steps Definition (JSON)
  Steps: string;                // JSON: IWorkflowStep[]

  // Variables (JSON)
  Variables?: string;           // JSON: IWorkflowVariable[]

  // Metadata
  Category?: string;
  Tags?: string;
  EstimatedDuration?: number;   // In hours

  // Statistics
  TimesUsed?: number;
  AverageCompletionTime?: number;
  SuccessRate?: number;

  // Audit
  PublishedDate?: Date;
  PublishedById?: number;
  PublishedBy?: IUser;
}

/**
 * Workflow Step - A single step in a workflow definition
 */
export interface IWorkflowStep {
  id: string;                   // Unique step ID (e.g., "STEP-001")
  name: string;                 // Display name
  description?: string;
  type: StepType;
  order: number;                // Execution order

  // Configuration specific to step type
  config: IStepConfig;

  // Entry conditions (all must be true to execute)
  conditions?: ICondition[];

  // What happens after this step
  onComplete: ITransition;

  // Timeout handling
  timeoutHours?: number;
  onTimeout?: ITransition;

  // SLA tracking
  sla?: ISLA;

  // Error handling configuration
  errorConfig?: IStepErrorConfig;

  // UI positioning (for designer)
  position?: { x: number; y: number };
}

/**
 * Step Configuration - Varies by step type
 */
export interface IStepConfig {
  // For AssignTasks / CreateTask
  taskTemplateId?: number;      // Reference to JML_Tasks
  taskTitle?: string;
  assigneeField?: string;       // Field name to get assignee (e.g., "ManagerId")
  assigneeId?: number | string; // Fixed assignee (ID or Entra ID)
  assigneeRole?: string;        // Role-based assignment
  assigneeType?: 'role' | 'specific';  // Assignment type (role or specific user)
  assigneeEmail?: string;       // Specific user email (from Entra ID)
  assigneeName?: string;        // Specific user display name
  dueDaysFromNow?: number;
  dueDaysField?: string;        // Field to calculate due date

  // Error handling options for task assignment
  allowEmptyTemplate?: boolean;   // Allow template with 0 tasks (default: false - fails)
  allowUnresolvedRole?: boolean;  // Allow unresolved role to skip task (default: false - fails)

  // For WaitForTasks
  waitForTaskIds?: string[];    // Step IDs that created tasks to wait for
  waitCondition?: 'all' | 'any';

  // For Approval
  approvalTemplateId?: number;
  approverField?: string;
  approverId?: number | string; // Fixed approver (ID or Entra ID)
  approverRole?: string;
  approverSource?: 'field' | 'specific';  // Approver source type
  approverEmail?: string;       // Specific approver email (from Entra ID)
  approverName?: string;        // Specific approver display name

  // For Condition
  conditionGroups?: IConditionGroup[];
  conditionMatch?: 'all' | 'any';           // Expression builder: AND/OR mode
  conditions?: IExpressionCondition[];      // Expression builder: condition rows
  trueBranch?: string;                      // Step ID/name for true path
  falseBranch?: string;                     // Step ID/name for false path

  // For Action
  actionType?: ActionType;
  actionConfig?: IActionConfig;

  // For Notification
  notificationType?: string;
  notificationSubject?: string;   // Email subject template
  recipientField?: string;
  recipientId?: number | string;
  recipientRole?: string;         // Role-based recipient (e.g., "HR Manager", "IT Admin")
  recipientSource?: 'field' | 'specific' | 'role';  // Recipient source type
  recipientEmails?: string[];    // Specific recipient emails (from Entra ID)
  recipientNames?: string[];     // Specific recipient display names
  recipientIds?: (number | string)[];  // Specific recipient IDs
  messageTemplate?: string;

  // For Wait
  waitHours?: number;
  waitUntilField?: string;

  // For WaitForTasks timeout configuration
  timeoutHours?: number;        // Timeout in hours (0 = no timeout)
  slaHours?: number;            // SLA duration in hours (alias for timeoutHours)
  onTimeout?: 'escalate' | 'skip' | 'fail';  // Action on timeout
  escalateToUserIds?: number[]; // User IDs to escalate to
  escalateToEmails?: string[];  // Emails to escalate to

  // For Parallel
  parallelStepIds?: string[];
  failOnAnyError?: boolean;     // Fail parallel execution if any step fails (default: false)

  // For SetVariable
  variableName?: string;
  variableValue?: string | number | boolean;
  variableExpression?: string;  // Dynamic value expression

  // For ForEach (Loop/Iterator)
  collectionPath?: string;        // Path to array variable (e.g., "employees", "tasks")
  itemVariable?: string;          // Variable name for current item (default: "item")
  indexVariable?: string;         // Variable name for current index (default: "index")
  innerSteps?: IWorkflowStep[];   // Steps to execute for each item
  parallelForEach?: boolean;      // Execute items in parallel (default: false)
  maxParallel?: number;           // Max concurrent executions if parallel

  // For CallWorkflow (Sub-workflow)
  subWorkflowCode?: string;       // Workflow code to call
  subWorkflowId?: number;         // Or workflow ID
  inputMappings?: Record<string, string>;   // Map parent vars to child input
  outputMappings?: Record<string, string>;  // Map child output to parent vars
  waitForSubWorkflow?: boolean;   // Wait for completion (default: true)

  // For Webhook
  webhookUrl?: string;
  webhookMethod?: 'GET' | 'POST' | 'PUT' | 'PATCH' | 'DELETE';
  webhookHeaders?: Record<string, string>;
  webhookBodyTemplate?: string;   // JSON template with {{variable}} tokens
  webhookResponseVariable?: string; // Store response in variable
  webhookTimeout?: number;        // Timeout in milliseconds

  // Enhanced error handling (for any step)
  onError?: IStepErrorConfig;
}

/**
 * Error handling configuration for steps
 */
export interface IStepErrorConfig {
  action: 'retry' | 'skip' | 'fail' | 'goto';
  retryCount?: number;            // Max retries (default: 3)
  retryDelayMinutes?: number;     // Delay between retries (default: 5)
  retryBackoffMultiplier?: number; // Exponential backoff (default: 2)
  gotoStepId?: string;            // Step to jump to on error
  notifyOnError?: string[];       // Email addresses to notify
  logLevel?: 'info' | 'warning' | 'error'; // How to log errors
}

/**
 * Condition for conditional logic
 */
export interface ICondition {
  id: string;
  field: string;                // Field to evaluate (e.g., "process.Department")
  operator: ConditionOperator;
  value?: string | number | boolean | Date | string[];
  valueField?: string;          // Compare to another field
}

/**
 * Group of conditions with AND/OR logic
 */
export interface IConditionGroup {
  conditions: ICondition[];
  logic: 'AND' | 'OR';
}

/**
 * Expression condition for visual condition builder
 */
export interface IExpressionCondition {
  id: string;
  field: string;
  operator: string;
  value: string;
}

/**
 * Transition to next step(s)
 */
export interface ITransition {
  type: TransitionType;
  targetStepId?: string;        // For goto
  branches?: IBranch[];         // For branch
  parallelStepIds?: string[];   // For parallel
}

/**
 * Branch in conditional transition
 */
export interface IBranch {
  name: string;
  conditions: IConditionGroup[];
  targetStepId: string;
  isDefault?: boolean;          // Fallback if no conditions match
}

/**
 * SLA configuration
 */
export interface ISLA {
  warningHours: number;
  breachHours: number;
  escalateTo?: string;          // Field or role
  escalateToId?: number;
}

/**
 * Trigger condition for auto-selecting workflow
 */
export interface ITriggerCondition {
  conditions: IConditionGroup[];
  priority: number;             // Higher priority = checked first
}

/**
 * Workflow variable
 */
export interface IWorkflowVariable {
  name: string;
  type: 'string' | 'number' | 'boolean' | 'date' | 'array';
  defaultValue?: string | number | boolean | Date;
  description?: string;
}

// ============================================================================
// ACTION CONFIGURATION INTERFACES
// ============================================================================

/**
 * Configuration for action execution
 */
export interface IActionConfig {
  // For UpdateListItem / CreateListItem
  listName?: string;
  itemId?: number;
  itemIdField?: string;         // Get ID from process field
  updates?: IFieldUpdate[];

  // For SendEmail
  to?: string[];
  toField?: string;
  cc?: string[];
  subject?: string;
  body?: string;
  templateId?: string;

  // For SendNotification
  recipientId?: number;         // Single recipient
  recipientIds?: number[];      // Multiple recipients
  recipientField?: string;
  notificationType?: string;
  message?: string;
  priority?: string;

  // For CallWebhook
  url?: string;
  method?: 'GET' | 'POST' | 'PUT' | 'PATCH';
  headers?: Record<string, string>;
  bodyTemplate?: string;

  // For SendTeamsMessage
  channelId?: string;
  teamId?: string;
  messageContent?: string;

  // For SLA/Escalation
  escalateTo?: string;          // User/role to escalate to

  // For Azure AD Actions
  userId?: string;              // User ID (Entra ID object ID)
  userIdField?: string;         // Get user ID from process field
  userEmail?: string;           // User email for lookup
  userEmailField?: string;      // Get user email from process field
  groupId?: string;             // Azure AD group ID
  groupIds?: string[];          // Multiple group IDs
  groupIdField?: string;        // Get group ID from process field
  groupName?: string;           // Azure AD group name for lookup
  groupNames?: string[];        // Multiple group names
  profileUpdates?: IAzureADProfileUpdate[];  // Profile field updates

  // For Calendar Actions
  calendarUserId?: string;      // Calendar owner (defaults to process owner)
  calendarUserIdField?: string; // Get calendar owner from process field
  eventTitle?: string;
  eventTitleTemplate?: string;  // Template with placeholders
  eventDescription?: string;
  eventDescriptionTemplate?: string;
  eventStartDate?: string;      // ISO date or expression
  eventStartDateField?: string; // Get start date from process field
  eventEndDate?: string;
  eventEndDateField?: string;
  eventDuration?: number;       // Duration in minutes
  eventLocation?: string;
  eventAttendees?: string[];    // Email addresses
  eventAttendeesField?: string; // Get attendees from process field
  eventIsOnline?: boolean;      // Create Teams meeting
  eventReminder?: number;       // Minutes before event

  // For Equipment/Asset Actions
  equipmentType?: string;       // Type of equipment (laptop, phone, badge, etc.)
  equipmentTypes?: string[];    // Multiple equipment types
  equipmentDescription?: string;
  assetTag?: string;           // Specific asset tag
  assetTagField?: string;      // Get asset tag from process field
  assetCategory?: string;      // Asset category filter
  licenseType?: string;        // License type (e.g., Office 365, Salesforce)
  licenseTypes?: string[];     // Multiple license types
  returnDeadline?: string;     // ISO date for asset return deadline
  returnDeadlineField?: string;

  // Phase 4: User Provisioning options
  displayName?: string;        // Display name for new user
  userPrincipalName?: string;  // UPN for new user
  mailNickname?: string;       // Mail nickname for new user
  password?: string;           // Initial password (optional - auto-generated if not provided)
  generatePassword?: boolean;  // Whether to auto-generate password (default: true)
  forceChangePasswordNextSignIn?: boolean;  // Force password change on first login
  accountEnabled?: boolean;    // Enable account on creation (default: true)
  department?: string;         // Department for new user
  jobTitle?: string;           // Job title for new user
  officeLocation?: string;     // Office location for new user
  manager?: string;            // Manager user ID or UPN
  usageLocation?: string;      // Usage location for licensing (e.g., 'GB', 'US')

  // Phase 4: License Management options
  licenseSkuId?: string;       // Single license SKU ID to assign
  licenseSkuIds?: string[];    // Multiple license SKU IDs
  licenseSkuIdField?: string;  // Field containing license SKU ID(s)
  disabledPlans?: string[];    // Service plans to disable within the license
  revokeAllLicenses?: boolean; // Revoke all licenses (for leaver workflows)
}

/**
 * Field update for list operations
 */
export interface IFieldUpdate {
  fieldName: string;
  value?: string | number | boolean | Date;
  valueField?: string;          // Get value from process field
  expression?: string;          // Dynamic expression
}

/**
 * Azure AD profile update configuration
 */
export interface IAzureADProfileUpdate {
  property: string;             // Azure AD property name (e.g., 'department', 'jobTitle')
  value?: string;               // Static value
  valueField?: string;          // Get value from process field
}

// ============================================================================
// WORKFLOW INSTANCE INTERFACES
// ============================================================================

/**
 * Workflow Instance - A running instance of a workflow
 */
export interface IWorkflowInstance extends IBaseListItem {
  // References
  WorkflowDefinitionId: number;
  WorkflowDefinition?: IWorkflowDefinition;
  ProcessId: number;            // Link to JML_Processes
  ProcessType?: ProcessType;    // Cached from process for notifications

  // Status
  Status: WorkflowInstanceStatus;
  CurrentStepId?: string;
  CurrentStepName?: string;

  // Progress
  TotalSteps: number;
  CompletedSteps: number;
  ProgressPercentage: number;

  // Timing
  StartedDate?: Date;
  CompletedDate?: Date;
  EstimatedCompletionDate?: Date;

  // Runtime data (JSON)
  Variables?: string;           // JSON: Record<string, any>
  Context?: string;             // JSON: IWorkflowContext

  // Error handling
  ErrorMessage?: string;
  ErrorStepId?: string;
  RetryCount?: number;
  LastRetryDate?: Date;

  // Audit
  StartedById?: number;
  StartedBy?: IUser;
  CompletedById?: number;
  CompletedBy?: IUser;
}

/**
 * Workflow execution context
 */
export interface IWorkflowContext {
  processId: number;
  processType: ProcessType;
  employeeName: string;
  employeeEmail: string;
  department: string;
  managerId?: number;
  managerEmail?: string;
  startedBy: string;
  startedAt: Date;
  customFields?: Record<string, unknown>;
}

/**
 * Step Status - Execution status of a step in an instance
 */
export interface IWorkflowStepStatus extends IBaseListItem {
  // References
  WorkflowInstanceId: number;
  StepId: string;
  StepName: string;

  // Status
  Status: StepStatus;
  Order: number;

  // Timing
  StartedDate?: Date;
  CompletedDate?: Date;
  Duration?: number;            // In minutes

  // Results (JSON)
  Result?: string;              // JSON: IStepResult
  OutputVariables?: string;     // JSON: Record<string, any>

  // Error handling
  ErrorMessage?: string;
  RetryCount?: number;

  // Related items
  TaskAssignmentIds?: string;   // Comma-separated IDs
  ApprovalId?: number;
}

/**
 * Result of step execution
 */
export interface IStepResult {
  success: boolean;
  message?: string;
  data?: Record<string, unknown>;
  createdItemIds?: number[];
  nextStepId?: string;
}

// ============================================================================
// WORKFLOW LOG INTERFACES
// ============================================================================

/**
 * Workflow execution log entry
 */
export interface IWorkflowLog extends IBaseListItem {
  WorkflowInstanceId: number;
  StepId?: string;
  StepName?: string;
  Action: string;
  Level: LogLevel;
  Message: string;
  Details?: string;             // JSON: additional details
  Timestamp: Date;
  UserId?: number;
}

// ============================================================================
// SCHEDULED ITEM INTERFACES
// ============================================================================

/**
 * Scheduled workflow item
 */
export interface IWorkflowScheduledItem extends IBaseListItem {
  WorkflowInstanceId: number;
  StepId: string;
  ActionType: ScheduledActionType;
  ScheduledDate: Date;
  Status: ScheduledItemStatus;
  ProcessedDate?: Date;
  ErrorMessage?: string;
  RetryCount?: number;
  ActionConfig?: string | IActionConfig;  // Object in code, JSON string when stored in SharePoint
}

// ============================================================================
// SERVICE INTERFACES
// ============================================================================

/**
 * Context passed to action handlers
 */
export interface IActionContext {
  workflowInstance: IWorkflowInstance;
  currentStep: IWorkflowStep;
  stepStatus: IWorkflowStepStatus;
  process: Record<string, unknown>;  // IJmlProcess data
  variables: Record<string, unknown>;
  services: IServiceContainer;
}

/**
 * Result from action handler
 */
export interface IActionResult {
  success: boolean;
  error?: string;
  outputVariables?: Record<string, unknown>;
  createdItemIds?: number[];
  nextAction?: 'continue' | 'wait' | 'retry' | 'fail';
  waitForItemType?: 'task' | 'approval';
  waitForItemIds?: number[];
}

/**
 * Service container for dependency injection
 */
export interface IServiceContainer {
  sp: unknown;                  // SPFI
  context: unknown;             // WebPartContext
  getService<T>(serviceName: string): T;
}

/**
 * Validation result
 */
export interface IValidationResult {
  valid: boolean;
  errors: IValidationError[];
  warnings?: IValidationError[];
}

/**
 * Validation error
 */
export interface IValidationError {
  field?: string;
  stepId?: string;
  code: string;
  message: string;
}

// ============================================================================
// WORKFLOW DEFINITION VIEW MODELS
// ============================================================================

/**
 * Summary view of workflow definition
 */
export interface IWorkflowDefinitionSummary {
  Id: number;
  Title: string;
  WorkflowCode: string;
  ProcessType: ProcessType;
  Version: string;
  IsActive: boolean;
  IsDefault: boolean;
  StepCount: number;
  TimesUsed: number;
  SuccessRate?: number;
}

/**
 * Summary view of workflow instance
 */
export interface IWorkflowInstanceSummary {
  Id: number;
  WorkflowName: string;
  ProcessId: number;
  EmployeeName: string;
  ProcessType: ProcessType;
  Status: WorkflowInstanceStatus;
  CurrentStepName?: string;
  ProgressPercentage: number;
  StartedDate?: Date;
  IsOverdue: boolean;
}

// ============================================================================
// UI VIEW MODELS (camelCase for React components)
// ============================================================================

/**
 * Workflow definition for UI components (Workflow Designer)
 * Uses camelCase naming convention for React compatibility
 */
export interface IWorkflowDefinitionUI {
  id: string;
  name: string;
  description?: string;
  processType: ProcessType;
  version: number;
  isActive: boolean;
  isDraft: boolean;
  steps: IWorkflowStepUI[];
  createdBy: string;
  createdDate: string;
  modifiedBy: string;
  modifiedDate: string;
}

// Alias for backward compatibility
export type IWorkflowDefinition_UI = IWorkflowDefinitionUI;

/**
 * Workflow step for UI components (Workflow Designer)
 */
export interface IWorkflowStepUI {
  id: string;
  name: string;
  description?: string;
  type: StepType;
  order: number;
  isRequired?: boolean;
  actions?: IWorkflowAction[];
  conditions?: IWorkflowCondition[];
  nextStepId?: string;

  // SLA settings
  slaWarningHours?: number;
  slaBreachHours?: number;

  // Error handling
  onError?: 'stop' | 'continue' | 'retry' | 'goto';
  maxRetries?: number;
  errorStepId?: string;

  // Delay settings
  delayHours?: number;

  // Approval settings
  approverType?: string;
  approver?: string;
  allowDelegation?: boolean;
}

/**
 * Condition rule for UI (supports string operators)
 */
export interface IConditionRuleUI {
  id: string;
  field: string;
  operator: string;
  value?: string | number | boolean;
  customField?: string;
}

/**
 * Workflow instance for UI components (Workflow Monitor)
 * Uses camelCase naming convention for React compatibility
 */
export interface IWorkflowInstanceUI {
  id: string;
  workflowDefinitionId: string;
  workflowName: string;
  processType: ProcessType;
  status: WorkflowInstanceStatusUI;

  // Employee info
  employeeId?: number;
  employeeName?: string;
  employeeEmail?: string;

  // Progress
  currentStepId?: string;
  currentStepName?: string;
  currentStepIndex?: number;
  totalSteps?: number;

  // Timing
  startedDate: string;
  completedDate?: string;
  estimatedCompletionDate?: string;

  // SLA
  slaStatus?: 'Healthy' | 'Warning' | 'Breached';

  // Metadata
  startedBy?: string;
}

/**
 * Step execution record for UI (Workflow Monitor details)
 */
export interface IStepExecutionUI {
  id: string;
  stepId: string;
  stepName: string;
  status: 'Pending' | 'InProgress' | 'Completed' | 'Failed' | 'Skipped';
  startedDate?: string;
  completedDate?: string;
  assignedTo?: string;
  errorMessage?: string;
}

/**
 * Simplified status enum for UI components
 */
export type WorkflowInstanceStatusUI =
  | 'NotStarted'
  | 'InProgress'
  | 'WaitingApproval'
  | 'Completed'
  | 'Failed'
  | 'Cancelled';

// Aliases for UI components - maps to the UI types
export type IWorkflowInstance_UI = IWorkflowInstanceUI;
export type IStepExecution = IStepExecutionUI;
