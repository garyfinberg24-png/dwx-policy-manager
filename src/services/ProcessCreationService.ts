// @ts-nocheck
/**
 * ProcessCreationService
 *
 * The critical missing link between Process Wizard and Task Generation.
 * This service handles the complete process creation workflow:
 * 1. Create process record
 * 2. Generate task assignments from templates
 * 3. Start workflow if configured
 * 4. Send notifications
 *
 * @author JML Development Team
 * @version 1.0.0
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/batching';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import {
  ProcessType,
  ProcessStatus,
  TaskStatus,
  Priority,
  TaskCategory,
  IUser
} from '../models/ICommon';
import { IJmlProcess, IJmlProcessForm } from '../models/IJmlProcess';
import { IJmlTask } from '../models/IJmlTask';
import { IJmlTaskAssignment } from '../models/IJmlTaskAssignment';
import { WorkflowEngineService, IStartWorkflowOptions } from './workflow/WorkflowEngineService';
import { WorkflowInstanceStatus } from '../models/IWorkflow';
import { logger } from './LoggingService';
import {
  retryWithDLQ,
  PROCESS_SYNC_RETRY_OPTIONS,
  workflowSyncDLQ,
  IRetryResult
} from '../utils/retryUtils';

// ============================================================================
// INTERFACES
// ============================================================================

/**
 * Input from Process Wizard
 */
export interface IProcessWizardInput {
  // Process Details
  processType: ProcessType;
  employeeName: string;
  employeeEmail: string;
  employeeId?: string;
  department: string;
  jobTitle: string;
  location: string;

  // Manager
  managerId?: number;
  managerEmail?: string;

  // Timing
  startDate: Date;
  targetCompletionDate: Date;

  // Priority & Options
  priority: Priority;

  // Template Selection (from Task Library)
  selectedTaskIds?: number[];          // Specific tasks selected
  taskTemplateId?: number;             // Checklist template ID
  useDefaultTasks?: boolean;           // Use default tasks for process type

  // Workflow
  workflowDefinitionId?: number;       // Specific workflow to use
  workflowDefinitionCode?: string;     // Or workflow by code
  autoStartWorkflow?: boolean;         // Default: true

  // Additional
  comments?: string;
  businessUnit?: string;
  costCenter?: string;
  contractType?: string;
  customFields?: Record<string, unknown>;

  // Creator context
  createdByUserId: number;
  createdByUserName: string;
}

/**
 * Result of process creation
 */
export interface IProcessCreationResult {
  success: boolean;
  processId?: number;
  process?: IJmlProcess;
  taskAssignments?: IJmlTaskAssignment[];
  workflowInstanceId?: number;
  errors?: string[];
  warnings?: string[];
}

/**
 * Task assignment input for batch creation
 */
interface ITaskAssignmentInput {
  taskId: number;
  taskTitle: string;
  taskCode: string;
  category: TaskCategory;
  assigneeId: number;
  dueDate: Date;
  priority: Priority;
  requiresApproval: boolean;
  approverId?: number;
  dependsOnTaskId?: number;
  slaHours?: number;
  estimatedHours?: number;
  order: number;
}

/**
 * Default task configuration by process type
 */
interface IDefaultTaskConfig {
  processType: ProcessType;
  categories: TaskCategory[];
  minTasks: number;
}

// ============================================================================
// CONSTANTS
// ============================================================================

const LIST_NAMES = {
  PROCESSES: 'JML_Processes',
  TASKS: 'JML_Tasks',
  TASK_ASSIGNMENTS: 'JML_TaskAssignments',
  CHECKLIST_TEMPLATES: 'JML_ChecklistTemplates',
  TEMPLATE_TASK_MAPPINGS: 'JML_TemplateTaskMappings'
};

const DEFAULT_TASK_CONFIG: IDefaultTaskConfig[] = [
  {
    processType: ProcessType.Joiner,
    categories: [
      TaskCategory.ITAccess,
      TaskCategory.ITEquipment,
      TaskCategory.HROnboarding,
      TaskCategory.FacilitiesAccess,
      TaskCategory.TrainingOrientation
    ],
    minTasks: 5
  },
  {
    processType: ProcessType.Mover,
    categories: [
      TaskCategory.ITAccess,
      TaskCategory.HRDocumentation,
      TaskCategory.FacilitiesAccess
    ],
    minTasks: 3
  },
  {
    processType: ProcessType.Leaver,
    categories: [
      TaskCategory.ITAccess,
      TaskCategory.ITEquipment,
      TaskCategory.HROffboarding,
      TaskCategory.FacilitiesAccess,
      TaskCategory.SecurityCompliance
    ],
    minTasks: 5
  }
];

// ============================================================================
// SERVICE CLASS
// ============================================================================

export class ProcessCreationService {
  private sp: SPFI;
  private context: WebPartContext;
  private workflowEngine: WorkflowEngineService;

  constructor(sp: SPFI, context: WebPartContext) {
    this.sp = sp;
    this.context = context;
    this.workflowEngine = new WorkflowEngineService(sp, context);

    // Register process status sync callback
    this.workflowEngine.onProcessStatusSync(this.syncProcessStatusFromWorkflow.bind(this));
  }

  // ============================================================================
  // MAIN ENTRY POINT
  // ============================================================================

  /**
   * Create a complete process with tasks and optionally start workflow
   * This is the main entry point called by the Process Wizard
   */
  public async createProcessWithTasks(
    input: IProcessWizardInput
  ): Promise<IProcessCreationResult> {
    const errors: string[] = [];
    const warnings: string[] = [];

    try {
      logger.info('ProcessCreationService', 'Starting process creation', {
        processType: input.processType,
        employeeName: input.employeeName
      });

      // Step 1: Validate input
      const validationErrors = this.validateInput(input);
      if (validationErrors.length > 0) {
        return {
          success: false,
          errors: validationErrors
        };
      }

      // Step 2: Create process record
      const process = await this.createProcessRecord(input);
      if (!process.Id) {
        throw new Error('Failed to create process record');
      }

      logger.info('ProcessCreationService', `Process created with ID: ${process.Id}`);

      // Step 3: Determine tasks to assign
      const tasksToAssign = await this.determineTasksToAssign(input);
      if (tasksToAssign.length === 0) {
        warnings.push('No tasks were assigned. Consider selecting a template or tasks.');
      }

      // Step 4: Create task assignments
      let taskAssignments: IJmlTaskAssignment[] = [];
      if (tasksToAssign.length > 0) {
        taskAssignments = await this.createTaskAssignments(
          process.Id,
          tasksToAssign,
          input
        );

        // Update process with task counts
        await this.updateProcessTaskCounts(process.Id, taskAssignments.length);

        logger.info('ProcessCreationService', `Created ${taskAssignments.length} task assignments`);
      }

      // Step 5: Start workflow (if enabled)
      let workflowInstanceId: number | undefined;
      const autoStartWorkflow = input.autoStartWorkflow !== false; // Default true

      if (autoStartWorkflow) {
        try {
          const workflowResult = await this.startWorkflow(process.Id, input);
          workflowInstanceId = workflowResult.instanceId;

          logger.info('ProcessCreationService', `Workflow started: ${workflowInstanceId}`);
        } catch (workflowError) {
          // Log but don't fail - workflow can be started later
          logger.warn('ProcessCreationService', 'Failed to auto-start workflow', workflowError);
          warnings.push('Workflow could not be started automatically. Please start manually.');
        }
      }

      // Step 6: Send launch notifications
      try {
        await this.sendLaunchNotifications(process, taskAssignments, input);
      } catch (notifyError) {
        logger.warn('ProcessCreationService', 'Failed to send notifications', notifyError);
        warnings.push('Some notifications could not be sent.');
      }

      // Step 7: Log audit entry
      await this.logProcessCreation(process.Id, input);

      return {
        success: true,
        processId: process.Id,
        process,
        taskAssignments,
        workflowInstanceId,
        warnings: warnings.length > 0 ? warnings : undefined
      };

    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      logger.error('ProcessCreationService', 'Process creation failed', error);

      errors.push(`Process creation failed: ${errorMessage}`);

      return {
        success: false,
        errors
      };
    }
  }

  // ============================================================================
  // VALIDATION
  // ============================================================================

  /**
   * Validate wizard input
   */
  private validateInput(input: IProcessWizardInput): string[] {
    const errors: string[] = [];

    // Required fields
    if (!input.processType) {
      errors.push('Process type is required');
    }
    if (!input.employeeName?.trim()) {
      errors.push('Employee name is required');
    }
    if (!input.employeeEmail?.trim()) {
      errors.push('Employee email is required');
    }
    if (!input.department?.trim()) {
      errors.push('Department is required');
    }
    if (!input.startDate) {
      errors.push('Start date is required');
    }
    if (!input.targetCompletionDate) {
      errors.push('Target completion date is required');
    }

    // Date validation
    if (input.startDate && input.targetCompletionDate) {
      if (new Date(input.startDate) > new Date(input.targetCompletionDate)) {
        errors.push('Start date cannot be after target completion date');
      }
    }

    // Email validation
    if (input.employeeEmail && !this.isValidEmail(input.employeeEmail)) {
      errors.push('Invalid employee email format');
    }

    return errors;
  }

  private isValidEmail(email: string): boolean {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
  }

  // ============================================================================
  // PROCESS CREATION
  // ============================================================================

  /**
   * Create the process record in SharePoint
   */
  private async createProcessRecord(input: IProcessWizardInput): Promise<IJmlProcess> {
    const processData: Partial<IJmlProcess> = {
      Title: `${input.processType} - ${input.employeeName}`,
      ProcessType: input.processType,
      ProcessStatus: ProcessStatus.InProgress,
      EmployeeName: input.employeeName,
      EmployeeEmail: input.employeeEmail,
      EmployeeID: input.employeeId,
      Department: input.department,
      JobTitle: input.jobTitle || '',
      Location: input.location || '',
      ManagerId: input.managerId,
      StartDate: input.startDate,
      TargetCompletionDate: input.targetCompletionDate,
      Priority: input.priority || Priority.Medium,
      ChecklistTemplateID: input.taskTemplateId?.toString(),
      Comments: input.comments,
      BusinessUnit: input.businessUnit,
      CostCenter: input.costCenter,
      ContractType: input.contractType,
      TotalTasks: 0,
      CompletedTasks: 0,
      ProgressPercentage: 0,
      OverdueTasks: 0,
      ApprovalRequired: false,
      IsDeleted: false,
      CustomFields: input.customFields ? JSON.stringify(input.customFields) : undefined
    };

    const result = await this.sp.web.lists
      .getByTitle(LIST_NAMES.PROCESSES)
      .items
      .add(processData);

    return {
      ...processData,
      Id: result.data.Id
    } as IJmlProcess;
  }

  // ============================================================================
  // TASK DETERMINATION
  // ============================================================================

  /**
   * Determine which tasks to assign based on wizard input
   */
  private async determineTasksToAssign(
    input: IProcessWizardInput
  ): Promise<IJmlTask[]> {
    let tasks: IJmlTask[] = [];

    // Priority 1: Specific tasks selected
    if (input.selectedTaskIds && input.selectedTaskIds.length > 0) {
      tasks = await this.getTasksByIds(input.selectedTaskIds);
    }
    // Priority 2: Template selected
    else if (input.taskTemplateId) {
      tasks = await this.getTasksFromTemplate(input.taskTemplateId);
    }
    // Priority 3: Use default tasks for process type
    else if (input.useDefaultTasks !== false) {
      tasks = await this.getDefaultTasksForProcessType(input.processType);
    }

    // Filter to only active tasks
    return tasks.filter(t => t.IsActive !== false);
  }

  /**
   * Get tasks by specific IDs
   */
  private async getTasksByIds(taskIds: number[]): Promise<IJmlTask[]> {
    if (taskIds.length === 0) return [];

    const filterQuery = taskIds.map(id => `Id eq ${id}`).join(' or ');

    const items = await this.sp.web.lists
      .getByTitle(LIST_NAMES.TASKS)
      .items
      .filter(filterQuery)
      .select(
        'Id', 'Title', 'TaskCode', 'Category', 'Description', 'Instructions',
        'Department', 'DefaultAssigneeId', 'AssigneeRole', 'SLAHours',
        'EstimatedHours', 'RequiresApproval', 'ApproverRole', 'DependsOn',
        'BlockingTask', 'IsActive', 'Priority', 'Tags'
      )();

    return items as IJmlTask[];
  }

  /**
   * Get tasks from a checklist template
   */
  private async getTasksFromTemplate(templateId: number): Promise<IJmlTask[]> {
    // Get task mappings for this template
    const mappings = await this.sp.web.lists
      .getByTitle(LIST_NAMES.TEMPLATE_TASK_MAPPINGS)
      .items
      .filter(`ChecklistTemplateId eq ${templateId}`)
      .select('TaskId', 'Order', 'IsRequired')
      .orderBy('Order')();

    if (mappings.length === 0) return [];

    // Get the actual tasks
    const taskIds = mappings.map((m: { TaskId: number }) => m.TaskId);
    return this.getTasksByIds(taskIds);
  }

  /**
   * Get default tasks for a process type
   */
  private async getDefaultTasksForProcessType(
    processType: ProcessType
  ): Promise<IJmlTask[]> {
    const config = DEFAULT_TASK_CONFIG.find(c => c.processType === processType);
    if (!config) return [];

    // Build category filter
    const categoryFilter = config.categories
      .map(cat => `Category eq '${cat}'`)
      .join(' or ');

    const items = await this.sp.web.lists
      .getByTitle(LIST_NAMES.TASKS)
      .items
      .filter(`IsActive eq 1 and (${categoryFilter})`)
      .select(
        'Id', 'Title', 'TaskCode', 'Category', 'Description', 'Instructions',
        'Department', 'DefaultAssigneeId', 'AssigneeRole', 'SLAHours',
        'EstimatedHours', 'RequiresApproval', 'ApproverRole', 'DependsOn',
        'BlockingTask', 'IsActive', 'Priority', 'Tags'
      )
      .orderBy('Category')
      .top(50)();

    return items as IJmlTask[];
  }

  // ============================================================================
  // TASK ASSIGNMENT CREATION
  // ============================================================================

  /**
   * Create task assignments for all determined tasks
   */
  private async createTaskAssignments(
    processId: number,
    tasks: IJmlTask[],
    input: IProcessWizardInput
  ): Promise<IJmlTaskAssignment[]> {
    const assignments: IJmlTaskAssignment[] = [];
    const startDate = new Date(input.startDate);
    const targetDate = new Date(input.targetCompletionDate);

    // Calculate days available for task distribution
    const totalDays = Math.ceil(
      (targetDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24)
    );
    const daysPerTask = Math.max(1, Math.floor(totalDays / tasks.length));

    // Prepare task assignment inputs
    const taskInputs: ITaskAssignmentInput[] = [];

    for (let i = 0; i < tasks.length; i++) {
      const task = tasks[i];

      // Calculate due date based on order and SLA
      let dueDate: Date;
      if (task.SLAHours) {
        // Use SLA hours from task
        dueDate = new Date(startDate);
        dueDate.setHours(dueDate.getHours() + task.SLAHours);
      } else {
        // Distribute evenly across timeline
        dueDate = new Date(startDate);
        dueDate.setDate(dueDate.getDate() + (i + 1) * daysPerTask);

        // Ensure due date doesn't exceed target
        if (dueDate > targetDate) {
          dueDate = new Date(targetDate);
        }
      }

      // Determine assignee
      const assigneeId = await this.resolveAssignee(task, input);

      taskInputs.push({
        taskId: task.Id!,
        taskTitle: task.Title,
        taskCode: task.TaskCode,
        category: task.Category,
        assigneeId,
        dueDate,
        priority: task.Priority || input.priority || Priority.Medium,
        requiresApproval: task.RequiresApproval || false,
        approverId: input.managerId, // Default approver is manager
        dependsOnTaskId: undefined, // Will be resolved after all created
        slaHours: task.SLAHours,
        estimatedHours: task.EstimatedHours,
        order: i + 1
      });
    }

    // Create assignments using batching for performance
    const batchSize = 20;
    for (let i = 0; i < taskInputs.length; i += batchSize) {
      const batch = taskInputs.slice(i, i + batchSize);
      const batchAssignments = await this.createTaskAssignmentBatch(processId, batch);
      assignments.push(...batchAssignments);
    }

    // Resolve dependencies after all tasks created
    await this.resolveDependencies(assignments, tasks);

    return assignments;
  }

  /**
   * Create a batch of task assignments
   */
  private async createTaskAssignmentBatch(
    processId: number,
    inputs: ITaskAssignmentInput[]
  ): Promise<IJmlTaskAssignment[]> {
    const list = this.sp.web.lists.getByTitle(LIST_NAMES.TASK_ASSIGNMENTS);
    const [batchedSp, execute] = this.sp.batched();
    const results: Promise<{ data: { Id: number } }>[] = [];

    for (const input of inputs) {
      const assignmentData = {
        Title: input.taskTitle,
        ProcessID: processId.toString(), // Text field until lookup conversion
        TaskID: input.taskId.toString(),  // Text field until lookup conversion
        AssignedToId: input.assigneeId,
        AssignedDate: new Date(),
        DueDate: input.dueDate,
        StartDate: new Date(),
        Status: TaskStatus.NotStarted,
        Priority: input.priority,
        PercentComplete: 0,
        RequiresApproval: input.requiresApproval,
        ApproverId: input.approverId,
        SLAHours: input.slaHours,
        IsDependentTask: false,
        IsBlocked: false,
        ReminderSent: false,
        EscalationSent: false,
        EscalationLevel: 0,
        IsDeleted: false
      };

      const promise = batchedSp.web.lists
        .getByTitle(LIST_NAMES.TASK_ASSIGNMENTS)
        .items
        .add(assignmentData);

      results.push(promise as unknown as Promise<{ data: { Id: number } }>);
    }

    await execute();

    // Wait for all results
    const resolvedResults = await Promise.all(results);

    return inputs.map((input, index) => ({
      Id: resolvedResults[index].data.Id,
      Title: input.taskTitle,
      ProcessID: processId.toString(),
      TaskID: input.taskId.toString(),
      AssignedToId: input.assigneeId,
      DueDate: input.dueDate,
      Status: TaskStatus.NotStarted,
      Priority: input.priority
    } as IJmlTaskAssignment));
  }

  /**
   * Resolve assignee for a task
   */
  private async resolveAssignee(
    task: IJmlTask,
    input: IProcessWizardInput
  ): Promise<number> {
    // Priority 1: Task has default assignee
    if (task.DefaultAssigneeId) {
      return task.DefaultAssigneeId;
    }

    // Priority 2: Resolve by role
    if (task.AssigneeRole) {
      const assignee = await this.resolveAssigneeByRole(task.AssigneeRole, input);
      if (assignee) return assignee;
    }

    // Priority 3: Use manager
    if (input.managerId) {
      return input.managerId;
    }

    // Fallback: Use creator
    return input.createdByUserId;
  }

  /**
   * Resolve assignee by role name
   */
  private async resolveAssigneeByRole(
    role: string,
    input: IProcessWizardInput
  ): Promise<number | undefined> {
    // Role mapping - in a real implementation, this would query a roles list
    const roleMap: Record<string, () => number | undefined> = {
      'Manager': () => input.managerId,
      'HR Admin': () => undefined, // Would query HR group
      'IT Admin': () => undefined, // Would query IT group
      'Facilities': () => undefined, // Would query Facilities group
      'Process Owner': () => input.createdByUserId
    };

    const resolver = roleMap[role];
    return resolver ? resolver() : undefined;
  }

  /**
   * Resolve task dependencies after creation
   */
  private async resolveDependencies(
    assignments: IJmlTaskAssignment[],
    tasks: IJmlTask[]
  ): Promise<void> {
    const taskCodeToAssignmentId = new Map<string, number>();

    // Build mapping of task codes to assignment IDs
    for (let i = 0; i < tasks.length; i++) {
      const task = tasks[i];
      const assignment = assignments[i];
      if (task.TaskCode && assignment.Id) {
        taskCodeToAssignmentId.set(task.TaskCode, assignment.Id);
      }
    }

    // Update dependencies
    for (let i = 0; i < tasks.length; i++) {
      const task = tasks[i];
      const assignment = assignments[i];

      if (task.DependsOn && assignment.Id) {
        const dependsOnAssignmentId = taskCodeToAssignmentId.get(task.DependsOn);

        if (dependsOnAssignmentId) {
          await this.sp.web.lists
            .getByTitle(LIST_NAMES.TASK_ASSIGNMENTS)
            .items
            .getById(assignment.Id)
            .update({
              IsDependentTask: true,
              DependsOnTaskId: dependsOnAssignmentId,
              IsBlocked: true,
              BlockedReason: `Waiting for ${task.DependsOn} to complete`
            });
        }
      }
    }
  }

  /**
   * Update process with task counts
   */
  private async updateProcessTaskCounts(
    processId: number,
    totalTasks: number
  ): Promise<void> {
    await this.sp.web.lists
      .getByTitle(LIST_NAMES.PROCESSES)
      .items
      .getById(processId)
      .update({
        TotalTasks: totalTasks,
        CompletedTasks: 0,
        ProgressPercentage: 0,
        OverdueTasks: 0
      });
  }

  // ============================================================================
  // WORKFLOW INTEGRATION
  // ============================================================================

  /**
   * Start workflow for the process
   */
  private async startWorkflow(
    processId: number,
    input: IProcessWizardInput
  ): Promise<{ instanceId: number }> {
    const workflowOptions: IStartWorkflowOptions = {
      definitionId: input.workflowDefinitionId,
      definitionCode: input.workflowDefinitionCode,
      processId,
      processType: input.processType,
      employeeName: input.employeeName,
      employeeEmail: input.employeeEmail,
      department: input.department,
      managerId: input.managerId,
      managerEmail: input.managerEmail,
      startedByUserId: input.createdByUserId,
      startedByUserName: input.createdByUserName,
      customContext: {
        jobTitle: input.jobTitle,
        location: input.location,
        businessUnit: input.businessUnit,
        costCenter: input.costCenter,
        ...input.customFields
      }
    };

    const result = await this.workflowEngine.startWorkflow(workflowOptions);

    if (!result.success) {
      throw new Error(result.error || 'Failed to start workflow');
    }

    return { instanceId: result.instanceId };
  }

  /**
   * Sync process status from workflow status changes
   * Enhanced with retry logic and dead letter queue for reliability
   */
  private async syncProcessStatusFromWorkflow(
    processId: number,
    workflowStatus: string,
    workflowInstanceId: number
  ): Promise<void> {
    // Map workflow status to process status
    const statusMap: Record<string, ProcessStatus> = {
      'Running': ProcessStatus.InProgress,
      'Completed': ProcessStatus.Completed,
      'Failed': ProcessStatus.OnHold,
      'Cancelled': ProcessStatus.Cancelled,
      'Paused': ProcessStatus.OnHold,
      'Waiting for Task': ProcessStatus.InProgress,
      'Waiting for Approval': ProcessStatus.InProgress,
      'Waiting for Input': ProcessStatus.InProgress
    };

    const processStatus = statusMap[workflowStatus];
    if (!processStatus) {
      logger.warn('ProcessCreationService', `Unknown workflow status: ${workflowStatus}`);
      return;
    }

    const syncPayload = {
      processId,
      workflowStatus,
      workflowInstanceId,
      processStatus,
      timestamp: new Date().toISOString()
    };

    const result = await retryWithDLQ<void>(
      async () => {
        await this.sp.web.lists
          .getByTitle(LIST_NAMES.PROCESSES)
          .items
          .getById(processId)
          .update({
            ProcessStatus: processStatus,
            WorkflowInstanceId: workflowInstanceId,
            WorkflowStatus: workflowStatus,
            ...(processStatus === ProcessStatus.Completed ? { ActualCompletionDate: new Date() } : {})
          });
      },
      'workflow-to-process-sync',
      syncPayload,
      PROCESS_SYNC_RETRY_OPTIONS,
      workflowSyncDLQ,
      {
        source: 'ProcessCreationService',
        operation: 'syncProcessStatusFromWorkflow'
      }
    );

    if (!result.success) {
      logger.error(
        'ProcessCreationService',
        `Failed to sync process ${processId} from workflow ${workflowInstanceId} after ${result.attempts} attempts. DLQ ID: ${result.deadLetterItemId}`,
        result.error
      );
    } else {
      logger.info(
        'ProcessCreationService',
        `Process ${processId} synced to ${processStatus} from workflow status ${workflowStatus} (${result.attempts} attempt(s))`
      );
    }
  }

  /**
   * Sync workflow status when process status changes (bidirectional sync)
   * Call this when process status is updated directly (not via workflow)
   */
  public async syncWorkflowFromProcessStatus(
    processId: number,
    newProcessStatus: ProcessStatus
  ): Promise<IRetryResult<void>> {
    // Map process status to workflow action
    const statusActionMap: Record<ProcessStatus, WorkflowInstanceStatus | null> = {
      [ProcessStatus.Draft]: null,
      [ProcessStatus.NotStarted]: null,
      [ProcessStatus.Pending]: null,
      [ProcessStatus.PendingApproval]: WorkflowInstanceStatus.WaitingForApproval,
      [ProcessStatus.InProgress]: WorkflowInstanceStatus.Running,
      [ProcessStatus.OnHold]: WorkflowInstanceStatus.Paused,
      [ProcessStatus.Completed]: WorkflowInstanceStatus.Completed,
      [ProcessStatus.Cancelled]: WorkflowInstanceStatus.Cancelled,
      [ProcessStatus.Archived]: null
    };

    const workflowStatus = statusActionMap[newProcessStatus];
    if (!workflowStatus) {
      return {
        success: true,
        attempts: 0,
        totalDurationMs: 0
      };
    }

    // Get the workflow instance for this process
    const syncPayload = {
      processId,
      newProcessStatus,
      targetWorkflowStatus: workflowStatus,
      timestamp: new Date().toISOString()
    };

    const result = await retryWithDLQ<void>(
      async () => {
        const workflowInstances = await this.sp.web.lists
          .getByTitle('JML_WorkflowInstances')
          .items
          .filter(`ProcessId eq ${processId} and (Status eq 'Running' or Status eq 'Paused' or Status eq 'Waiting for Task' or Status eq 'Waiting for Approval')`)
          .select('Id', 'Status')
          .top(1)();

        if (workflowInstances.length === 0) {
          logger.info('ProcessCreationService', `No active workflow for process ${processId} to sync`);
          return;
        }

        const instanceId = workflowInstances[0].Id;

        // Update workflow instance status
        await this.sp.web.lists
          .getByTitle('JML_WorkflowInstances')
          .items
          .getById(instanceId)
          .update({
            Status: workflowStatus,
            ...(workflowStatus === WorkflowInstanceStatus.Completed ? { CompletedDate: new Date() } : {}),
            ...(workflowStatus === WorkflowInstanceStatus.Cancelled ? { CompletedDate: new Date() } : {})
          });

        logger.info('ProcessCreationService', `Workflow ${instanceId} synced to ${workflowStatus} from process status ${newProcessStatus}`);
      },
      'process-to-workflow-sync',
      syncPayload,
      PROCESS_SYNC_RETRY_OPTIONS,
      workflowSyncDLQ,
      {
        source: 'ProcessCreationService',
        operation: 'syncWorkflowFromProcessStatus'
      }
    );

    if (!result.success) {
      logger.error(
        'ProcessCreationService',
        `Failed to sync workflow from process ${processId} after ${result.attempts} attempts. DLQ ID: ${result.deadLetterItemId}`,
        result.error
      );
    }

    return result;
  }

  /**
   * Force sync both directions - use when sync state is uncertain
   */
  public async forceBidirectionalSync(processId: number): Promise<{
    processToWorkflow: IRetryResult<void>;
    workflowToProcess: IRetryResult<void>;
  }> {
    // Get current process status
    const process = await this.sp.web.lists
      .getByTitle(LIST_NAMES.PROCESSES)
      .items
      .getById(processId)
      .select('Id', 'ProcessStatus', 'WorkflowInstanceId')() as { Id: number; ProcessStatus: ProcessStatus; WorkflowInstanceId?: number };

    // Get current workflow status if exists
    let workflowToProcessResult: IRetryResult<void> = {
      success: true,
      attempts: 0,
      totalDurationMs: 0
    };

    if (process.WorkflowInstanceId) {
      const workflow = await this.sp.web.lists
        .getByTitle('JML_WorkflowInstances')
        .items
        .getById(process.WorkflowInstanceId)
        .select('Id', 'Status')() as { Id: number; Status: string };

      // Sync workflow status to process
      await this.syncProcessStatusFromWorkflow(processId, workflow.Status, workflow.Id);
      workflowToProcessResult = { success: true, attempts: 1, totalDurationMs: 0 };
    }

    // Sync process status to workflow
    const processToWorkflowResult = await this.syncWorkflowFromProcessStatus(processId, process.ProcessStatus);

    return {
      processToWorkflow: processToWorkflowResult,
      workflowToProcess: workflowToProcessResult
    };
  }

  /**
   * Get sync failure statistics
   */
  public getSyncFailureStats(): { total: number; byType: Record<string, number> } {
    return workflowSyncDLQ.getStats();
  }

  /**
   * Get failed sync operations for this service
   */
  public getFailedSyncOperations(): Array<{
    id: string;
    operationType: string;
    payload: unknown;
    error: string;
    attempts: number;
  }> {
    return workflowSyncDLQ.getAll()
      .filter(item =>
        item.operationType === 'workflow-to-process-sync' ||
        item.operationType === 'process-to-workflow-sync'
      );
  }

  // ============================================================================
  // NOTIFICATIONS
  // ============================================================================

  /**
   * Send notifications after process launch
   */
  private async sendLaunchNotifications(
    process: IJmlProcess,
    taskAssignments: IJmlTaskAssignment[],
    input: IProcessWizardInput
  ): Promise<void> {
    // Get unique assignee IDs
    const assigneeIds = Array.from(new Set(taskAssignments.map(t => t.AssignedToId)));

    // Create in-app notifications for assignees
    for (const assigneeId of assigneeIds) {
      const assigneeTasks = taskAssignments.filter(t => t.AssignedToId === assigneeId);

      await this.createNotification({
        recipientId: assigneeId,
        title: `New ${input.processType} Process Tasks`,
        message: `You have been assigned ${assigneeTasks.length} task(s) for ${input.employeeName}'s ${input.processType.toLowerCase()} process.`,
        type: 'TaskAssigned',
        priority: input.priority,
        linkUrl: `/sites/JML/SitePages/MyTasks.aspx?processId=${process.Id}`,
        processId: process.Id
      });
    }

    // Notify manager if different from creator
    if (input.managerId && input.managerId !== input.createdByUserId) {
      await this.createNotification({
        recipientId: input.managerId,
        title: `${input.processType} Process Started`,
        message: `A ${input.processType.toLowerCase()} process has been started for ${input.employeeName} in your team.`,
        type: 'ProcessStarted',
        priority: input.priority,
        linkUrl: `/sites/JML/SitePages/ProcessDetails.aspx?processId=${process.Id}`,
        processId: process.Id
      });
    }
  }

  /**
   * Create a notification record
   */
  private async createNotification(params: {
    recipientId: number;
    title: string;
    message: string;
    type: string;
    priority: Priority;
    linkUrl?: string;
    processId?: number;
  }): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle('JML_Notifications')
        .items
        .add({
          Title: params.title,
          Message: params.message,
          RecipientId: params.recipientId,
          NotificationType: params.type,
          Priority: params.priority,
          LinkUrl: params.linkUrl,
          ProcessId: params.processId,
          IsRead: false,
          SentDate: new Date()
        });
    } catch (error) {
      // Log but don't throw - notification failure shouldn't fail process
      logger.warn('ProcessCreationService', 'Failed to create notification', error);
    }
  }

  // ============================================================================
  // AUDIT LOGGING
  // ============================================================================

  /**
   * Log process creation for audit
   */
  private async logProcessCreation(
    processId: number,
    input: IProcessWizardInput
  ): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle('JML_AuditLog')
        .items
        .add({
          Title: `Process Created: ${processId}`,
          Action: 'ProcessCreated',
          EntityType: 'Process',
          EntityId: processId,
          PerformedById: input.createdByUserId,
          Details: JSON.stringify({
            processType: input.processType,
            employeeName: input.employeeName,
            department: input.department,
            startDate: input.startDate,
            targetCompletionDate: input.targetCompletionDate
          }),
          Timestamp: new Date()
        });
    } catch (error) {
      // Log but don't throw - audit failure shouldn't fail process
      logger.warn('ProcessCreationService', 'Failed to create audit log', error);
    }
  }

  // ============================================================================
  // UTILITY METHODS
  // ============================================================================

  /**
   * Get process with all related data
   */
  public async getProcessWithDetails(processId: number): Promise<{
    process: IJmlProcess;
    taskAssignments: IJmlTaskAssignment[];
    workflowInstanceId?: number;
  }> {
    // Get process
    const process = await this.sp.web.lists
      .getByTitle(LIST_NAMES.PROCESSES)
      .items
      .getById(processId)
      .select('*', 'Manager/Id', 'Manager/Title', 'Manager/EMail')
      .expand('Manager')() as IJmlProcess;

    // Get task assignments
    const taskAssignments = await this.sp.web.lists
      .getByTitle(LIST_NAMES.TASK_ASSIGNMENTS)
      .items
      .filter(`ProcessID eq '${processId}'`)
      .select('*', 'AssignedTo/Id', 'AssignedTo/Title', 'AssignedTo/EMail')
      .expand('AssignedTo')
      .orderBy('DueDate')() as IJmlTaskAssignment[];

    // Get workflow instance (if any)
    const workflowInstances = await this.sp.web.lists
      .getByTitle('JML_WorkflowInstances')
      .items
      .filter(`ProcessId eq ${processId}`)
      .select('Id')
      .top(1)();

    return {
      process,
      taskAssignments,
      workflowInstanceId: workflowInstances.length > 0 ? workflowInstances[0].Id : undefined
    };
  }

  /**
   * Recalculate and update process progress
   */
  public async recalculateProcessProgress(processId: number): Promise<void> {
    // Get all task assignments for process
    const assignments = await this.sp.web.lists
      .getByTitle(LIST_NAMES.TASK_ASSIGNMENTS)
      .items
      .filter(`ProcessID eq '${processId}' and IsDeleted ne 1`)
      .select('Id', 'Status', 'DueDate')();

    const totalTasks = assignments.length;
    const completedTasks = assignments.filter(
      (a: { Status: TaskStatus }) => a.Status === TaskStatus.Completed
    ).length;
    const overdueTasks = assignments.filter(
      (a: { Status: TaskStatus; DueDate: string }) =>
        a.Status !== TaskStatus.Completed &&
        new Date(a.DueDate) < new Date()
    ).length;

    const progressPercentage = totalTasks > 0
      ? Math.round((completedTasks / totalTasks) * 100)
      : 0;

    // Determine if process should auto-complete
    let processStatus: ProcessStatus | undefined;
    if (totalTasks > 0 && completedTasks === totalTasks) {
      processStatus = ProcessStatus.Completed;
    }

    await this.sp.web.lists
      .getByTitle(LIST_NAMES.PROCESSES)
      .items
      .getById(processId)
      .update({
        TotalTasks: totalTasks,
        CompletedTasks: completedTasks,
        ProgressPercentage: progressPercentage,
        OverdueTasks: overdueTasks,
        ...(processStatus ? {
          ProcessStatus: processStatus,
          ActualCompletionDate: new Date()
        } : {})
      });
  }
}
