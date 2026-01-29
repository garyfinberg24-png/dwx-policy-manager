// @ts-nocheck
/**
 * ProcessOrchestrationService
 * Master coordinator that connects Process management with Workflow execution
 *
 * This service is the single entry point for all process lifecycle operations,
 * ensuring that process status and workflow status remain synchronized.
 *
 * Key responsibilities:
 * - Create processes and automatically start workflows
 * - Sync process status with workflow status bidirectionally
 * - Handle task completion callbacks to update workflow steps
 * - Manage process lifecycle transitions
 */

import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import { logger } from './LoggingService';
import { SPService } from './SPService';
import { WorkflowEngineService, IStartWorkflowOptions, IExecutionResult } from './workflow/WorkflowEngineService';
import { WorkflowInstanceService } from './workflow/WorkflowInstanceService';
import { ApprovalService } from './ApprovalService';
import { IJmlProcess } from '../models/IJmlProcess';
import { IJmlTaskAssignment } from '../models/IJmlTaskAssignment';
import { ProcessType, ProcessStatus, TaskStatus } from '../models/ICommon';
import { WorkflowInstanceStatus, StepType } from '../models/IWorkflow';

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Extract numeric process ID from ProcessID field which can be string or lookup object
 */
function extractProcessId(processId: string | { Id: number; Title: string } | undefined): number | undefined {
  if (!processId) return undefined;
  if (typeof processId === 'string') {
    const id = parseInt(processId, 10);
    return isNaN(id) ? undefined : id;
  }
  if (typeof processId === 'object' && 'Id' in processId) {
    return processId.Id;
  }
  return undefined;
}

// ============================================================================
// INTERFACES
// ============================================================================

/**
 * Options for initiating a new process
 */
export interface IProcessInitiationOptions {
  // Process details
  processType: ProcessType;
  employeeName: string;
  employeeEmail: string;
  department: string;
  jobTitle: string;
  location?: string;
  startDate: Date;
  targetCompletionDate?: Date;
  priority?: string;
  managerId?: number;
  managerEmail?: string;
  processOwnerId?: number;

  // Template and tasks
  templateId?: string;
  templateName?: string;
  tasks?: ITaskInitiation[];

  // Workflow options
  workflowDefinitionId?: number;
  workflowDefinitionCode?: string;
  autoStartWorkflow?: boolean;

  // Automation options
  enableAutomation?: boolean;
  sendWelcomeEmail?: boolean;
  notifyStakeholders?: boolean;
  scheduleKickoffMeeting?: boolean;

  // Additional data
  comments?: string;
  customFields?: Record<string, unknown>;
}

/**
 * Task initiation data
 */
export interface ITaskInitiation {
  title: string;
  description?: string;
  assigneeId?: number;
  dueDate?: Date;
  priority?: string;
  category?: string;
  requiresApproval?: boolean;
}

/**
 * Result of process initiation
 */
export interface IProcessInitiationResult {
  success: boolean;
  processId?: number;
  process?: IJmlProcess;
  workflowInstanceId?: number;
  workflowStarted: boolean;
  taskCount?: number;
  error?: string;
  warnings: string[];
}

/**
 * Result of task completion handling
 */
export interface ITaskCompletionResult {
  success: boolean;
  processUpdated: boolean;
  workflowResumed: boolean;
  processCompleted: boolean;
  error?: string;
}

/**
 * Process status with workflow info
 */
export interface IProcessWithWorkflow extends IJmlProcess {
  workflowInstanceId?: number;
  workflowStatus?: WorkflowInstanceStatus;
  workflowCurrentStep?: string;
  workflowProgress?: number;
}

// ============================================================================
// PROCESS ORCHESTRATION SERVICE
// ============================================================================

export class ProcessOrchestrationService {
  private sp: SPFI;
  private context: WebPartContext;
  private spService: SPService;
  private workflowEngine: WorkflowEngineService;
  private workflowInstanceService: WorkflowInstanceService;
  private approvalService: ApprovalService;

  // List name for process-workflow mapping
  private readonly PROCESS_WORKFLOW_MAP_LIST = 'PM_ProcessWorkflowMap';

  constructor(sp: SPFI, context: WebPartContext) {
    this.sp = sp;
    this.context = context;
    this.spService = new SPService(sp);
    this.workflowEngine = new WorkflowEngineService(sp, context);
    this.workflowInstanceService = new WorkflowInstanceService(sp);
    this.approvalService = new ApprovalService(sp);
  }

  // ============================================================================
  // PROCESS LIFECYCLE
  // ============================================================================

  /**
   * Initiate a new JML process with automatic workflow start
   * This is the primary entry point for creating processes
   */
  public async initiateProcess(options: IProcessInitiationOptions): Promise<IProcessInitiationResult> {
    const warnings: string[] = [];
    let processId: number | undefined;
    let workflowInstanceId: number | undefined;
    let workflowStarted = false;

    try {
      logger.info('ProcessOrchestrationService', 'Initiating process', {
        processType: options.processType,
        employeeName: options.employeeName
      });

      // Step 1: Create the process record
      const process = await this.createProcessRecord(options);
      processId = process.Id;
      logger.info('ProcessOrchestrationService', `Process created with ID: ${processId}`);

      // Step 2: Create tasks if provided
      let taskCount = 0;
      if (options.tasks && options.tasks.length > 0) {
        taskCount = await this.createProcessTasks(processId!, options.tasks, options.startDate);
        await this.updateProcessTaskCount(processId!, taskCount);
        logger.info('ProcessOrchestrationService', `Created ${taskCount} tasks for process ${processId}`);
      }

      // Step 3: Start workflow (if enabled)
      if (options.autoStartWorkflow !== false) {
        try {
          const workflowResult = await this.startWorkflowForProcess(processId!, options);

          if (workflowResult.success) {
            workflowInstanceId = workflowResult.instanceId;
            workflowStarted = true;

            // Link process to workflow
            await this.linkProcessToWorkflow(processId!, workflowInstanceId);

            // Update process status based on workflow
            await this.syncProcessStatusFromWorkflow(processId!, workflowResult.status);

            logger.info('ProcessOrchestrationService', `Workflow ${workflowInstanceId} started for process ${processId}`);
          } else {
            warnings.push(`Workflow could not be started: ${workflowResult.error}`);
            logger.warn('ProcessOrchestrationService', `Workflow start failed: ${workflowResult.error}`);
          }
        } catch (workflowError) {
          const errorMsg = workflowError instanceof Error ? workflowError.message : 'Unknown workflow error';
          warnings.push(`Workflow engine error: ${errorMsg}`);
          logger.error('ProcessOrchestrationService', 'Error starting workflow', workflowError);
          // Don't fail process creation - workflow can be started manually
        }
      }

      // Step 4: Send notifications if enabled
      if (options.enableAutomation) {
        await this.sendProcessNotifications(processId!, options, warnings);
      }

      // Step 5: Create audit log
      await this.createAuditLog(processId!, 'ProcessInitiated', {
        processType: options.processType,
        employeeName: options.employeeName,
        department: options.department,
        taskCount
      });

      return {
        success: true,
        processId,
        process,
        workflowInstanceId,
        workflowStarted,
        taskCount,
        warnings
      };

    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : 'Unknown error';
      logger.error('ProcessOrchestrationService', 'Error initiating process', error);

      return {
        success: false,
        processId,
        workflowInstanceId,
        workflowStarted,
        error: errorMsg,
        warnings
      };
    }
  }

  /**
   * Handle task completion and update workflow accordingly
   * Called when a task in a process is marked as complete
   */
  public async handleTaskCompletion(
    taskAssignmentId: number,
    completedByUserId: number,
    result?: Record<string, unknown>
  ): Promise<ITaskCompletionResult> {
    try {
      // Get the task assignment
      const task = await this.getTaskAssignment(taskAssignmentId);
      if (!task) {
        return {
          success: false,
          processUpdated: false,
          workflowResumed: false,
          processCompleted: false,
          error: 'Task assignment not found'
        };
      }

      const processId = extractProcessId(task.ProcessID);
      if (!processId) {
        return {
          success: false,
          processUpdated: false,
          workflowResumed: false,
          processCompleted: false,
          error: 'Invalid process ID'
        };
      }

      logger.info('ProcessOrchestrationService', `Handling task completion: ${taskAssignmentId} for process ${processId}`);

      // Update process progress
      const progressResult = await this.updateProcessProgress(processId);
      const processUpdated = progressResult.success;
      let processCompleted = progressResult.allTasksComplete;

      // Check if workflow needs to be notified
      let workflowResumed = false;
      const workflowInstanceId = await this.getWorkflowInstanceForProcess(processId);

      if (workflowInstanceId) {
        try {
          // Get the workflow instance
          const workflowInstance = await this.workflowInstanceService.getById(workflowInstanceId);

          // If workflow is waiting for tasks, try to resume it
          if (workflowInstance.Status === WorkflowInstanceStatus.WaitingForTask) {
            const resumeResult = await this.workflowEngine.completeWaitingStep(
              workflowInstanceId,
              workflowInstance.CurrentStepId || '',
              {
                taskId: taskAssignmentId,
                completedBy: completedByUserId,
                completedAt: new Date().toISOString(),
                ...result
              }
            );

            workflowResumed = resumeResult.success;

            // Check if workflow completed
            if (resumeResult.status === WorkflowInstanceStatus.Completed) {
              processCompleted = true;
              await this.completeProcess(processId);
            }

            // Sync process status
            await this.syncProcessStatusFromWorkflow(processId, resumeResult.status);
          }
        } catch (workflowError) {
          logger.warn('ProcessOrchestrationService', `Error resuming workflow for task ${taskAssignmentId}`, workflowError);
        }
      }

      // If all tasks complete and no workflow, complete the process
      if (progressResult.allTasksComplete && !workflowInstanceId) {
        await this.completeProcess(processId);
        processCompleted = true;
      }

      return {
        success: true,
        processUpdated,
        workflowResumed,
        processCompleted
      };

    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : 'Unknown error';
      logger.error('ProcessOrchestrationService', 'Error handling task completion', error);
      return {
        success: false,
        processUpdated: false,
        workflowResumed: false,
        processCompleted: false,
        error: errorMsg
      };
    }
  }

  /**
   * Cancel a process and its associated workflow
   * GAP FIX: Now also cancels pending approvals to prevent orphaned data
   */
  public async cancelProcess(processId: number, reason?: string, cancelledByUserId?: number): Promise<boolean> {
    try {
      logger.info('ProcessOrchestrationService', `Cancelling process ${processId}`, { reason });

      // Cancel the workflow first (if exists)
      const workflowInstanceId = await this.getWorkflowInstanceForProcess(processId);
      if (workflowInstanceId) {
        await this.workflowInstanceService.cancel(workflowInstanceId, reason);
      }

      // GAP FIX: Cancel any pending approvals for this process
      // This prevents orphaned approval requests that would confuse approvers
      try {
        const approvalResult = await this.approvalService.cancelPendingApprovalsForProcess(processId, reason);
        if (approvalResult.cancelled > 0) {
          logger.info('ProcessOrchestrationService',
            `Cancelled ${approvalResult.cancelled} pending approvals for process ${processId}`);
        }
      } catch (approvalError) {
        // Log but don't fail - process cancellation should continue
        logger.warn('ProcessOrchestrationService',
          `Failed to cancel pending approvals for process ${processId}`, approvalError);
      }

      // Update process status
      await this.spService.updateProcess(processId, {
        ProcessStatus: ProcessStatus.Cancelled,
        ActualCompletionDate: new Date(),
        Comments: reason ? `Cancelled: ${reason}` : 'Cancelled'
      });

      // Create audit log
      await this.createAuditLog(processId, 'ProcessCancelled', { reason, cancelledByUserId });

      return true;
    } catch (error) {
      logger.error('ProcessOrchestrationService', `Error cancelling process ${processId}`, error);
      return false;
    }
  }

  /**
   * Put a process on hold
   */
  public async holdProcess(processId: number, reason?: string): Promise<boolean> {
    try {
      logger.info('ProcessOrchestrationService', `Putting process ${processId} on hold`, { reason });

      // Pause the workflow (if exists)
      const workflowInstanceId = await this.getWorkflowInstanceForProcess(processId);
      if (workflowInstanceId) {
        await this.workflowInstanceService.pause(workflowInstanceId);
      }

      // Update process status
      await this.spService.updateProcess(processId, {
        ProcessStatus: ProcessStatus.OnHold,
        Comments: reason ? `On Hold: ${reason}` : 'On Hold'
      });

      return true;
    } catch (error) {
      logger.error('ProcessOrchestrationService', `Error holding process ${processId}`, error);
      return false;
    }
  }

  /**
   * Resume a held process
   */
  public async resumeProcess(processId: number): Promise<boolean> {
    try {
      logger.info('ProcessOrchestrationService', `Resuming process ${processId}`);

      // Resume the workflow (if exists)
      const workflowInstanceId = await this.getWorkflowInstanceForProcess(processId);
      if (workflowInstanceId) {
        await this.workflowInstanceService.resume(workflowInstanceId);
      }

      // Update process status
      await this.spService.updateProcess(processId, {
        ProcessStatus: ProcessStatus.InProgress
      });

      return true;
    } catch (error) {
      logger.error('ProcessOrchestrationService', `Error resuming process ${processId}`, error);
      return false;
    }
  }

  /**
   * Complete a process
   */
  public async completeProcess(processId: number): Promise<boolean> {
    try {
      logger.info('ProcessOrchestrationService', `Completing process ${processId}`);

      await this.spService.updateProcess(processId, {
        ProcessStatus: ProcessStatus.Completed,
        ActualCompletionDate: new Date(),
        ProgressPercentage: 100
      });

      await this.createAuditLog(processId, 'ProcessCompleted', {});

      return true;
    } catch (error) {
      logger.error('ProcessOrchestrationService', `Error completing process ${processId}`, error);
      return false;
    }
  }

  // ============================================================================
  // WORKFLOW OPERATIONS
  // ============================================================================

  /**
   * Start workflow for an existing process
   */
  public async startWorkflowForProcess(
    processId: number,
    options: IProcessInitiationOptions
  ): Promise<IExecutionResult> {
    const currentUser = this.context.pageContext?.user;
    const startedByUserId = currentUser?.loginName
      ? await this.resolveUserId(currentUser.loginName)
      : 0;

    const workflowOptions: IStartWorkflowOptions = {
      definitionId: options.workflowDefinitionId,
      definitionCode: options.workflowDefinitionCode,
      processId: processId,
      processType: options.processType,
      employeeName: options.employeeName,
      employeeEmail: options.employeeEmail,
      department: options.department,
      managerId: options.managerId,
      managerEmail: options.managerEmail,
      startedByUserId: startedByUserId,
      startedByUserName: currentUser?.displayName || 'System',
      customContext: {
        jobTitle: options.jobTitle,
        location: options.location,
        startDate: options.startDate.toISOString(),
        templateId: options.templateId,
        templateName: options.templateName,
        priority: options.priority,
        ...options.customFields
      }
    };

    return await this.workflowEngine.startWorkflow(workflowOptions);
  }

  /**
   * Manually start workflow for a process that doesn't have one
   */
  public async attachWorkflowToProcess(
    processId: number,
    workflowDefinitionId?: number,
    workflowDefinitionCode?: string
  ): Promise<IExecutionResult> {
    // Get process details
    const process = await this.spService.getProcessById(processId);

    const options: IProcessInitiationOptions = {
      processType: process.ProcessType,
      employeeName: process.EmployeeName,
      employeeEmail: process.EmployeeEmail,
      department: process.Department,
      jobTitle: process.JobTitle,
      location: process.Location,
      startDate: new Date(process.StartDate),
      managerId: process.ManagerId,
      workflowDefinitionId,
      workflowDefinitionCode
    };

    const result = await this.startWorkflowForProcess(processId, options);

    if (result.success) {
      await this.linkProcessToWorkflow(processId, result.instanceId);
    }

    return result;
  }

  /**
   * Get process with workflow information
   */
  public async getProcessWithWorkflow(processId: number): Promise<IProcessWithWorkflow> {
    const process = await this.spService.getProcessById(processId);
    const workflowInstanceId = await this.getWorkflowInstanceForProcess(processId);

    const result: IProcessWithWorkflow = { ...process };

    if (workflowInstanceId) {
      try {
        const workflowInstance = await this.workflowInstanceService.getById(workflowInstanceId);
        result.workflowInstanceId = workflowInstanceId;
        result.workflowStatus = workflowInstance.Status;
        result.workflowCurrentStep = workflowInstance.CurrentStepName;
        result.workflowProgress = workflowInstance.ProgressPercentage;
      } catch (error) {
        logger.warn('ProcessOrchestrationService', `Error fetching workflow for process ${processId}`, error);
      }
    }

    return result;
  }

  // ============================================================================
  // STATUS SYNCHRONIZATION
  // ============================================================================

  /**
   * Sync process status from workflow status
   */
  public async syncProcessStatusFromWorkflow(
    processId: number,
    workflowStatus: WorkflowInstanceStatus
  ): Promise<void> {
    let processStatus: ProcessStatus;

    switch (workflowStatus) {
      case WorkflowInstanceStatus.Pending:
        processStatus = ProcessStatus.Pending;
        break;
      case WorkflowInstanceStatus.Running:
      case WorkflowInstanceStatus.WaitingForInput:
      case WorkflowInstanceStatus.WaitingForApproval:
      case WorkflowInstanceStatus.WaitingForTask:
        processStatus = ProcessStatus.InProgress;
        break;
      case WorkflowInstanceStatus.Paused:
        processStatus = ProcessStatus.OnHold;
        break;
      case WorkflowInstanceStatus.Completed:
        processStatus = ProcessStatus.Completed;
        break;
      case WorkflowInstanceStatus.Failed:
      case WorkflowInstanceStatus.Cancelled:
        processStatus = ProcessStatus.Cancelled;
        break;
      default:
        processStatus = ProcessStatus.InProgress;
    }

    await this.spService.updateProcess(processId, {
      ProcessStatus: processStatus
    });
  }

  /**
   * Sync workflow from process status change
   */
  public async syncWorkflowFromProcessStatus(
    processId: number,
    newProcessStatus: ProcessStatus
  ): Promise<void> {
    const workflowInstanceId = await this.getWorkflowInstanceForProcess(processId);
    if (!workflowInstanceId) return;

    switch (newProcessStatus) {
      case ProcessStatus.OnHold:
        await this.workflowInstanceService.pause(workflowInstanceId);
        break;
      case ProcessStatus.Cancelled:
        await this.workflowInstanceService.cancel(workflowInstanceId, 'Process cancelled');
        break;
      case ProcessStatus.InProgress:
        // Check if workflow is paused
        const instance = await this.workflowInstanceService.getById(workflowInstanceId);
        if (instance.Status === WorkflowInstanceStatus.Paused) {
          await this.workflowInstanceService.resume(workflowInstanceId);
        }
        break;
    }
  }

  // ============================================================================
  // PRIVATE HELPER METHODS
  // ============================================================================

  /**
   * Create the process record in SharePoint
   */
  private async createProcessRecord(options: IProcessInitiationOptions): Promise<IJmlProcess> {
    const processData: Partial<IJmlProcess> = {
      Title: `${options.employeeName} - ${options.processType}`,
      ProcessType: options.processType,
      ProcessStatus: ProcessStatus.InProgress,
      EmployeeName: options.employeeName,
      EmployeeEmail: options.employeeEmail,
      Department: options.department,
      JobTitle: options.jobTitle,
      Location: options.location,
      StartDate: options.startDate,
      TargetCompletionDate: options.targetCompletionDate || this.calculateDefaultTargetDate(options.startDate, options.processType),
      Priority: options.priority as IJmlProcess['Priority'],
      ManagerId: options.managerId,
      ProcessOwnerId: options.processOwnerId,
      ChecklistTemplateID: options.templateId,
      Comments: options.comments,
      TotalTasks: 0,
      CompletedTasks: 0,
      ProgressPercentage: 0
    };

    return await this.spService.createProcess(processData);
  }

  /**
   * Calculate default target completion date based on process type
   */
  private calculateDefaultTargetDate(startDate: Date, processType: ProcessType): Date {
    const daysMap: Record<ProcessType, number> = {
      [ProcessType.Joiner]: 90,  // 3 months for onboarding
      [ProcessType.Mover]: 30,   // 1 month for transfers
      [ProcessType.Leaver]: 14   // 2 weeks for offboarding
    };

    const days = daysMap[processType] || 30;
    const targetDate = new Date(startDate);
    targetDate.setDate(targetDate.getDate() + days);
    return targetDate;
  }

  /**
   * Create tasks for the process
   */
  private async createProcessTasks(
    processId: number,
    tasks: ITaskInitiation[],
    startDate: Date
  ): Promise<number> {
    let createdCount = 0;

    for (const task of tasks) {
      try {
        await this.sp.web.lists.getByTitle('PM_TaskAssignments').items.add({
          Title: task.title,
          ProcessID: processId.toString(),
          AssignedToId: task.assigneeId,
          DueDate: task.dueDate || startDate,
          StartDate: startDate,
          Status: TaskStatus.NotStarted,
          Priority: task.priority || 'Medium',
          Notes: task.description,
          Category: task.category,
          RequiresApproval: task.requiresApproval || false,
          PercentComplete: 0,
          IsBlocked: false
        });
        createdCount++;
      } catch (error) {
        logger.warn('ProcessOrchestrationService', `Error creating task: ${task.title}`, error);
      }
    }

    return createdCount;
  }

  /**
   * Update process task count
   */
  private async updateProcessTaskCount(processId: number, taskCount: number): Promise<void> {
    await this.spService.updateProcess(processId, {
      TotalTasks: taskCount,
      CompletedTasks: 0,
      ProgressPercentage: 0
    });
  }

  /**
   * Link process to workflow instance
   */
  private async linkProcessToWorkflow(processId: number, workflowInstanceId: number): Promise<void> {
    try {
      // Try to use dedicated mapping list
      await this.sp.web.lists.getByTitle(this.PROCESS_WORKFLOW_MAP_LIST).items.add({
        Title: `Process ${processId} - Workflow ${workflowInstanceId}`,
        ProcessId: processId,
        WorkflowInstanceId: workflowInstanceId,
        LinkedDate: new Date().toISOString(),
        IsActive: true
      });
    } catch {
      // List might not exist - use custom field on process instead
      logger.info('ProcessOrchestrationService', 'Using fallback process-workflow linking');
      // Store in process CustomFields
      const process = await this.spService.getProcessById(processId);
      const customFields = process.CustomFields ? JSON.parse(process.CustomFields) : {};
      customFields.workflowInstanceId = workflowInstanceId;

      await this.spService.updateProcess(processId, {
        CustomFields: JSON.stringify(customFields)
      });
    }
  }

  /**
   * Get workflow instance ID for a process
   */
  private async getWorkflowInstanceForProcess(processId: number): Promise<number | undefined> {
    try {
      // Try dedicated mapping list first
      const mappings = await this.sp.web.lists.getByTitle(this.PROCESS_WORKFLOW_MAP_LIST).items
        .filter(`ProcessId eq ${processId} and IsActive eq 1`)
        .top(1)();

      if (mappings.length > 0) {
        return mappings[0].WorkflowInstanceId;
      }
    } catch {
      // Mapping list doesn't exist - try CustomFields
    }

    // Fallback: check CustomFields
    try {
      const process = await this.spService.getProcessById(processId);
      if (process.CustomFields) {
        const customFields = JSON.parse(process.CustomFields);
        return customFields.workflowInstanceId;
      }
    } catch {
      // Ignore parsing errors
    }

    // Final fallback: query workflow instances
    try {
      const instance = await this.workflowInstanceService.getActiveForProcess(processId);
      return instance?.Id;
    } catch {
      return undefined;
    }
  }

  /**
   * Get task assignment by ID
   */
  private async getTaskAssignment(taskId: number): Promise<IJmlTaskAssignment | null> {
    try {
      const item = await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
        .getById(taskId)
        .select('Id', 'Title', 'ProcessID', 'Status', 'AssignedToId')();
      return item as IJmlTaskAssignment;
    } catch {
      return null;
    }
  }

  /**
   * Update process progress based on task completion
   */
  private async updateProcessProgress(processId: number): Promise<{
    success: boolean;
    totalTasks: number;
    completedTasks: number;
    allTasksComplete: boolean;
  }> {
    try {
      // Get all tasks for this process
      const tasks = await this.spService.getTaskAssignmentsByProcess(processId);
      const totalTasks = tasks.length;
      const completedTasks = tasks.filter(t =>
        t.Status === TaskStatus.Completed || t.Status === TaskStatus.Skipped
      ).length;
      const progressPercentage = totalTasks > 0 ? Math.round((completedTasks / totalTasks) * 100) : 0;
      const overdueTasks = tasks.filter(t => {
        if (t.Status === TaskStatus.Completed || t.Status === TaskStatus.Skipped) return false;
        if (!t.DueDate) return false;
        return new Date(t.DueDate) < new Date();
      }).length;

      await this.spService.updateProcess(processId, {
        TotalTasks: totalTasks,
        CompletedTasks: completedTasks,
        ProgressPercentage: progressPercentage,
        OverdueTasks: overdueTasks
      });

      return {
        success: true,
        totalTasks,
        completedTasks,
        allTasksComplete: totalTasks > 0 && completedTasks === totalTasks
      };
    } catch (error) {
      logger.error('ProcessOrchestrationService', `Error updating progress for process ${processId}`, error);
      return {
        success: false,
        totalTasks: 0,
        completedTasks: 0,
        allTasksComplete: false
      };
    }
  }

  /**
   * Send process notifications
   */
  private async sendProcessNotifications(
    processId: number,
    options: IProcessInitiationOptions,
    warnings: string[]
  ): Promise<void> {
    try {
      // Notify manager if required
      if (options.notifyStakeholders && options.managerId) {
        await this.sp.web.lists.getByTitle('PM_Notifications').items.add({
          Title: `New ${options.processType} Process: ${options.employeeName}`,
          NotificationType: 'ProcessStarted',
          MessageBody: `A new ${options.processType} process has been started for ${options.employeeName}.`,
          Priority: 'Normal',
          RecipientId: options.managerId,
          ProcessId: processId.toString(),
          Status: 'Pending'
        });
      }
    } catch (error) {
      warnings.push('Failed to send some notifications');
      logger.warn('ProcessOrchestrationService', 'Error sending notifications', error);
    }
  }

  /**
   * Create audit log entry
   */
  private async createAuditLog(
    processId: number,
    eventType: string,
    data: Record<string, unknown>
  ): Promise<void> {
    try {
      await this.spService.createAuditLog({
        Title: `${eventType}: Process ${processId}`,
        EventType: eventType,
        EntityType: 'Process',
        EntityId: processId,
        ProcessId: processId,
        Action: eventType,
        Description: JSON.stringify(data),
        AdditionalData: JSON.stringify(data)
      });
    } catch (error) {
      logger.warn('ProcessOrchestrationService', 'Error creating audit log', error);
    }
  }

  /**
   * Resolve SharePoint user ID from login name
   */
  private async resolveUserId(loginName: string): Promise<number> {
    try {
      const user = await this.sp.web.ensureUser(loginName);
      return user.data.Id;
    } catch {
      return 0;
    }
  }
}

export default ProcessOrchestrationService;
