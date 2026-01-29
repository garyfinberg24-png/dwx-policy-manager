// @ts-nocheck
/* eslint-disable */
/**
 * WorkflowInstanceService
 * Service for managing workflow instances in SharePoint
 * Handles running instances, step status tracking, and execution context
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/batching';

import {
  IWorkflowInstance,
  IWorkflowInstanceSummary,
  IWorkflowInstanceUI,
  IWorkflowStepStatus,
  IWorkflowLog,
  IWorkflowContext,
  WorkflowInstanceStatus,
  StepStatus,
  LogLevel
} from '../../models/IWorkflow';
import { ProcessType, ProcessStatus } from '../../models/ICommon';
import { logger } from '../LoggingService';

// List name constants
const INSTANCES_LIST = 'PM_WorkflowInstances';
const STEP_STATUS_LIST = 'PM_WorkflowStepStatus';
const LOGS_LIST = 'PM_WorkflowLogs';

// Select fields for instances
const INSTANCE_SELECT = [
  'Id', 'Title', 'WorkflowDefinitionId', 'ProcessId',
  'Status', 'CurrentStepId', 'CurrentStepName',
  'TotalSteps', 'CompletedSteps', 'ProgressPercentage',
  'StartedDate', 'CompletedDate', 'EstimatedCompletionDate',
  'Variables', 'Context',
  'ErrorMessage', 'ErrorStepId', 'RetryCount', 'LastRetryDate',
  'StartedById', 'CompletedById',
  'Created', 'Modified'
].join(',');

export class WorkflowInstanceService {
  private sp: SPFI;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ============================================================================
  // INSTANCE READ OPERATIONS
  // ============================================================================

  /**
   * Get all workflow instances
   */
  public async getAll(filter?: string, top?: number): Promise<IWorkflowInstance[]> {
    try {
      let query = this.sp.web.lists.getByTitle(INSTANCES_LIST).items
        .select(INSTANCE_SELECT)
        .orderBy('Created', false);

      if (filter) {
        query = query.filter(filter);
      }
      if (top) {
        query = query.top(top);
      }

      const items = await query();
      return items as IWorkflowInstance[];
    } catch (error) {
      logger.error('WorkflowInstanceService', 'Error fetching workflow instances', error);
      throw new Error('Unable to retrieve workflow instances.');
    }
  }

  /**
   * Get workflow instance summaries for dashboard
   */
  public async getSummaries(options?: {
    status?: WorkflowInstanceStatus;
    processType?: ProcessType;
    top?: number;
  }): Promise<IWorkflowInstanceSummary[]> {
    try {
      let filterParts: string[] = [];

      if (options?.status) {
        filterParts.push(`Status eq '${options.status}'`);
      }

      let query = this.sp.web.lists.getByTitle(INSTANCES_LIST).items
        .select('Id', 'Title', 'ProcessId', 'Status', 'CurrentStepName', 'ProgressPercentage', 'StartedDate', 'EstimatedCompletionDate', 'Context')
        .orderBy('Created', false);

      if (filterParts.length > 0) {
        query = query.filter(filterParts.join(' and '));
      }
      if (options?.top) {
        query = query.top(options.top);
      }

      const items = await query();

      return items.map((item: IWorkflowInstance) => {
        let context: Partial<IWorkflowContext> = {};
        try {
          context = item.Context ? JSON.parse(item.Context) : {};
        } catch { /* ignore */ }

        const now = new Date();
        const estimated = item.EstimatedCompletionDate ? new Date(item.EstimatedCompletionDate) : null;

        return {
          Id: item.Id,
          WorkflowName: item.Title,
          ProcessId: item.ProcessId,
          EmployeeName: context.employeeName || 'Unknown',
          ProcessType: context.processType || ProcessType.Joiner,
          Status: item.Status,
          CurrentStepName: item.CurrentStepName,
          ProgressPercentage: item.ProgressPercentage || 0,
          StartedDate: item.StartedDate,
          IsOverdue: estimated ? now > estimated : false
        } as IWorkflowInstanceSummary;
      });
    } catch (error) {
      logger.error('WorkflowInstanceService', 'Error fetching workflow summaries', error);
      throw new Error('Unable to retrieve workflow summaries.');
    }
  }

  /**
   * Get workflow instance by ID
   */
  public async getById(id: number): Promise<IWorkflowInstance> {
    try {
      const item = await this.sp.web.lists.getByTitle(INSTANCES_LIST).items
        .getById(id)
        .select(INSTANCE_SELECT)();

      return item as IWorkflowInstance;
    } catch (error) {
      logger.error('WorkflowInstanceService', `Error fetching workflow instance ${id}`, error);
      throw new Error('Workflow instance not found.');
    }
  }

  /**
   * Get active workflow instance for a process
   */
  public async getActiveForProcess(processId: number): Promise<IWorkflowInstance | undefined> {
    try {
      const activeStatuses = [
        WorkflowInstanceStatus.Pending,
        WorkflowInstanceStatus.Running,
        WorkflowInstanceStatus.Paused,
        WorkflowInstanceStatus.WaitingForInput,
        WorkflowInstanceStatus.WaitingForApproval,
        WorkflowInstanceStatus.WaitingForTask
      ];

      const statusFilter = activeStatuses.map(s => `Status eq '${s}'`).join(' or ');

      const items = await this.sp.web.lists.getByTitle(INSTANCES_LIST).items
        .select(INSTANCE_SELECT)
        .filter(`ProcessId eq ${processId} and (${statusFilter})`)
        .top(1)();

      return items.length > 0 ? items[0] as IWorkflowInstance : undefined;
    } catch (error) {
      logger.error('WorkflowInstanceService', `Error fetching active workflow for process ${processId}`, error);
      throw new Error('Unable to retrieve workflow instance.');
    }
  }

  /**
   * Get instances for a workflow definition
   */
  public async getByDefinition(definitionId: number, top?: number): Promise<IWorkflowInstance[]> {
    try {
      let query = this.sp.web.lists.getByTitle(INSTANCES_LIST).items
        .select(INSTANCE_SELECT)
        .filter(`WorkflowDefinitionId eq ${definitionId}`)
        .orderBy('Created', false);

      if (top) {
        query = query.top(top);
      }

      const items = await query();
      return items as IWorkflowInstance[];
    } catch (error) {
      logger.error('WorkflowInstanceService', `Error fetching instances for definition ${definitionId}`, error);
      throw new Error('Unable to retrieve workflow instances.');
    }
  }

  /**
   * Get running instances count
   */
  public async getRunningCount(): Promise<number> {
    try {
      const items = await this.sp.web.lists.getByTitle(INSTANCES_LIST).items
        .filter(`Status eq '${WorkflowInstanceStatus.Running}'`)
        .select('Id')();

      return items.length;
    } catch (error) {
      logger.error('WorkflowInstanceService', 'Error fetching running count', error);
      return 0;
    }
  }

  // ============================================================================
  // INSTANCE CREATE/UPDATE OPERATIONS
  // ============================================================================

  /**
   * Create new workflow instance
   */
  public async create(instance: Partial<IWorkflowInstance>): Promise<IWorkflowInstance> {
    try {
      const itemData = {
        Title: instance.Title,
        WorkflowDefinitionId: instance.WorkflowDefinitionId,
        ProcessId: instance.ProcessId,
        Status: instance.Status || WorkflowInstanceStatus.Pending,
        CurrentStepId: instance.CurrentStepId,
        CurrentStepName: instance.CurrentStepName,
        TotalSteps: instance.TotalSteps || 0,
        CompletedSteps: 0,
        ProgressPercentage: 0,
        StartedDate: new Date().toISOString(),
        EstimatedCompletionDate: instance.EstimatedCompletionDate,
        Variables: typeof instance.Variables === 'string'
          ? instance.Variables
          : JSON.stringify(instance.Variables || {}),
        Context: typeof instance.Context === 'string'
          ? instance.Context
          : JSON.stringify(instance.Context || {}),
        StartedById: instance.StartedById,
        RetryCount: 0
      };

      const result = await this.sp.web.lists.getByTitle(INSTANCES_LIST).items.add(itemData);

      logger.info('WorkflowInstanceService', `Created workflow instance: ${result.data.Id}`);
      return await this.getById(result.data.Id);
    } catch (error) {
      logger.error('WorkflowInstanceService', 'Error creating workflow instance', error);
      throw new Error('Unable to create workflow instance.');
    }
  }

  /**
   * Update workflow instance
   */
  public async update(id: number, updates: Partial<IWorkflowInstance>): Promise<void> {
    try {
      const itemData: Record<string, unknown> = { ...updates };

      // Stringify JSON fields if objects
      if (updates.Variables !== undefined && typeof updates.Variables !== 'string') {
        itemData.Variables = JSON.stringify(updates.Variables);
      }
      if (updates.Context !== undefined && typeof updates.Context !== 'string') {
        itemData.Context = JSON.stringify(updates.Context);
      }

      await this.sp.web.lists.getByTitle(INSTANCES_LIST).items
        .getById(id)
        .update(itemData);

      logger.info('WorkflowInstanceService', `Updated workflow instance: ${id}`);
    } catch (error) {
      logger.error('WorkflowInstanceService', `Error updating workflow instance ${id}`, error);
      throw new Error('Unable to update workflow instance.');
    }
  }

  /**
   * Update instance status
   * INTEGRATION FIX: Now also syncs process status bidirectionally
   */
  public async updateStatus(id: number, status: WorkflowInstanceStatus, errorMessage?: string): Promise<void> {
    try {
      const updates: Record<string, unknown> = { Status: status };

      if (errorMessage) {
        updates.ErrorMessage = errorMessage;
      }

      if (status === WorkflowInstanceStatus.Completed || status === WorkflowInstanceStatus.Failed || status === WorkflowInstanceStatus.Cancelled) {
        updates.CompletedDate = new Date().toISOString();
      }

      await this.sp.web.lists.getByTitle(INSTANCES_LIST).items
        .getById(id)
        .update(updates);

      logger.info('WorkflowInstanceService', `Updated status for instance ${id} to ${status}`);

      // INTEGRATION FIX: Sync process status when workflow status changes
      await this.syncProcessStatusFromWorkflow(id, status);
    } catch (error) {
      logger.error('WorkflowInstanceService', `Error updating status for instance ${id}`, error);
      throw new Error('Unable to update workflow status.');
    }
  }

  /**
   * INTEGRATION FIX: Sync process status when workflow status changes
   * This ensures bidirectional sync between process and workflow
   */
  private async syncProcessStatusFromWorkflow(instanceId: number, workflowStatus: WorkflowInstanceStatus): Promise<void> {
    try {
      // Get the workflow instance to find the process ID
      const instance = await this.sp.web.lists.getByTitle(INSTANCES_LIST).items
        .getById(instanceId)
        .select('ProcessId')();

      const processId = instance.ProcessId;
      if (!processId) {
        logger.info('WorkflowInstanceService', `No ProcessId found for instance ${instanceId}, skipping sync`);
        return;
      }

      // Map workflow status to process status
      let processStatus: ProcessStatus | null = null;

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
          processStatus = null;
      }

      if (processStatus) {
        const processUpdates: Record<string, unknown> = {
          ProcessStatus: processStatus
        };

        if (processStatus === ProcessStatus.Completed) {
          processUpdates.ActualCompletionDate = new Date().toISOString();
          processUpdates.ProgressPercentage = 100;
        }

        await this.sp.web.lists.getByTitle('PM_Processes').items
          .getById(processId)
          .update(processUpdates);

        logger.info('WorkflowInstanceService',
          `Synced process ${processId} to status ${processStatus} from workflow status ${workflowStatus}`);
      }
    } catch (error) {
      // Don't throw - workflow update was successful, process sync is secondary
      logger.warn('WorkflowInstanceService',
        `Failed to sync process status from workflow instance ${instanceId}`, error);
    }
  }

  /**
   * Update progress
   */
  public async updateProgress(id: number, currentStepId: string, currentStepName: string, completedSteps: number, totalSteps: number): Promise<void> {
    try {
      const progressPercentage = totalSteps > 0 ? Math.round((completedSteps / totalSteps) * 100) : 0;

      await this.sp.web.lists.getByTitle(INSTANCES_LIST).items
        .getById(id)
        .update({
          CurrentStepId: currentStepId,
          CurrentStepName: currentStepName,
          CompletedSteps: completedSteps,
          TotalSteps: totalSteps,
          ProgressPercentage: progressPercentage
        });
    } catch (error) {
      logger.error('WorkflowInstanceService', `Error updating progress for instance ${id}`, error);
      throw new Error('Unable to update workflow progress.');
    }
  }

  /**
   * Update instance variables
   */
  public async updateVariables(id: number, variables: Record<string, unknown>): Promise<void> {
    try {
      const instance = await this.getById(id);
      let currentVars: Record<string, unknown> = {};

      try {
        currentVars = instance.Variables ? JSON.parse(instance.Variables) : {};
      } catch { /* ignore */ }

      const mergedVars = { ...currentVars, ...variables };

      await this.sp.web.lists.getByTitle(INSTANCES_LIST).items
        .getById(id)
        .update({
          Variables: JSON.stringify(mergedVars)
        });
    } catch (error) {
      logger.error('WorkflowInstanceService', `Error updating variables for instance ${id}`, error);
      throw new Error('Unable to update workflow variables.');
    }
  }

  /**
   * Set error state
   */
  public async setError(id: number, stepId: string, errorMessage: string): Promise<void> {
    try {
      const instance = await this.getById(id);

      await this.sp.web.lists.getByTitle(INSTANCES_LIST).items
        .getById(id)
        .update({
          Status: WorkflowInstanceStatus.Failed,
          ErrorStepId: stepId,
          ErrorMessage: errorMessage,
          RetryCount: (instance.RetryCount || 0) + 1,
          LastRetryDate: new Date().toISOString()
        });

      logger.warn('WorkflowInstanceService', `Set error state for instance ${id}: ${errorMessage}`);
    } catch (error) {
      logger.error('WorkflowInstanceService', `Error setting error state for instance ${id}`, error);
      throw new Error('Unable to set workflow error state.');
    }
  }

  /**
   * Cancel workflow instance
   */
  public async cancel(id: number, reason?: string): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(INSTANCES_LIST).items
        .getById(id)
        .update({
          Status: WorkflowInstanceStatus.Cancelled,
          CompletedDate: new Date().toISOString(),
          ErrorMessage: reason || 'Cancelled by user'
        });

      // Log cancellation
      await this.addLog(id, undefined, undefined, 'Workflow Cancelled', LogLevel.Info, reason || 'Cancelled by user');

      logger.info('WorkflowInstanceService', `Cancelled workflow instance: ${id}`);
    } catch (error) {
      logger.error('WorkflowInstanceService', `Error cancelling workflow instance ${id}`, error);
      throw new Error('Unable to cancel workflow.');
    }
  }

  /**
   * Pause workflow instance
   */
  public async pause(id: number): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(INSTANCES_LIST).items
        .getById(id)
        .update({
          Status: WorkflowInstanceStatus.Paused
        });

      await this.addLog(id, undefined, undefined, 'Workflow Paused', LogLevel.Info, 'Workflow paused by user');
      logger.info('WorkflowInstanceService', `Paused workflow instance: ${id}`);
    } catch (error) {
      logger.error('WorkflowInstanceService', `Error pausing workflow instance ${id}`, error);
      throw new Error('Unable to pause workflow.');
    }
  }

  /**
   * Resume paused workflow instance
   */
  public async resume(id: number): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(INSTANCES_LIST).items
        .getById(id)
        .update({
          Status: WorkflowInstanceStatus.Running
        });

      await this.addLog(id, undefined, undefined, 'Workflow Resumed', LogLevel.Info, 'Workflow resumed by user');
      logger.info('WorkflowInstanceService', `Resumed workflow instance: ${id}`);
    } catch (error) {
      logger.error('WorkflowInstanceService', `Error resuming workflow instance ${id}`, error);
      throw new Error('Unable to resume workflow.');
    }
  }

  // ============================================================================
  // STEP STATUS OPERATIONS
  // ============================================================================

  /**
   * Get step statuses for an instance
   */
  public async getStepStatuses(instanceId: number): Promise<IWorkflowStepStatus[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(STEP_STATUS_LIST).items
        .filter(`WorkflowInstanceId eq ${instanceId}`)
        .orderBy('Order', true)();

      return items as IWorkflowStepStatus[];
    } catch (error) {
      logger.error('WorkflowInstanceService', `Error fetching step statuses for instance ${instanceId}`, error);
      throw new Error('Unable to retrieve step statuses.');
    }
  }

  /**
   * Get step status by step ID
   */
  public async getStepStatus(instanceId: number, stepId: string): Promise<IWorkflowStepStatus | undefined> {
    try {
      const items = await this.sp.web.lists.getByTitle(STEP_STATUS_LIST).items
        .filter(`WorkflowInstanceId eq ${instanceId} and StepId eq '${stepId}'`)
        .top(1)();

      return items.length > 0 ? items[0] as IWorkflowStepStatus : undefined;
    } catch (error) {
      logger.error('WorkflowInstanceService', `Error fetching step status for ${stepId}`, error);
      return undefined;
    }
  }

  /**
   * Create step status record
   */
  public async createStepStatus(status: Partial<IWorkflowStepStatus>): Promise<IWorkflowStepStatus> {
    try {
      const itemData = {
        Title: status.StepName,
        WorkflowInstanceId: status.WorkflowInstanceId,
        StepId: status.StepId,
        StepName: status.StepName,
        Status: status.Status || StepStatus.Pending,
        Order: status.Order,
        Result: typeof status.Result === 'string' ? status.Result : JSON.stringify(status.Result || {}),
        OutputVariables: typeof status.OutputVariables === 'string' ? status.OutputVariables : JSON.stringify(status.OutputVariables || {}),
        RetryCount: 0
      };

      const result = await this.sp.web.lists.getByTitle(STEP_STATUS_LIST).items.add(itemData);
      return result.data as IWorkflowStepStatus;
    } catch (error) {
      logger.error('WorkflowInstanceService', 'Error creating step status', error);
      throw new Error('Unable to create step status.');
    }
  }

  /**
   * Update step status
   */
  public async updateStepStatus(statusId: number, updates: Partial<IWorkflowStepStatus>): Promise<void> {
    try {
      const itemData: Record<string, unknown> = { ...updates };

      if (updates.Result !== undefined && typeof updates.Result !== 'string') {
        itemData.Result = JSON.stringify(updates.Result);
      }
      if (updates.OutputVariables !== undefined && typeof updates.OutputVariables !== 'string') {
        itemData.OutputVariables = JSON.stringify(updates.OutputVariables);
      }

      await this.sp.web.lists.getByTitle(STEP_STATUS_LIST).items
        .getById(statusId)
        .update(itemData);
    } catch (error) {
      logger.error('WorkflowInstanceService', `Error updating step status ${statusId}`, error);
      throw new Error('Unable to update step status.');
    }
  }

  /**
   * Start step execution
   */
  public async startStep(instanceId: number, stepId: string): Promise<void> {
    const status = await this.getStepStatus(instanceId, stepId);
    if (status) {
      await this.updateStepStatus(status.Id, {
        Status: StepStatus.InProgress,
        StartedDate: new Date()
      });
    }
  }

  /**
   * Complete step execution
   */
  public async completeStep(instanceId: number, stepId: string, result?: Record<string, unknown>, outputVars?: Record<string, unknown>): Promise<void> {
    const status = await this.getStepStatus(instanceId, stepId);
    if (status) {
      const now = new Date();
      const startedDate = status.StartedDate ? new Date(status.StartedDate) : now;
      const durationMinutes = Math.round((now.getTime() - startedDate.getTime()) / 60000);

      await this.updateStepStatus(status.Id, {
        Status: StepStatus.Completed,
        CompletedDate: now,
        Duration: durationMinutes,
        Result: JSON.stringify(result || { success: true }),
        OutputVariables: JSON.stringify(outputVars || {})
      });
    }
  }

  /**
   * Fail step execution
   */
  public async failStep(instanceId: number, stepId: string, errorMessage: string): Promise<void> {
    const status = await this.getStepStatus(instanceId, stepId);
    if (status) {
      await this.updateStepStatus(status.Id, {
        Status: StepStatus.Failed,
        ErrorMessage: errorMessage,
        RetryCount: (status.RetryCount || 0) + 1
      });
    }
  }

  /**
   * Skip step
   */
  public async skipStep(instanceId: number, stepId: string, reason?: string): Promise<void> {
    const status = await this.getStepStatus(instanceId, stepId);
    if (status) {
      await this.updateStepStatus(status.Id, {
        Status: StepStatus.Skipped,
        Result: JSON.stringify({ skipped: true, reason: reason || 'Condition not met' })
      });
    }
  }

  // ============================================================================
  // LOGGING OPERATIONS
  // ============================================================================

  /**
   * Add log entry
   */
  public async addLog(
    instanceId: number,
    stepId?: string,
    stepName?: string,
    action: string = 'Unknown',
    level: LogLevel = LogLevel.Info,
    message: string = '',
    details?: Record<string, unknown>,
    userId?: number
  ): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(LOGS_LIST).items.add({
        Title: action,
        WorkflowInstanceId: instanceId,
        StepId: stepId,
        StepName: stepName,
        Action: action,
        Level: level,
        Message: message,
        Details: details ? JSON.stringify(details) : undefined,
        Timestamp: new Date().toISOString(),
        UserId: userId
      });
    } catch (error) {
      // Don't throw on logging errors - just warn
      logger.warn('WorkflowInstanceService', `Failed to add log entry for instance ${instanceId}`, error);
    }
  }

  /**
   * Get logs for instance
   */
  public async getLogs(instanceId: number, level?: LogLevel, top?: number): Promise<IWorkflowLog[]> {
    try {
      let query = this.sp.web.lists.getByTitle(LOGS_LIST).items
        .filter(`WorkflowInstanceId eq ${instanceId}`)
        .orderBy('Timestamp', false);

      if (level) {
        query = query.filter(`Level eq '${level}'`);
      }
      if (top) {
        query = query.top(top);
      }

      const items = await query();
      return items as IWorkflowLog[];
    } catch (error) {
      logger.error('WorkflowInstanceService', `Error fetching logs for instance ${instanceId}`, error);
      return [];
    }
  }

  /**
   * Get error logs for instance
   */
  public async getErrorLogs(instanceId: number): Promise<IWorkflowLog[]> {
    return this.getLogs(instanceId, LogLevel.Error);
  }

  // ============================================================================
  // STATISTICS
  // ============================================================================

  /**
   * Get statistics for a workflow definition
   */
  public async getDefinitionStats(definitionId: number): Promise<{
    totalInstances: number;
    completed: number;
    failed: number;
    cancelled: number;
    running: number;
    successRate: number;
    averageDurationHours: number;
  }> {
    try {
      const instances = await this.getByDefinition(definitionId);

      const completed = instances.filter(i => i.Status === WorkflowInstanceStatus.Completed);
      const failed = instances.filter(i => i.Status === WorkflowInstanceStatus.Failed);
      const cancelled = instances.filter(i => i.Status === WorkflowInstanceStatus.Cancelled);
      const running = instances.filter(i =>
        i.Status === WorkflowInstanceStatus.Running ||
        i.Status === WorkflowInstanceStatus.Pending ||
        i.Status === WorkflowInstanceStatus.Paused
      );

      // Calculate average duration for completed instances
      let totalDurationMs = 0;
      let countWithDuration = 0;

      completed.forEach(i => {
        if (i.StartedDate && i.CompletedDate) {
          const start = new Date(i.StartedDate).getTime();
          const end = new Date(i.CompletedDate).getTime();
          totalDurationMs += end - start;
          countWithDuration++;
        }
      });

      const avgDurationHours = countWithDuration > 0
        ? (totalDurationMs / countWithDuration) / (1000 * 60 * 60)
        : 0;

      const successRate = (completed.length + failed.length) > 0
        ? Math.round((completed.length / (completed.length + failed.length)) * 100)
        : 0;

      return {
        totalInstances: instances.length,
        completed: completed.length,
        failed: failed.length,
        cancelled: cancelled.length,
        running: running.length,
        successRate,
        averageDurationHours: Math.round(avgDurationHours * 10) / 10
      };
    } catch (error) {
      logger.error('WorkflowInstanceService', `Error calculating stats for definition ${definitionId}`, error);
      throw new Error('Unable to calculate workflow statistics.');
    }
  }

  // ============================================================================
  // UI VIEW MODEL OPERATIONS
  // ============================================================================

  /**
   * Get all workflow instances in UI-friendly format (camelCase)
   */
  public async getAllUI(): Promise<IWorkflowInstanceUI[]> {
    try {
      const instances = await this.getAll();
      return instances.map(instance => this.mapToUI(instance));
    } catch (error) {
      logger.error('WorkflowInstanceService', 'Error fetching UI workflow instances', error);
      throw new Error('Unable to retrieve workflow instances.');
    }
  }

  /**
   * Get workflow instance by ID in UI-friendly format
   */
  public async getByIdUI(id: number): Promise<IWorkflowInstanceUI> {
    try {
      const instance = await this.getById(id);
      return this.mapToUI(instance);
    } catch (error) {
      logger.error('WorkflowInstanceService', `Error fetching UI workflow instance ${id}`, error);
      throw new Error('Workflow instance not found.');
    }
  }

  /**
   * Map SharePoint instance to UI-friendly format
   */
  private mapToUI(instance: IWorkflowInstance): IWorkflowInstanceUI {
    // Parse context if available
    let context: IWorkflowContext | null = null;
    try {
      if (instance.Context) {
        context = typeof instance.Context === 'string'
          ? JSON.parse(instance.Context)
          : instance.Context;
      }
    } catch {
      context = null;
    }

    // Map status to UI-friendly values
    const mapStatus = (status: WorkflowInstanceStatus): IWorkflowInstanceUI['status'] => {
      switch (status) {
        case WorkflowInstanceStatus.Pending: return 'NotStarted';
        case WorkflowInstanceStatus.Running: return 'InProgress';
        case WorkflowInstanceStatus.WaitingForApproval: return 'WaitingApproval';
        case WorkflowInstanceStatus.Completed: return 'Completed';
        case WorkflowInstanceStatus.Failed: return 'Failed';
        case WorkflowInstanceStatus.Cancelled: return 'Cancelled';
        default: return 'NotStarted';
      }
    };

    // Calculate SLA status
    const calculateSLAStatus = (): 'Healthy' | 'Warning' | 'Breached' | undefined => {
      if (!instance.EstimatedCompletionDate) return undefined;
      const now = new Date();
      const estimatedDate = new Date(instance.EstimatedCompletionDate);
      const hoursRemaining = (estimatedDate.getTime() - now.getTime()) / (1000 * 60 * 60);

      if (hoursRemaining < 0) return 'Breached';
      if (hoursRemaining < 24) return 'Warning';
      return 'Healthy';
    };

    return {
      id: String(instance.Id),
      workflowDefinitionId: String(instance.WorkflowDefinitionId),
      workflowName: instance.Title || 'Unnamed Workflow',
      processType: context?.processType || 'Joiner' as ProcessType,
      status: mapStatus(instance.Status),
      employeeId: context?.processId,
      employeeName: context?.employeeName,
      employeeEmail: context?.employeeEmail,
      currentStepId: instance.CurrentStepId,
      currentStepName: instance.CurrentStepName,
      currentStepIndex: instance.CompletedSteps ? instance.CompletedSteps + 1 : 1,
      totalSteps: instance.TotalSteps,
      startedDate: instance.StartedDate?.toISOString() || new Date().toISOString(),
      completedDate: instance.CompletedDate?.toISOString(),
      estimatedCompletionDate: instance.EstimatedCompletionDate?.toISOString(),
      slaStatus: calculateSLAStatus(),
      startedBy: context?.startedBy
    };
  }
}
