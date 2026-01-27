// @ts-nocheck
/**
 * ProcessService
 * Enhanced process management service with workflow integration
 *
 * This service provides CRUD operations for JML processes with automatic
 * workflow synchronization. It wraps SPService and adds workflow-aware
 * functionality.
 */

import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import { logger } from './LoggingService';
import { SPService } from './SPService';
import { IJmlProcess } from '../models/IJmlProcess';
import { ProcessType, ProcessStatus, TaskStatus } from '../models/ICommon';
import { WorkflowInstanceStatus } from '../models/IWorkflow';
import { WorkflowInstanceService } from './workflow/WorkflowInstanceService';

// ============================================================================
// INTERFACES
// ============================================================================

/**
 * Extended process with workflow information
 */
export interface IProcessExtended extends IJmlProcess {
  // Workflow info
  hasWorkflow: boolean;
  workflowInstanceId?: number;
  workflowStatus?: WorkflowInstanceStatus;
  workflowCurrentStep?: string;
  workflowProgress?: number;

  // Computed fields
  isOverdue: boolean;
  daysRemaining: number;
  daysElapsed: number;
  completionRate: number;
}

/**
 * Process filter options
 */
export interface IProcessFilterOptions {
  processType?: ProcessType | ProcessType[];
  status?: ProcessStatus | ProcessStatus[];
  department?: string;
  managerId?: number;
  processOwnerId?: number;
  dateFrom?: Date;
  dateTo?: Date;
  isOverdue?: boolean;
  searchTerm?: string;
}

/**
 * Process summary for dashboards
 */
export interface IProcessSummary {
  id: number;
  title: string;
  employeeName: string;
  processType: ProcessType;
  status: ProcessStatus;
  progress: number;
  startDate: Date;
  targetDate: Date;
  isOverdue: boolean;
  daysRemaining: number;
  hasWorkflow: boolean;
  workflowStatus?: WorkflowInstanceStatus;
}

/**
 * Process statistics
 */
export interface IProcessStatistics {
  total: number;
  byStatus: Record<ProcessStatus, number>;
  byType: Record<ProcessType, number>;
  overdue: number;
  completedThisMonth: number;
  averageCompletionDays: number;
  onTimeCompletionRate: number;
}

// ============================================================================
// PROCESS SERVICE
// ============================================================================

export class ProcessService {
  private sp: SPFI;
  private context: WebPartContext;
  private spService: SPService;
  private workflowInstanceService: WorkflowInstanceService;

  constructor(sp: SPFI, context: WebPartContext) {
    this.sp = sp;
    this.context = context;
    this.spService = new SPService(sp);
    this.workflowInstanceService = new WorkflowInstanceService(sp);
  }

  // ============================================================================
  // READ OPERATIONS
  // ============================================================================

  /**
   * Get all processes with optional filtering
   */
  public async getAll(options?: IProcessFilterOptions): Promise<IJmlProcess[]> {
    const filter = this.buildFilterString(options);
    return await this.spService.getProcesses(filter, 'Created desc');
  }

  /**
   * Get processes with extended workflow information
   */
  public async getAllExtended(options?: IProcessFilterOptions): Promise<IProcessExtended[]> {
    const processes = await this.getAll(options);
    return await Promise.all(processes.map(p => this.enrichWithWorkflow(p)));
  }

  /**
   * Get process summaries for dashboard display
   */
  public async getSummaries(options?: IProcessFilterOptions, top?: number): Promise<IProcessSummary[]> {
    const filter = this.buildFilterString(options);
    const processes = await this.spService.getProcesses(filter, 'Created desc', top || 50);

    const summaries: IProcessSummary[] = [];

    for (const process of processes) {
      const workflowInstance = await this.getWorkflowForProcess(process.Id!);
      const now = new Date();
      const targetDate = new Date(process.TargetCompletionDate);
      const daysRemaining = Math.ceil((targetDate.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));

      summaries.push({
        id: process.Id!,
        title: process.Title,
        employeeName: process.EmployeeName,
        processType: process.ProcessType,
        status: process.ProcessStatus,
        progress: process.ProgressPercentage || 0,
        startDate: new Date(process.StartDate),
        targetDate: targetDate,
        isOverdue: daysRemaining < 0 && process.ProcessStatus !== ProcessStatus.Completed,
        daysRemaining: Math.max(0, daysRemaining),
        hasWorkflow: !!workflowInstance,
        workflowStatus: workflowInstance?.Status
      });
    }

    return summaries;
  }

  /**
   * Get process by ID with workflow information
   */
  public async getById(id: number): Promise<IProcessExtended> {
    const process = await this.spService.getProcessById(id);
    return await this.enrichWithWorkflow(process);
  }

  /**
   * Get processes by status
   */
  public async getByStatus(status: ProcessStatus | ProcessStatus[]): Promise<IJmlProcess[]> {
    return await this.getAll({ status });
  }

  /**
   * Get processes by type
   */
  public async getByType(processType: ProcessType | ProcessType[]): Promise<IJmlProcess[]> {
    return await this.getAll({ processType });
  }

  /**
   * Get processes for a manager
   */
  public async getByManager(managerId: number): Promise<IJmlProcess[]> {
    return await this.getAll({ managerId });
  }

  /**
   * Get overdue processes
   */
  public async getOverdue(): Promise<IJmlProcess[]> {
    return await this.getAll({ isOverdue: true });
  }

  /**
   * Search processes
   */
  public async search(searchTerm: string): Promise<IJmlProcess[]> {
    return await this.getAll({ searchTerm });
  }

  // ============================================================================
  // WRITE OPERATIONS
  // ============================================================================

  /**
   * Create a new process
   */
  public async create(process: Partial<IJmlProcess>): Promise<IJmlProcess> {
    return await this.spService.createProcess(process);
  }

  /**
   * Update a process
   */
  public async update(id: number, updates: Partial<IJmlProcess>): Promise<void> {
    await this.spService.updateProcess(id, updates);

    // If status changed, sync with workflow
    if (updates.ProcessStatus) {
      await this.syncWorkflowStatus(id, updates.ProcessStatus);
    }
  }

  /**
   * Update process status with workflow sync
   */
  public async updateStatus(id: number, status: ProcessStatus, reason?: string): Promise<void> {
    const updates: Partial<IJmlProcess> = {
      ProcessStatus: status
    };

    if (status === ProcessStatus.Completed) {
      updates.ActualCompletionDate = new Date();
      updates.ProgressPercentage = 100;
    }

    if (reason) {
      const process = await this.spService.getProcessById(id);
      updates.Comments = process.Comments
        ? `${process.Comments}\n${new Date().toISOString()}: ${reason}`
        : `${new Date().toISOString()}: ${reason}`;
    }

    await this.update(id, updates);
  }

  /**
   * Delete a process (soft delete)
   */
  public async delete(id: number): Promise<void> {
    // Cancel any associated workflow first
    const workflowInstance = await this.getWorkflowForProcess(id);
    if (workflowInstance) {
      await this.workflowInstanceService.cancel(workflowInstance.Id, 'Process deleted');
    }

    // Soft delete by updating status
    await this.update(id, {
      ProcessStatus: ProcessStatus.Cancelled,
      IsDeleted: true
    });
  }

  /**
   * Hard delete a process (use with caution)
   */
  public async hardDelete(id: number): Promise<void> {
    // Cancel workflow first
    const workflowInstance = await this.getWorkflowForProcess(id);
    if (workflowInstance) {
      await this.workflowInstanceService.cancel(workflowInstance.Id, 'Process hard deleted');
    }

    await this.spService.deleteProcess(id);
  }

  // ============================================================================
  // PROGRESS OPERATIONS
  // ============================================================================

  /**
   * Recalculate process progress from tasks
   */
  public async recalculateProgress(processId: number): Promise<{
    totalTasks: number;
    completedTasks: number;
    progressPercentage: number;
    overdueTasks: number;
  }> {
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

    return { totalTasks, completedTasks, progressPercentage, overdueTasks };
  }

  /**
   * Check if all tasks are complete
   */
  public async areAllTasksComplete(processId: number): Promise<boolean> {
    const tasks = await this.spService.getTaskAssignmentsByProcess(processId);
    if (tasks.length === 0) return false;

    return tasks.every(t =>
      t.Status === TaskStatus.Completed || t.Status === TaskStatus.Skipped
    );
  }

  // ============================================================================
  // STATISTICS
  // ============================================================================

  /**
   * Get process statistics
   */
  public async getStatistics(options?: IProcessFilterOptions): Promise<IProcessStatistics> {
    const processes = await this.getAll(options);
    const now = new Date();
    const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);

    const byStatus: Record<ProcessStatus, number> = {
      [ProcessStatus.Draft]: 0,
      [ProcessStatus.NotStarted]: 0,
      [ProcessStatus.Pending]: 0,
      [ProcessStatus.PendingApproval]: 0,
      [ProcessStatus.InProgress]: 0,
      [ProcessStatus.OnHold]: 0,
      [ProcessStatus.Completed]: 0,
      [ProcessStatus.Cancelled]: 0,
      [ProcessStatus.Archived]: 0
    };

    const byType: Record<ProcessType, number> = {
      [ProcessType.Joiner]: 0,
      [ProcessType.Mover]: 0,
      [ProcessType.Leaver]: 0
    };

    let overdue = 0;
    let completedThisMonth = 0;
    let totalCompletionDays = 0;
    let onTimeCount = 0;
    let completedCount = 0;

    for (const process of processes) {
      byStatus[process.ProcessStatus]++;
      byType[process.ProcessType]++;

      // Check overdue
      if (process.ProcessStatus !== ProcessStatus.Completed && process.ProcessStatus !== ProcessStatus.Cancelled) {
        const targetDate = new Date(process.TargetCompletionDate);
        if (targetDate < now) {
          overdue++;
        }
      }

      // Check completed this month
      if (process.ProcessStatus === ProcessStatus.Completed && process.ActualCompletionDate) {
        const completedDate = new Date(process.ActualCompletionDate);
        if (completedDate >= startOfMonth) {
          completedThisMonth++;
        }

        // Calculate completion time
        const startDate = new Date(process.StartDate);
        const days = Math.ceil((completedDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24));
        totalCompletionDays += days;
        completedCount++;

        // Check if on time
        const targetDate = new Date(process.TargetCompletionDate);
        if (completedDate <= targetDate) {
          onTimeCount++;
        }
      }
    }

    return {
      total: processes.length,
      byStatus,
      byType,
      overdue,
      completedThisMonth,
      averageCompletionDays: completedCount > 0 ? Math.round(totalCompletionDays / completedCount) : 0,
      onTimeCompletionRate: completedCount > 0 ? Math.round((onTimeCount / completedCount) * 100) : 0
    };
  }

  // ============================================================================
  // WORKFLOW INTEGRATION
  // ============================================================================

  /**
   * Get workflow instance for a process
   */
  public async getWorkflowForProcess(processId: number): Promise<{ Id: number; Status: WorkflowInstanceStatus } | null> {
    try {
      const instance = await this.workflowInstanceService.getActiveForProcess(processId);
      return instance ? { Id: instance.Id, Status: instance.Status } : null;
    } catch {
      return null;
    }
  }

  /**
   * Check if process has an active workflow
   */
  public async hasActiveWorkflow(processId: number): Promise<boolean> {
    const workflow = await this.getWorkflowForProcess(processId);
    return !!workflow;
  }

  /**
   * Sync workflow status when process status changes
   */
  private async syncWorkflowStatus(processId: number, newStatus: ProcessStatus): Promise<void> {
    const workflowInstance = await this.getWorkflowForProcess(processId);
    if (!workflowInstance) return;

    try {
      switch (newStatus) {
        case ProcessStatus.OnHold:
          await this.workflowInstanceService.pause(workflowInstance.Id);
          break;
        case ProcessStatus.Cancelled:
          await this.workflowInstanceService.cancel(workflowInstance.Id, 'Process cancelled');
          break;
        case ProcessStatus.InProgress:
          if (workflowInstance.Status === WorkflowInstanceStatus.Paused) {
            await this.workflowInstanceService.resume(workflowInstance.Id);
          }
          break;
      }
    } catch (error) {
      logger.warn('ProcessService', `Error syncing workflow status for process ${processId}`, error);
    }
  }

  // ============================================================================
  // PRIVATE HELPERS
  // ============================================================================

  /**
   * Build filter string from options
   */
  private buildFilterString(options?: IProcessFilterOptions): string | undefined {
    if (!options) return undefined;

    const filters: string[] = [];

    // Exclude deleted by default
    filters.push('IsDeleted ne 1');

    // Process type filter
    if (options.processType) {
      const types = Array.isArray(options.processType) ? options.processType : [options.processType];
      const typeFilters = types.map(t => `ProcessType eq '${t}'`);
      filters.push(`(${typeFilters.join(' or ')})`);
    }

    // Status filter
    if (options.status) {
      const statuses = Array.isArray(options.status) ? options.status : [options.status];
      const statusFilters = statuses.map(s => `ProcessStatus eq '${s}'`);
      filters.push(`(${statusFilters.join(' or ')})`);
    }

    // Department filter
    if (options.department) {
      filters.push(`Department eq '${options.department}'`);
    }

    // Manager filter
    if (options.managerId) {
      filters.push(`ManagerId eq ${options.managerId}`);
    }

    // Process owner filter
    if (options.processOwnerId) {
      filters.push(`ProcessOwnerId eq ${options.processOwnerId}`);
    }

    // Date range filters
    if (options.dateFrom) {
      filters.push(`StartDate ge datetime'${options.dateFrom.toISOString()}'`);
    }
    if (options.dateTo) {
      filters.push(`StartDate le datetime'${options.dateTo.toISOString()}'`);
    }

    // Overdue filter
    if (options.isOverdue) {
      const today = new Date().toISOString();
      filters.push(`TargetCompletionDate lt datetime'${today}'`);
      filters.push(`ProcessStatus ne '${ProcessStatus.Completed}'`);
      filters.push(`ProcessStatus ne '${ProcessStatus.Cancelled}'`);
    }

    // Search term (title or employee name)
    if (options.searchTerm) {
      const term = options.searchTerm.replace(/'/g, "''");
      filters.push(`(substringof('${term}', Title) or substringof('${term}', EmployeeName))`);
    }

    return filters.length > 0 ? filters.join(' and ') : undefined;
  }

  /**
   * Enrich process with workflow information
   */
  private async enrichWithWorkflow(process: IJmlProcess): Promise<IProcessExtended> {
    const now = new Date();
    const startDate = new Date(process.StartDate);
    const targetDate = new Date(process.TargetCompletionDate);
    const daysRemaining = Math.ceil((targetDate.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));
    const daysElapsed = Math.ceil((now.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24));

    const extended: IProcessExtended = {
      ...process,
      hasWorkflow: false,
      isOverdue: daysRemaining < 0 && process.ProcessStatus !== ProcessStatus.Completed,
      daysRemaining: Math.max(0, daysRemaining),
      daysElapsed: Math.max(0, daysElapsed),
      completionRate: process.TotalTasks && process.TotalTasks > 0
        ? Math.round(((process.CompletedTasks || 0) / process.TotalTasks) * 100)
        : 0
    };

    // Try to get workflow info
    try {
      const workflowInstance = await this.workflowInstanceService.getActiveForProcess(process.Id!);
      if (workflowInstance) {
        extended.hasWorkflow = true;
        extended.workflowInstanceId = workflowInstance.Id;
        extended.workflowStatus = workflowInstance.Status;
        extended.workflowCurrentStep = workflowInstance.CurrentStepName;
        extended.workflowProgress = workflowInstance.ProgressPercentage;
      }
    } catch {
      // Ignore workflow errors
    }

    return extended;
  }
}

export default ProcessService;
