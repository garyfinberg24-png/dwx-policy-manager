// @ts-nocheck
/* eslint-disable */
/**
 * WorkflowSchedulerService
 * Handles timer-based workflow processing including:
 * - Scheduled step execution
 * - SLA warnings and breaches
 * - Reminders and escalations
 * - Resuming waiting workflows
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import {
  IWorkflowScheduledItem,
  IWorkflowInstance,
  IWorkflowStepStatus,
  ScheduledActionType,
  ScheduledItemStatus,
  WorkflowInstanceStatus,
  StepStatus,
  LogLevel
} from '../../models/IWorkflow';
import { WorkflowInstanceService } from './WorkflowInstanceService';
import { NotificationActionHandler } from './handlers/NotificationActionHandler';
import { logger } from '../LoggingService';

// List name constants
const SCHEDULE_LIST = 'JML_WorkflowSchedule';
const INSTANCES_LIST = 'JML_WorkflowInstances';
const STEP_STATUS_LIST = 'JML_WorkflowStepStatus';

export interface ISchedulerResult {
  processed: number;
  succeeded: number;
  failed: number;
  errors: string[];
}

export class WorkflowSchedulerService {
  private sp: SPFI;
  private context: WebPartContext;
  private instanceService: WorkflowInstanceService;
  private notificationHandler: NotificationActionHandler;

  constructor(sp: SPFI, context: WebPartContext) {
    this.sp = sp;
    this.context = context;
    this.instanceService = new WorkflowInstanceService(sp);
    this.notificationHandler = new NotificationActionHandler(sp, context);
  }

  // ============================================================================
  // SCHEDULED ITEM MANAGEMENT
  // ============================================================================

  /**
   * Schedule a new item for future processing
   */
  public async scheduleItem(item: Partial<IWorkflowScheduledItem>): Promise<number> {
    try {
      const itemData = {
        Title: `${item.ActionType} - Instance ${item.WorkflowInstanceId}`,
        WorkflowInstanceId: item.WorkflowInstanceId,
        StepId: item.StepId,
        ActionType: item.ActionType,
        ScheduledDate: item.ScheduledDate?.toISOString() || new Date().toISOString(),
        Status: ScheduledItemStatus.Pending,
        RetryCount: 0,
        ActionConfig: item.ActionConfig ? JSON.stringify(item.ActionConfig) : undefined
      };

      const result = await this.sp.web.lists.getByTitle(SCHEDULE_LIST).items.add(itemData);

      logger.info('WorkflowSchedulerService', `Scheduled ${item.ActionType} for instance ${item.WorkflowInstanceId}`);
      return result.data.Id;
    } catch (error) {
      logger.error('WorkflowSchedulerService', 'Error scheduling item', error);
      throw new Error('Unable to schedule workflow item.');
    }
  }

  /**
   * Schedule SLA warning
   */
  public async scheduleSLAWarning(
    instanceId: number,
    stepId: string,
    warningDate: Date,
    escalateTo?: string
  ): Promise<number> {
    return this.scheduleItem({
      WorkflowInstanceId: instanceId,
      StepId: stepId,
      ActionType: ScheduledActionType.SLAWarning,
      ScheduledDate: warningDate,
      ActionConfig: { escalateTo }
    });
  }

  /**
   * Schedule SLA breach
   */
  public async scheduleSLABreach(
    instanceId: number,
    stepId: string,
    breachDate: Date,
    escalateTo?: string
  ): Promise<number> {
    return this.scheduleItem({
      WorkflowInstanceId: instanceId,
      StepId: stepId,
      ActionType: ScheduledActionType.SLABreach,
      ScheduledDate: breachDate,
      ActionConfig: { escalateTo }
    });
  }

  /**
   * Schedule a reminder
   */
  public async scheduleReminder(
    instanceId: number,
    stepId: string,
    reminderDate: Date,
    recipientId: number,
    message: string
  ): Promise<number> {
    return this.scheduleItem({
      WorkflowInstanceId: instanceId,
      StepId: stepId,
      ActionType: ScheduledActionType.Reminder,
      ScheduledDate: reminderDate,
      ActionConfig: { recipientId, message }
    });
  }

  /**
   * Schedule step execution (for wait/delay steps)
   */
  public async scheduleStepExecution(
    instanceId: number,
    stepId: string,
    executeDate: Date
  ): Promise<number> {
    return this.scheduleItem({
      WorkflowInstanceId: instanceId,
      StepId: stepId,
      ActionType: ScheduledActionType.ExecuteStep,
      ScheduledDate: executeDate
    });
  }

  /**
   * Cancel scheduled items for an instance
   */
  public async cancelScheduledItems(instanceId: number, stepId?: string): Promise<void> {
    try {
      let filter = `WorkflowInstanceId eq ${instanceId} and Status eq '${ScheduledItemStatus.Pending}'`;
      if (stepId) {
        filter += ` and StepId eq '${stepId}'`;
      }

      const items = await this.sp.web.lists.getByTitle(SCHEDULE_LIST).items
        .filter(filter)
        .select('Id')();

      for (const item of items) {
        await this.sp.web.lists.getByTitle(SCHEDULE_LIST).items
          .getById(item.Id)
          .update({ Status: ScheduledItemStatus.Cancelled });
      }

      logger.info('WorkflowSchedulerService', `Cancelled ${items.length} scheduled items for instance ${instanceId}`);
    } catch (error) {
      logger.error('WorkflowSchedulerService', 'Error cancelling scheduled items', error);
    }
  }

  // ============================================================================
  // PROCESSING METHODS
  // ============================================================================

  /**
   * Process all due scheduled items
   * This should be called periodically (e.g., via Power Automate or Azure Function)
   */
  public async processDueItems(): Promise<ISchedulerResult> {
    const result: ISchedulerResult = {
      processed: 0,
      succeeded: 0,
      failed: 0,
      errors: []
    };

    try {
      const now = new Date().toISOString();

      // Get all pending items that are due
      const dueItems = await this.sp.web.lists.getByTitle(SCHEDULE_LIST).items
        .filter(`Status eq '${ScheduledItemStatus.Pending}' and ScheduledDate le datetime'${now}'`)
        .orderBy('ScheduledDate', true)
        .top(50)(); // Process in batches

      logger.info('WorkflowSchedulerService', `Found ${dueItems.length} due scheduled items`);

      for (const item of dueItems) {
        result.processed++;

        try {
          // Mark as processing
          await this.sp.web.lists.getByTitle(SCHEDULE_LIST).items
            .getById(item.Id)
            .update({ Status: ScheduledItemStatus.Processing });

          // Process based on action type
          await this.processScheduledItem(item);

          // Mark as completed
          await this.sp.web.lists.getByTitle(SCHEDULE_LIST).items
            .getById(item.Id)
            .update({
              Status: ScheduledItemStatus.Completed,
              ProcessedDate: new Date().toISOString()
            });

          result.succeeded++;
        } catch (error) {
          const errorMessage = error instanceof Error ? error.message : 'Unknown error';
          result.failed++;
          result.errors.push(`Item ${item.Id}: ${errorMessage}`);

          // Mark as failed with retry count
          await this.sp.web.lists.getByTitle(SCHEDULE_LIST).items
            .getById(item.Id)
            .update({
              Status: item.RetryCount >= 3 ? ScheduledItemStatus.Failed : ScheduledItemStatus.Pending,
              RetryCount: (item.RetryCount || 0) + 1,
              ErrorMessage: errorMessage
            });

          logger.error('WorkflowSchedulerService', `Failed to process item ${item.Id}`, error);
        }
      }

      logger.info('WorkflowSchedulerService', `Processed ${result.processed} items: ${result.succeeded} succeeded, ${result.failed} failed`);
      return result;
    } catch (error) {
      logger.error('WorkflowSchedulerService', 'Error processing due items', error);
      throw error;
    }
  }

  /**
   * Process a single scheduled item
   */
  private async processScheduledItem(item: IWorkflowScheduledItem): Promise<void> {
    // Verify workflow instance is still active
    const instance = await this.instanceService.getById(item.WorkflowInstanceId);
    if (!instance || this.isTerminalStatus(instance.Status)) {
      logger.info('WorkflowSchedulerService', `Skipping item for terminated instance ${item.WorkflowInstanceId}`);
      return;
    }

    let config: Record<string, unknown> = {};
    try {
      if (item.ActionConfig) {
        config = typeof item.ActionConfig === 'string'
          ? JSON.parse(item.ActionConfig)
          : item.ActionConfig as Record<string, unknown>;
      }
    } catch { /* ignore */ }

    switch (item.ActionType) {
      case ScheduledActionType.ExecuteStep:
        await this.executeScheduledStep(instance, item.StepId);
        break;

      case ScheduledActionType.Reminder:
        await this.sendReminder(instance, item.StepId, config);
        break;

      case ScheduledActionType.SLAWarning:
        await this.processSLAWarning(instance, item.StepId, config);
        break;

      case ScheduledActionType.SLABreach:
        await this.processSLABreach(instance, item.StepId, config);
        break;

      case ScheduledActionType.Escalation:
        await this.processEscalation(instance, item.StepId, config);
        break;

      default:
        logger.warn('WorkflowSchedulerService', `Unknown action type: ${item.ActionType}`);
    }
  }

  /**
   * Execute a scheduled step (resume from wait)
   */
  private async executeScheduledStep(instance: IWorkflowInstance, stepId: string): Promise<void> {
    // Update instance status to running
    await this.instanceService.updateStatus(instance.Id, WorkflowInstanceStatus.Running);

    // Log the resumption
    await this.instanceService.addLog(
      instance.Id,
      stepId,
      undefined,
      'Step Resumed',
      LogLevel.Info,
      'Workflow step resumed from scheduled wait'
    );

    // Note: The actual step execution would be triggered by the WorkflowEngineService
    // This just marks the instance as ready to continue
    logger.info('WorkflowSchedulerService', `Marked instance ${instance.Id} as ready to resume at step ${stepId}`);
  }

  /**
   * Send reminder notification
   */
  private async sendReminder(
    instance: IWorkflowInstance,
    stepId: string,
    config: Record<string, unknown>
  ): Promise<void> {
    const recipientId = config.recipientId as number;
    const message = config.message as string || 'You have a pending workflow task.';

    if (!recipientId) {
      logger.warn('WorkflowSchedulerService', 'Reminder recipient not specified');
      return;
    }

    // Create reminder notification
    await this.sp.web.lists.getByTitle('JML_Notifications').items.add({
      Title: 'Workflow Reminder',
      Message: message,
      RecipientId: recipientId,
      Type: 'Reminder',
      Priority: 'Normal',
      IsRead: false,
      RelatedItemType: 'WorkflowInstance',
      RelatedItemId: instance.Id,
      WorkflowInstanceId: instance.Id,
      WorkflowStepId: stepId
    });

    // Log the reminder
    await this.instanceService.addLog(
      instance.Id,
      stepId,
      undefined,
      'Reminder Sent',
      LogLevel.Info,
      `Reminder sent to user ${recipientId}`
    );

    logger.info('WorkflowSchedulerService', `Sent reminder for instance ${instance.Id}`);
  }

  /**
   * Process SLA warning
   */
  private async processSLAWarning(
    instance: IWorkflowInstance,
    stepId: string,
    config: Record<string, unknown>
  ): Promise<void> {
    // Get step status to find assignee
    const stepStatus = await this.instanceService.getStepStatus(instance.Id, stepId);
    if (!stepStatus || stepStatus.Status === StepStatus.Completed) {
      return; // Step already complete, no warning needed
    }

    // Send warning notification
    // In a real implementation, you'd look up the step's assignee
    await this.instanceService.addLog(
      instance.Id,
      stepId,
      stepStatus?.StepName,
      'SLA Warning',
      LogLevel.Warning,
      'Step is approaching SLA deadline'
    );

    // If escalateTo is specified, notify that person
    if (config.escalateTo) {
      logger.info('WorkflowSchedulerService', `SLA warning for instance ${instance.Id}, escalate to: ${config.escalateTo}`);
    }
  }

  /**
   * Process SLA breach
   */
  private async processSLABreach(
    instance: IWorkflowInstance,
    stepId: string,
    config: Record<string, unknown>
  ): Promise<void> {
    // Get step status
    const stepStatus = await this.instanceService.getStepStatus(instance.Id, stepId);
    if (!stepStatus || stepStatus.Status === StepStatus.Completed) {
      return; // Step already complete, no breach
    }

    // Log the breach
    await this.instanceService.addLog(
      instance.Id,
      stepId,
      stepStatus?.StepName,
      'SLA Breach',
      LogLevel.Error,
      'Step has breached SLA deadline'
    );

    // If escalateTo is specified, escalate
    if (config.escalateTo) {
      await this.processEscalation(instance, stepId, config);
    }

    logger.warn('WorkflowSchedulerService', `SLA breached for instance ${instance.Id}, step ${stepId}`);
  }

  /**
   * Process escalation
   */
  private async processEscalation(
    instance: IWorkflowInstance,
    stepId: string,
    config: Record<string, unknown>
  ): Promise<void> {
    const escalateTo = config.escalateTo as string;
    const escalateToId = config.escalateToId as number;

    // Log the escalation
    await this.instanceService.addLog(
      instance.Id,
      stepId,
      undefined,
      'Escalation',
      LogLevel.Warning,
      `Task escalated to: ${escalateTo || escalateToId}`
    );

    // Create escalation notification
    if (escalateToId) {
      await this.sp.web.lists.getByTitle('JML_Notifications').items.add({
        Title: 'Workflow Escalation',
        Message: `A workflow task has been escalated to you. Please review.`,
        RecipientId: escalateToId,
        Type: 'Escalation',
        Priority: 'High',
        IsRead: false,
        RelatedItemType: 'WorkflowInstance',
        RelatedItemId: instance.Id,
        WorkflowInstanceId: instance.Id,
        WorkflowStepId: stepId
      });
    }

    logger.info('WorkflowSchedulerService', `Escalated instance ${instance.Id} to ${escalateTo || escalateToId}`);
  }

  // ============================================================================
  // WAITING WORKFLOW PROCESSING
  // ============================================================================

  /**
   * Check and resume waiting workflows
   * Checks for completed tasks/approvals that should trigger workflow resumption
   */
  public async processWaitingWorkflows(): Promise<ISchedulerResult> {
    const result: ISchedulerResult = {
      processed: 0,
      succeeded: 0,
      failed: 0,
      errors: []
    };

    try {
      // Get all waiting instances
      const waitingStatuses = [
        WorkflowInstanceStatus.WaitingForTask,
        WorkflowInstanceStatus.WaitingForApproval,
        WorkflowInstanceStatus.WaitingForInput
      ];

      const statusFilter = waitingStatuses.map(s => `Status eq '${s}'`).join(' or ');

      const waitingInstances = await this.sp.web.lists.getByTitle(INSTANCES_LIST).items
        .filter(statusFilter)
        .select('Id', 'Status', 'CurrentStepId', 'ProcessId')
        .top(50)();

      logger.info('WorkflowSchedulerService', `Found ${waitingInstances.length} waiting workflows`);

      for (const instance of waitingInstances) {
        result.processed++;

        try {
          const shouldResume = await this.checkWaitCondition(instance);

          if (shouldResume) {
            await this.instanceService.updateStatus(instance.Id, WorkflowInstanceStatus.Running);
            await this.instanceService.addLog(
              instance.Id,
              instance.CurrentStepId,
              undefined,
              'Wait Condition Met',
              LogLevel.Info,
              'Workflow ready to resume - wait condition satisfied'
            );
            result.succeeded++;
            logger.info('WorkflowSchedulerService', `Instance ${instance.Id} ready to resume`);
          }
        } catch (error) {
          const errorMessage = error instanceof Error ? error.message : 'Unknown error';
          result.failed++;
          result.errors.push(`Instance ${instance.Id}: ${errorMessage}`);
          logger.error('WorkflowSchedulerService', `Error checking wait condition for ${instance.Id}`, error);
        }
      }

      return result;
    } catch (error) {
      logger.error('WorkflowSchedulerService', 'Error processing waiting workflows', error);
      throw error;
    }
  }

  /**
   * Check if wait condition is met for an instance
   */
  private async checkWaitCondition(instance: IWorkflowInstance): Promise<boolean> {
    // Get step status to determine what we're waiting for
    const stepStatus = await this.instanceService.getStepStatus(instance.Id, instance.CurrentStepId);
    if (!stepStatus) return false;

    switch (instance.Status) {
      case WorkflowInstanceStatus.WaitingForTask:
        return await this.checkTasksComplete(instance.Id);

      case WorkflowInstanceStatus.WaitingForApproval:
        return await this.checkApprovalComplete(instance.Id);

      case WorkflowInstanceStatus.WaitingForInput:
        // Input completion would be handled by a form/UI
        return false;

      default:
        return false;
    }
  }

  /**
   * Check if all required tasks are complete
   */
  private async checkTasksComplete(instanceId: number): Promise<boolean> {
    try {
      const tasks = await this.sp.web.lists.getByTitle('JML_TaskAssignments').items
        .filter(`WorkflowInstanceId eq ${instanceId}`)
        .select('Id', 'Status')();

      if (tasks.length === 0) return true;

      const pendingTasks = tasks.filter(t =>
        t.Status !== 'Completed' && t.Status !== 'Skipped'
      );

      return pendingTasks.length === 0;
    } catch (error) {
      logger.error('WorkflowSchedulerService', `Error checking tasks for instance ${instanceId}`, error);
      return false;
    }
  }

  /**
   * Check if approval is complete
   */
  private async checkApprovalComplete(instanceId: number): Promise<boolean> {
    try {
      const approvals = await this.sp.web.lists.getByTitle('JML_Approvals').items
        .filter(`WorkflowInstanceId eq ${instanceId} and Status eq 'Pending'`)
        .select('Id')
        .top(1)();

      return approvals.length === 0; // No pending approvals means complete
    } catch (error) {
      logger.error('WorkflowSchedulerService', `Error checking approvals for instance ${instanceId}`, error);
      return false;
    }
  }

  // ============================================================================
  // SLA MONITORING
  // ============================================================================

  /**
   * Check for SLA violations across all running workflows
   */
  public async checkSLAViolations(): Promise<{
    warnings: number;
    breaches: number;
  }> {
    let warnings = 0;
    let breaches = 0;

    try {
      // Get all in-progress step statuses
      const inProgressSteps = await this.sp.web.lists.getByTitle(STEP_STATUS_LIST).items
        .filter(`Status eq '${StepStatus.InProgress}'`)
        .select('Id', 'WorkflowInstanceId', 'StepId', 'StepName', 'StartedDate')();

      const now = new Date();

      for (const step of inProgressSteps) {
        if (!step.StartedDate) continue;

        const startedDate = new Date(step.StartedDate);
        const hoursElapsed = (now.getTime() - startedDate.getTime()) / (1000 * 60 * 60);

        // These thresholds should come from step configuration
        // For now using defaults
        const warningThreshold = 24; // hours
        const breachThreshold = 48; // hours

        if (hoursElapsed >= breachThreshold) {
          breaches++;
          await this.instanceService.addLog(
            step.WorkflowInstanceId,
            step.StepId,
            step.StepName,
            'SLA Breach Detected',
            LogLevel.Error,
            `Step has been in progress for ${Math.round(hoursElapsed)} hours (breach threshold: ${breachThreshold}h)`
          );
        } else if (hoursElapsed >= warningThreshold) {
          warnings++;
          await this.instanceService.addLog(
            step.WorkflowInstanceId,
            step.StepId,
            step.StepName,
            'SLA Warning',
            LogLevel.Warning,
            `Step has been in progress for ${Math.round(hoursElapsed)} hours (warning threshold: ${warningThreshold}h)`
          );
        }
      }

      logger.info('WorkflowSchedulerService', `SLA check complete: ${warnings} warnings, ${breaches} breaches`);
      return { warnings, breaches };
    } catch (error) {
      logger.error('WorkflowSchedulerService', 'Error checking SLA violations', error);
      throw error;
    }
  }

  // ============================================================================
  // HELPER METHODS
  // ============================================================================

  /**
   * Check if status is terminal (workflow ended)
   */
  private isTerminalStatus(status: WorkflowInstanceStatus): boolean {
    return [
      WorkflowInstanceStatus.Completed,
      WorkflowInstanceStatus.Failed,
      WorkflowInstanceStatus.Cancelled
    ].includes(status);
  }

  /**
   * Get pending scheduled items count
   */
  public async getPendingCount(): Promise<number> {
    try {
      const items = await this.sp.web.lists.getByTitle(SCHEDULE_LIST).items
        .filter(`Status eq '${ScheduledItemStatus.Pending}'`)
        .select('Id')();
      return items.length;
    } catch (error) {
      logger.error('WorkflowSchedulerService', 'Error getting pending count', error);
      return 0;
    }
  }

  /**
   * Get overdue items count
   */
  public async getOverdueCount(): Promise<number> {
    try {
      const now = new Date().toISOString();
      const items = await this.sp.web.lists.getByTitle(SCHEDULE_LIST).items
        .filter(`Status eq '${ScheduledItemStatus.Pending}' and ScheduledDate lt datetime'${now}'`)
        .select('Id')();
      return items.length;
    } catch (error) {
      logger.error('WorkflowSchedulerService', 'Error getting overdue count', error);
      return 0;
    }
  }
}
