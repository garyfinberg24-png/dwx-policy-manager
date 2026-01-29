// @ts-nocheck
/**
 * TaskEscalationScheduler
 * Provides scheduling capabilities for task escalation processing
 *
 * INTEGRATION FIX: This service bridges the gap between TaskNotificationService
 * (which has escalation logic) and the system (which needs to run it periodically).
 *
 * Can be used in multiple ways:
 * 1. Client-side interval (for SPFx webparts that stay open)
 * 2. On-demand execution (triggered by user action)
 * 3. Azure Function timer trigger (recommended for production)
 * 4. Power Automate scheduled flow
 */

import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import { TaskNotificationService } from './TaskNotificationService';
import { ApprovalNotificationService } from './ApprovalNotificationService';
import { ApprovalStatus, TaskNotificationType, ITaskNotificationQueueItem } from '../models';
import { TaskStatus } from '../models/ICommon';
import { logger } from './LoggingService';

/**
 * Escalation run result
 */
export interface IEscalationRunResult {
  success: boolean;
  runId: string;
  startTime: Date;
  endTime: Date;
  durationMs: number;
  tasksProcessed: number;
  taskNotificationsSent: number;
  /** P2 INTEGRATION FIX: Due date reminder notifications sent */
  taskDueDateRemindersSent: number;
  approvalsProcessed: number;
  approvalNotificationsSent: number;
  errors: string[];
}

/**
 * Scheduler configuration
 */
export interface ISchedulerConfig {
  /** Interval in minutes for client-side scheduling (default: 15) */
  intervalMinutes?: number;
  /** Whether to process task escalations (default: true) */
  processTaskEscalations?: boolean;
  /** Whether to process approval reminders (default: true) */
  processApprovalReminders?: boolean;
  /** P2 INTEGRATION FIX: Whether to process task due date reminders (default: true) */
  processTaskDueDateReminders?: boolean;
  /** P2 INTEGRATION FIX: Hours before due date to send reminder (default: 24) */
  dueDateReminderHours?: number;
  /** Maximum tasks to process per run (default: 500) */
  maxTasksPerRun?: number;
  /** Maximum approvals to process per run (default: 200) */
  maxApprovalsPerRun?: number;
}

/**
 * Default configuration
 */
const DEFAULT_CONFIG: Required<ISchedulerConfig> = {
  intervalMinutes: 15,
  processTaskEscalations: true,
  processApprovalReminders: true,
  processTaskDueDateReminders: true,
  dueDateReminderHours: 24,
  maxTasksPerRun: 500,
  maxApprovalsPerRun: 200
};

export class TaskEscalationScheduler {
  private sp: SPFI;
  private context: WebPartContext;
  private siteUrl: string;
  private taskNotificationService: TaskNotificationService;
  private approvalNotificationService: ApprovalNotificationService;
  private config: Required<ISchedulerConfig>;

  // Client-side scheduling
  private intervalId: number | null = null;
  private isRunning: boolean = false;
  private lastRunResult: IEscalationRunResult | null = null;

  constructor(sp: SPFI, context: WebPartContext, config?: ISchedulerConfig) {
    this.sp = sp;
    this.context = context;
    this.siteUrl = context.pageContext.web.absoluteUrl;
    this.config = { ...DEFAULT_CONFIG, ...config };

    // Initialize services
    this.taskNotificationService = new TaskNotificationService(sp, context);
    this.approvalNotificationService = new ApprovalNotificationService(sp, this.siteUrl);
  }

  /**
   * Run escalation processing once (on-demand)
   * This is the main entry point for escalation processing
   */
  public async runEscalationProcessing(): Promise<IEscalationRunResult> {
    if (this.isRunning) {
      logger.warn('TaskEscalationScheduler', 'Escalation processing already running, skipping');
      return {
        success: false,
        runId: this.generateRunId(),
        startTime: new Date(),
        endTime: new Date(),
        durationMs: 0,
        tasksProcessed: 0,
        taskNotificationsSent: 0,
        taskDueDateRemindersSent: 0,
        approvalsProcessed: 0,
        approvalNotificationsSent: 0,
        errors: ['Escalation processing already in progress']
      };
    }

    this.isRunning = true;
    const runId = this.generateRunId();
    const startTime = new Date();
    const errors: string[] = [];

    let taskNotificationsSent = 0;
    let taskDueDateRemindersSent = 0;
    let approvalNotificationsSent = 0;
    let tasksProcessed = 0;
    let approvalsProcessed = 0;

    try {
      logger.info('TaskEscalationScheduler', `Starting escalation run: ${runId}`);

      // Process task escalations
      if (this.config.processTaskEscalations) {
        try {
          taskNotificationsSent = await this.taskNotificationService.processAllTasksForEscalation();
          tasksProcessed = taskNotificationsSent; // Approximation - could be refined
          logger.info('TaskEscalationScheduler', `Task escalations: ${taskNotificationsSent} notifications sent`);
        } catch (taskError) {
          const errorMsg = `Task escalation error: ${taskError instanceof Error ? taskError.message : 'Unknown error'}`;
          errors.push(errorMsg);
          logger.error('TaskEscalationScheduler', errorMsg, taskError);
        }
      }

      // Process approval reminders
      if (this.config.processApprovalReminders) {
        try {
          const approvalResult = await this.processApprovalReminders();
          approvalNotificationsSent = approvalResult.notificationsSent;
          approvalsProcessed = approvalResult.approvalsProcessed;
          logger.info('TaskEscalationScheduler', `Approval reminders: ${approvalNotificationsSent} notifications sent`);
        } catch (approvalError) {
          const errorMsg = `Approval reminder error: ${approvalError instanceof Error ? approvalError.message : 'Unknown error'}`;
          errors.push(errorMsg);
          logger.error('TaskEscalationScheduler', errorMsg, approvalError);
        }
      }

      // P2 INTEGRATION FIX: Process task due date reminders
      if (this.config.processTaskDueDateReminders) {
        try {
          const dueDateResult = await this.processTaskDueDateReminders();
          taskDueDateRemindersSent = dueDateResult.remindersSent;
          logger.info('TaskEscalationScheduler', `Task due date reminders: ${taskDueDateRemindersSent} notifications sent`);
        } catch (dueDateError) {
          const errorMsg = `Task due date reminder error: ${dueDateError instanceof Error ? dueDateError.message : 'Unknown error'}`;
          errors.push(errorMsg);
          logger.error('TaskEscalationScheduler', errorMsg, dueDateError);
        }
      }

      // Log run to SharePoint for audit/monitoring
      await this.logEscalationRun(runId, startTime, errors, taskNotificationsSent, approvalNotificationsSent, taskDueDateRemindersSent);

    } finally {
      this.isRunning = false;
    }

    const endTime = new Date();
    const result: IEscalationRunResult = {
      success: errors.length === 0,
      runId,
      startTime,
      endTime,
      durationMs: endTime.getTime() - startTime.getTime(),
      tasksProcessed,
      taskNotificationsSent,
      taskDueDateRemindersSent,
      approvalsProcessed,
      approvalNotificationsSent,
      errors
    };

    this.lastRunResult = result;

    logger.info(
      'TaskEscalationScheduler',
      `Escalation run ${runId} completed in ${result.durationMs}ms: ` +
      `${taskNotificationsSent} task escalations, ${taskDueDateRemindersSent} due date reminders, ` +
      `${approvalNotificationsSent} approval notifications, ${errors.length} errors`
    );

    return result;
  }

  /**
   * Process approval reminders for pending approvals
   */
  private async processApprovalReminders(): Promise<{
    approvalsProcessed: number;
    notificationsSent: number;
  }> {
    let approvalsProcessed = 0;
    let notificationsSent = 0;

    try {
      // Get pending approvals
      const pendingApprovals = await this.sp.web.lists
        .getByTitle('PM_Approvals')
        .items
        .filter(`Status eq '${ApprovalStatus.Pending}'`)
        .select(
          'Id', 'Title', 'ProcessId', 'ApproverId', 'Status', 'ApprovalLevel',
          'DueDate', 'RequestedDate', 'LastReminderSent',
          'Approver/Id', 'Approver/Title', 'Approver/EMail'
        )
        .expand('Approver')
        .top(this.config.maxApprovalsPerRun)();

      const now = new Date();

      for (const approvalItem of pendingApprovals) {
        approvalsProcessed++;

        // Skip if no due date
        if (!approvalItem.DueDate) {
          continue;
        }

        const dueDate = new Date(approvalItem.DueDate);
        const daysToDue = Math.ceil((dueDate.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));

        // Check if we should send a reminder
        // Send reminder: 1 day before, on due date, and daily after overdue
        const shouldRemind = daysToDue <= 1;

        // Check if we already sent a reminder today
        const lastReminder = approvalItem.LastReminderSent
          ? new Date(approvalItem.LastReminderSent)
          : null;
        const alreadyRemindedToday = lastReminder &&
          lastReminder.toDateString() === now.toDateString();

        if (shouldRemind && !alreadyRemindedToday) {
          try {
            // Build approval object for notification service
            const approval = {
              Id: approvalItem.Id,
              Title: approvalItem.Title,
              ProcessId: approvalItem.ProcessId,
              ProcessTitle: approvalItem.Title,
              ProcessType: 'Unknown',
              ApproverId: approvalItem.ApproverId,
              Approver: {
                Id: approvalItem.Approver?.Id || approvalItem.ApproverId,
                Title: approvalItem.Approver?.Title || ''
              },
              Status: ApprovalStatus.Pending,
              ApprovalLevel: approvalItem.ApprovalLevel || 1,
              DueDate: dueDate,
              RequestedDate: new Date(approvalItem.RequestedDate)
            };

            await this.approvalNotificationService.sendReminderNotification(approval as any);
            notificationsSent++;

            // Update last reminder sent date
            await this.sp.web.lists.getByTitle('PM_Approvals').items
              .getById(approvalItem.Id)
              .update({
                LastReminderSent: now.toISOString()
              });

          } catch (notifyError) {
            logger.warn(
              'TaskEscalationScheduler',
              `Failed to send reminder for approval ${approvalItem.Id}`,
              notifyError
            );
          }
        }
      }
    } catch (error) {
      logger.error('TaskEscalationScheduler', 'Error processing approval reminders', error);
      throw error;
    }

    return { approvalsProcessed, notificationsSent };
  }

  /**
   * P2 INTEGRATION FIX: Process task due date reminders
   * Sends reminder notifications for tasks approaching their due date
   */
  private async processTaskDueDateReminders(): Promise<{
    tasksChecked: number;
    remindersSent: number;
  }> {
    let tasksChecked = 0;
    let remindersSent = 0;

    try {
      const now = new Date();
      const reminderThreshold = new Date(now.getTime() + (this.config.dueDateReminderHours * 60 * 60 * 1000));

      // Get active tasks with due dates approaching
      const activeTasks = await this.sp.web.lists
        .getByTitle('PM_TaskAssignments')
        .items
        .filter(`Status ne '${TaskStatus.Completed}' and Status ne '${TaskStatus.Cancelled}'`)
        .select(
          'Id', 'Title', 'ProcessIDId', 'TaskIDId', 'AssignedToId', 'Status', 'Priority',
          'DueDate', 'ReminderSent', 'LastReminderDate',
          'AssignedTo/Id', 'AssignedTo/Title', 'AssignedTo/EMail'
        )
        .expand('AssignedTo')
        .top(this.config.maxTasksPerRun)();

      for (const taskItem of activeTasks) {
        tasksChecked++;

        // Skip if no due date
        if (!taskItem.DueDate) {
          continue;
        }

        const dueDate = new Date(taskItem.DueDate);
        const hoursToDue = (dueDate.getTime() - now.getTime()) / (1000 * 60 * 60);

        // Check if task is within reminder threshold (approaching due date but not yet overdue)
        // Overdue tasks are handled by the escalation process
        const shouldRemind = hoursToDue > 0 && hoursToDue <= this.config.dueDateReminderHours;

        // Check if we already sent a reminder today
        const lastReminder = taskItem.LastReminderDate
          ? new Date(taskItem.LastReminderDate)
          : null;
        const alreadyRemindedToday = lastReminder &&
          lastReminder.toDateString() === now.toDateString();

        if (shouldRemind && !alreadyRemindedToday) {
          try {
            // Build notification queue item
            const notification: ITaskNotificationQueueItem = {
              TaskAssignmentId: taskItem.Id,
              NotificationType: TaskNotificationType.Reminder,
              ScheduledFor: new Date(),
              Priority: (taskItem.Priority as 'Low' | 'Normal' | 'High' | 'Urgent') || 'Normal',
              Recipients: taskItem.AssignedToId ? [taskItem.AssignedToId] : [],
              Message: `Task "${taskItem.Title}" is due in ${Math.round(hoursToDue)} hours`,
              IsProcessed: false
            };

            // Send via TaskNotificationService
            await this.taskNotificationService.sendNotification(notification);
            remindersSent++;

            // Update last reminder sent date on the task
            await this.sp.web.lists.getByTitle('PM_TaskAssignments').items
              .getById(taskItem.Id)
              .update({
                ReminderSent: true,
                LastReminderDate: now.toISOString()
              });

            logger.debug(
              'TaskEscalationScheduler',
              `Sent due date reminder for task ${taskItem.Id}: ${taskItem.Title}`
            );

          } catch (notifyError) {
            logger.warn(
              'TaskEscalationScheduler',
              `Failed to send due date reminder for task ${taskItem.Id}`,
              notifyError
            );
          }
        }
      }
    } catch (error) {
      logger.error('TaskEscalationScheduler', 'Error processing task due date reminders', error);
      throw error;
    }

    return { tasksChecked, remindersSent };
  }

  /**
   * Log escalation run to SharePoint for monitoring
   */
  private async logEscalationRun(
    runId: string,
    startTime: Date,
    errors: string[],
    taskNotifications: number,
    approvalNotifications: number,
    dueDateReminders: number = 0
  ): Promise<void> {
    try {
      // Try to log to PM_SystemLogs if it exists
      await this.sp.web.lists.getByTitle('PM_SystemLogs').items.add({
        Title: `Escalation Run: ${runId}`,
        LogType: 'EscalationRun',
        LogLevel: errors.length > 0 ? 'Warning' : 'Info',
        Message: `Processed: ${taskNotifications} task escalations, ${dueDateReminders} due date reminders, ${approvalNotifications} approval notifications`,
        Details: JSON.stringify({
          runId,
          startTime: startTime.toISOString(),
          taskNotifications,
          dueDateReminders,
          approvalNotifications,
          errors
        }),
        Created: new Date()
      });
    } catch {
      // PM_SystemLogs might not exist - that's OK, just log locally
      logger.debug('TaskEscalationScheduler', 'Could not log escalation run to SharePoint (PM_SystemLogs list may not exist)');
    }
  }

  /**
   * Start client-side scheduled processing
   * Note: This only works while the webpart is open in the browser
   */
  public startScheduledProcessing(): void {
    if (this.intervalId !== null) {
      logger.warn('TaskEscalationScheduler', 'Scheduled processing already running');
      return;
    }

    const intervalMs = this.config.intervalMinutes * 60 * 1000;

    logger.info('TaskEscalationScheduler', `Starting scheduled processing every ${this.config.intervalMinutes} minutes`);

    // Run immediately
    this.runEscalationProcessing().catch(error => {
      logger.error('TaskEscalationScheduler', 'Initial escalation run failed', error);
    });

    // Schedule recurring runs
    this.intervalId = window.setInterval(() => {
      this.runEscalationProcessing().catch(error => {
        logger.error('TaskEscalationScheduler', 'Scheduled escalation run failed', error);
      });
    }, intervalMs);
  }

  /**
   * Stop client-side scheduled processing
   */
  public stopScheduledProcessing(): void {
    if (this.intervalId !== null) {
      window.clearInterval(this.intervalId);
      this.intervalId = null;
      logger.info('TaskEscalationScheduler', 'Stopped scheduled processing');
    }
  }

  /**
   * Check if scheduler is currently running
   */
  public isSchedulerRunning(): boolean {
    return this.intervalId !== null;
  }

  /**
   * Check if escalation processing is currently in progress
   */
  public isProcessing(): boolean {
    return this.isRunning;
  }

  /**
   * Get the last run result
   */
  public getLastRunResult(): IEscalationRunResult | null {
    return this.lastRunResult;
  }

  /**
   * Get scheduler status
   */
  public getStatus(): {
    isSchedulerRunning: boolean;
    isProcessing: boolean;
    intervalMinutes: number;
    lastRunResult: IEscalationRunResult | null;
  } {
    return {
      isSchedulerRunning: this.intervalId !== null,
      isProcessing: this.isRunning,
      intervalMinutes: this.config.intervalMinutes,
      lastRunResult: this.lastRunResult
    };
  }

  /**
   * Generate unique run ID
   */
  private generateRunId(): string {
    return `ESC-${Date.now()}-${Math.random().toString(36).substring(2, 8)}`;
  }

  /**
   * Cleanup resources
   */
  public dispose(): void {
    this.stopScheduledProcessing();
  }
}

/**
 * Factory function to create TaskEscalationScheduler
 * Use this for easy instantiation in webparts
 */
export function createTaskEscalationScheduler(
  sp: SPFI,
  context: WebPartContext,
  config?: ISchedulerConfig
): TaskEscalationScheduler {
  return new TaskEscalationScheduler(sp, context, config);
}
