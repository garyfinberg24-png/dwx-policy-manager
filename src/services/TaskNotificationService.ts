// @ts-nocheck
// TaskNotificationService - Handles task escalation and notification logic
// Evaluates escalation rules, queues notifications, and sends via Email/Teams

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/items/get-all';
import '@pnp/sp/site-users';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import {
  IJmlTaskEscalationRule,
  IJmlTaskEscalationLog,
  IJmlTaskAssignment,
  ITaskNotificationQueueItem,
  EscalationTrigger,
  TaskNotificationType
} from '../models';
import { TaskStatus, Priority } from '../models/ICommon';
import { logger } from './LoggingService';
import { ToastService } from './ToastService';
import {
  NotificationPreferencesService,
  NotificationEventType,
  NotificationChannel,
  DigestFrequency,
  IResolvedDeliverySettings
} from './workflow/NotificationPreferencesService';
// INTEGRATION FIX: Add retry utilities for notification resilience
import {
  retryWithDLQ,
  notificationDLQ,
  NOTIFICATION_RETRY_OPTIONS,
  IRetryResult
} from '../utils/retryUtils';

/**
 * Notification delivery options
 */
export interface INotificationDeliveryOptions {
  sendEmail?: boolean;
  sendTeams?: boolean;
  sendInApp?: boolean;
}

/**
 * Result of notification delivery
 * INTEGRATION FIX: Enhanced with retry tracking
 */
export interface INotificationDeliveryResult {
  success: boolean;
  emailSent?: boolean;
  teamsSent?: boolean;
  inAppSent?: boolean;
  errors?: string[];
  // INTEGRATION FIX: Retry tracking for failed notifications
  retryAttempts?: number;
  deadLetterItemId?: string;
  queuedForRetry?: boolean;
}

export class TaskNotificationService {
  private sp: SPFI;
  private context?: WebPartContext;
  private graphClient?: MSGraphClientV3;
  private preferencesService: NotificationPreferencesService;
  private rulesListTitle = 'JML_TaskEscalationRules';
  private logListTitle = 'JML_TaskEscalationLog';
  private tasksListTitle = 'JML_TaskAssignments';
  private notificationsListTitle = 'JML_Notifications';

  constructor(sp: SPFI, context?: WebPartContext) {
    this.sp = sp;
    this.context = context;
    // INTEGRATION FIX P1: Initialize preferences service for user preference checking
    this.preferencesService = new NotificationPreferencesService(sp);
  }

  /**
   * Initialize Graph client for email/Teams functionality
   */
  private async getGraphClient(): Promise<MSGraphClientV3 | undefined> {
    if (!this.context) {
      logger.warn('TaskNotificationService', 'WebPartContext not provided, Graph operations unavailable');
      return undefined;
    }
    if (!this.graphClient) {
      this.graphClient = await this.context.msGraphClientFactory.getClient('3');
    }
    return this.graphClient;
  }

  /**
   * Get all active escalation rules
   */
  public async getActiveRules(): Promise<IJmlTaskEscalationRule[]> {
    try {
      logger.info('TaskNotificationService', 'Fetching active escalation rules');

      const items = await this.sp.web.lists
        .getByTitle(this.rulesListTitle)
        .items.filter('IsActive eq 1')
        .select(
          'Id', 'TaskId', 'IsGlobalRule', 'AppliesToCategoryFilter', 'ApplicesToDepartment',
          'EscalationTrigger', 'TriggerValue', 'NotifyRoles', 'NotifySpecificUsers',
          'EscalationLevel', 'AutoReassign', 'ReassignToRole', 'AutoChangeStatus',
          'NewStatus', 'IsActive', 'Created', 'Modified'
        )
        .orderBy('EscalationLevel', true)();

      const rules: IJmlTaskEscalationRule[] = items.map((item: any) => ({
        Id: item.Id,
        Title: item.Title || `Escalation Rule ${item.Id}`,
        TaskId: item.TaskId,
        IsGlobalRule: item.IsGlobalRule || false,
        AppliesToCategoryFilter: item.AppliesToCategoryFilter,
        ApplicesToDepartment: item.ApplicesToDepartment,
        EscalationTrigger: item.EscalationTrigger as EscalationTrigger,
        TriggerValue: item.TriggerValue,
        NotifyRoles: item.NotifyRoles,
        NotifySpecificUsers: item.NotifySpecificUsers,
        EscalationLevel: item.EscalationLevel,
        AutoReassign: item.AutoReassign || false,
        ReassignToRole: item.ReassignToRole,
        AutoChangeStatus: item.AutoChangeStatus || false,
        NewStatus: item.NewStatus,
        IsActive: item.IsActive || false,
        Created: new Date(item.Created),
        Modified: new Date(item.Modified)
      }));

      logger.info('TaskNotificationService', `Retrieved ${rules.length} active rules`);
      return rules;
    } catch (error) {
      logger.error('TaskNotificationService', 'Error fetching escalation rules', error);
      throw error;
    }
  }

  /**
   * Evaluate a task against all escalation rules
   */
  public async evaluateTaskEscalation(task: IJmlTaskAssignment): Promise<ITaskNotificationQueueItem[]> {
    try {
      logger.info('TaskNotificationService', `Evaluating escalation for task ${task.Id}`);

      // Skip completed tasks
      if (task.Status === TaskStatus.Completed) {
        return [];
      }

      const rules = await this.getActiveRules();
      const notifications: ITaskNotificationQueueItem[] = [];
      const now = new Date();

      for (const rule of rules) {
        // Check if rule applies to this task
        if (!this.ruleAppliesToTask(rule, task)) {
          continue;
        }

        // Check if escalation trigger is met
        const triggerMet = this.checkTrigger(rule, task, now);
        if (!triggerMet.isMet) {
          continue;
        }

        // Check if we've already escalated for this rule recently
        const recentlyEscalated = await this.wasRecentlyEscalated(task.Id, rule.Id!);
        if (recentlyEscalated) {
          logger.info('TaskNotificationService', `Task ${task.Id} recently escalated for rule ${rule.Id}, skipping`);
          continue;
        }

        // Build notification
        const recipients = await this.getRecipients(rule, task);
        if (recipients.length === 0) {
          logger.warn('TaskNotificationService', `No recipients found for rule ${rule.Id}`);
          continue;
        }

        const notification: ITaskNotificationQueueItem = {
          TaskAssignmentId: task.Id,
          NotificationType: TaskNotificationType.Escalation,
          ScheduledFor: now,
          Priority: this.determinePriority(rule.EscalationLevel, task.Priority),
          Recipients: recipients,
          Message: this.buildEscalationMessage(rule, task, triggerMet.reason),
          IsProcessed: false
        };

        notifications.push(notification);

        // Log the escalation
        await this.logEscalation(task, rule, recipients, notification.Message);

        // Perform auto-actions if configured
        if (rule.AutoReassign && rule.ReassignToRole) {
          await this.autoReassignTask(task, rule);
        }
        if (rule.AutoChangeStatus && rule.NewStatus) {
          await this.autoChangeStatus(task, rule);
        }
      }

      logger.info('TaskNotificationService', `Generated ${notifications.length} escalation notifications`);
      return notifications;
    } catch (error) {
      logger.error('TaskNotificationService', 'Error evaluating task escalation', error);
      throw error;
    }
  }

  /**
   * Check if rule applies to task
   */
  private ruleAppliesToTask(rule: IJmlTaskEscalationRule, task: IJmlTaskAssignment): boolean {
    // Task-specific rule
    if (rule.TaskId && rule.TaskId === task.TaskIDId) {
      return true;
    }

    // Global rule
    if (rule.IsGlobalRule) {
      // Check category filter
      if (rule.AppliesToCategoryFilter) {
        try {
          const categories = JSON.parse(rule.AppliesToCategoryFilter);
          // Category info is in TaskID lookup after Script 09
          const taskCategory = typeof task.TaskID === 'object' ? (task.TaskID as any).Category : '';
          if (taskCategory && !categories.includes(taskCategory)) {
            return false;
          }
        } catch {
          logger.warn('TaskNotificationService', `Invalid category filter JSON for rule ${rule.Id}`);
        }
      }

      // Check department filter - would need to get from ProcessID
      // For now, skip department filter as it's not directly on task assignment
      if (rule.ApplicesToDepartment) {
        // Department info would be on ProcessID lookup, not directly on task
        // Skip this check for now
      }

      return true;
    }

    return false;
  }

  /**
   * Check if escalation trigger is met
   */
  private checkTrigger(
    rule: IJmlTaskEscalationRule,
    task: IJmlTaskAssignment,
    now: Date
  ): { isMet: boolean; reason?: string } {
    const triggerValue = rule.TriggerValue;

    switch (rule.EscalationTrigger) {
      case EscalationTrigger.OverdueBy:
        if (!task.DueDate) return { isMet: false };
        const dueDate = new Date(task.DueDate);
        const hoursOverdue = (now.getTime() - dueDate.getTime()) / (1000 * 60 * 60);
        if (hoursOverdue >= triggerValue) {
          return {
            isMet: true,
            reason: `Task is ${Math.round(hoursOverdue)} hours overdue (threshold: ${triggerValue}h)`
          };
        }
        return { isMet: false };

      case EscalationTrigger.NotStartedAfter:
        if (!task.AssignedDate) return { isMet: false };
        if (task.Status !== TaskStatus.NotStarted) return { isMet: false };
        const assignedDate = new Date(task.AssignedDate);
        const hoursNotStarted = (now.getTime() - assignedDate.getTime()) / (1000 * 60 * 60);
        if (hoursNotStarted >= triggerValue) {
          return {
            isMet: true,
            reason: `Task not started ${Math.round(hoursNotStarted)} hours after assignment (threshold: ${triggerValue}h)`
          };
        }
        return { isMet: false };

      case EscalationTrigger.StuckInStatus:
        if (!task.Modified) return { isMet: false };
        const lastModified = new Date(task.Modified);
        const hoursStuck = (now.getTime() - lastModified.getTime()) / (1000 * 60 * 60);
        if (hoursStuck >= triggerValue) {
          return {
            isMet: true,
            reason: `No updates in ${Math.round(hoursStuck)} hours (threshold: ${triggerValue}h)`
          };
        }
        return { isMet: false };

      case EscalationTrigger.ApproachingDue:
        if (!task.DueDate) return { isMet: false };
        const dueDateApproaching = new Date(task.DueDate);
        const hoursUntilDue = (dueDateApproaching.getTime() - now.getTime()) / (1000 * 60 * 60);
        if (hoursUntilDue > 0 && hoursUntilDue <= triggerValue) {
          return {
            isMet: true,
            reason: `Task due in ${Math.round(hoursUntilDue)} hours (threshold: ${triggerValue}h)`
          };
        }
        return { isMet: false };

      case EscalationTrigger.HighPriorityOverdue:
        if (task.Priority !== 'High' && task.Priority !== 'Critical') return { isMet: false };
        if (!task.DueDate) return { isMet: false };
        const highPriorityDue = new Date(task.DueDate);
        const isOverdue = now.getTime() > highPriorityDue.getTime();
        if (isOverdue) {
          return {
            isMet: true,
            reason: `High/Critical priority task is overdue`
          };
        }
        return { isMet: false };

      default:
        return { isMet: false };
    }
  }

  /**
   * Check if task was recently escalated for this rule
   */
  private async wasRecentlyEscalated(taskId: number, ruleId: number): Promise<boolean> {
    try {
      // Check escalation log for entries in last 24 hours
      const yesterday = new Date();
      yesterday.setHours(yesterday.getHours() - 24);

      const items = await this.sp.web.lists
        .getByTitle(this.logListTitle)
        .items.filter(
          `TaskAssignmentId eq ${taskId} and EscalationRuleId eq ${ruleId} and NotificationSentDate ge datetime'${yesterday.toISOString()}'`
        )
        .top(1)();

      return items.length > 0;
    } catch (error) {
      logger.warn('TaskNotificationService', 'Error checking escalation log, assuming not escalated', error);
      return false;
    }
  }

  /**
   * Get notification recipients
   */
  private async getRecipients(rule: IJmlTaskEscalationRule, task: IJmlTaskAssignment): Promise<number[]> {
    const recipients: number[] = [];

    // Add specific users
    if (rule.NotifySpecificUsers) {
      try {
        const userIds = JSON.parse(rule.NotifySpecificUsers);
        recipients.push(...userIds);
      } catch {
        logger.warn('TaskNotificationService', `Invalid NotifySpecificUsers JSON for rule ${rule.Id}`);
      }
    }

    // Add role-based recipients
    if (rule.NotifyRoles) {
      try {
        const roles = JSON.parse(rule.NotifyRoles);
        for (const role of roles) {
          const roleRecipients = await this.getRoleRecipients(role, task);
          recipients.push(...roleRecipients);
        }
      } catch {
        logger.warn('TaskNotificationService', `Invalid NotifyRoles JSON for rule ${rule.Id}`);
      }
    }

    // Remove duplicates - convert Set to Array manually for ES5 compatibility
    const uniqueRecipients: number[] = [];
    const seen = new Set<number>();
    for (let i = 0; i < recipients.length; i++) {
      if (!seen.has(recipients[i])) {
        seen.add(recipients[i]);
        uniqueRecipients.push(recipients[i]);
      }
    }
    return uniqueRecipients;
  }

  /**
   * Get recipients for a specific role
   */
  private async getRoleRecipients(role: string, task: IJmlTaskAssignment): Promise<number[]> {
    const recipients: number[] = [];

    switch (role) {
      case 'Manager':
        // Add task assignee's manager
        if (task.AssignedToId) {
          // In real implementation, get manager from user profile
          // For now, add assignee as fallback
          recipients.push(task.AssignedToId);
        }
        break;

      case 'ProcessOwner':
        // Add process owner
        if (task.ProcessIDId) {
          // Get process owner from JML_ProcessInstances
          try {
            const process = await this.sp.web.lists
              .getByTitle('JML_ProcessInstances')
              .items.getById(task.ProcessIDId)
              .select('EmployeeId')();
            if (process.EmployeeId) {
              recipients.push(process.EmployeeId);
            }
          } catch {
            logger.warn('TaskNotificationService', `Could not find process owner for process ${task.ProcessIDId}`);
          }
        }
        break;

      case 'HR':
        // Add HR team members (would need to be configured in settings)
        // For now, add task creator if available
        if (task.AuthorId) {
          recipients.push(task.AuthorId);
        }
        break;

      case 'Assignee':
        // Add current assignee
        if (task.AssignedToId) {
          recipients.push(task.AssignedToId);
        }
        break;

      default:
        logger.warn('TaskNotificationService', `Unknown role: ${role}`);
    }

    return recipients;
  }

  /**
   * Determine notification priority
   */
  private determinePriority(
    escalationLevel: number,
    taskPriority?: string
  ): 'Low' | 'Normal' | 'High' | 'Urgent' {
    if (escalationLevel >= 3 || taskPriority === 'Critical') {
      return 'Urgent';
    } else if (escalationLevel >= 2 || taskPriority === 'High') {
      return 'High';
    } else if (escalationLevel >= 1) {
      return 'Normal';
    }
    return 'Low';
  }

  /**
   * Build escalation message
   */
  private buildEscalationMessage(
    rule: IJmlTaskEscalationRule,
    task: IJmlTaskAssignment,
    reason?: string
  ): string {
    const level = rule.EscalationLevel === 1 ? '⚠️ Level 1' :
      rule.EscalationLevel === 2 ? '⚠️⚠️ Level 2' : '🔴 Level 3';

    let message = `${level} Escalation: ${task.Title}\n\n`;
    message += `Reason: ${reason || 'Escalation threshold met'}\n`;
    message += `Status: ${task.Status}\n`;
    if (task.Priority) {
      message += `Priority: ${task.Priority}\n`;
    }
    if (task.DueDate) {
      message += `Due Date: ${new Date(task.DueDate).toLocaleDateString()}\n`;
    }
    if (task.AssignedTo?.Title) {
      message += `Assigned To: ${task.AssignedTo.Title}\n`;
    }

    return message;
  }

  /**
   * Log escalation to history
   */
  private async logEscalation(
    task: IJmlTaskAssignment,
    rule: IJmlTaskEscalationRule,
    recipients: number[],
    message: string
  ): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.logListTitle)
        .items.add({
          TaskAssignmentId: task.Id,
          EscalationRuleId: rule.Id,
          EscalationTrigger: rule.EscalationTrigger,
          EscalationLevel: rule.EscalationLevel,
          NotifiedUsers: JSON.stringify(recipients),
          NotificationSentDate: new Date().toISOString(),
          NotificationMethod: 'Email',
          WasReassigned: false,
          StatusChanged: false,
          IsResolved: false
        });

      logger.info('TaskNotificationService', `Logged escalation for task ${task.Id}`);
    } catch (error) {
      logger.error('TaskNotificationService', 'Error logging escalation', error);
      // Don't throw - logging failure shouldn't prevent escalation
    }
  }

  /**
   * Auto-reassign task
   */
  private async autoReassignTask(task: IJmlTaskAssignment, rule: IJmlTaskEscalationRule): Promise<void> {
    try {
      logger.info('TaskNotificationService', `Auto-reassigning task ${task.Id}`);

      // In real implementation, resolve role to user ID
      // For now, just log the action
      logger.info('TaskNotificationService', `Would reassign to role: ${rule.ReassignToRole}`);

      // Update escalation log
      const logs = await this.sp.web.lists
        .getByTitle(this.logListTitle)
        .items.filter(`TaskAssignmentId eq ${task.Id} and EscalationRuleId eq ${rule.Id}`)
        .orderBy('Created', false)
        .top(1)();

      if (logs.length > 0) {
        await this.sp.web.lists
          .getByTitle(this.logListTitle)
          .items.getById(logs[0].Id)
          .update({
            WasReassigned: true
          });
      }
    } catch (error) {
      logger.error('TaskNotificationService', 'Error auto-reassigning task', error);
    }
  }

  /**
   * Auto-change task status
   */
  private async autoChangeStatus(task: IJmlTaskAssignment, rule: IJmlTaskEscalationRule): Promise<void> {
    try {
      logger.info('TaskNotificationService', `Auto-changing status for task ${task.Id} to ${rule.NewStatus}`);

      await this.sp.web.lists
        .getByTitle(this.tasksListTitle)
        .items.getById(task.Id)
        .update({
          Status: rule.NewStatus
        });

      // Update escalation log
      const logs = await this.sp.web.lists
        .getByTitle(this.logListTitle)
        .items.filter(`TaskAssignmentId eq ${task.Id} and EscalationRuleId eq ${rule.Id}`)
        .orderBy('Created', false)
        .top(1)();

      if (logs.length > 0) {
        await this.sp.web.lists
          .getByTitle(this.logListTitle)
          .items.getById(logs[0].Id)
          .update({
            StatusChanged: true,
            PreviousStatus: task.Status,
            NewStatus: rule.NewStatus
          });
      }

      logger.info('TaskNotificationService', `Status changed successfully`);
    } catch (error) {
      logger.error('TaskNotificationService', 'Error auto-changing status', error);
    }
  }

  /**
   * Process all tasks for escalation (to be run on schedule)
   */
  public async processAllTasksForEscalation(): Promise<number> {
    try {
      logger.info('TaskNotificationService', 'Processing all tasks for escalation');

      // Get all incomplete tasks
      const tasks = await this.sp.web.lists
        .getByTitle(this.tasksListTitle)
        .items.filter(`Status ne '${TaskStatus.Completed}'`)
        .select(
          'Id', 'Title', 'TaskId', 'Status', 'Priority', 'DueDate', 'AssignedToId',
          'AssignedDate', 'ProcessIDId', 'Category', 'Department', 'Modified',
          'AssignedTo/Id', 'AssignedTo/Title', 'AssignedTo/EMail',
          'Author/Id', 'Author/Title'
        )
        .expand('AssignedTo', 'Author')
        .top(500)();

      let totalNotifications = 0;

      for (const taskItem of tasks) {
        const task: IJmlTaskAssignment = {
          Id: taskItem.Id,
          Title: taskItem.Title,
          Status: taskItem.Status as TaskStatus,
          Priority: taskItem.Priority,
          DueDate: taskItem.DueDate ? new Date(taskItem.DueDate) : undefined,
          AssignedToId: taskItem.AssignedToId,
          AssignedTo: taskItem.AssignedTo ? {
            Id: taskItem.AssignedTo.Id,
            Title: taskItem.AssignedTo.Title,
            EMail: taskItem.AssignedTo.EMail
          } : undefined,
          AssignedDate: taskItem.AssignedDate ? new Date(taskItem.AssignedDate) : undefined,
          ProcessIDId: taskItem.ProcessIDId,
          TaskIDId: taskItem.TaskIDId,
          TaskID: taskItem.TaskID,
          Modified: new Date(taskItem.Modified),
          AuthorId: taskItem.Author?.Id,
          Created: new Date()
        };

        const notifications = await this.evaluateTaskEscalation(task);
        totalNotifications += notifications.length;

        // In real implementation, queue notifications for processing
        // For now, just log
        if (notifications.length > 0) {
          logger.info('TaskNotificationService', `Task ${task.Id} generated ${notifications.length} escalations`);
        }
      }

      logger.info('TaskNotificationService', `Processed ${tasks.length} tasks, generated ${totalNotifications} notifications`);
      return totalNotifications;
    } catch (error) {
      logger.error('TaskNotificationService', 'Error processing tasks for escalation', error);
      throw error;
    }
  }

  /**
   * Send notification with multi-channel delivery (Email, Teams, In-App)
   * INTEGRATION FIX P1: Now respects user notification preferences
   */
  public async sendNotification(
    notification: ITaskNotificationQueueItem,
    options: INotificationDeliveryOptions = { sendEmail: true, sendTeams: true, sendInApp: true }
  ): Promise<INotificationDeliveryResult> {
    const result: INotificationDeliveryResult = {
      success: true,
      errors: []
    };

    try {
      // INTEGRATION FIX P1: Map TaskNotificationType to NotificationEventType for preferences
      const eventType = this.mapToNotificationEventType(notification.NotificationType);
      const priority = this.mapPriorityString(notification.Priority);

      // Process each recipient with their preferences
      for (const recipientId of notification.Recipients) {
        try {
          // Get user email for preference lookup
          const user = await this.sp.web.siteUsers.getById(recipientId)();
          const userEmail = user.Email || '';

          // INTEGRATION FIX P1: Check user preferences before sending
          const deliverySettings = await this.preferencesService.resolveDeliverySettings(
            recipientId,
            userEmail,
            eventType,
            priority
          );

          // Skip if user has disabled this notification type
          if (!deliverySettings.shouldDeliver) {
            logger.info('TaskNotificationService',
              `Skipping notification for user ${recipientId}: ${deliverySettings.reason}`);
            continue;
          }

          // Queue for digest if user prefers digest delivery
          if (deliverySettings.isDigest && deliverySettings.digestFrequency !== DigestFrequency.Immediate) {
            await this.preferencesService.queueForDigest(
              recipientId,
              eventType,
              this.getNotificationTitle(notification.NotificationType),
              notification.Message,
              priority,
              deliverySettings.digestFrequency,
              notification.TaskAssignmentId,
              'TaskAssignment'
            );
            logger.info('TaskNotificationService',
              `Queued notification for user ${recipientId} for ${deliverySettings.digestFrequency} digest`);
            continue;
          }

          // Send via user's preferred channels
          const shouldSendEmail = options.sendEmail &&
            deliverySettings.channels.includes(NotificationChannel.Email);
          const shouldSendTeams = options.sendTeams &&
            deliverySettings.channels.includes(NotificationChannel.Teams);
          const shouldSendInApp = options.sendInApp &&
            deliverySettings.channels.includes(NotificationChannel.InApp);

          // Send in-app notification
          if (shouldSendInApp) {
            try {
              await this.sendInAppNotificationToUser(notification, recipientId);
              result.inAppSent = true;
            } catch (error) {
              result.errors?.push(`In-app notification failed for user ${recipientId}: ${error instanceof Error ? error.message : 'Unknown error'}`);
            }
          }

          // Send email notification (with retry and DLQ)
          if (shouldSendEmail && userEmail) {
            try {
              const emailResult = await this.sendEmailNotification(notification, [userEmail]);
              if (emailResult.success) {
                result.emailSent = true;
                result.retryAttempts = (result.retryAttempts || 0) + emailResult.attempts;
              } else {
                result.errors?.push(`Email notification failed for ${userEmail} after ${emailResult.attempts} retries`);
                result.deadLetterItemId = emailResult.deadLetterItemId;
                result.queuedForRetry = true;
              }
            } catch (error) {
              result.errors?.push(`Email notification failed for ${userEmail}: ${error instanceof Error ? error.message : 'Unknown error'}`);
            }
          }

          // Send Teams notification (with retry and DLQ)
          if (shouldSendTeams && userEmail) {
            try {
              const teamsResult = await this.sendTeamsNotification(notification, [userEmail]);
              if (teamsResult.successCount > 0) {
                result.teamsSent = true;
              }
              if (teamsResult.failedCount > 0) {
                result.errors?.push(`Teams notification partially failed for ${userEmail}: ${teamsResult.failedCount} failed`);
                if (teamsResult.dlqItems.length > 0) {
                  result.deadLetterItemId = teamsResult.dlqItems[0]; // First DLQ item
                  result.queuedForRetry = true;
                }
              }
            } catch (error) {
              result.errors?.push(`Teams notification failed for ${userEmail}: ${error instanceof Error ? error.message : 'Unknown error'}`);
            }
          }

        } catch (userError) {
          result.errors?.push(`Failed to process notification for user ${recipientId}: ${userError instanceof Error ? userError.message : 'Unknown error'}`);
        }
      }

      result.success = (result.errors?.length || 0) === 0;
      return result;
    } catch (error) {
      logger.error('TaskNotificationService', 'Error sending notification', error);
      return {
        success: false,
        errors: [error instanceof Error ? error.message : 'Unknown error']
      };
    }
  }

  /**
   * Map TaskNotificationType to NotificationEventType for preferences
   * INTEGRATION FIX P1: Enables preference checking
   */
  private mapToNotificationEventType(type: TaskNotificationType): NotificationEventType {
    const mapping: Record<TaskNotificationType, NotificationEventType> = {
      [TaskNotificationType.Assigned]: NotificationEventType.TaskAssigned,
      [TaskNotificationType.Reminder]: NotificationEventType.TaskDueSoon,
      [TaskNotificationType.Overdue]: NotificationEventType.TaskOverdue,
      [TaskNotificationType.Escalation]: NotificationEventType.TaskOverdue,
      [TaskNotificationType.Reassigned]: NotificationEventType.TaskAssigned,
      [TaskNotificationType.Mentioned]: NotificationEventType.Reminder,
      [TaskNotificationType.StatusChange]: NotificationEventType.WorkflowStepComplete,
      [TaskNotificationType.Completed]: NotificationEventType.TaskCompleted,
      [TaskNotificationType.SLAWarning]: NotificationEventType.TaskDueSoon,
      [TaskNotificationType.SLABreach]: NotificationEventType.TaskOverdue
    };
    return mapping[type] || NotificationEventType.Reminder;
  }

  /**
   * Map priority string to Priority enum
   * INTEGRATION FIX P1: Type-safe priority handling
   */
  private mapPriorityString(priorityStr: string): Priority {
    switch (priorityStr?.toLowerCase()) {
      case 'critical':
      case 'urgent':
        return Priority.Critical;
      case 'high':
        return Priority.High;
      case 'low':
        return Priority.Low;
      default:
        return Priority.Medium;
    }
  }

  /**
   * Send in-app notification to a single user
   * INTEGRATION FIX P1: Per-user notification delivery
   */
  private async sendInAppNotificationToUser(
    notification: ITaskNotificationQueueItem,
    recipientId: number
  ): Promise<void> {
    await this.sp.web.lists.getByTitle(this.notificationsListTitle).items.add({
      Title: this.getNotificationTitle(notification.NotificationType),
      Message: notification.Message,
      RecipientId: recipientId,
      Type: notification.NotificationType,
      Priority: notification.Priority,
      IsRead: false,
      RelatedItemType: 'TaskAssignment',
      RelatedItemId: notification.TaskAssignmentId
    });
  }

  /**
   * Resolve user IDs to email addresses
   */
  private async resolveRecipientEmails(userIds: number[]): Promise<string[]> {
    const emails: string[] = [];

    for (const userId of userIds) {
      try {
        const user = await this.sp.web.siteUsers.getById(userId)();
        if (user.Email) {
          emails.push(user.Email);
        }
      } catch (error) {
        logger.warn('TaskNotificationService', `Could not resolve email for user ${userId}`, error);
      }
    }

    return emails;
  }

  /**
   * Send in-app notifications to JML_Notifications list
   */
  private async sendInAppNotifications(notification: ITaskNotificationQueueItem): Promise<void> {
    for (const recipientId of notification.Recipients) {
      await this.sp.web.lists.getByTitle(this.notificationsListTitle).items.add({
        Title: this.getNotificationTitle(notification.NotificationType),
        Message: notification.Message,
        RecipientId: recipientId,
        Type: notification.NotificationType,
        Priority: notification.Priority,
        IsRead: false,
        RelatedItemType: 'TaskAssignment',
        RelatedItemId: notification.TaskAssignmentId
      });
    }

    logger.info('TaskNotificationService', `Sent ${notification.Recipients.length} in-app notifications`);
  }

  /**
   * Send email notification via Microsoft Graph
   * INTEGRATION FIX: Enhanced with retry and dead letter queue for resilience
   */
  private async sendEmailNotification(
    notification: ITaskNotificationQueueItem,
    recipientEmails: string[]
  ): Promise<IRetryResult<void>> {
    const graphClient = await this.getGraphClient();
    if (!graphClient) {
      throw new Error('Graph client not available - WebPartContext required');
    }

    const subject = this.getEmailSubject(notification);
    const htmlBody = this.buildEmailHtml(notification);

    const emailMessage = {
      message: {
        subject,
        body: {
          contentType: 'HTML',
          content: htmlBody
        },
        toRecipients: recipientEmails.map(email => ({
          emailAddress: { address: email }
        }))
      },
      saveToSentItems: false
    };

    // INTEGRATION FIX: Wrap API call with retry and DLQ fallback
    const result = await retryWithDLQ<void>(
      async () => {
        await graphClient.api('/me/sendMail').post(emailMessage);
      },
      'email_notification',
      {
        notificationType: notification.NotificationType,
        taskId: notification.TaskAssignmentId,
        recipientCount: recipientEmails.length,
        recipients: recipientEmails,
        subject: subject
      },
      NOTIFICATION_RETRY_OPTIONS,
      notificationDLQ,
      {
        channel: 'email',
        priority: notification.Priority
      }
    );

    if (result.success) {
      logger.info('TaskNotificationService', `Sent email to ${recipientEmails.length} recipients after ${result.attempts} attempt(s)`);
    } else {
      logger.warn('TaskNotificationService', `Email delivery failed after ${result.attempts} attempts, queued to DLQ: ${result.deadLetterItemId}`);
    }

    return result;
  }

  /**
   * Send Teams chat message via Microsoft Graph
   * INTEGRATION FIX: Enhanced with retry and dead letter queue for resilience
   */
  private async sendTeamsNotification(
    notification: ITaskNotificationQueueItem,
    recipientEmails: string[]
  ): Promise<{ successCount: number; failedCount: number; dlqItems: string[] }> {
    const graphClient = await this.getGraphClient();
    if (!graphClient) {
      throw new Error('Graph client not available - WebPartContext required');
    }

    let successCount = 0;
    let failedCount = 0;
    const dlqItems: string[] = [];

    // Send individual chat messages to each recipient with retry
    for (const recipientEmail of recipientEmails) {
      // INTEGRATION FIX: Wrap each recipient's delivery with retry and DLQ
      const result = await retryWithDLQ<void>(
        async () => {
          // Get user ID from email
          const user = await graphClient.api(`/users/${recipientEmail}`).select('id').get();

          // Create or get 1:1 chat with the user
          const chat = {
            chatType: 'oneOnOne',
            members: [
              {
                '@odata.type': '#microsoft.graph.aadUserConversationMember',
                roles: ['owner'],
                'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${this.context?.pageContext?.user?.loginName}`
              },
              {
                '@odata.type': '#microsoft.graph.aadUserConversationMember',
                roles: ['owner'],
                'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${user.id}`
              }
            ]
          };

          const createdChat = await graphClient.api('/chats').post(chat);

          // Send message to the chat
          const message = {
            body: {
              contentType: 'html',
              content: this.buildTeamsMessageHtml(notification)
            }
          };

          await graphClient.api(`/chats/${createdChat.id}/messages`).post(message);
        },
        'teams_notification',
        {
          notificationType: notification.NotificationType,
          taskId: notification.TaskAssignmentId,
          recipient: recipientEmail
        },
        NOTIFICATION_RETRY_OPTIONS,
        notificationDLQ,
        {
          channel: 'teams',
          priority: notification.Priority
        }
      );

      if (result.success) {
        successCount++;
        logger.info('TaskNotificationService', `Sent Teams message to ${recipientEmail} after ${result.attempts} attempt(s)`);
      } else {
        failedCount++;
        if (result.deadLetterItemId) {
          dlqItems.push(result.deadLetterItemId);
        }
        logger.warn('TaskNotificationService', `Teams delivery to ${recipientEmail} failed after ${result.attempts} attempts, queued to DLQ: ${result.deadLetterItemId}`);
      }
    }

    logger.info('TaskNotificationService', `Teams notification summary: ${successCount} sent, ${failedCount} failed`);
    return { successCount, failedCount, dlqItems };
  }

  /**
   * Get notification title based on type
   */
  private getNotificationTitle(type: TaskNotificationType): string {
    switch (type) {
      case TaskNotificationType.Reminder:
        return 'Task Reminder';
      case TaskNotificationType.Overdue:
        return 'Task Overdue';
      case TaskNotificationType.Escalation:
        return 'Task Escalation';
      case TaskNotificationType.Reassigned:
        return 'Task Reassigned';
      case TaskNotificationType.StatusChange:
        return 'Task Status Changed';
      case TaskNotificationType.Completed:
        return 'Task Completed';
      case TaskNotificationType.SLAWarning:
        return 'SLA Warning';
      case TaskNotificationType.SLABreach:
        return 'SLA Breach Alert';
      default:
        return 'Task Notification';
    }
  }

  /**
   * Get email subject based on notification type
   */
  private getEmailSubject(notification: ITaskNotificationQueueItem): string {
    const priorityPrefix = notification.Priority === 'Urgent' ? '🔴 URGENT: ' :
      notification.Priority === 'High' ? '⚠️ ' : '';

    return `${priorityPrefix}${this.getNotificationTitle(notification.NotificationType)} - Task #${notification.TaskAssignmentId}`;
  }

  /**
   * Build HTML email body
   */
  private buildEmailHtml(notification: ITaskNotificationQueueItem): string {
    const priorityColor = notification.Priority === 'Urgent' ? '#d13438' :
      notification.Priority === 'High' ? '#ffaa44' :
      notification.Priority === 'Normal' ? '#0078d4' : '#6b6b6b';

    return `
      <!DOCTYPE html>
      <html>
      <head>
        <style>
          body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 0; background-color: #f5f5f5; }
          .container { max-width: 600px; margin: 20px auto; background: white; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); overflow: hidden; }
          .header { background: ${priorityColor}; color: white; padding: 20px; }
          .header h1 { margin: 0; font-size: 20px; }
          .content { padding: 24px; }
          .message { white-space: pre-line; line-height: 1.6; color: #323130; }
          .priority-badge { display: inline-block; padding: 4px 12px; border-radius: 4px; font-size: 12px; font-weight: 600; color: white; background: ${priorityColor}; }
          .footer { background: #f8f8f8; padding: 16px 24px; border-top: 1px solid #e1e1e1; font-size: 12px; color: #6b6b6b; }
          .button { display: inline-block; padding: 10px 20px; background: #0078d4; color: white; text-decoration: none; border-radius: 4px; margin-top: 16px; }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h1>${this.getNotificationTitle(notification.NotificationType)}</h1>
          </div>
          <div class="content">
            <div style="margin-bottom: 16px;">
              <span class="priority-badge">${notification.Priority} Priority</span>
            </div>
            <div class="message">${notification.Message}</div>
            <a href="#" class="button">View Task</a>
          </div>
          <div class="footer">
            <p>This notification was sent by the JML Task Management System.</p>
            <p>Notification Type: ${notification.NotificationType}</p>
          </div>
        </div>
      </body>
      </html>
    `;
  }

  /**
   * Build Teams message HTML
   */
  private buildTeamsMessageHtml(notification: ITaskNotificationQueueItem): string {
    const priorityEmoji = notification.Priority === 'Urgent' ? '🔴' :
      notification.Priority === 'High' ? '⚠️' :
      notification.Priority === 'Normal' ? '📋' : '📌';

    return `
      <div>
        <h3>${priorityEmoji} ${this.getNotificationTitle(notification.NotificationType)}</h3>
        <p><strong>Priority:</strong> ${notification.Priority}</p>
        <p>${notification.Message.replace(/\n/g, '<br>')}</p>
        <p><em>Task #${notification.TaskAssignmentId}</em></p>
      </div>
    `;
  }

  /**
   * Send task assignment notification to assignee
   * INTEGRATION FIX: Critical for TaskActionHandler integration
   */
  public async sendTaskAssignmentNotification(
    task: IJmlTaskAssignment,
    processTitle?: string,
    options?: INotificationDeliveryOptions
  ): Promise<INotificationDeliveryResult> {
    if (!task.AssignedToId) {
      logger.warn('TaskNotificationService', `Cannot send assignment notification - task ${task.Id} has no assignee`);
      return { success: false, errors: ['No assignee specified'] };
    }

    const notification: ITaskNotificationQueueItem = {
      TaskAssignmentId: task.Id!,
      NotificationType: TaskNotificationType.Assigned,
      ScheduledFor: new Date(),
      Priority: task.Priority === 'Critical' ? 'Urgent' : task.Priority === 'High' ? 'High' : 'Normal',
      Recipients: [task.AssignedToId],
      Message: this.buildAssignmentMessage(task, processTitle),
      IsProcessed: false
    };

    logger.info('TaskNotificationService', `Sending assignment notification for task ${task.Id} to user ${task.AssignedToId}`);
    return this.sendNotification(notification, options);
  }

  /**
   * Send task completion notification
   * INTEGRATION FIX: Notify relevant stakeholders when task is completed
   */
  public async sendTaskCompletionNotification(
    task: IJmlTaskAssignment,
    completedByUserId: number,
    notifyAssignee: boolean = false,
    additionalRecipients?: number[],
    options?: INotificationDeliveryOptions
  ): Promise<INotificationDeliveryResult> {
    const recipients: number[] = [];

    // Optionally notify the assignee (if someone else completed it)
    if (notifyAssignee && task.AssignedToId && task.AssignedToId !== completedByUserId) {
      recipients.push(task.AssignedToId);
    }

    // Add any additional recipients (manager, process owner, etc.)
    if (additionalRecipients && additionalRecipients.length > 0) {
      recipients.push(...additionalRecipients.filter(r => !recipients.includes(r)));
    }

    if (recipients.length === 0) {
      logger.info('TaskNotificationService', `No recipients for completion notification for task ${task.Id}`);
      return { success: true }; // No error, just no recipients
    }

    const notification: ITaskNotificationQueueItem = {
      TaskAssignmentId: task.Id!,
      NotificationType: TaskNotificationType.Completed,
      ScheduledFor: new Date(),
      Priority: 'Normal',
      Recipients: recipients,
      Message: this.buildCompletionMessage(task, completedByUserId),
      IsProcessed: false
    };

    logger.info('TaskNotificationService', `Sending completion notification for task ${task.Id} to ${recipients.length} recipients`);
    return this.sendNotification(notification, options);
  }

  /**
   * Send reminder notification for upcoming task
   */
  public async sendTaskReminder(
    task: IJmlTaskAssignment,
    hoursUntilDue: number,
    options?: INotificationDeliveryOptions
  ): Promise<INotificationDeliveryResult> {
    const notification: ITaskNotificationQueueItem = {
      TaskAssignmentId: task.Id!,
      NotificationType: TaskNotificationType.Reminder,
      ScheduledFor: new Date(),
      Priority: task.Priority === 'Critical' ? 'Urgent' : task.Priority === 'High' ? 'High' : 'Normal',
      Recipients: task.AssignedToId ? [task.AssignedToId] : [],
      Message: this.buildReminderMessage(task, hoursUntilDue),
      IsProcessed: false
    };

    return this.sendNotification(notification, options);
  }

  /**
   * Send overdue notification for task
   */
  public async sendOverdueNotification(
    task: IJmlTaskAssignment,
    hoursOverdue: number,
    options?: INotificationDeliveryOptions
  ): Promise<INotificationDeliveryResult> {
    const notification: ITaskNotificationQueueItem = {
      TaskAssignmentId: task.Id!,
      NotificationType: TaskNotificationType.Overdue,
      ScheduledFor: new Date(),
      Priority: hoursOverdue > 48 ? 'Urgent' : hoursOverdue > 24 ? 'High' : 'Normal',
      Recipients: task.AssignedToId ? [task.AssignedToId] : [],
      Message: this.buildOverdueMessage(task, hoursOverdue),
      IsProcessed: false
    };

    return this.sendNotification(notification, options);
  }

  /**
   * Send SLA warning notification
   */
  public async sendSLAWarning(
    task: IJmlTaskAssignment,
    hoursRemaining: number,
    options?: INotificationDeliveryOptions
  ): Promise<INotificationDeliveryResult> {
    const notification: ITaskNotificationQueueItem = {
      TaskAssignmentId: task.Id!,
      NotificationType: TaskNotificationType.SLAWarning,
      ScheduledFor: new Date(),
      Priority: 'High',
      Recipients: task.AssignedToId ? [task.AssignedToId] : [],
      Message: this.buildSLAWarningMessage(task, hoursRemaining),
      IsProcessed: false
    };

    return this.sendNotification(notification, options);
  }

  /**
   * Send SLA breach notification
   */
  public async sendSLABreach(
    task: IJmlTaskAssignment,
    hoursOverSLA: number,
    options?: INotificationDeliveryOptions
  ): Promise<INotificationDeliveryResult> {
    const notification: ITaskNotificationQueueItem = {
      TaskAssignmentId: task.Id!,
      NotificationType: TaskNotificationType.SLABreach,
      ScheduledFor: new Date(),
      Priority: 'Urgent',
      Recipients: task.AssignedToId ? [task.AssignedToId] : [],
      Message: this.buildSLABreachMessage(task, hoursOverSLA),
      IsProcessed: false
    };

    return this.sendNotification(notification, options);
  }

  /**
   * Build task assignment message
   * INTEGRATION FIX: Email notification for new task assignments
   */
  private buildAssignmentMessage(task: IJmlTaskAssignment, processTitle?: string): string {
    const dueDateStr = task.DueDate
      ? new Date(task.DueDate).toLocaleDateString()
      : 'Not set';

    return `📋 New Task Assigned: "${task.Title}"

You have been assigned a new task${processTitle ? ` for ${processTitle}` : ''}.

Task Details:
- Status: ${task.Status || 'Not Started'}
- Priority: ${task.Priority || 'Medium'}
- Due Date: ${dueDateStr}
${task.Notes ? `\nDescription: ${task.Notes}` : ''}

Please review and complete this task by the due date.`;
  }

  /**
   * Build task completion message
   * INTEGRATION FIX: Email notification when task is completed
   */
  private buildCompletionMessage(task: IJmlTaskAssignment, completedByUserId: number): string {
    const completedDate = new Date().toLocaleDateString();

    return `✅ Task Completed: "${task.Title}"

A task has been marked as completed.

Task Details:
- Completed By: User ID ${completedByUserId}
- Completed On: ${completedDate}
- Priority: ${task.Priority || 'Medium'}

${task.Notes ? `Notes: ${task.Notes}` : ''}`;
  }

  /**
   * Build reminder message
   */
  private buildReminderMessage(task: IJmlTaskAssignment, hoursUntilDue: number): string {
    const daysUntil = Math.floor(hoursUntilDue / 24);
    const timeStr = daysUntil >= 1 ? `${daysUntil} day(s)` : `${Math.round(hoursUntilDue)} hours`;

    return `📅 Task Reminder: "${task.Title}"

Your task is due in ${timeStr}.

Status: ${task.Status}
Priority: ${task.Priority}
Due Date: ${task.DueDate ? new Date(task.DueDate).toLocaleDateString() : 'Not set'}

Please complete this task before the due date.`;
  }

  /**
   * Build overdue message
   */
  private buildOverdueMessage(task: IJmlTaskAssignment, hoursOverdue: number): string {
    const daysOverdue = Math.floor(hoursOverdue / 24);
    const timeStr = daysOverdue >= 1 ? `${daysOverdue} day(s)` : `${Math.round(hoursOverdue)} hours`;

    return `⚠️ Overdue Task Alert: "${task.Title}"

This task is now ${timeStr} overdue.

Status: ${task.Status}
Priority: ${task.Priority}
Due Date: ${task.DueDate ? new Date(task.DueDate).toLocaleDateString() : 'Not set'}

Please complete this task immediately or update its status.`;
  }

  /**
   * Build SLA warning message
   */
  private buildSLAWarningMessage(task: IJmlTaskAssignment, hoursRemaining: number): string {
    return `⏰ SLA Warning: "${task.Title}"

This task is approaching its SLA deadline.
Time remaining: ${Math.round(hoursRemaining)} hours

Status: ${task.Status}
Priority: ${task.Priority}
SLA Target: ${task.SLAHours || 'N/A'} hours

Please take action to avoid an SLA breach.`;
  }

  /**
   * Build SLA breach message
   */
  private buildSLABreachMessage(task: IJmlTaskAssignment, hoursOverSLA: number): string {
    return `🔴 SLA BREACH: "${task.Title}"

This task has exceeded its SLA by ${Math.round(hoursOverSLA)} hours.

Status: ${task.Status}
Priority: ${task.Priority}
SLA Target: ${task.SLAHours || 'N/A'} hours

Immediate action required. This breach has been logged.`;
  }
}
