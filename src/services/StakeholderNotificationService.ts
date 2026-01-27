// @ts-nocheck
/**
 * StakeholderNotificationService
 * P3 INTEGRATION FIX: Centralized stakeholder notification management
 *
 * Consolidates stakeholder notification logic that was previously scattered
 * across ProcessOrchestrationService, WorkflowEngineService, and ApprovalService.
 *
 * Features:
 * - Identifies all relevant stakeholders for a process
 * - Sends notifications for key process events
 * - Respects user notification preferences (via NotificationPreferencesService)
 * - Supports multiple notification channels (email, Teams, in-app)
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users';

import { logger } from './LoggingService';
import {
  NotificationPreferencesService,
  NotificationEventType,
  NotificationChannel,
  DigestFrequency
} from './workflow/NotificationPreferencesService';
import { Priority } from '../models/ICommon';
// INTEGRATION FIX P5: IT Admin notification confirmation tracking
import { NotificationConfirmationStatus } from '../models/IJmlNotification';
import { v4 as uuidv4 } from 'uuid';

/**
 * Stakeholder roles in a JML process
 */
export enum StakeholderRole {
  Employee = 'Employee',
  Manager = 'Manager',
  HRAdmin = 'HRAdmin',
  ITAdmin = 'ITAdmin',
  Buddy = 'Buddy',
  Recruiter = 'Recruiter',
  ProcessOwner = 'ProcessOwner',
  Approver = 'Approver',
  TaskAssignee = 'TaskAssignee'
}

/**
 * Process event types for stakeholder notifications
 */
export enum ProcessEventType {
  ProcessStarted = 'ProcessStarted',
  ProcessCompleted = 'ProcessCompleted',
  ProcessFailed = 'ProcessFailed',
  ProcessCancelled = 'ProcessCancelled',
  MilestoneReached = 'MilestoneReached',
  ApprovalRequired = 'ApprovalRequired',
  ApprovalCompleted = 'ApprovalCompleted',
  ApprovalRejected = 'ApprovalRejected',
  TaskAssigned = 'TaskAssigned',
  TaskCompleted = 'TaskCompleted',
  TaskOverdue = 'TaskOverdue',
  WorkflowStepCompleted = 'WorkflowStepCompleted',
  EscalationTriggered = 'EscalationTriggered'
}

/**
 * Stakeholder information
 */
export interface IStakeholder {
  id: number;
  email: string;
  name: string;
  role: StakeholderRole;
  isPrimary: boolean;
}

/**
 * Process context for stakeholder identification
 */
export interface IProcessContext {
  processId: number;
  processType: string;
  employeeId?: number;
  employeeName?: string;
  employeeEmail?: string;
  managerId?: number;
  managerEmail?: string;
  hrAdminId?: number;
  itAdminId?: number;
  buddyId?: number;
  recruiterId?: number;
  processOwnerId?: number;
  additionalStakeholderIds?: number[];
}

/**
 * Notification event data
 */
export interface INotificationEventData {
  eventType: ProcessEventType;
  processContext: IProcessContext;
  eventTitle: string;
  eventMessage: string;
  priority: Priority;
  metadata?: Record<string, unknown>;
  targetRoles?: StakeholderRole[];
  excludeRoles?: StakeholderRole[];
  actionUrl?: string;
}

/**
 * Notification result
 */
export interface IStakeholderNotificationResult {
  success: boolean;
  notificationsSent: number;
  stakeholdersNotified: string[];
  errors: string[];
}

/**
 * Role-to-event mapping for default notifications
 */
const DEFAULT_ROLE_EVENT_MAP: Record<ProcessEventType, StakeholderRole[]> = {
  [ProcessEventType.ProcessStarted]: [
    StakeholderRole.Manager,
    StakeholderRole.HRAdmin,
    StakeholderRole.ITAdmin
  ],
  [ProcessEventType.ProcessCompleted]: [
    StakeholderRole.Employee,
    StakeholderRole.Manager,
    StakeholderRole.HRAdmin
  ],
  [ProcessEventType.ProcessFailed]: [
    StakeholderRole.Manager,
    StakeholderRole.HRAdmin,
    StakeholderRole.ProcessOwner
  ],
  [ProcessEventType.ProcessCancelled]: [
    StakeholderRole.Employee,
    StakeholderRole.Manager,
    StakeholderRole.HRAdmin
  ],
  [ProcessEventType.MilestoneReached]: [
    StakeholderRole.Manager,
    StakeholderRole.HRAdmin
  ],
  [ProcessEventType.ApprovalRequired]: [
    StakeholderRole.Approver
  ],
  [ProcessEventType.ApprovalCompleted]: [
    StakeholderRole.ProcessOwner,
    StakeholderRole.Manager
  ],
  [ProcessEventType.ApprovalRejected]: [
    StakeholderRole.ProcessOwner,
    StakeholderRole.Manager,
    StakeholderRole.HRAdmin
  ],
  [ProcessEventType.TaskAssigned]: [
    StakeholderRole.TaskAssignee
  ],
  [ProcessEventType.TaskCompleted]: [
    StakeholderRole.Manager
  ],
  [ProcessEventType.TaskOverdue]: [
    StakeholderRole.TaskAssignee,
    StakeholderRole.Manager
  ],
  [ProcessEventType.WorkflowStepCompleted]: [
    StakeholderRole.ProcessOwner
  ],
  [ProcessEventType.EscalationTriggered]: [
    StakeholderRole.Manager,
    StakeholderRole.HRAdmin
  ]
};

export class StakeholderNotificationService {
  private sp: SPFI;
  private siteUrl: string;
  private preferencesService: NotificationPreferencesService;

  constructor(sp: SPFI, siteUrl: string) {
    this.sp = sp;
    this.siteUrl = siteUrl;
    this.preferencesService = new NotificationPreferencesService(sp);
  }

  /**
   * Notify stakeholders of a process event
   * Main entry point for stakeholder notifications
   */
  public async notifyStakeholders(
    eventData: INotificationEventData
  ): Promise<IStakeholderNotificationResult> {
    const result: IStakeholderNotificationResult = {
      success: true,
      notificationsSent: 0,
      stakeholdersNotified: [],
      errors: []
    };

    try {
      // Get stakeholders for this event
      const stakeholders = await this.getStakeholdersForEvent(eventData);

      if (stakeholders.length === 0) {
        logger.warn(
          'StakeholderNotificationService',
          `No stakeholders found for event ${eventData.eventType} on process ${eventData.processContext.processId}`
        );
        return result;
      }

      // Notify each stakeholder
      for (const stakeholder of stakeholders) {
        try {
          const sent = await this.notifyStakeholder(stakeholder, eventData);
          if (sent) {
            result.notificationsSent++;
            result.stakeholdersNotified.push(`${stakeholder.name} (${stakeholder.role})`);
          }
        } catch (error) {
          const errorMsg = `Failed to notify ${stakeholder.name}: ${error instanceof Error ? error.message : 'Unknown error'}`;
          result.errors.push(errorMsg);
          logger.warn('StakeholderNotificationService', errorMsg, error);
        }
      }

      result.success = result.errors.length === 0;

      logger.info(
        'StakeholderNotificationService',
        `Notified ${result.notificationsSent} stakeholders for ${eventData.eventType} on process ${eventData.processContext.processId}`
      );

    } catch (error) {
      result.success = false;
      result.errors.push(`Error notifying stakeholders: ${error instanceof Error ? error.message : 'Unknown error'}`);
      logger.error('StakeholderNotificationService', 'Error notifying stakeholders', error);
    }

    return result;
  }

  /**
   * Get stakeholders for a specific event
   */
  private async getStakeholdersForEvent(
    eventData: INotificationEventData
  ): Promise<IStakeholder[]> {
    const stakeholders: IStakeholder[] = [];
    const context = eventData.processContext;

    // Determine which roles to notify
    const targetRoles = eventData.targetRoles || DEFAULT_ROLE_EVENT_MAP[eventData.eventType] || [];
    const excludeRoles = eventData.excludeRoles || [];

    // Filter roles
    const rolesToNotify = targetRoles.filter(role => !excludeRoles.includes(role));

    // Build stakeholder list from context
    for (const role of rolesToNotify) {
      const stakeholder = await this.getStakeholderForRole(role, context);
      if (stakeholder) {
        stakeholders.push(stakeholder);
      }
    }

    // Add any additional stakeholders from context
    if (context.additionalStakeholderIds) {
      for (const userId of context.additionalStakeholderIds) {
        try {
          const user = await this.sp.web.siteUsers.getById(userId)();
          if (user && user.Email) {
            // Check if not already in list
            if (!stakeholders.find(s => s.id === userId)) {
              stakeholders.push({
                id: userId,
                email: user.Email,
                name: user.Title || 'Unknown',
                role: StakeholderRole.ProcessOwner,
                isPrimary: false
              });
            }
          }
        } catch {
          logger.debug('StakeholderNotificationService', `Could not resolve additional stakeholder ${userId}`);
        }
      }
    }

    return stakeholders;
  }

  /**
   * Get stakeholder details for a specific role
   */
  private async getStakeholderForRole(
    role: StakeholderRole,
    context: IProcessContext
  ): Promise<IStakeholder | null> {
    let userId: number | undefined;
    let userEmail: string | undefined;
    let isPrimary = false;

    switch (role) {
      case StakeholderRole.Employee:
        userId = context.employeeId;
        userEmail = context.employeeEmail;
        isPrimary = true;
        break;
      case StakeholderRole.Manager:
        userId = context.managerId;
        userEmail = context.managerEmail;
        isPrimary = true;
        break;
      case StakeholderRole.HRAdmin:
        userId = context.hrAdminId;
        break;
      case StakeholderRole.ITAdmin:
        userId = context.itAdminId;
        break;
      case StakeholderRole.Buddy:
        userId = context.buddyId;
        break;
      case StakeholderRole.Recruiter:
        userId = context.recruiterId;
        break;
      case StakeholderRole.ProcessOwner:
        userId = context.processOwnerId;
        break;
      default:
        return null;
    }

    if (!userId && !userEmail) {
      return null;
    }

    // Resolve user details if we have ID but no email
    if (userId && !userEmail) {
      try {
        const user = await this.sp.web.siteUsers.getById(userId)();
        userEmail = user.Email;
        return {
          id: userId,
          email: userEmail || '',
          name: user.Title || 'Unknown',
          role,
          isPrimary
        };
      } catch {
        logger.debug('StakeholderNotificationService', `Could not resolve user ${userId} for role ${role}`);
        return null;
      }
    }

    // If we have email but no ID, try to resolve
    if (userEmail && !userId) {
      try {
        const user = await this.sp.web.ensureUser(userEmail);
        return {
          id: user.data.Id,
          email: userEmail,
          name: user.data.Title || userEmail,
          role,
          isPrimary
        };
      } catch {
        // Return with just email
        return {
          id: 0,
          email: userEmail,
          name: userEmail,
          role,
          isPrimary
        };
      }
    }

    return null;
  }

  /**
   * Notify a single stakeholder
   * Checks preferences and sends via appropriate channels
   */
  private async notifyStakeholder(
    stakeholder: IStakeholder,
    eventData: INotificationEventData
  ): Promise<boolean> {
    // Map process event to notification event type
    const notificationEventType = this.mapToNotificationEventType(eventData.eventType);

    // Check user preferences
    const deliverySettings = await this.preferencesService.resolveDeliverySettings(
      stakeholder.id,
      stakeholder.email,
      notificationEventType,
      eventData.priority
    );

    // Skip if user has disabled this notification type
    if (!deliverySettings.shouldDeliver) {
      logger.debug(
        'StakeholderNotificationService',
        `Notification disabled for ${stakeholder.email} event ${eventData.eventType}`
      );
      return false;
    }

    // If digest is preferred, queue for digest
    if (deliverySettings.isDigest && deliverySettings.digestFrequency !== DigestFrequency.Immediate) {
      await this.preferencesService.queueForDigest(
        stakeholder.id,
        notificationEventType,
        eventData.eventTitle,
        eventData.eventMessage,
        eventData.priority,
        deliverySettings.digestFrequency,
        eventData.processContext.processId,
        'Process'
      );
      return true;
    }

    // Send immediately via preferred channels
    let sent = false;
    for (const channel of deliverySettings.channels) {
      try {
        switch (channel) {
          case NotificationChannel.Email:
            await this.sendEmailNotification(stakeholder, eventData);
            sent = true;
            break;
          case NotificationChannel.Teams:
            await this.sendTeamsNotification(stakeholder, eventData);
            sent = true;
            break;
          case NotificationChannel.InApp:
            await this.sendInAppNotification(stakeholder, eventData);
            sent = true;
            break;
        }
      } catch (error) {
        logger.warn(
          'StakeholderNotificationService',
          `Failed to send via ${channel} to ${stakeholder.email}`,
          error
        );
      }
    }

    return sent;
  }

  /**
   * Map process event type to notification event type
   */
  private mapToNotificationEventType(eventType: ProcessEventType): NotificationEventType {
    const mapping: Record<ProcessEventType, NotificationEventType> = {
      [ProcessEventType.ProcessStarted]: NotificationEventType.ProcessStarted,
      [ProcessEventType.ProcessCompleted]: NotificationEventType.ProcessCompleted,
      [ProcessEventType.ProcessFailed]: NotificationEventType.WorkflowError,
      [ProcessEventType.ProcessCancelled]: NotificationEventType.ProcessBlocked,
      [ProcessEventType.MilestoneReached]: NotificationEventType.WorkflowStepComplete,
      [ProcessEventType.ApprovalRequired]: NotificationEventType.ApprovalRequired,
      [ProcessEventType.ApprovalCompleted]: NotificationEventType.ApprovalCompleted,
      [ProcessEventType.ApprovalRejected]: NotificationEventType.ApprovalRejected,
      [ProcessEventType.TaskAssigned]: NotificationEventType.TaskAssigned,
      [ProcessEventType.TaskCompleted]: NotificationEventType.TaskCompleted,
      [ProcessEventType.TaskOverdue]: NotificationEventType.TaskOverdue,
      [ProcessEventType.WorkflowStepCompleted]: NotificationEventType.WorkflowStepComplete,
      [ProcessEventType.EscalationTriggered]: NotificationEventType.ApprovalEscalated
    };

    return mapping[eventType] || NotificationEventType.Reminder;
  }

  /**
   * Send email notification
   */
  private async sendEmailNotification(
    stakeholder: IStakeholder,
    eventData: INotificationEventData
  ): Promise<void> {
    // Queue email via EmailQueueService pattern
    await this.sp.web.lists.getByTitle('JML_EmailQueue').items.add({
      Title: eventData.eventTitle,
      RecipientEmail: stakeholder.email,
      RecipientName: stakeholder.name,
      Subject: eventData.eventTitle,
      Body: this.formatEmailBody(eventData),
      Priority: eventData.priority,
      Status: 'Pending',
      ProcessId: eventData.processContext.processId?.toString(),
      NotificationType: eventData.eventType,
      Created: new Date()
    });
  }

  /**
   * Send Teams notification (via Power Automate webhook or adaptive card)
   */
  private async sendTeamsNotification(
    stakeholder: IStakeholder,
    eventData: INotificationEventData
  ): Promise<void> {
    // Queue for Teams delivery (typically handled by Power Automate)
    await this.sp.web.lists.getByTitle('JML_Notifications').items.add({
      Title: eventData.eventTitle,
      NotificationType: eventData.eventType,
      MessageBody: eventData.eventMessage,
      Priority: eventData.priority,
      RecipientId: stakeholder.id > 0 ? stakeholder.id : undefined,
      RecipientEmail: stakeholder.email,
      ProcessId: eventData.processContext.processId?.toString(),
      Status: 'Pending',
      DeliveryChannel: 'Teams',
      ActionUrl: eventData.actionUrl || '',
      Created: new Date()
    });
  }

  /**
   * Send in-app notification
   */
  private async sendInAppNotification(
    stakeholder: IStakeholder,
    eventData: INotificationEventData
  ): Promise<void> {
    await this.sp.web.lists.getByTitle('JML_Notifications').items.add({
      Title: eventData.eventTitle,
      NotificationType: eventData.eventType,
      MessageBody: eventData.eventMessage,
      Priority: eventData.priority,
      RecipientId: stakeholder.id > 0 ? stakeholder.id : undefined,
      RecipientEmail: stakeholder.email,
      ProcessId: eventData.processContext.processId?.toString(),
      Status: 'Unread',
      DeliveryChannel: 'InApp',
      ActionUrl: eventData.actionUrl || '',
      IsRead: false,
      Created: new Date()
    });
  }

  /**
   * Format email body with event details
   */
  private formatEmailBody(eventData: INotificationEventData): string {
    const context = eventData.processContext;

    let body = `<div style="font-family: 'Segoe UI', sans-serif; padding: 20px;">`;
    body += `<h2 style="color: #0078d4;">${eventData.eventTitle}</h2>`;
    body += `<p>${eventData.eventMessage}</p>`;

    body += `<div style="background-color: #f3f2f1; padding: 16px; border-radius: 4px; margin: 16px 0;">`;
    body += `<p><strong>Process Type:</strong> ${context.processType}</p>`;
    if (context.employeeName) {
      body += `<p><strong>Employee:</strong> ${context.employeeName}</p>`;
    }
    body += `<p><strong>Process ID:</strong> ${context.processId}</p>`;
    body += `</div>`;

    if (eventData.actionUrl) {
      body += `<p><a href="${eventData.actionUrl}" style="color: #0078d4;">View Details</a></p>`;
    }

    body += `<hr style="border: none; border-top: 1px solid #edebe9; margin: 20px 0;" />`;
    body += `<p style="color: #605e5c; font-size: 12px;">This is an automated notification from the JML Employee Lifecycle Management System.</p>`;
    body += `</div>`;

    return body;
  }

  /**
   * Convenience method: Notify process started
   */
  public async notifyProcessStarted(
    context: IProcessContext,
    additionalMessage?: string
  ): Promise<IStakeholderNotificationResult> {
    return this.notifyStakeholders({
      eventType: ProcessEventType.ProcessStarted,
      processContext: context,
      eventTitle: `New ${context.processType} Process Started`,
      eventMessage: additionalMessage || `A new ${context.processType} process has been started for ${context.employeeName || 'an employee'}.`,
      priority: Priority.Medium,
      actionUrl: `${this.siteUrl}/SitePages/ProcessDetails.aspx?processId=${context.processId}`
    });
  }

  /**
   * Convenience method: Notify process completed
   */
  public async notifyProcessCompleted(
    context: IProcessContext,
    additionalMessage?: string
  ): Promise<IStakeholderNotificationResult> {
    return this.notifyStakeholders({
      eventType: ProcessEventType.ProcessCompleted,
      processContext: context,
      eventTitle: `${context.processType} Process Completed`,
      eventMessage: additionalMessage || `The ${context.processType} process for ${context.employeeName || 'the employee'} has been successfully completed.`,
      priority: Priority.Medium,
      actionUrl: `${this.siteUrl}/SitePages/ProcessDetails.aspx?processId=${context.processId}`
    });
  }

  /**
   * Convenience method: Notify process failed
   */
  public async notifyProcessFailed(
    context: IProcessContext,
    errorMessage: string
  ): Promise<IStakeholderNotificationResult> {
    return this.notifyStakeholders({
      eventType: ProcessEventType.ProcessFailed,
      processContext: context,
      eventTitle: `${context.processType} Process Failed`,
      eventMessage: `The ${context.processType} process for ${context.employeeName || 'the employee'} has encountered an error: ${errorMessage}`,
      priority: Priority.High,
      actionUrl: `${this.siteUrl}/SitePages/ProcessDetails.aspx?processId=${context.processId}`
    });
  }

  /**
   * Convenience method: Notify milestone reached
   */
  public async notifyMilestoneReached(
    context: IProcessContext,
    milestoneName: string,
    milestoneDescription?: string
  ): Promise<IStakeholderNotificationResult> {
    return this.notifyStakeholders({
      eventType: ProcessEventType.MilestoneReached,
      processContext: context,
      eventTitle: `Milestone Reached: ${milestoneName}`,
      eventMessage: milestoneDescription || `The milestone "${milestoneName}" has been reached in the ${context.processType} process for ${context.employeeName || 'the employee'}.`,
      priority: Priority.Low,
      metadata: { milestoneName },
      actionUrl: `${this.siteUrl}/SitePages/ProcessDetails.aspx?processId=${context.processId}`
    });
  }

  /**
   * Convenience method: Notify approval required
   */
  public async notifyApprovalRequired(
    context: IProcessContext,
    approverIds: number[],
    approvalTitle: string,
    dueDate?: Date
  ): Promise<IStakeholderNotificationResult> {
    const extendedContext: IProcessContext = {
      ...context,
      additionalStakeholderIds: approverIds
    };

    let message = `Your approval is required for: ${approvalTitle}`;
    if (dueDate) {
      message += ` (Due: ${dueDate.toLocaleDateString()})`;
    }

    return this.notifyStakeholders({
      eventType: ProcessEventType.ApprovalRequired,
      processContext: extendedContext,
      eventTitle: `Approval Required: ${approvalTitle}`,
      eventMessage: message,
      priority: Priority.High,
      targetRoles: [StakeholderRole.Approver],
      actionUrl: `${this.siteUrl}/SitePages/ApprovalCenter.aspx?processId=${context.processId}`
    });
  }

  /**
   * Convenience method: Notify escalation triggered
   */
  public async notifyEscalationTriggered(
    context: IProcessContext,
    escalationLevel: number,
    reason: string
  ): Promise<IStakeholderNotificationResult> {
    return this.notifyStakeholders({
      eventType: ProcessEventType.EscalationTriggered,
      processContext: context,
      eventTitle: `Escalation Level ${escalationLevel} Triggered`,
      eventMessage: `An escalation has been triggered for the ${context.processType} process: ${reason}`,
      priority: Priority.High,
      metadata: { escalationLevel, reason },
      actionUrl: `${this.siteUrl}/SitePages/ProcessDetails.aspx?processId=${context.processId}`
    });
  }

  // ============================================================================
  // INTEGRATION FIX P5: IT Admin Notification Confirmation Tracking
  // ============================================================================

  /**
   * Send IT Admin notification with confirmation tracking
   * Used for provisioning tasks that require IT Admin acknowledgment
   */
  public async sendITAdminNotificationWithConfirmation(
    context: IProcessContext,
    actionRequired: string,
    taskDescription: string,
    confirmationExpiryHours: number = 48
  ): Promise<{ notificationId: number; confirmationToken: string } | null> {
    if (!context.itAdminId) {
      logger.warn('StakeholderNotificationService', 'No IT Admin configured for process', { processId: context.processId });
      return null;
    }

    try {
      // Get IT Admin details
      const itAdmin = await this.getStakeholderForRole(StakeholderRole.ITAdmin, context);
      if (!itAdmin) {
        logger.warn('StakeholderNotificationService', 'Could not resolve IT Admin user');
        return null;
      }

      // Generate confirmation token
      const confirmationToken = uuidv4();
      const expiresAt = new Date();
      expiresAt.setHours(expiresAt.getHours() + confirmationExpiryHours);

      // Build confirmation URL
      const confirmationUrl = `${this.siteUrl}/SitePages/ConfirmITAction.aspx?token=${confirmationToken}&processId=${context.processId}`;

      // Create notification with confirmation tracking
      const notificationItem = await this.sp.web.lists.getByTitle('JML_Notifications').items.add({
        Title: `IT Action Required: ${actionRequired}`,
        NotificationType: 'ITProvisioning',
        MessageBody: this.formatITAdminNotificationBody(context, actionRequired, taskDescription, confirmationUrl),
        Priority: Priority.High,
        RecipientId: itAdmin.id,
        RecipientEmail: itAdmin.email,
        ProcessId: context.processId?.toString(),
        Status: 'Pending',
        DeliveryChannel: 'Email,InApp',
        ActionUrl: confirmationUrl,
        // INTEGRATION FIX P5: Confirmation tracking fields
        RequiresConfirmation: true,
        ConfirmationStatus: NotificationConfirmationStatus.Pending,
        ConfirmationToken: confirmationToken,
        ConfirmationExpiresAt: expiresAt,
        Created: new Date()
      });

      logger.info('StakeholderNotificationService', 'Sent IT Admin notification with confirmation tracking', {
        notificationId: notificationItem.data.Id,
        itAdminEmail: itAdmin.email,
        processId: context.processId,
        expiresAt
      });

      return {
        notificationId: notificationItem.data.Id,
        confirmationToken
      };
    } catch (error) {
      logger.error('StakeholderNotificationService', 'Failed to send IT Admin notification', error);
      return null;
    }
  }

  /**
   * Confirm an IT Admin notification by token
   */
  public async confirmITAdminNotification(
    confirmationToken: string,
    confirmedById: number,
    notes?: string
  ): Promise<{ success: boolean; processId?: number; error?: string }> {
    try {
      // Find notification by token
      const items = await this.sp.web.lists.getByTitle('JML_Notifications').items
        .filter(`ConfirmationToken eq '${confirmationToken}'`)
        .select('Id', 'ProcessId', 'ConfirmationStatus', 'ConfirmationExpiresAt')
        .top(1)();

      if (items.length === 0) {
        return { success: false, error: 'Invalid confirmation token' };
      }

      const notification = items[0];

      // Check if already confirmed
      if (notification.ConfirmationStatus === NotificationConfirmationStatus.Confirmed) {
        return { success: false, error: 'Notification already confirmed' };
      }

      // Check if expired
      if (notification.ConfirmationExpiresAt && new Date(notification.ConfirmationExpiresAt) < new Date()) {
        await this.sp.web.lists.getByTitle('JML_Notifications').items.getById(notification.Id).update({
          ConfirmationStatus: NotificationConfirmationStatus.Expired
        });
        return { success: false, error: 'Confirmation token has expired' };
      }

      // Mark as confirmed
      await this.sp.web.lists.getByTitle('JML_Notifications').items.getById(notification.Id).update({
        ConfirmationStatus: NotificationConfirmationStatus.Confirmed,
        ConfirmedAt: new Date(),
        ConfirmedById: confirmedById,
        ConfirmationNotes: notes || '',
        Status: 'Read'
      });

      logger.info('StakeholderNotificationService', 'IT Admin notification confirmed', {
        notificationId: notification.Id,
        processId: notification.ProcessId,
        confirmedById
      });

      return {
        success: true,
        processId: notification.ProcessId ? parseInt(notification.ProcessId, 10) : undefined
      };
    } catch (error) {
      logger.error('StakeholderNotificationService', 'Failed to confirm IT Admin notification', error);
      return { success: false, error: error instanceof Error ? error.message : 'Unknown error' };
    }
  }

  /**
   * Get pending IT Admin confirmations for a process
   */
  public async getPendingITAdminConfirmations(
    processId: number
  ): Promise<Array<{
    notificationId: number;
    title: string;
    createdAt: Date;
    expiresAt: Date;
    itAdminEmail: string;
  }>> {
    try {
      const items = await this.sp.web.lists.getByTitle('JML_Notifications').items
        .filter(`ProcessId eq '${processId}' and RequiresConfirmation eq 1 and ConfirmationStatus eq 'Pending'`)
        .select('Id', 'Title', 'Created', 'ConfirmationExpiresAt', 'RecipientEmail')
        .orderBy('Created', false)();

      return items.map(item => ({
        notificationId: item.Id,
        title: item.Title,
        createdAt: new Date(item.Created),
        expiresAt: new Date(item.ConfirmationExpiresAt),
        itAdminEmail: item.RecipientEmail
      }));
    } catch (error) {
      logger.error('StakeholderNotificationService', 'Failed to get pending confirmations', error);
      return [];
    }
  }

  /**
   * Check if all IT Admin confirmations are complete for a process
   */
  public async areAllITConfirmationsComplete(processId: number): Promise<boolean> {
    const pending = await this.getPendingITAdminConfirmations(processId);
    return pending.length === 0;
  }

  /**
   * Format IT Admin notification body with confirmation link
   */
  private formatITAdminNotificationBody(
    context: IProcessContext,
    actionRequired: string,
    taskDescription: string,
    confirmationUrl: string
  ): string {
    let body = `<div style="font-family: 'Segoe UI', sans-serif; padding: 20px;">`;
    body += `<h2 style="color: #0078d4;">IT Action Required</h2>`;
    body += `<p style="font-size: 16px; font-weight: 600; color: #323130;">${actionRequired}</p>`;
    body += `<p>${taskDescription}</p>`;

    body += `<div style="background-color: #f3f2f1; padding: 16px; border-radius: 4px; margin: 16px 0;">`;
    body += `<p><strong>Process Type:</strong> ${context.processType}</p>`;
    if (context.employeeName) {
      body += `<p><strong>Employee:</strong> ${context.employeeName}</p>`;
    }
    body += `<p><strong>Process ID:</strong> ${context.processId}</p>`;
    body += `</div>`;

    body += `<div style="margin: 24px 0;">`;
    body += `<a href="${confirmationUrl}" style="background-color: #0078d4; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; font-weight: 600;">`;
    body += `Confirm Action Complete`;
    body += `</a>`;
    body += `</div>`;

    body += `<p style="color: #605e5c; font-size: 12px;">Please click the button above to confirm when you have completed this action. This confirmation is required for the process to proceed.</p>`;

    body += `<hr style="border: none; border-top: 1px solid #edebe9; margin: 20px 0;" />`;
    body += `<p style="color: #605e5c; font-size: 12px;">This is an automated notification from the JML Employee Lifecycle Management System.</p>`;
    body += `</div>`;

    return body;
  }
}

/**
 * Factory function to create StakeholderNotificationService
 */
export function createStakeholderNotificationService(
  sp: SPFI,
  siteUrl: string
): StakeholderNotificationService {
  return new StakeholderNotificationService(sp, siteUrl);
}
