// @ts-nocheck
// Approval Notification Service
// Sends email notifications and reminders for approvals
// INTEGRATION FIX P1: Now respects user notification preferences

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';
import '@pnp/sp/sputilities';
import {
  IJmlApproval,
  IApprovalNotification,
  ApprovalStatus
} from '../models';
import { Priority } from '../models/ICommon';
import { logger } from './LoggingService';
import {
  NotificationPreferencesService,
  NotificationEventType,
  NotificationChannel,
  DigestFrequency
} from './workflow/NotificationPreferencesService';

export class ApprovalNotificationService {
  private sp: SPFI;
  private siteUrl: string;
  private preferencesService: NotificationPreferencesService;

  /**
   * Convert Date to display format
   */
  private formatDate(date: Date | string): string {
    const d = typeof date === 'string' ? new Date(date) : date;
    return d.toLocaleDateString();
  }

  constructor(sp: SPFI, siteUrl: string) {
    this.sp = sp;
    this.siteUrl = siteUrl;
    // INTEGRATION FIX P1: Initialize preferences service
    this.preferencesService = new NotificationPreferencesService(sp);
  }

  /**
   * Send notification for new approval
   */
  public async sendNewApprovalNotification(approval: Partial<IJmlApproval>): Promise<void> {
    try {
      const notification: IApprovalNotification = {
        approvalId: approval.Id,
        recipientId: approval.ApproverId,
        notificationType: 'NewApproval',
        subject: `New Approval Request: ${approval.ProcessTitle}`,
        body: this.buildNewApprovalEmail(approval),
        sendEmail: true,
        sendInApp: true
      };

      await this.sendNotification(notification);
    } catch (error) {
      logger.error('ApprovalNotificationService', 'Failed to send new approval notification:', error);
    }
  }

  /**
   * Send reminder for pending approval
   */
  public async sendReminderNotification(approval: Partial<IJmlApproval>): Promise<void> {
    try {
      const daysOverdue = this.calculateDaysOverdue(approval.DueDate);

      const notification: IApprovalNotification = {
        approvalId: approval.Id,
        recipientId: approval.ApproverId,
        notificationType: 'Reminder',
        subject: `Reminder: Approval Required - ${approval.ProcessTitle}`,
        body: this.buildReminderEmail(approval, daysOverdue),
        sendEmail: true,
        sendInApp: true
      };

      await this.sendNotification(notification);
    } catch (error) {
      logger.error('ApprovalNotificationService', 'Failed to send reminder notification:', error);
    }
  }

  /**
   * Send escalation notification
   */
  public async sendEscalationNotification(
    approval: Partial<IJmlApproval>,
    escalatedToId: number
  ): Promise<void> {
    try {
      const notification: IApprovalNotification = {
        approvalId: approval.Id,
        recipientId: escalatedToId,
        notificationType: 'Escalation',
        subject: `Escalated Approval: ${approval.ProcessTitle}`,
        body: this.buildEscalationEmail(approval),
        sendEmail: true,
        sendInApp: true
      };

      await this.sendNotification(notification);
    } catch (error) {
      logger.error('ApprovalNotificationService', 'Failed to send escalation notification:', error);
    }
  }

  /**
   * Send delegation notification
   */
  public async sendDelegationNotification(
    approval: Partial<IJmlApproval>,
    delegatedToId: number,
    reason: string
  ): Promise<void> {
    try {
      const notification: IApprovalNotification = {
        approvalId: approval.Id,
        recipientId: delegatedToId,
        notificationType: 'Delegated',
        subject: `Delegated Approval: ${approval.ProcessTitle}`,
        body: this.buildDelegationEmail(approval, reason),
        sendEmail: true,
        sendInApp: true
      };

      await this.sendNotification(notification);
    } catch (error) {
      logger.error('ApprovalNotificationService', 'Failed to send delegation notification:', error);
    }
  }

  /**
   * Send completion notification
   */
  public async sendCompletionNotification(
    approval: Partial<IJmlApproval>,
    requesterId: number
  ): Promise<void> {
    try {
      const statusText = approval.Status === ApprovalStatus.Approved ? 'Approved' : 'Rejected';

      const notification: IApprovalNotification = {
        approvalId: approval.Id,
        recipientId: requesterId,
        notificationType: 'Completed',
        subject: `Approval ${statusText}: ${approval.ProcessTitle}`,
        body: this.buildCompletionEmail(approval),
        sendEmail: true,
        sendInApp: true
      };

      await this.sendNotification(notification);
    } catch (error) {
      logger.error('ApprovalNotificationService', 'Failed to send completion notification:', error);
    }
  }

  /**
   * Send notification via SharePoint utilities
   */
  /**
   * Send notification respecting user preferences
   * INTEGRATION FIX P1: Now checks user preferences before sending
   */
  private async sendNotification(notification: IApprovalNotification): Promise<void> {
    try {
      // Get recipient email
      const user = await this.sp.web.siteUsers.getById(notification.recipientId)();

      if (!user || !user.Email) {
        logger.warn('ApprovalNotificationService', 'User email not found', { recipientId: notification.recipientId });
        return;
      }

      // INTEGRATION FIX P1: Map notification type to event type for preferences
      const eventType = this.mapToNotificationEventType(notification.notificationType);

      // Check user preferences
      const deliverySettings = await this.preferencesService.resolveDeliverySettings(
        notification.recipientId,
        user.Email,
        eventType,
        Priority.High // Approvals are high priority by default
      );

      // Skip if user has disabled this notification type
      if (!deliverySettings.shouldDeliver) {
        logger.info('ApprovalNotificationService',
          `Skipping approval notification for user ${notification.recipientId}: ${deliverySettings.reason}`);
        return;
      }

      // Queue for digest if user prefers digest delivery
      if (deliverySettings.isDigest && deliverySettings.digestFrequency !== DigestFrequency.Immediate) {
        await this.preferencesService.queueForDigest(
          notification.recipientId,
          eventType,
          notification.subject,
          notification.body,
          Priority.High,
          deliverySettings.digestFrequency,
          notification.approvalId,
          'Approval'
        );
        logger.info('ApprovalNotificationService',
          `Queued approval notification for user ${notification.recipientId} for ${deliverySettings.digestFrequency} digest`);
        return;
      }

      // Send email if enabled in preferences
      if (notification.sendEmail && deliverySettings.channels.includes(NotificationChannel.Email)) {
        const emailProps = {
          To: [user.Email],
          Subject: notification.subject,
          Body: notification.body
        };
        await this.sp.utility.sendEmail(emailProps);
        logger.debug('ApprovalNotificationService', 'Email sent successfully', { to: user.Email, subject: notification.subject });
      }

      // Send in-app notification if enabled
      if (notification.sendInApp && deliverySettings.channels.includes(NotificationChannel.InApp)) {
        await this.sendInAppNotification(notification, notification.recipientId);
      }

      logger.debug('ApprovalNotificationService', 'Notification sent', { type: notification.notificationType, to: user.Email });
    } catch (error) {
      logger.error('ApprovalNotificationService', 'Failed to send notification:', error);
      throw error;
    }
  }

  /**
   * Send in-app notification to JML_Notifications list
   * INTEGRATION FIX P1: Separate in-app notification delivery
   */
  private async sendInAppNotification(notification: IApprovalNotification, recipientId: number): Promise<void> {
    await this.sp.web.lists.getByTitle('JML_Notifications').items.add({
      Title: notification.subject,
      Message: notification.body.replace(/<[^>]*>/g, '').substring(0, 500), // Strip HTML, limit length
      RecipientId: recipientId,
      Type: 'Approval',
      Priority: 'High',
      IsRead: false,
      RelatedItemType: 'Approval',
      RelatedItemId: notification.approvalId
    });
  }

  /**
   * Map notification type to NotificationEventType for preferences
   * INTEGRATION FIX P1: Enables preference checking for approvals
   */
  private mapToNotificationEventType(notificationType: string): NotificationEventType {
    const mapping: Record<string, NotificationEventType> = {
      'NewApproval': NotificationEventType.ApprovalRequired,
      'Reminder': NotificationEventType.Reminder,
      'Escalation': NotificationEventType.ApprovalEscalated,
      'Delegated': NotificationEventType.ApprovalRequired,
      'Completion': NotificationEventType.ApprovalCompleted,
      'Approved': NotificationEventType.ApprovalCompleted,
      'Rejected': NotificationEventType.ApprovalRejected
    };
    return mapping[notificationType] || NotificationEventType.ApprovalRequired;
  }

  /**
   * Calculate days overdue
   */
  private calculateDaysOverdue(dueDate: Date | string): number {
    const now = new Date();
    const due = typeof dueDate === 'string' ? new Date(dueDate) : dueDate;
    return Math.ceil((now.getTime() - due.getTime()) / (1000 * 60 * 60 * 24));
  }

  /**
   * Build new approval email
   */
  private buildNewApprovalEmail(approval: Partial<IJmlApproval>): string {
    const approvalUrl = `${this.siteUrl}/SitePages/ApprovalCenter.aspx?approvalId=${approval.Id}`;

    return `
      <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #323130;">
          <h2 style="color: #0078d4;">New Approval Request</h2>

          <p>You have a new approval request that requires your attention.</p>

          <table style="border-collapse: collapse; width: 100%; max-width: 600px; margin: 20px 0;">
            <tr>
              <td style="padding: 10px; background-color: #f3f2f1; font-weight: 600;">Process:</td>
              <td style="padding: 10px;">${approval.ProcessTitle}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background-color: #f3f2f1; font-weight: 600;">Type:</td>
              <td style="padding: 10px;">${approval.ProcessType}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background-color: #f3f2f1; font-weight: 600;">Level:</td>
              <td style="padding: 10px;">Level ${approval.ApprovalLevel}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background-color: #f3f2f1; font-weight: 600;">Due Date:</td>
              <td style="padding: 10px;">${this.formatDate(approval.DueDate)}</td>
            </tr>
          </table>

          <p style="margin-top: 20px;">
            <a href="${approvalUrl}"
               style="background-color: #0078d4; color: white; padding: 12px 24px;
                      text-decoration: none; border-radius: 4px; display: inline-block;">
              Review Approval
            </a>
          </p>

          <p style="color: #605e5c; font-size: 12px; margin-top: 30px;">
            This is an automated notification from the JML Approval System.
          </p>
        </body>
      </html>
    `;
  }

  /**
   * Build reminder email
   */
  private buildReminderEmail(approval: Partial<IJmlApproval>, daysOverdue: number): string {
    const approvalUrl = `${this.siteUrl}/SitePages/ApprovalCenter.aspx?approvalId=${approval.Id}`;
    const urgency = daysOverdue > 0 ? 'OVERDUE' : 'DUE SOON';

    return `
      <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #323130;">
          <h2 style="color: ${daysOverdue > 0 ? '#d13438' : '#f7630c'};">Approval Reminder - ${urgency}</h2>

          <p>This is a reminder that the following approval is ${daysOverdue > 0 ? 'overdue' : 'due soon'}.</p>

          <table style="border-collapse: collapse; width: 100%; max-width: 600px; margin: 20px 0;">
            <tr>
              <td style="padding: 10px; background-color: #f3f2f1; font-weight: 600;">Process:</td>
              <td style="padding: 10px;">${approval.ProcessTitle}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background-color: #f3f2f1; font-weight: 600;">Due Date:</td>
              <td style="padding: 10px; ${daysOverdue > 0 ? 'color: #d13438; font-weight: 600;' : ''}">
                ${this.formatDate(approval.DueDate)}
                ${daysOverdue > 0 ? `(${daysOverdue} days overdue)` : ''}
              </td>
            </tr>
            <tr>
              <td style="padding: 10px; background-color: #f3f2f1; font-weight: 600;">Requested:</td>
              <td style="padding: 10px;">${this.formatDate(approval.RequestedDate)}</td>
            </tr>
          </table>

          <p style="margin-top: 20px;">
            <a href="${approvalUrl}"
               style="background-color: ${daysOverdue > 0 ? '#d13438' : '#f7630c'}; color: white;
                      padding: 12px 24px; text-decoration: none; border-radius: 4px; display: inline-block;">
              Review Approval Now
            </a>
          </p>

          <p style="color: #605e5c; font-size: 12px; margin-top: 30px;">
            This is an automated reminder from the JML Approval System.
          </p>
        </body>
      </html>
    `;
  }

  /**
   * Build escalation email
   */
  private buildEscalationEmail(approval: Partial<IJmlApproval>): string {
    const approvalUrl = `${this.siteUrl}/SitePages/ApprovalCenter.aspx?approvalId=${approval.Id}`;

    return `
      <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #323130;">
          <h2 style="color: #d13438;">Escalated Approval Request</h2>

          <p>An approval has been escalated to you due to no response from the original approver.</p>

          <table style="border-collapse: collapse; width: 100%; max-width: 600px; margin: 20px 0;">
            <tr>
              <td style="padding: 10px; background-color: #f3f2f1; font-weight: 600;">Process:</td>
              <td style="padding: 10px;">${approval.ProcessTitle}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background-color: #f3f2f1; font-weight: 600;">Original Approver:</td>
              <td style="padding: 10px;">${approval.OriginalApprover?.Title || approval.Approver.Title}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background-color: #f3f2f1; font-weight: 600;">Original Due Date:</td>
              <td style="padding: 10px;">${this.formatDate(approval.DueDate)}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background-color: #f3f2f1; font-weight: 600;">Escalation Level:</td>
              <td style="padding: 10px;">${approval.EscalationLevel}</td>
            </tr>
          </table>

          <p style="margin-top: 20px;">
            <a href="${approvalUrl}"
               style="background-color: #d13438; color: white; padding: 12px 24px;
                      text-decoration: none; border-radius: 4px; display: inline-block;">
              Review Escalated Approval
            </a>
          </p>

          <p style="color: #605e5c; font-size: 12px; margin-top: 30px;">
            This is an automated escalation from the JML Approval System.
          </p>
        </body>
      </html>
    `;
  }

  /**
   * Build delegation email
   */
  private buildDelegationEmail(approval: Partial<IJmlApproval>, reason: string): string {
    const approvalUrl = `${this.siteUrl}/SitePages/ApprovalCenter.aspx?approvalId=${approval.Id}`;

    return `
      <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #323130;">
          <h2 style="color: #0078d4;">Delegated Approval Request</h2>

          <p>An approval has been delegated to you by ${approval.DelegatedBy?.Title || approval.OriginalApprover?.Title}.</p>

          ${reason ? `<p style="background-color: #f3f2f1; padding: 12px; border-left: 4px solid #0078d4;">
            <strong>Delegation Reason:</strong> ${reason}
          </p>` : ''}

          <table style="border-collapse: collapse; width: 100%; max-width: 600px; margin: 20px 0;">
            <tr>
              <td style="padding: 10px; background-color: #f3f2f1; font-weight: 600;">Process:</td>
              <td style="padding: 10px;">${approval.ProcessTitle}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background-color: #f3f2f1; font-weight: 600;">Original Approver:</td>
              <td style="padding: 10px;">${approval.OriginalApprover?.Title || 'N/A'}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background-color: #f3f2f1; font-weight: 600;">Due Date:</td>
              <td style="padding: 10px;">${this.formatDate(approval.DueDate)}</td>
            </tr>
          </table>

          <p style="margin-top: 20px;">
            <a href="${approvalUrl}"
               style="background-color: #0078d4; color: white; padding: 12px 24px;
                      text-decoration: none; border-radius: 4px; display: inline-block;">
              Review Delegated Approval
            </a>
          </p>

          <p style="color: #605e5c; font-size: 12px; margin-top: 30px;">
            This is an automated notification from the JML Approval System.
          </p>
        </body>
      </html>
    `;
  }

  /**
   * Build completion email
   */
  private buildCompletionEmail(approval: Partial<IJmlApproval>): string {
    const statusColor = approval.Status === ApprovalStatus.Approved ? '#107c10' : '#d13438';
    const statusText = approval.Status === ApprovalStatus.Approved ? 'Approved' : 'Rejected';

    return `
      <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #323130;">
          <h2 style="color: ${statusColor};">Approval ${statusText}</h2>

          <p>Your approval request has been ${statusText.toLowerCase()} by ${approval.Approver.Title}.</p>

          <table style="border-collapse: collapse; width: 100%; max-width: 600px; margin: 20px 0;">
            <tr>
              <td style="padding: 10px; background-color: #f3f2f1; font-weight: 600;">Process:</td>
              <td style="padding: 10px;">${approval.ProcessTitle}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background-color: #f3f2f1; font-weight: 600;">Decision:</td>
              <td style="padding: 10px; color: ${statusColor}; font-weight: 600;">${statusText}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background-color: #f3f2f1; font-weight: 600;">Approver:</td>
              <td style="padding: 10px;">${approval.Approver.Title}</td>
            </tr>
            <tr>
              <td style="padding: 10px; background-color: #f3f2f1; font-weight: 600;">Date:</td>
              <td style="padding: 10px;">${approval.ActualCompletionDate ? this.formatDate(approval.ActualCompletionDate) : 'N/A'}</td>
            </tr>
          </table>

          ${approval.Notes ? `<p style="background-color: #f3f2f1; padding: 12px; border-left: 4px solid ${statusColor};">
            <strong>Comments:</strong> ${approval.Notes}
          </p>` : ''}

          <p style="color: #605e5c; font-size: 12px; margin-top: 30px;">
            This is an automated notification from the JML Approval System.
          </p>
        </body>
      </html>
    `;
  }

  /**
   * Process reminder queue (to be called by timer job)
   */
  public async processReminders(approvals: IJmlApproval[]): Promise<void> {
    const now = new Date();

    for (let i = 0; i < approvals.length; i++) {
      const approval = approvals[i];

      if (approval.Status !== ApprovalStatus.Pending) {
        continue;
      }

      const dueDate = typeof approval.DueDate === 'string' ? new Date(approval.DueDate) : approval.DueDate;
      const daysToDue = Math.ceil((dueDate.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));

      // Send reminder 1 day before due, on due date, and every day after
      if (daysToDue === 1 || daysToDue === 0 || daysToDue < 0) {
        await this.sendReminderNotification(approval);
      }
    }
  }
}
