// @ts-nocheck
/**
 * NotificationActionHandler
 * Handles notification-related workflow actions
 * Sends in-app notifications, emails, and Teams messages
 * Integrates with WorkflowNotificationService for consistent templates
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import {
  IStepConfig,
  IActionContext,
  IActionResult,
  IActionConfig
} from '../../../models/IWorkflow';
import { WorkflowConditionEvaluator } from '../WorkflowConditionEvaluator';
import { WorkflowNotificationService, WorkflowNotificationEvent, NotificationChannel } from '../WorkflowNotificationService';
import { logger } from '../../LoggingService';
import {
  retryWithDLQ,
  NOTIFICATION_RETRY_OPTIONS,
  notificationDLQ,
  IRetryResult
} from '../../../utils/retryUtils';

/**
 * Notification delivery result with verification
 */
export interface INotificationDeliveryResult {
  success: boolean;
  notificationId?: number;
  recipientId?: number;
  recipientEmail?: string;
  channel: 'inApp' | 'email' | 'teams';
  deliveryStatus: 'sent' | 'delivered' | 'failed' | 'queued';
  attempts: number;
  error?: string;
  dlqItemId?: string;
}

/**
 * Batch notification result
 */
export interface IBatchNotificationResult {
  totalRecipients: number;
  successCount: number;
  failureCount: number;
  queuedCount: number;
  deliveryResults: INotificationDeliveryResult[];
}

export class NotificationActionHandler {
  private sp: SPFI;
  private context: WebPartContext;
  private conditionEvaluator: WorkflowConditionEvaluator;
  private notificationService: WorkflowNotificationService;

  constructor(sp: SPFI, context: WebPartContext) {
    this.sp = sp;
    this.context = context;
    this.conditionEvaluator = new WorkflowConditionEvaluator();
    this.notificationService = new WorkflowNotificationService(sp, context);
  }

  /**
   * Send notification (in-app + email for certain types like Welcome)
   * For Welcome, Transfer, Completion notification types, also sends email
   */
  public async sendNotification(config: IStepConfig, context: IActionContext): Promise<IActionResult> {
    try {
      const recipientId = config.recipientId;
      const messageTemplate = config.messageTemplate || `Notification from workflow step: ${context.currentStep.name}`;
      const notificationType = config.notificationType || 'Info';

      // Process message template
      const evalContext = {
        ...context.process,
        ...context.variables,
        workflowInstance: context.workflowInstance,
        currentStep: context.currentStep
      };
      const message = this.conditionEvaluator.replaceTokens(messageTemplate, evalContext);

      // Resolve recipient email from field if specified
      let recipientEmail: string | undefined;
      if (config.recipientField) {
        const fieldValue = this.conditionEvaluator.evaluateExpression(config.recipientField, evalContext);
        if (typeof fieldValue === 'string' && fieldValue.includes('@')) {
          recipientEmail = fieldValue;
        }
      }

      const createdItemIds: number[] = [];
      let emailSent = false;

      // Create in-app notification if we have a recipientId
      if (recipientId) {
        const notificationData = {
          Title: notificationType,
          Message: message,
          RecipientId: recipientId,
          Type: notificationType,
          Priority: 'Normal',
          IsRead: false,
          RelatedItemType: 'WorkflowInstance',
          RelatedItemId: context.workflowInstance.Id,
          WorkflowInstanceId: context.workflowInstance.Id,
          WorkflowStepId: context.currentStep.id
        };

        const result = await this.sp.web.lists.getByTitle('PM_Notifications').items.add(notificationData);
        createdItemIds.push(result.data.Id);
        logger.info('NotificationActionHandler', `Created in-app notification: ${result.data.Id}`);
      }

      // For certain notification types, also send email
      const emailNotificationTypes = ['Welcome', 'Completion', 'Transfer', 'Farewell', 'Reminder'];
      if (recipientEmail && emailNotificationTypes.includes(notificationType)) {
        try {
          await this.sendWelcomeEmail(notificationType, recipientEmail, message, evalContext);
          emailSent = true;
          logger.info('NotificationActionHandler', `Sent ${notificationType} email to ${recipientEmail}`);
        } catch (emailError) {
          logger.warn('NotificationActionHandler', `Failed to send ${notificationType} email`, emailError);
          // Don't fail the step for email errors - in-app notification was created
        }
      }

      return {
        success: true,
        nextAction: 'continue',
        createdItemIds,
        outputVariables: {
          notificationId: createdItemIds[0],
          emailSent,
          notificationType
        }
      };
    } catch (error) {
      logger.error('NotificationActionHandler', 'Error sending notification', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to send notification'
      };
    }
  }

  /**
   * Send welcome/completion email via Microsoft Graph
   * Used for Welcome, Completion, Transfer notification types
   */
  private async sendWelcomeEmail(
    notificationType: string,
    recipientEmail: string,
    message: string,
    evalContext: Record<string, unknown>
  ): Promise<void> {
    const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');

    const employeeName = evalContext.employeeName as string || 'Team Member';
    const processType = evalContext.processType as string || 'Process';

    // Build subject based on notification type
    let subject: string;
    let headerColor: string;
    let headerIcon: string;

    switch (notificationType) {
      case 'Welcome':
        subject = `Welcome to the Team, ${employeeName}!`;
        headerColor = '#107c10'; // Green
        headerIcon = '🎉';
        break;
      case 'Completion':
        subject = `${processType} Process Complete - ${employeeName}`;
        headerColor = '#0078d4'; // Blue
        headerIcon = '✅';
        break;
      case 'Transfer':
        subject = `Transfer Update - ${employeeName}`;
        headerColor = '#8764b8'; // Purple
        headerIcon = '🔄';
        break;
      case 'Farewell':
        subject = `Farewell - ${employeeName}`;
        headerColor = '#605e5c'; // Gray
        headerIcon = '👋';
        break;
      default:
        subject = `JML Notification - ${employeeName}`;
        headerColor = '#0078d4';
        headerIcon = '📬';
    }

    const htmlBody = this.buildWelcomeEmailHtml(
      notificationType,
      employeeName,
      message,
      headerColor,
      headerIcon,
      evalContext
    );

    const emailMessage = {
      message: {
        subject,
        body: {
          contentType: 'HTML',
          content: htmlBody
        },
        toRecipients: [
          { emailAddress: { address: recipientEmail } }
        ]
      },
      saveToSentItems: false
    };

    await graphClient.api('/me/sendMail').post(emailMessage);
  }

  /**
   * Build welcome/completion email HTML with rich branding
   * Uses table-based layout for Outlook/Gmail compatibility
   */
  private buildWelcomeEmailHtml(
    notificationType: string,
    employeeName: string,
    message: string,
    headerColor: string,
    headerIcon: string,
    evalContext: Record<string, unknown>
  ): string {
    const branding = this.getBrandingConfig();
    const managerName = evalContext.managerName as string || 'Your Manager';
    const department = evalContext.department as string || '';
    const jobTitle = evalContext.jobTitle as string || '';
    const startDate = evalContext.startDate as string || '';

    // Build additional content sections based on notification type
    let detailsSection = '';
    let nextStepsSection = '';

    if (notificationType === 'Welcome' || notificationType === 'Transfer') {
      detailsSection = `
        <!-- Details Card -->
        <tr>
          <td style="padding: 0 32px 24px 32px;">
            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="background-color: #f8f9fa; border-radius: 8px; border-left: 4px solid ${branding.primaryColor};">
              <tr>
                <td style="padding: 20px 24px;">
                  <h3 style="margin: 0 0 16px 0; color: #323130; font-size: 16px; font-weight: 600;">Your Details</h3>
                  <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
                    ${jobTitle ? `
                    <tr>
                      <td style="padding: 6px 0; color: #605e5c; font-size: 14px; width: 120px;"><strong>Role:</strong></td>
                      <td style="padding: 6px 0; color: #323130; font-size: 14px;">${jobTitle}</td>
                    </tr>` : ''}
                    ${department ? `
                    <tr>
                      <td style="padding: 6px 0; color: #605e5c; font-size: 14px;"><strong>Department:</strong></td>
                      <td style="padding: 6px 0; color: #323130; font-size: 14px;">${department}</td>
                    </tr>` : ''}
                    ${managerName ? `
                    <tr>
                      <td style="padding: 6px 0; color: #605e5c; font-size: 14px;"><strong>Manager:</strong></td>
                      <td style="padding: 6px 0; color: #323130; font-size: 14px;">${managerName}</td>
                    </tr>` : ''}
                    ${startDate ? `
                    <tr>
                      <td style="padding: 6px 0; color: #605e5c; font-size: 14px;"><strong>Start Date:</strong></td>
                      <td style="padding: 6px 0; color: #323130; font-size: 14px;">${startDate}</td>
                    </tr>` : ''}
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>`;

      if (notificationType === 'Welcome') {
        nextStepsSection = `
        <!-- Next Steps Card -->
        <tr>
          <td style="padding: 0 32px 24px 32px;">
            <h3 style="margin: 0 0 16px 0; color: #323130; font-size: 16px; font-weight: 600;">What's Next?</h3>
            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%">
              <tr>
                <td style="padding: 8px 0; font-size: 14px; color: #323130;">
                  <span style="color: ${branding.primaryColor}; font-weight: bold;">✓</span>&nbsp;&nbsp;Check your JML portal for onboarding tasks
                </td>
              </tr>
              <tr>
                <td style="padding: 8px 0; font-size: 14px; color: #323130;">
                  <span style="color: ${branding.primaryColor}; font-weight: bold;">✓</span>&nbsp;&nbsp;Connect with your manager and team members
                </td>
              </tr>
              <tr>
                <td style="padding: 8px 0; font-size: 14px; color: #323130;">
                  <span style="color: ${branding.primaryColor}; font-weight: bold;">✓</span>&nbsp;&nbsp;Review the employee handbook and company policies
                </td>
              </tr>
              <tr>
                <td style="padding: 8px 0; font-size: 14px; color: #323130;">
                  <span style="color: ${branding.primaryColor}; font-weight: bold;">✓</span>&nbsp;&nbsp;Complete required training and compliance modules
                </td>
              </tr>
            </table>
          </td>
        </tr>`;
      }
    }

    // Determine heading text
    let headingText = notificationType;
    if (notificationType === 'Welcome') {
      headingText = `Welcome, ${employeeName}!`;
    } else if (notificationType === 'Completion') {
      headingText = `Process Complete`;
    } else if (notificationType === 'Farewell') {
      headingText = `Farewell, ${employeeName}`;
    }

    return `
      <!DOCTYPE html>
      <html lang="en" xmlns="http://www.w3.org/1999/xhtml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
      <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <title>${headingText}</title>
        <!--[if mso]>
        <style type="text/css">
          table { border-collapse: collapse; }
          .button-link { padding: 14px 28px !important; }
        </style>
        <![endif]-->
      </head>
      <body style="margin: 0; padding: 0; background-color: #f5f5f5; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">
        <!-- Wrapper Table -->
        <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="background-color: #f5f5f5;">
          <tr>
            <td align="center" style="padding: 24px 10px;">
              <!-- Main Container -->
              <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="600" style="max-width: 600px; background-color: #ffffff; border-radius: 12px; overflow: hidden;">

                <!-- Hero Header with Gradient Effect -->
                <tr>
                  <td style="background: linear-gradient(135deg, ${headerColor} 0%, ${this.adjustColor(headerColor, -20)} 100%); background-color: ${headerColor}; padding: 40px 32px; text-align: center;">
                    ${branding.logoUrl ? `<img src="${branding.logoUrl}" alt="${branding.companyName}" style="max-height: 36px; margin-bottom: 20px;">` : ''}
                    <div style="font-size: 48px; margin-bottom: 16px;">${headerIcon}</div>
                    <h1 style="margin: 0; color: #ffffff; font-size: 28px; font-weight: 600; line-height: 1.3;">${headingText}</h1>
                    ${notificationType === 'Welcome' ? `<p style="margin: 12px 0 0 0; color: rgba(255,255,255,0.9); font-size: 16px;">We're thrilled to have you join us!</p>` : ''}
                    ${notificationType === 'Completion' ? `<p style="margin: 12px 0 0 0; color: rgba(255,255,255,0.9); font-size: 16px;">${employeeName}'s process has been completed successfully</p>` : ''}
                  </td>
                </tr>

                <!-- Main Message -->
                <tr>
                  <td style="padding: 32px; color: #323130; font-size: 15px; line-height: 1.8;">
                    <p style="margin: 0;">${message.replace(/\n/g, '<br>')}</p>
                  </td>
                </tr>

                ${detailsSection}
                ${nextStepsSection}

                <!-- CTA Button -->
                <tr>
                  <td style="padding: 0 32px 32px 32px; text-align: center;">
                    <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="margin: 0 auto;">
                      <tr>
                        <td style="border-radius: 6px; background-color: ${branding.primaryColor};">
                          <a href="${branding.portalUrl}" target="_blank" class="button-link" style="display: inline-block; padding: 14px 32px; color: #ffffff; font-size: 15px; font-weight: 600; text-decoration: none; text-align: center;">
                            Open Policy Manager &rarr;
                          </a>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>

                <!-- Divider -->
                <tr>
                  <td style="padding: 0 32px;">
                    <hr style="border: none; border-top: 1px solid #edebe9; margin: 0;">
                  </td>
                </tr>

                <!-- Footer -->
                <tr>
                  <td style="padding: 24px 32px; background-color: #faf9f8; text-align: center;">
                    <p style="margin: 0 0 8px 0; font-size: 13px; color: #323130; font-weight: 600;">
                      ${branding.companyName}
                    </p>
                    <p style="margin: 0 0 12px 0; font-size: 12px; color: #605e5c;">
                      This email was sent automatically. Please do not reply directly.
                    </p>
                    <p style="margin: 0; font-size: 11px; color: #8a8886;">
                      Questions? Contact <a href="mailto:${branding.supportEmail}" style="color: ${branding.primaryColor};">HR Support</a> or your manager.
                    </p>
                  </td>
                </tr>

              </table>

              <!-- Email Footer Note -->
              <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="600" style="max-width: 600px;">
                <tr>
                  <td style="padding: 16px 32px; text-align: center;">
                    <p style="margin: 0; font-size: 11px; color: #a19f9d;">
                      &copy; ${new Date().getFullYear()} ${branding.companyName}. All rights reserved.
                    </p>
                  </td>
                </tr>
              </table>

            </td>
          </tr>
        </table>
      </body>
      </html>
    `;
  }

  /**
   * Adjust hex color brightness
   */
  private adjustColor(hex: string, amount: number): string {
    const clamp = (val: number): number => Math.min(255, Math.max(0, val));
    const num = parseInt(hex.replace('#', ''), 16);
    const r = clamp((num >> 16) + amount);
    const g = clamp(((num >> 8) & 0x00FF) + amount);
    const b = clamp((num & 0x0000FF) + amount);
    return `#${(1 << 24 | r << 16 | g << 8 | b).toString(16).slice(1)}`;
  }

  /**
   * Send email via Microsoft Graph
   */
  public async sendEmail(config: IActionConfig, context: IActionContext): Promise<IActionResult> {
    try {
      const toAddresses = config.to as string[] || [];
      const toField = config.toField as string;
      const ccAddresses = config.cc as string[] || [];
      const subject = config.subject as string || 'Workflow Notification';
      const bodyTemplate = config.body as string || '';

      // Resolve recipient from field if specified
      if (toField && toAddresses.length === 0) {
        const fieldValue = context.process[toField] || context.variables[toField];
        if (typeof fieldValue === 'string' && fieldValue.includes('@')) {
          toAddresses.push(fieldValue);
        }
      }

      if (toAddresses.length === 0) {
        return { success: false, error: 'No email recipients specified' };
      }

      // Process body template
      const evalContext = {
        ...context.process,
        ...context.variables,
        workflowInstance: context.workflowInstance
      };
      const body = this.conditionEvaluator.replaceTokens(bodyTemplate, evalContext);
      const processedSubject = this.conditionEvaluator.replaceTokens(subject, evalContext);

      // Build email message
      const emailMessage = {
        message: {
          subject: processedSubject,
          body: {
            contentType: 'HTML',
            content: body
          },
          toRecipients: toAddresses.map(email => ({
            emailAddress: { address: email }
          })),
          ccRecipients: ccAddresses.map(email => ({
            emailAddress: { address: email }
          }))
        },
        saveToSentItems: true
      };

      // Send via Graph API
      const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
      await graphClient.api('/me/sendMail').post(emailMessage);

      logger.info('NotificationActionHandler', `Sent email to ${toAddresses.join(', ')}`);

      return {
        success: true,
        nextAction: 'continue',
        outputVariables: {
          emailSent: true,
          emailRecipients: toAddresses
        }
      };
    } catch (error) {
      logger.error('NotificationActionHandler', 'Error sending email', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to send email'
      };
    }
  }

  /**
   * Send Teams message via Microsoft Graph
   */
  public async sendTeamsMessage(config: IActionConfig, context: IActionContext): Promise<IActionResult> {
    try {
      const channelId = config.channelId as string;
      const teamId = config.teamId as string;
      const messageContent = config.messageContent as string || '';

      if (!teamId || !channelId) {
        // Try to send as chat message to user
        return await this.sendTeamsChatMessage(config as Record<string, unknown>, context);
      }

      // Process message template
      const evalContext = {
        ...context.process,
        ...context.variables,
        workflowInstance: context.workflowInstance
      };
      const content = this.conditionEvaluator.replaceTokens(messageContent, evalContext);

      // Build Teams message
      const message = {
        body: {
          contentType: 'html',
          content
        }
      };

      // Send via Graph API
      const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
      await graphClient.api(`/teams/${teamId}/channels/${channelId}/messages`).post(message);

      logger.info('NotificationActionHandler', `Sent Teams message to channel ${channelId}`);

      return {
        success: true,
        nextAction: 'continue',
        outputVariables: {
          teamsMessageSent: true,
          channelId
        }
      };
    } catch (error) {
      logger.error('NotificationActionHandler', 'Error sending Teams message', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to send Teams message'
      };
    }
  }

  /**
   * Send Teams chat message to a user
   */
  private async sendTeamsChatMessage(config: Record<string, unknown>, context: IActionContext): Promise<IActionResult> {
    try {
      const recipientEmail = config.recipientEmail as string;
      const messageContent = config.messageContent as string || '';

      if (!recipientEmail) {
        return { success: false, error: 'Recipient email not specified for Teams chat' };
      }

      // Process message template
      const evalContext = {
        ...context.process,
        ...context.variables,
        workflowInstance: context.workflowInstance
      };
      const content = this.conditionEvaluator.replaceTokens(messageContent, evalContext);

      const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');

      // Get user ID
      const user = await graphClient.api(`/users/${recipientEmail}`).select('id').get();

      // Create or get chat
      const chat = {
        chatType: 'oneOnOne',
        members: [
          {
            '@odata.type': '#microsoft.graph.aadUserConversationMember',
            roles: ['owner'],
            'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${this.context.pageContext.user.loginName}`
          },
          {
            '@odata.type': '#microsoft.graph.aadUserConversationMember',
            roles: ['owner'],
            'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${user.id}`
          }
        ]
      };

      const createdChat = await graphClient.api('/chats').post(chat);

      // Send message
      await graphClient.api(`/chats/${createdChat.id}/messages`).post({
        body: {
          content
        }
      });

      logger.info('NotificationActionHandler', `Sent Teams chat message to ${recipientEmail}`);

      return {
        success: true,
        nextAction: 'continue',
        outputVariables: {
          teamsChatSent: true,
          recipientEmail
        }
      };
    } catch (error) {
      logger.error('NotificationActionHandler', 'Error sending Teams chat message', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to send Teams chat message'
      };
    }
  }

  /**
   * Send notification to multiple recipients
   */
  public async sendBulkNotification(
    recipientIds: number[],
    title: string,
    message: string,
    type: string = 'Info',
    context: IActionContext
  ): Promise<IActionResult> {
    try {
      const createdIds: number[] = [];

      for (const recipientId of recipientIds) {
        const notificationData = {
          Title: title,
          Message: message,
          RecipientId: recipientId,
          Type: type,
          Priority: 'Normal',
          IsRead: false,
          RelatedItemType: 'WorkflowInstance',
          RelatedItemId: context.workflowInstance.Id,
          WorkflowInstanceId: context.workflowInstance.Id
        };

        const result = await this.sp.web.lists.getByTitle('PM_Notifications').items.add(notificationData);
        createdIds.push(result.data.Id);
      }

      logger.info('NotificationActionHandler', `Sent ${createdIds.length} bulk notifications`);

      return {
        success: true,
        nextAction: 'continue',
        createdItemIds: createdIds,
        outputVariables: {
          notificationsSent: createdIds.length
        }
      };
    } catch (error) {
      logger.error('NotificationActionHandler', 'Error sending bulk notifications', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to send bulk notifications'
      };
    }
  }

  /**
   * Send workflow event notification using unified service
   * Provides rich HTML templates and multi-channel delivery
   * Extended notification properties should be passed via config.actionConfig
   */
  public async sendWorkflowEventNotification(
    config: IStepConfig,
    context: IActionContext
  ): Promise<IActionResult> {
    try {
      // Access extended config via actionConfig
      const extConfig = (config.actionConfig || {}) as Record<string, unknown>;

      const eventType = (extConfig.workflowEvent as WorkflowNotificationEvent) || WorkflowNotificationEvent.Custom;
      const channel = (extConfig.notificationChannel as NotificationChannel) || NotificationChannel.All;

      // Resolve recipients
      const recipientIds: number[] = [];
      const recipientEmails: string[] = [];

      // From config - handle both number and string IDs
      if (config.recipientId) {
        if (typeof config.recipientId === 'number') {
          recipientIds.push(config.recipientId);
        } else if (typeof config.recipientId === 'string' && config.recipientId.includes('@')) {
          recipientEmails.push(config.recipientId);
        }
      }
      // Handle new recipientEmails from People Picker
      if (config.recipientEmails) {
        recipientEmails.push(...config.recipientEmails);
      }
      if (extConfig.recipientIds) {
        recipientIds.push(...(extConfig.recipientIds as number[]));
      }
      if (extConfig.recipientEmails) {
        recipientEmails.push(...(extConfig.recipientEmails as string[]));
      }

      // From field references
      if (config.recipientField) {
        const evalContext = { ...context.process, ...context.variables };
        const fieldValue = this.conditionEvaluator.evaluateExpression(config.recipientField, evalContext);
        if (typeof fieldValue === 'number') {
          recipientIds.push(fieldValue);
        } else if (typeof fieldValue === 'string' && fieldValue.includes('@')) {
          recipientEmails.push(fieldValue);
        }
      }

      if (recipientIds.length === 0 && recipientEmails.length === 0) {
        return { success: false, error: 'No recipients specified for workflow event notification' };
      }

      // Process custom message if provided
      let customMessage: string | undefined;
      let customSubject: string | undefined;
      if (config.messageTemplate) {
        const evalContext = {
          ...context.process,
          ...context.variables,
          workflowInstance: context.workflowInstance,
          currentStep: context.currentStep
        };
        customMessage = this.conditionEvaluator.replaceTokens(config.messageTemplate, evalContext);
      }
      if (extConfig.subjectTemplate) {
        const evalContext = {
          ...context.process,
          ...context.variables,
          workflowInstance: context.workflowInstance
        };
        customSubject = this.conditionEvaluator.replaceTokens(extConfig.subjectTemplate as string, evalContext);
      }

      // Build additional data from extended config
      const additionalData: Record<string, unknown> = {};
      if (extConfig.taskId) additionalData.taskId = extConfig.taskId;
      if (extConfig.taskTitle) additionalData.taskTitle = extConfig.taskTitle;
      if (extConfig.approvalId) additionalData.approvalId = extConfig.approvalId;
      if (extConfig.approvalTitle) additionalData.approvalTitle = extConfig.approvalTitle;
      if (extConfig.hoursRemaining !== undefined) additionalData.hoursRemaining = extConfig.hoursRemaining;
      if (extConfig.hoursOverdue !== undefined) additionalData.hoursOverdue = extConfig.hoursOverdue;
      if (extConfig.stepName) additionalData.stepName = extConfig.stepName;
      if (extConfig.escalationLevel) additionalData.escalationLevel = extConfig.escalationLevel;
      if (extConfig.reason) additionalData.reason = extConfig.reason;

      // Teams configuration
      const teamsConfig = extConfig.teamsConfig as { teamId?: string; channelId?: string; mentionUsers?: boolean } | undefined;

      // Send via unified notification service
      const result = await this.notificationService.sendNotification({
        event: eventType,
        recipientIds,
        recipientEmails,
        workflowInstance: context.workflowInstance,
        stepConfig: config,
        channel,
        customSubject,
        customMessage,
        additionalData,
        teamsConfig: teamsConfig ? {
          teamId: teamsConfig.teamId,
          channelId: teamsConfig.channelId,
          mentionUsers: teamsConfig.mentionUsers
        } : undefined
      });

      if (!result.success) {
        return {
          success: false,
          error: result.errors?.join('; ') || 'Failed to send workflow event notification'
        };
      }

      logger.info('NotificationActionHandler', `Sent workflow event notification: ${eventType}`);

      return {
        success: true,
        nextAction: 'continue',
        createdItemIds: result.notificationIds,
        outputVariables: {
          notificationsSent: (result.notificationIds?.length || 0) + (result.emailsSent || 0) + (result.teamsMessagesSent || 0),
          emailsSent: result.emailsSent,
          teamsMessagesSent: result.teamsMessagesSent
        }
      };
    } catch (error) {
      logger.error('NotificationActionHandler', 'Error sending workflow event notification', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Failed to send workflow event notification'
      };
    }
  }

  // ============================================================================
  // BATCH NOTIFICATIONS WITH RETRY AND DELIVERY VERIFICATION
  // ============================================================================

  /**
   * Send batch notifications with retry logic and delivery verification
   * Supports multiple recipients and multiple channels with proper error handling
   */
  public async sendBatchNotificationWithRetry(
    recipients: Array<{ id?: number; email?: string; name?: string }>,
    notification: {
      title: string;
      message: string;
      type: string;
      priority?: string;
      linkUrl?: string;
    },
    context: IActionContext,
    options?: {
      channels?: ('inApp' | 'email' | 'teams')[];
      trackDelivery?: boolean;
      failOnFirstError?: boolean;
    }
  ): Promise<IBatchNotificationResult> {
    const channels = options?.channels || ['inApp'];
    const trackDelivery = options?.trackDelivery !== false;
    const deliveryResults: INotificationDeliveryResult[] = [];

    let successCount = 0;
    let failureCount = 0;
    let queuedCount = 0;

    for (const recipient of recipients) {
      for (const channel of channels) {
        let result: INotificationDeliveryResult;

        switch (channel) {
          case 'inApp':
            if (recipient.id) {
              result = await this.sendInAppNotificationWithRetry(
                recipient.id,
                notification,
                context,
                trackDelivery
              );
            } else {
              result = {
                success: false,
                channel: 'inApp',
                deliveryStatus: 'failed',
                attempts: 0,
                error: 'Recipient ID required for in-app notification'
              };
            }
            break;

          case 'email':
            if (recipient.email) {
              result = await this.sendEmailNotificationWithRetry(
                recipient.email,
                notification,
                context,
                trackDelivery
              );
            } else {
              result = {
                success: false,
                channel: 'email',
                deliveryStatus: 'failed',
                attempts: 0,
                error: 'Recipient email required for email notification'
              };
            }
            break;

          case 'teams':
            if (recipient.email) {
              result = await this.sendTeamsNotificationWithRetry(
                recipient.email,
                notification,
                context,
                trackDelivery
              );
            } else {
              result = {
                success: false,
                channel: 'teams',
                deliveryStatus: 'failed',
                attempts: 0,
                error: 'Recipient email required for Teams notification'
              };
            }
            break;

          default:
            result = {
              success: false,
              channel: channel as 'inApp',
              deliveryStatus: 'failed',
              attempts: 0,
              error: `Unknown channel: ${channel}`
            };
        }

        deliveryResults.push(result);

        if (result.success) {
          successCount++;
        } else if (result.deliveryStatus === 'queued') {
          queuedCount++;
        } else {
          failureCount++;
          if (options?.failOnFirstError) {
            break;
          }
        }
      }
    }

    const batchResult: IBatchNotificationResult = {
      totalRecipients: recipients.length,
      successCount,
      failureCount,
      queuedCount,
      deliveryResults
    };

    logger.info('NotificationActionHandler',
      `Batch notification complete: ${successCount} sent, ${failureCount} failed, ${queuedCount} queued`);

    return batchResult;
  }

  /**
   * Send in-app notification with retry and DLQ
   */
  private async sendInAppNotificationWithRetry(
    recipientId: number,
    notification: { title: string; message: string; type: string; priority?: string; linkUrl?: string },
    context: IActionContext,
    trackDelivery: boolean
  ): Promise<INotificationDeliveryResult> {
    const payload = {
      recipientId,
      notification,
      workflowInstanceId: context.workflowInstance.Id,
      stepId: context.currentStep.id
    };

    const result = await retryWithDLQ<number>(
      async () => {
        const notificationData = {
          Title: notification.title,
          Message: notification.message,
          RecipientId: recipientId,
          Type: notification.type,
          Priority: notification.priority || 'Normal',
          IsRead: false,
          LinkUrl: notification.linkUrl,
          RelatedItemType: 'WorkflowInstance',
          RelatedItemId: context.workflowInstance.Id,
          WorkflowInstanceId: context.workflowInstance.Id,
          WorkflowStepId: context.currentStep.id,
          DeliveryStatus: 'Sent',
          SentDate: new Date().toISOString()
        };

        const addResult = await this.sp.web.lists.getByTitle('PM_Notifications').items.add(notificationData);
        return addResult.data.Id;
      },
      'notification-inapp',
      payload,
      NOTIFICATION_RETRY_OPTIONS,
      notificationDLQ,
      {
        source: 'NotificationActionHandler',
        operation: 'sendInAppNotification'
      }
    );

    if (result.success && result.result) {
      // Update delivery status if tracking enabled
      if (trackDelivery) {
        await this.updateDeliveryStatus(result.result, 'Delivered');
      }

      return {
        success: true,
        notificationId: result.result,
        recipientId,
        channel: 'inApp',
        deliveryStatus: 'delivered',
        attempts: result.attempts
      };
    }

    return {
      success: false,
      recipientId,
      channel: 'inApp',
      deliveryStatus: result.deadLetterItemId ? 'queued' : 'failed',
      attempts: result.attempts,
      error: result.error?.message,
      dlqItemId: result.deadLetterItemId
    };
  }

  /**
   * Send email notification with retry and DLQ
   */
  private async sendEmailNotificationWithRetry(
    recipientEmail: string,
    notification: { title: string; message: string; type: string; priority?: string; linkUrl?: string },
    context: IActionContext,
    _trackDelivery: boolean
  ): Promise<INotificationDeliveryResult> {
    const payload = {
      recipientEmail,
      notification,
      workflowInstanceId: context.workflowInstance.Id
    };

    const result = await retryWithDLQ<void>(
      async () => {
        const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');

        const emailMessage = {
          message: {
            subject: notification.title,
            body: {
              contentType: 'HTML',
              content: this.buildSimpleEmailHtml(notification.title, notification.message, notification.linkUrl)
            },
            toRecipients: [{ emailAddress: { address: recipientEmail } }]
          },
          saveToSentItems: false
        };

        await graphClient.api('/me/sendMail').post(emailMessage);
      },
      'notification-email',
      payload,
      NOTIFICATION_RETRY_OPTIONS,
      notificationDLQ,
      {
        source: 'NotificationActionHandler',
        operation: 'sendEmailNotification'
      }
    );

    if (result.success) {
      return {
        success: true,
        recipientEmail,
        channel: 'email',
        deliveryStatus: 'sent', // Email delivery is async, we can only confirm sent
        attempts: result.attempts
      };
    }

    return {
      success: false,
      recipientEmail,
      channel: 'email',
      deliveryStatus: result.deadLetterItemId ? 'queued' : 'failed',
      attempts: result.attempts,
      error: result.error?.message,
      dlqItemId: result.deadLetterItemId
    };
  }

  /**
   * Send Teams notification with retry and DLQ
   */
  private async sendTeamsNotificationWithRetry(
    recipientEmail: string,
    notification: { title: string; message: string; type: string; priority?: string; linkUrl?: string },
    context: IActionContext,
    _trackDelivery: boolean
  ): Promise<INotificationDeliveryResult> {
    const payload = {
      recipientEmail,
      notification,
      workflowInstanceId: context.workflowInstance.Id
    };

    const result = await retryWithDLQ<void>(
      async () => {
        const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');

        // Get user ID
        const user = await graphClient.api(`/users/${recipientEmail}`).select('id').get();

        // Create or get chat
        const chat = {
          chatType: 'oneOnOne',
          members: [
            {
              '@odata.type': '#microsoft.graph.aadUserConversationMember',
              roles: ['owner'],
              'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${this.context.pageContext.user.loginName}`
            },
            {
              '@odata.type': '#microsoft.graph.aadUserConversationMember',
              roles: ['owner'],
              'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${user.id}`
            }
          ]
        };

        const createdChat = await graphClient.api('/chats').post(chat);

        // Send message with link if provided
        let content = `<strong>${notification.title}</strong><br/><br/>${notification.message}`;
        if (notification.linkUrl) {
          content += `<br/><br/><a href="${notification.linkUrl}">View Details</a>`;
        }

        await graphClient.api(`/chats/${createdChat.id}/messages`).post({
          body: { content }
        });
      },
      'notification-teams',
      payload,
      NOTIFICATION_RETRY_OPTIONS,
      notificationDLQ,
      {
        source: 'NotificationActionHandler',
        operation: 'sendTeamsNotification'
      }
    );

    if (result.success) {
      return {
        success: true,
        recipientEmail,
        channel: 'teams',
        deliveryStatus: 'sent',
        attempts: result.attempts
      };
    }

    return {
      success: false,
      recipientEmail,
      channel: 'teams',
      deliveryStatus: result.deadLetterItemId ? 'queued' : 'failed',
      attempts: result.attempts,
      error: result.error?.message,
      dlqItemId: result.deadLetterItemId
    };
  }

  /**
   * Update delivery status for tracking
   */
  private async updateDeliveryStatus(notificationId: number, status: string): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('PM_Notifications').items
        .getById(notificationId)
        .update({
          DeliveryStatus: status,
          DeliveredDate: new Date().toISOString()
        });
    } catch (error) {
      logger.warn('NotificationActionHandler', `Failed to update delivery status for ${notificationId}`, error);
    }
  }

  // ============================================================================
  // BRANDED EMAIL TEMPLATES - Outlook/Gmail Compatible
  // ============================================================================

  /**
   * Company branding configuration for email templates
   * Can be extended to fetch from SharePoint list PM_Config_Settings
   */
  private getBrandingConfig(): {
    companyName: string;
    primaryColor: string;
    accentColor: string;
    logoUrl?: string;
    supportEmail: string;
    portalUrl: string;
  } {
    // TODO: In production, fetch from PM_Config_Settings list
    return {
      companyName: 'JML Employee Lifecycle',
      primaryColor: '#03787C',  // JML Teal
      accentColor: '#0078d4',   // Microsoft Blue
      logoUrl: undefined,       // Set your company logo URL
      supportEmail: 'hr@company.com',
      portalUrl: '/sites/JML'
    };
  }

  /**
   * Build simple email HTML for notifications
   * Uses table-based layout for maximum email client compatibility
   */
  private buildSimpleEmailHtml(title: string, message: string, linkUrl?: string): string {
    const branding = this.getBrandingConfig();

    return `
      <!DOCTYPE html>
      <html lang="en" xmlns="http://www.w3.org/1999/xhtml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
      <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <title>${title}</title>
        <!--[if mso]>
        <style type="text/css">
          table { border-collapse: collapse; }
          .button-link { padding: 12px 24px !important; }
        </style>
        <![endif]-->
      </head>
      <body style="margin: 0; padding: 0; background-color: #f5f5f5; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">
        <!-- Wrapper Table -->
        <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="background-color: #f5f5f5;">
          <tr>
            <td align="center" style="padding: 20px 10px;">
              <!-- Main Container -->
              <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="600" style="max-width: 600px; background-color: #ffffff; border-radius: 8px; overflow: hidden;">

                <!-- Brand Header -->
                <tr>
                  <td style="background-color: ${branding.primaryColor}; padding: 24px 32px; text-align: center;">
                    ${branding.logoUrl ? `<img src="${branding.logoUrl}" alt="${branding.companyName}" style="max-height: 40px; margin-bottom: 12px;">` : ''}
                    <h1 style="margin: 0; color: #ffffff; font-size: 22px; font-weight: 600; line-height: 1.3;">${title}</h1>
                  </td>
                </tr>

                <!-- Content -->
                <tr>
                  <td style="padding: 32px; color: #323130; font-size: 15px; line-height: 1.7;">
                    <p style="margin: 0 0 20px 0;">${message.replace(/\n/g, '<br>')}</p>
                    ${linkUrl ? `
                    <!-- CTA Button -->
                    <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="margin: 24px auto 0 auto;">
                      <tr>
                        <td style="border-radius: 4px; background-color: ${branding.primaryColor};">
                          <a href="${linkUrl}" target="_blank" class="button-link" style="display: inline-block; padding: 12px 28px; color: #ffffff; font-size: 14px; font-weight: 600; text-decoration: none; text-align: center;">
                            View Details &rarr;
                          </a>
                        </td>
                      </tr>
                    </table>
                    ` : ''}
                  </td>
                </tr>

                <!-- Divider -->
                <tr>
                  <td style="padding: 0 32px;">
                    <hr style="border: none; border-top: 1px solid #edebe9; margin: 0;">
                  </td>
                </tr>

                <!-- Footer -->
                <tr>
                  <td style="padding: 24px 32px; background-color: #faf9f8; text-align: center;">
                    <p style="margin: 0 0 8px 0; font-size: 12px; color: #605e5c;">
                      This is an automated notification from <strong>${branding.companyName}</strong>
                    </p>
                    <p style="margin: 0; font-size: 11px; color: #8a8886;">
                      Questions? Contact <a href="mailto:${branding.supportEmail}" style="color: ${branding.primaryColor};">HR Support</a>
                    </p>
                  </td>
                </tr>

              </table>
            </td>
          </tr>
        </table>
      </body>
      </html>
    `;
  }

  /**
   * Get notification delivery statistics
   */
  public getDeliveryStats(): { total: number; byType: Record<string, number> } {
    return notificationDLQ.getStats();
  }

  /**
   * Get failed notifications for retry
   */
  public getFailedNotifications(): Array<{
    id: string;
    operationType: string;
    payload: unknown;
    error: string;
    attempts: number;
  }> {
    return notificationDLQ.getAll();
  }

  /**
   * Retry a failed notification
   */
  public async retryFailedNotification(dlqItemId: string): Promise<IRetryResult<void>> {
    const items = notificationDLQ.getAll();
    const item = items.find(i => i.id === dlqItemId);

    if (!item) {
      return {
        success: false,
        error: new Error(`DLQ item ${dlqItemId} not found`),
        attempts: 0,
        totalDurationMs: 0
      };
    }

    // Update attempt count
    notificationDLQ.updateAttempt(dlqItemId);

    // Retry based on operation type
    const payload = item.payload as Record<string, unknown>;

    const result = await retryWithDLQ<void>(
      async () => {
        if (item.operationType === 'notification-email' && payload.recipientEmail) {
          const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
          const notification = payload.notification as { title: string; message: string };

          await graphClient.api('/me/sendMail').post({
            message: {
              subject: notification.title,
              body: { contentType: 'HTML', content: notification.message },
              toRecipients: [{ emailAddress: { address: payload.recipientEmail as string } }]
            },
            saveToSentItems: false
          });
        } else if (item.operationType === 'notification-inapp' && payload.recipientId) {
          const notification = payload.notification as { title: string; message: string; type: string };
          await this.sp.web.lists.getByTitle('PM_Notifications').items.add({
            Title: notification.title,
            Message: notification.message,
            RecipientId: payload.recipientId,
            Type: notification.type,
            IsRead: false,
            SentDate: new Date().toISOString()
          });
        }
      },
      `${item.operationType}-retry`,
      payload,
      { ...NOTIFICATION_RETRY_OPTIONS, maxRetries: 1 },
      notificationDLQ
    );

    if (result.success) {
      notificationDLQ.remove(dlqItemId);
      logger.info('NotificationActionHandler', `Successfully retried notification ${dlqItemId}`);
    }

    return result;
  }
}
