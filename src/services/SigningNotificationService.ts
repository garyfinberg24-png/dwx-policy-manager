// @ts-nocheck
// Signing Notification Service
// Handles all notifications for the Signing Service module

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users';

import {
  ISigningRequest,
  ISigner,
  SigningNotificationType,
  SigningRequestStatus
} from '../models/ISigning';
import { logger } from './LoggingService';

/**
 * Notification template interface
 */
export interface INotificationTemplate {
  subject: string;
  body: string;
  variables: string[];
}

/**
 * Notification request interface
 */
export interface INotificationRequest {
  to: string[];
  cc?: string[];
  subject: string;
  body: string;
  isHtml?: boolean;
  priority?: 'Low' | 'Normal' | 'High';
}

/**
 * Signing Notification Service
 */
export class SigningNotificationService {
  private sp: SPFI;
  private readonly REQUESTS_LIST = 'PM_SigningRequests';
  private readonly SIGNERS_LIST = 'PM_Signers';
  private readonly NOTIFICATIONS_LIST = 'PM_Notifications';

  // Base URL for signing portal (would be configured in settings)
  private signingPortalUrl: string = '';

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Initialize the service
   */
  public async initialize(signingPortalUrl?: string): Promise<void> {
    if (signingPortalUrl) {
      this.signingPortalUrl = signingPortalUrl;
    } else {
      // Get from site URL
      const web = await this.sp.web.select('Url')();
      this.signingPortalUrl = `${web.Url}/SitePages/SigningService.aspx`;
    }
  }

  // ============================================
  // SIGNATURE REQUEST NOTIFICATIONS
  // ============================================

  /**
   * Send signature request notification to a signer
   */
  public async sendSignatureRequestNotification(requestId: number, signerId: number): Promise<void> {
    try {
      const { request, signer } = await this.getRequestAndSigner(requestId, signerId);

      const signingUrl = this.buildSigningUrl(requestId, signerId);

      const notification: INotificationRequest = {
        to: [signer.SignerEmail],
        subject: request.EmailSubject || `Please sign: ${request.Title}`,
        body: this.buildSignatureRequestEmail(request, signer, signingUrl),
        isHtml: true,
        priority: request.Priority === 'Critical' ? 'High' : 'Normal'
      };

      await this.sendEmailNotification(notification);
      await this.createInAppNotification(signer.SignerUserId, {
        type: SigningNotificationType.SignatureRequested,
        title: `Signature Required: ${request.Title}`,
        message: `You have been requested to sign "${request.Title}". Please review and sign the document.`,
        actionUrl: signingUrl,
        requestId: requestId
      });

      logger.info('SigningNotificationService', `Sent signature request to ${signer.SignerEmail}`);
    } catch (error) {
      logger.error('SigningNotificationService', 'Failed to send signature request notification:', error);
    }
  }

  /**
   * Send reminder notification to a signer
   */
  public async sendReminderNotification(requestId: number, signerId: number): Promise<void> {
    try {
      const { request, signer } = await this.getRequestAndSigner(requestId, signerId);

      const signingUrl = this.buildSigningUrl(requestId, signerId);
      const daysOverdue = this.calculateDaysOverdue(signer.SentDate);

      const notification: INotificationRequest = {
        to: [signer.SignerEmail],
        subject: `Reminder: Please sign "${request.Title}"`,
        body: this.buildReminderEmail(request, signer, signingUrl, daysOverdue),
        isHtml: true,
        priority: daysOverdue > 5 ? 'High' : 'Normal'
      };

      await this.sendEmailNotification(notification);
      await this.createInAppNotification(signer.SignerUserId, {
        type: SigningNotificationType.Reminder,
        title: `Reminder: Signature Required`,
        message: `Please sign "${request.Title}". This request has been pending for ${daysOverdue} day(s).`,
        actionUrl: signingUrl,
        requestId: requestId
      });

      logger.info('SigningNotificationService', `Sent reminder to ${signer.SignerEmail}`);
    } catch (error) {
      logger.error('SigningNotificationService', 'Failed to send reminder notification:', error);
    }
  }

  /**
   * Send completion notification
   */
  public async sendCompletionNotification(requestId: number): Promise<void> {
    try {
      const request = await this.getRequest(requestId);
      const signers = await this.getSigners(requestId);
      const requester = await this.getRequester(request.RequesterId);

      // Notify requester
      const notification: INotificationRequest = {
        to: [requester.Email],
        subject: `Signing Complete: ${request.Title}`,
        body: this.buildCompletionEmail(request, signers),
        isHtml: true
      };

      await this.sendEmailNotification(notification);
      await this.createInAppNotification(request.RequesterId, {
        type: SigningNotificationType.RequestCompleted,
        title: `Signing Complete: ${request.Title}`,
        message: `All signatures have been collected for "${request.Title}".`,
        actionUrl: this.buildRequestUrl(requestId),
        requestId: requestId
      });

      // Notify all signers
      for (const signer of signers) {
        if (signer.SignerEmail !== requester.Email) {
          await this.sendEmailNotification({
            to: [signer.SignerEmail],
            subject: `Signing Complete: ${request.Title}`,
            body: this.buildSignerCompletionEmail(request, signer),
            isHtml: true
          });
        }
      }

      logger.info('SigningNotificationService', `Sent completion notifications for request ${requestId}`);
    } catch (error) {
      logger.error('SigningNotificationService', 'Failed to send completion notification:', error);
    }
  }

  /**
   * Send decline notification
   */
  public async sendDeclineNotification(requestId: number, signerId: number): Promise<void> {
    try {
      const { request, signer } = await this.getRequestAndSigner(requestId, signerId);
      const requester = await this.getRequester(request.RequesterId);

      const notification: INotificationRequest = {
        to: [requester.Email],
        subject: `Signing Declined: ${request.Title}`,
        body: this.buildDeclineEmail(request, signer),
        isHtml: true,
        priority: 'High'
      };

      await this.sendEmailNotification(notification);
      await this.createInAppNotification(request.RequesterId, {
        type: SigningNotificationType.RequestDeclined,
        title: `Signing Declined: ${request.Title}`,
        message: `${signer.SignerName} has declined to sign "${request.Title}".`,
        actionUrl: this.buildRequestUrl(requestId),
        requestId: requestId
      });

      logger.info('SigningNotificationService', `Sent decline notification for request ${requestId}`);
    } catch (error) {
      logger.error('SigningNotificationService', 'Failed to send decline notification:', error);
    }
  }

  /**
   * Send escalation notification
   */
  public async sendEscalationNotification(requestId: number): Promise<void> {
    try {
      const request = await this.getRequest(requestId);
      const pendingSigners = await this.getPendingSigners(requestId);
      const requester = await this.getRequester(request.RequesterId);

      // Notify requester
      const notification: INotificationRequest = {
        to: [requester.Email],
        subject: `Escalation: ${request.Title}`,
        body: this.buildEscalationEmail(request, pendingSigners),
        isHtml: true,
        priority: 'High'
      };

      await this.sendEmailNotification(notification);
      await this.createInAppNotification(request.RequesterId, {
        type: SigningNotificationType.Escalated,
        title: `Escalation: ${request.Title}`,
        message: `"${request.Title}" has been escalated due to pending signatures.`,
        actionUrl: this.buildRequestUrl(requestId),
        requestId: requestId
      });

      // Also send urgent reminders to pending signers
      for (const signer of pendingSigners) {
        await this.sendUrgentReminderNotification(requestId, signer.Id);
      }

      logger.info('SigningNotificationService', `Sent escalation notification for request ${requestId}`);
    } catch (error) {
      logger.error('SigningNotificationService', 'Failed to send escalation notification:', error);
    }
  }

  /**
   * Send requester escalation notification
   */
  public async sendRequesterEscalationNotification(requestId: number): Promise<void> {
    await this.sendEscalationNotification(requestId);
  }

  /**
   * Send expiration notification
   */
  public async sendExpirationNotification(requestId: number): Promise<void> {
    try {
      const request = await this.getRequest(requestId);
      const signers = await this.getSigners(requestId);
      const requester = await this.getRequester(request.RequesterId);

      // Notify requester
      const notification: INotificationRequest = {
        to: [requester.Email],
        subject: `Signing Request Expired: ${request.Title}`,
        body: this.buildExpirationEmail(request, signers),
        isHtml: true
      };

      await this.sendEmailNotification(notification);
      await this.createInAppNotification(request.RequesterId, {
        type: SigningNotificationType.RequestExpired,
        title: `Expired: ${request.Title}`,
        message: `The signing request for "${request.Title}" has expired.`,
        actionUrl: this.buildRequestUrl(requestId),
        requestId: requestId
      });

      logger.info('SigningNotificationService', `Sent expiration notification for request ${requestId}`);
    } catch (error) {
      logger.error('SigningNotificationService', 'Failed to send expiration notification:', error);
    }
  }

  /**
   * Send expiration warning notification
   */
  public async sendExpirationWarningNotification(requestId: number): Promise<void> {
    try {
      const request = await this.getRequest(requestId);
      const pendingSigners = await this.getPendingSigners(requestId);
      const requester = await this.getRequester(request.RequesterId);

      const daysUntilExpiration = this.calculateDaysUntilExpiration(request.ExpirationDate);

      // Notify requester
      const notification: INotificationRequest = {
        to: [requester.Email],
        subject: `Expiring Soon: ${request.Title}`,
        body: this.buildExpirationWarningEmail(request, pendingSigners, daysUntilExpiration),
        isHtml: true,
        priority: 'High'
      };

      await this.sendEmailNotification(notification);

      // Notify pending signers
      for (const signer of pendingSigners) {
        await this.sendEmailNotification({
          to: [signer.SignerEmail],
          subject: `Urgent: Signature Required - Expires in ${daysUntilExpiration} days`,
          body: this.buildSignerExpirationWarningEmail(request, signer, daysUntilExpiration),
          isHtml: true,
          priority: 'High'
        });
      }

      logger.info('SigningNotificationService', `Sent expiration warning for request ${requestId}`);
    } catch (error) {
      logger.error('SigningNotificationService', 'Failed to send expiration warning:', error);
    }
  }

  /**
   * Send delegation notification
   */
  public async sendDelegationNotification(
    requestId: number,
    originalSignerId: number,
    delegateEmail: string,
    delegateName: string
  ): Promise<void> {
    try {
      const { request, signer } = await this.getRequestAndSigner(requestId, originalSignerId);

      // Notify delegate
      const signingUrl = this.buildSigningUrl(requestId, originalSignerId);

      await this.sendEmailNotification({
        to: [delegateEmail],
        subject: `Signature Delegated: ${request.Title}`,
        body: this.buildDelegationEmail(request, signer, delegateName, signingUrl),
        isHtml: true
      });

      // Notify requester
      const requester = await this.getRequester(request.RequesterId);
      await this.sendEmailNotification({
        to: [requester.Email],
        subject: `Signature Delegated: ${request.Title}`,
        body: this.buildDelegationNotifyEmail(request, signer, delegateName),
        isHtml: true
      });

      logger.info('SigningNotificationService', `Sent delegation notification to ${delegateEmail}`);
    } catch (error) {
      logger.error('SigningNotificationService', 'Failed to send delegation notification:', error);
    }
  }

  /**
   * Send urgent reminder (for escalations)
   */
  private async sendUrgentReminderNotification(requestId: number, signerId: number): Promise<void> {
    try {
      const { request, signer } = await this.getRequestAndSigner(requestId, signerId);
      const signingUrl = this.buildSigningUrl(requestId, signerId);

      await this.sendEmailNotification({
        to: [signer.SignerEmail],
        subject: `URGENT: Please sign "${request.Title}" immediately`,
        body: this.buildUrgentReminderEmail(request, signer, signingUrl),
        isHtml: true,
        priority: 'High'
      });
    } catch (error) {
      logger.error('SigningNotificationService', 'Failed to send urgent reminder:', error);
    }
  }

  // ============================================
  // EMAIL TEMPLATES
  // ============================================

  private buildSignatureRequestEmail(request: any, signer: any, signingUrl: string): string {
    return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .header { background: #0078d4; color: white; padding: 20px; text-align: center; border-radius: 8px 8px 0 0; }
    .content { background: #f5f5f5; padding: 20px; border-radius: 0 0 8px 8px; }
    .button { display: inline-block; background: #0078d4; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; margin: 20px 0; }
    .details { background: white; padding: 15px; border-radius: 4px; margin: 15px 0; }
    .footer { text-align: center; color: #666; font-size: 12px; margin-top: 20px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>Signature Required</h1>
    </div>
    <div class="content">
      <p>Hello ${signer.SignerName},</p>

      <p>You have been requested to sign the following document:</p>

      <div class="details">
        <strong>Document:</strong> ${request.Title}<br>
        <strong>Request #:</strong> ${request.RequestNumber}<br>
        ${request.DueDate ? `<strong>Due Date:</strong> ${new Date(request.DueDate).toLocaleDateString()}<br>` : ''}
        ${request.EmailMessage ? `<p><strong>Message:</strong> ${request.EmailMessage}</p>` : ''}
      </div>

      <p style="text-align: center;">
        <a href="${signingUrl}" class="button">Review & Sign Document</a>
      </p>

      <p>If you have any questions, please contact the requester.</p>
    </div>
    <div class="footer">
      <p>This is an automated message from JML Signing Service.</p>
    </div>
  </div>
</body>
</html>`;
  }

  private buildReminderEmail(request: any, signer: any, signingUrl: string, daysOverdue: number): string {
    return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .header { background: #ff8c00; color: white; padding: 20px; text-align: center; border-radius: 8px 8px 0 0; }
    .content { background: #f5f5f5; padding: 20px; border-radius: 0 0 8px 8px; }
    .button { display: inline-block; background: #0078d4; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; margin: 20px 0; }
    .warning { background: #fff4ce; border-left: 4px solid #ff8c00; padding: 10px 15px; margin: 15px 0; }
    .footer { text-align: center; color: #666; font-size: 12px; margin-top: 20px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>Reminder: Signature Required</h1>
    </div>
    <div class="content">
      <p>Hello ${signer.SignerName},</p>

      <div class="warning">
        This signing request has been pending for <strong>${daysOverdue} day(s)</strong>.
      </div>

      <p>Please review and sign the following document:</p>

      <p><strong>${request.Title}</strong><br>
      Request #: ${request.RequestNumber}</p>

      <p style="text-align: center;">
        <a href="${signingUrl}" class="button">Sign Now</a>
      </p>
    </div>
    <div class="footer">
      <p>This is an automated reminder from JML Signing Service.</p>
    </div>
  </div>
</body>
</html>`;
  }

  private buildCompletionEmail(request: any, signers: any[]): string {
    const signerList = signers.map(s =>
      `<li>${s.SignerName} - Signed on ${new Date(s.SignedDate).toLocaleDateString()}</li>`
    ).join('');

    return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .header { background: #107c10; color: white; padding: 20px; text-align: center; border-radius: 8px 8px 0 0; }
    .content { background: #f5f5f5; padding: 20px; border-radius: 0 0 8px 8px; }
    .success { background: #dff6dd; border-left: 4px solid #107c10; padding: 10px 15px; margin: 15px 0; }
    .footer { text-align: center; color: #666; font-size: 12px; margin-top: 20px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>Signing Complete</h1>
    </div>
    <div class="content">
      <div class="success">
        All signatures have been collected for <strong>${request.Title}</strong>.
      </div>

      <p><strong>Request #:</strong> ${request.RequestNumber}</p>
      <p><strong>Completed:</strong> ${new Date().toLocaleDateString()}</p>

      <p><strong>Signers:</strong></p>
      <ul>${signerList}</ul>

      <p>The signed documents and certificate of completion are now available.</p>
    </div>
    <div class="footer">
      <p>This is an automated message from JML Signing Service.</p>
    </div>
  </div>
</body>
</html>`;
  }

  private buildSignerCompletionEmail(request: any, signer: any): string {
    return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .header { background: #107c10; color: white; padding: 20px; text-align: center; border-radius: 8px 8px 0 0; }
    .content { background: #f5f5f5; padding: 20px; border-radius: 0 0 8px 8px; }
    .footer { text-align: center; color: #666; font-size: 12px; margin-top: 20px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>Signing Complete</h1>
    </div>
    <div class="content">
      <p>Hello ${signer.SignerName},</p>

      <p>The signing process for <strong>${request.Title}</strong> has been completed.</p>

      <p>All parties have signed the document. A copy of the signed document will be provided for your records.</p>

      <p>Thank you for your participation.</p>
    </div>
    <div class="footer">
      <p>This is an automated message from JML Signing Service.</p>
    </div>
  </div>
</body>
</html>`;
  }

  private buildDeclineEmail(request: any, signer: any): string {
    return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .header { background: #d13438; color: white; padding: 20px; text-align: center; border-radius: 8px 8px 0 0; }
    .content { background: #f5f5f5; padding: 20px; border-radius: 0 0 8px 8px; }
    .alert { background: #fde7e9; border-left: 4px solid #d13438; padding: 10px 15px; margin: 15px 0; }
    .footer { text-align: center; color: #666; font-size: 12px; margin-top: 20px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>Signing Declined</h1>
    </div>
    <div class="content">
      <div class="alert">
        <strong>${signer.SignerName}</strong> has declined to sign <strong>${request.Title}</strong>.
      </div>

      <p><strong>Request #:</strong> ${request.RequestNumber}</p>
      <p><strong>Declined by:</strong> ${signer.SignerName} (${signer.SignerEmail})</p>
      <p><strong>Reason:</strong> ${signer.DeclineReason || 'No reason provided'}</p>

      <p>Please review and take appropriate action.</p>
    </div>
    <div class="footer">
      <p>This is an automated message from JML Signing Service.</p>
    </div>
  </div>
</body>
</html>`;
  }

  private buildEscalationEmail(request: any, pendingSigners: any[]): string {
    const signerList = pendingSigners.map(s =>
      `<li>${s.SignerName} (${s.SignerEmail})</li>`
    ).join('');

    return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .header { background: #ff8c00; color: white; padding: 20px; text-align: center; border-radius: 8px 8px 0 0; }
    .content { background: #f5f5f5; padding: 20px; border-radius: 0 0 8px 8px; }
    .warning { background: #fff4ce; border-left: 4px solid #ff8c00; padding: 10px 15px; margin: 15px 0; }
    .footer { text-align: center; color: #666; font-size: 12px; margin-top: 20px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>Escalation Notice</h1>
    </div>
    <div class="content">
      <div class="warning">
        Your signing request has been escalated due to pending signatures.
      </div>

      <p><strong>Document:</strong> ${request.Title}</p>
      <p><strong>Request #:</strong> ${request.RequestNumber}</p>

      <p><strong>Pending Signers:</strong></p>
      <ul>${signerList}</ul>

      <p>Urgent reminders have been sent to the pending signers.</p>
    </div>
    <div class="footer">
      <p>This is an automated message from JML Signing Service.</p>
    </div>
  </div>
</body>
</html>`;
  }

  private buildExpirationEmail(request: any, signers: any[]): string {
    const pendingSigners = signers.filter(s => s.Status !== 'Signed');
    const signerList = pendingSigners.map(s =>
      `<li>${s.SignerName} - ${s.Status}</li>`
    ).join('');

    return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .header { background: #797775; color: white; padding: 20px; text-align: center; border-radius: 8px 8px 0 0; }
    .content { background: #f5f5f5; padding: 20px; border-radius: 0 0 8px 8px; }
    .footer { text-align: center; color: #666; font-size: 12px; margin-top: 20px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>Signing Request Expired</h1>
    </div>
    <div class="content">
      <p>The signing request for <strong>${request.Title}</strong> has expired.</p>

      <p><strong>Request #:</strong> ${request.RequestNumber}</p>
      <p><strong>Expired:</strong> ${new Date().toLocaleDateString()}</p>

      ${pendingSigners.length > 0 ? `
      <p><strong>Signers who did not complete:</strong></p>
      <ul>${signerList}</ul>
      ` : ''}

      <p>If you still need signatures, please create a new request.</p>
    </div>
    <div class="footer">
      <p>This is an automated message from JML Signing Service.</p>
    </div>
  </div>
</body>
</html>`;
  }

  private buildExpirationWarningEmail(request: any, pendingSigners: any[], daysUntilExpiration: number): string {
    const signerList = pendingSigners.map(s =>
      `<li>${s.SignerName} (${s.SignerEmail})</li>`
    ).join('');

    return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .header { background: #ff8c00; color: white; padding: 20px; text-align: center; border-radius: 8px 8px 0 0; }
    .content { background: #f5f5f5; padding: 20px; border-radius: 0 0 8px 8px; }
    .warning { background: #fff4ce; border-left: 4px solid #ff8c00; padding: 10px 15px; margin: 15px 0; }
    .footer { text-align: center; color: #666; font-size: 12px; margin-top: 20px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>Expiring Soon</h1>
    </div>
    <div class="content">
      <div class="warning">
        Your signing request will expire in <strong>${daysUntilExpiration} day(s)</strong>.
      </div>

      <p><strong>Document:</strong> ${request.Title}</p>
      <p><strong>Request #:</strong> ${request.RequestNumber}</p>

      <p><strong>Still waiting for:</strong></p>
      <ul>${signerList}</ul>

      <p>Please follow up with the pending signers to ensure completion before expiration.</p>
    </div>
    <div class="footer">
      <p>This is an automated message from JML Signing Service.</p>
    </div>
  </div>
</body>
</html>`;
  }

  private buildSignerExpirationWarningEmail(request: any, signer: any, daysUntilExpiration: number): string {
    const signingUrl = this.buildSigningUrl(request.Id, signer.Id);

    return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .header { background: #d13438; color: white; padding: 20px; text-align: center; border-radius: 8px 8px 0 0; }
    .content { background: #f5f5f5; padding: 20px; border-radius: 0 0 8px 8px; }
    .button { display: inline-block; background: #d13438; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; margin: 20px 0; }
    .alert { background: #fde7e9; border-left: 4px solid #d13438; padding: 10px 15px; margin: 15px 0; }
    .footer { text-align: center; color: #666; font-size: 12px; margin-top: 20px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>URGENT: Signature Required</h1>
    </div>
    <div class="content">
      <p>Hello ${signer.SignerName},</p>

      <div class="alert">
        This signing request will expire in <strong>${daysUntilExpiration} day(s)</strong>.
      </div>

      <p>Please sign <strong>${request.Title}</strong> immediately.</p>

      <p style="text-align: center;">
        <a href="${signingUrl}" class="button">Sign Now</a>
      </p>
    </div>
    <div class="footer">
      <p>This is an automated message from JML Signing Service.</p>
    </div>
  </div>
</body>
</html>`;
  }

  private buildDelegationEmail(request: any, originalSigner: any, delegateName: string, signingUrl: string): string {
    return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .header { background: #0078d4; color: white; padding: 20px; text-align: center; border-radius: 8px 8px 0 0; }
    .content { background: #f5f5f5; padding: 20px; border-radius: 0 0 8px 8px; }
    .button { display: inline-block; background: #0078d4; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; margin: 20px 0; }
    .info { background: #e6f2ff; border-left: 4px solid #0078d4; padding: 10px 15px; margin: 15px 0; }
    .footer { text-align: center; color: #666; font-size: 12px; margin-top: 20px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>Signature Delegated to You</h1>
    </div>
    <div class="content">
      <p>Hello ${delegateName},</p>

      <div class="info">
        <strong>${originalSigner.SignerName}</strong> has delegated a signature request to you.
      </div>

      <p><strong>Document:</strong> ${request.Title}</p>
      <p><strong>Request #:</strong> ${request.RequestNumber}</p>

      <p style="text-align: center;">
        <a href="${signingUrl}" class="button">Review & Sign</a>
      </p>
    </div>
    <div class="footer">
      <p>This is an automated message from JML Signing Service.</p>
    </div>
  </div>
</body>
</html>`;
  }

  private buildDelegationNotifyEmail(request: any, originalSigner: any, delegateName: string): string {
    return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .header { background: #0078d4; color: white; padding: 20px; text-align: center; border-radius: 8px 8px 0 0; }
    .content { background: #f5f5f5; padding: 20px; border-radius: 0 0 8px 8px; }
    .info { background: #e6f2ff; border-left: 4px solid #0078d4; padding: 10px 15px; margin: 15px 0; }
    .footer { text-align: center; color: #666; font-size: 12px; margin-top: 20px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>Signature Delegated</h1>
    </div>
    <div class="content">
      <div class="info">
        <strong>${originalSigner.SignerName}</strong> has delegated their signature to <strong>${delegateName}</strong>.
      </div>

      <p><strong>Document:</strong> ${request.Title}</p>
      <p><strong>Request #:</strong> ${request.RequestNumber}</p>

      <p>The delegate has been notified and can now sign on behalf of ${originalSigner.SignerName}.</p>
    </div>
    <div class="footer">
      <p>This is an automated message from JML Signing Service.</p>
    </div>
  </div>
</body>
</html>`;
  }

  private buildUrgentReminderEmail(request: any, signer: any, signingUrl: string): string {
    return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .header { background: #d13438; color: white; padding: 20px; text-align: center; border-radius: 8px 8px 0 0; }
    .content { background: #f5f5f5; padding: 20px; border-radius: 0 0 8px 8px; }
    .button { display: inline-block; background: #d13438; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; margin: 20px 0; }
    .alert { background: #fde7e9; border-left: 4px solid #d13438; padding: 10px 15px; margin: 15px 0; }
    .footer { text-align: center; color: #666; font-size: 12px; margin-top: 20px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>URGENT: Immediate Action Required</h1>
    </div>
    <div class="content">
      <p>Hello ${signer.SignerName},</p>

      <div class="alert">
        This signing request has been escalated and requires your immediate attention.
      </div>

      <p>Please sign <strong>${request.Title}</strong> as soon as possible.</p>

      <p style="text-align: center;">
        <a href="${signingUrl}" class="button">Sign Now</a>
      </p>
    </div>
    <div class="footer">
      <p>This is an automated message from JML Signing Service.</p>
    </div>
  </div>
</body>
</html>`;
  }

  // ============================================
  // HELPER METHODS
  // ============================================

  private async getRequest(requestId: number): Promise<any> {
    return await this.sp.web.lists
      .getByTitle(this.REQUESTS_LIST)
      .items.getById(requestId)();
  }

  private async getSigners(requestId: number): Promise<any[]> {
    return await this.sp.web.lists
      .getByTitle(this.SIGNERS_LIST)
      .items.filter(`RequestId eq ${requestId}`)();
  }

  private async getPendingSigners(requestId: number): Promise<any[]> {
    return await this.sp.web.lists
      .getByTitle(this.SIGNERS_LIST)
      .items.filter(`RequestId eq ${requestId} and (Status eq 'Sent' or Status eq 'Viewed')`)();
  }

  private async getRequestAndSigner(requestId: number, signerId: number): Promise<{ request: any; signer: any }> {
    const [request, signer] = await Promise.all([
      this.getRequest(requestId),
      this.sp.web.lists.getByTitle(this.SIGNERS_LIST).items.getById(signerId)()
    ]);

    return { request, signer };
  }

  private async getRequester(requesterId: number): Promise<{ Email: string; Title: string }> {
    try {
      const user = await this.sp.web.siteUsers.getById(requesterId)();
      return { Email: user.Email, Title: user.Title };
    } catch {
      return { Email: '', Title: 'Unknown' };
    }
  }

  private buildSigningUrl(requestId: number, signerId: number): string {
    return `${this.signingPortalUrl}?requestId=${requestId}&signerId=${signerId}&action=sign`;
  }

  private buildRequestUrl(requestId: number): string {
    return `${this.signingPortalUrl}?requestId=${requestId}&view=details`;
  }

  private calculateDaysOverdue(sentDate?: Date): number {
    if (!sentDate) return 0;
    const sent = new Date(sentDate);
    const now = new Date();
    return Math.floor((now.getTime() - sent.getTime()) / (1000 * 60 * 60 * 24));
  }

  private calculateDaysUntilExpiration(expirationDate?: Date): number {
    if (!expirationDate) return 0;
    const exp = new Date(expirationDate);
    const now = new Date();
    return Math.max(0, Math.floor((exp.getTime() - now.getTime()) / (1000 * 60 * 60 * 24)));
  }

  /**
   * Send email notification
   * In a real implementation, this would integrate with your email service
   */
  private async sendEmailNotification(notification: INotificationRequest): Promise<void> {
    try {
      // This would integrate with your email service (e.g., Graph API, SMTP, etc.)
      // For now, we'll log and store in a notifications list

      logger.info('SigningNotificationService', `Sending email to: ${notification.to.join(', ')}`);

      // Store notification record
      await this.sp.web.lists
        .getByTitle(this.NOTIFICATIONS_LIST)
        .items.add({
          Title: notification.subject,
          To: notification.to.join(';'),
          CC: notification.cc?.join(';') || '',
          Subject: notification.subject,
          Body: notification.body,
          Priority: notification.priority || 'Normal',
          Status: 'Sent',
          SentDate: new Date().toISOString()
        });
    } catch (error) {
      logger.error('SigningNotificationService', 'Failed to send email:', error);
    }
  }

  /**
   * Create in-app notification
   */
  private async createInAppNotification(
    userId: number | undefined,
    notification: {
      type: SigningNotificationType;
      title: string;
      message: string;
      actionUrl: string;
      requestId: number;
    }
  ): Promise<void> {
    if (!userId) return;

    try {
      await this.sp.web.lists
        .getByTitle(this.NOTIFICATIONS_LIST)
        .items.add({
          Title: notification.title,
          UserId: userId,
          NotificationType: notification.type,
          Message: notification.message,
          ActionUrl: notification.actionUrl,
          RelatedItemId: notification.requestId,
          IsRead: false,
          Created: new Date().toISOString()
        });
    } catch (error) {
      logger.error('SigningNotificationService', 'Failed to create in-app notification:', error);
    }
  }
}

export default SigningNotificationService;
