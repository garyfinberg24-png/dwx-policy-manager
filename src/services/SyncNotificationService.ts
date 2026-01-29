// @ts-nocheck
/**
 * SyncNotificationService
 *
 * Handles email notifications for user sync operations.
 * Sends formatted emails with sync results, statistics, and error details.
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { graphfi, SPFx as GraphSPFx } from '@pnp/graph';
import { GraphFI } from '@pnp/graph';
import '@pnp/graph/users';

import {
  INotificationSettings,
  INotificationRecipient,
  ISyncAnalytics
} from '../models/IUserSyncConfig';
import { ISyncSummary } from '../models/IEntraUser';

/**
 * Service for sending sync notification emails
 */
export class SyncNotificationService {
  private readonly graph: GraphFI;
  private readonly context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
    this.graph = graphfi().using(GraphSPFx(context));
  }

  /**
   * Sends notification email based on sync results
   */
  public async sendSyncNotification(
    summary: ISyncSummary,
    settings: INotificationSettings,
    addedUsers?: string[],
    errorDetails?: string[]
  ): Promise<boolean> {
    if (!settings.enabled || settings.recipients.length === 0) {
      return false;
    }

    // Determine if we should send based on status
    const status = summary.status.toLowerCase();
    const shouldSend =
      (status === 'completed' && settings.onSuccess) ||
      (status === 'completedwitherrors' && settings.onError) ||
      (status === 'failed' && settings.onFailure);

    if (!shouldSend) {
      return false;
    }

    // Filter recipients based on notification preferences
    const notifyType =
      status === 'completed' ? 'success' : status === 'failed' ? 'failure' : 'error';

    const recipients = settings.recipients.filter(r =>
      r.notifyOn.includes(notifyType)
    );

    if (recipients.length === 0) {
      return false;
    }

    // Build email content
    const subject = this.buildSubject(summary, settings.template.subject);
    const body = this.buildEmailBody(summary, settings, addedUsers, errorDetails);

    // Send email
    try {
      await this.sendEmail(recipients, subject, body);
      return true;
    } catch (error) {
      console.error('Failed to send sync notification:', error);
      return false;
    }
  }

  /**
   * Sends a sync report email with analytics
   */
  public async sendSyncReport(
    analytics: ISyncAnalytics,
    recipients: INotificationRecipient[],
    reportType: 'daily' | 'weekly' | 'monthly'
  ): Promise<boolean> {
    if (recipients.length === 0) {
      return false;
    }

    const subject = `JML User Sync ${reportType.charAt(0).toUpperCase() + reportType.slice(1)} Report - ${new Date().toLocaleDateString()}`;
    const body = this.buildReportEmailBody(analytics, reportType);

    try {
      await this.sendEmail(recipients, subject, body);
      return true;
    } catch (error) {
      console.error('Failed to send sync report:', error);
      return false;
    }
  }

  /**
   * Builds the email subject line
   */
  private buildSubject(summary: ISyncSummary, template: string): string {
    const statusText =
      summary.status === 'Completed'
        ? 'Successful'
        : summary.status === 'CompletedWithErrors'
        ? 'Completed with Errors'
        : 'Failed';

    return template
      .replace('{status}', statusText)
      .replace('{date}', new Date().toLocaleDateString())
      .replace('{syncId}', summary.syncId)
      .replace('{added}', summary.added.toString())
      .replace('{updated}', summary.updated.toString())
      .replace('{errors}', summary.errors.toString());
  }

  /**
   * Builds the email body HTML
   */
  private buildEmailBody(
    summary: ISyncSummary,
    settings: INotificationSettings,
    addedUsers?: string[],
    errorDetails?: string[]
  ): string {
    const statusColor =
      summary.status === 'Completed'
        ? '#107c10'
        : summary.status === 'CompletedWithErrors'
        ? '#d83b01'
        : '#a80000';

    const statusText =
      summary.status === 'Completed'
        ? 'Successful'
        : summary.status === 'CompletedWithErrors'
        ? 'Completed with Errors'
        : 'Failed';

    const duration = summary.completedAt
      ? Math.round(
          (summary.completedAt.getTime() - summary.startedAt.getTime()) / 1000
        )
      : 0;

    let html = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; }
    .header { background: linear-gradient(135deg, #0078d4, #106ebe); color: white; padding: 30px; border-radius: 8px 8px 0 0; }
    .header h1 { margin: 0; font-size: 24px; }
    .header p { margin: 5px 0 0; opacity: 0.9; }
    .content { background: #ffffff; border: 1px solid #e1e1e1; border-top: none; padding: 30px; border-radius: 0 0 8px 8px; }
    .status-badge { display: inline-block; padding: 6px 16px; border-radius: 20px; font-weight: 600; font-size: 14px; }
    .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin: 20px 0; }
    .stat-card { background: #f8f9fa; padding: 15px; border-radius: 8px; text-align: center; }
    .stat-value { font-size: 28px; font-weight: 700; color: #0078d4; }
    .stat-label { font-size: 12px; color: #666; margin-top: 4px; }
    .section { margin: 25px 0; }
    .section-title { font-size: 16px; font-weight: 600; color: #333; margin-bottom: 10px; border-bottom: 2px solid #0078d4; padding-bottom: 5px; }
    .user-list { background: #f8f9fa; padding: 15px; border-radius: 8px; max-height: 200px; overflow-y: auto; }
    .user-item { padding: 5px 0; border-bottom: 1px solid #e1e1e1; }
    .user-item:last-child { border-bottom: none; }
    .error-list { background: #fef2f2; padding: 15px; border-radius: 8px; border-left: 4px solid #a80000; }
    .error-item { padding: 5px 0; color: #a80000; font-size: 13px; }
    .footer { margin-top: 30px; padding-top: 20px; border-top: 1px solid #e1e1e1; font-size: 12px; color: #666; }
    .info-row { display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px solid #f0f0f0; }
    .info-label { color: #666; }
    .info-value { font-weight: 500; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>JML User Sync Report</h1>
      <p>Sync ID: ${summary.syncId}</p>
    </div>
    <div class="content">
      <div style="text-align: center; margin-bottom: 20px;">
        <span class="status-badge" style="background-color: ${statusColor}20; color: ${statusColor};">
          ${statusText}
        </span>
      </div>`;

    // Stats section
    if (settings.template.includeStats) {
      html += `
      <div class="stats-grid">
        <div class="stat-card">
          <div class="stat-value">${summary.totalProcessed}</div>
          <div class="stat-label">Processed</div>
        </div>
        <div class="stat-card">
          <div class="stat-value" style="color: #107c10;">${summary.added}</div>
          <div class="stat-label">Added</div>
        </div>
        <div class="stat-card">
          <div class="stat-value" style="color: #0078d4;">${summary.updated}</div>
          <div class="stat-label">Updated</div>
        </div>
        <div class="stat-card">
          <div class="stat-value" style="color: ${summary.errors > 0 ? '#a80000' : '#666'};">${summary.errors}</div>
          <div class="stat-label">Errors</div>
        </div>
      </div>

      <div class="section">
        <div class="section-title">Sync Details</div>
        <div class="info-row">
          <span class="info-label">Started</span>
          <span class="info-value">${summary.startedAt.toLocaleString()}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Completed</span>
          <span class="info-value">${summary.completedAt?.toLocaleString() || 'N/A'}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Duration</span>
          <span class="info-value">${duration} seconds</span>
        </div>
        <div class="info-row">
          <span class="info-label">Deactivated</span>
          <span class="info-value">${summary.deactivated}</span>
        </div>
        <div class="info-row">
          <span class="info-label">Skipped</span>
          <span class="info-value">${summary.skipped}</span>
        </div>
      </div>`;
    }

    // Added users section
    if (settings.template.includeAddedUsers && addedUsers && addedUsers.length > 0) {
      const displayUsers = addedUsers.slice(0, settings.template.maxUsersToList);
      const hasMore = addedUsers.length > settings.template.maxUsersToList;

      html += `
      <div class="section">
        <div class="section-title">New Users Added (${summary.added})</div>
        <div class="user-list">
          ${displayUsers.map(u => `<div class="user-item">${u}</div>`).join('')}
          ${hasMore ? `<div class="user-item" style="color: #666; font-style: italic;">... and ${addedUsers.length - settings.template.maxUsersToList} more</div>` : ''}
        </div>
      </div>`;
    }

    // Errors section
    if (settings.template.includeErrors && errorDetails && errorDetails.length > 0) {
      const displayErrors = errorDetails.slice(0, 10);
      const hasMoreErrors = errorDetails.length > 10;

      html += `
      <div class="section">
        <div class="section-title">Errors (${summary.errors})</div>
        <div class="error-list">
          ${displayErrors.map(e => `<div class="error-item">${e}</div>`).join('')}
          ${hasMoreErrors ? `<div class="error-item" style="color: #666;">... and ${errorDetails.length - 10} more errors</div>` : ''}
        </div>
      </div>`;
    }

    // Footer
    html += `
      <div class="footer">
        <p>${settings.template.footerText}</p>
        <p>
          <a href="${this.context.pageContext.web.absoluteUrl}/Lists/PM_Sync_Log" style="color: #0078d4;">
            View Full Sync History
          </a>
        </p>
      </div>
    </div>
  </div>
</body>
</html>`;

    return html;
  }

  /**
   * Builds the analytics report email body
   */
  private buildReportEmailBody(
    analytics: ISyncAnalytics,
    reportType: 'daily' | 'weekly' | 'monthly'
  ): string {
    const periodText = reportType.charAt(0).toUpperCase() + reportType.slice(1);

    return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.6; color: #333; }
    .container { max-width: 700px; margin: 0 auto; padding: 20px; }
    .header { background: linear-gradient(135deg, #0078d4, #106ebe); color: white; padding: 30px; border-radius: 8px 8px 0 0; }
    .content { background: #ffffff; border: 1px solid #e1e1e1; border-top: none; padding: 30px; border-radius: 0 0 8px 8px; }
    .stats-row { display: flex; gap: 15px; margin: 20px 0; flex-wrap: wrap; }
    .stat-box { flex: 1; min-width: 120px; background: #f8f9fa; padding: 20px; border-radius: 8px; text-align: center; }
    .stat-value { font-size: 32px; font-weight: 700; color: #0078d4; }
    .stat-label { font-size: 12px; color: #666; }
    .section { margin: 25px 0; }
    .section-title { font-size: 18px; font-weight: 600; color: #333; margin-bottom: 15px; }
    table { width: 100%; border-collapse: collapse; }
    th, td { padding: 12px; text-align: left; border-bottom: 1px solid #e1e1e1; }
    th { background: #f8f9fa; font-weight: 600; }
    .success-rate { font-size: 48px; font-weight: 700; }
    .trend-up { color: #107c10; }
    .trend-down { color: #a80000; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>JML User Sync ${periodText} Report</h1>
      <p>${new Date().toLocaleDateString()}</p>
    </div>
    <div class="content">
      <div class="section">
        <div class="section-title">Summary</div>
        <div class="stats-row">
          <div class="stat-box">
            <div class="stat-value">${analytics.summary.totalSyncs}</div>
            <div class="stat-label">Total Syncs</div>
          </div>
          <div class="stat-box">
            <div class="stat-value">${analytics.summary.totalUsersSynced}</div>
            <div class="stat-label">Users Synced</div>
          </div>
          <div class="stat-box">
            <div class="success-rate" style="color: ${analytics.summary.successRate >= 95 ? '#107c10' : analytics.summary.successRate >= 80 ? '#d83b01' : '#a80000'};">
              ${analytics.summary.successRate}%
            </div>
            <div class="stat-label">Success Rate</div>
          </div>
          <div class="stat-box">
            <div class="stat-value">${analytics.summary.avgDuration}s</div>
            <div class="stat-label">Avg Duration</div>
          </div>
        </div>
      </div>

      <div class="section">
        <div class="section-title">This ${periodText}</div>
        <div class="stats-row">
          <div class="stat-box">
            <div class="stat-value" style="color: #107c10;">${analytics.summary.usersAddedThisMonth}</div>
            <div class="stat-label">Users Added</div>
          </div>
          <div class="stat-box">
            <div class="stat-value" style="color: #0078d4;">${analytics.summary.usersUpdatedThisMonth}</div>
            <div class="stat-label">Users Updated</div>
          </div>
        </div>
      </div>

      <div class="section">
        <div class="section-title">Department Breakdown</div>
        <table>
          <thead>
            <tr>
              <th>Department</th>
              <th>Total</th>
              <th>Active</th>
              <th>Coverage</th>
            </tr>
          </thead>
          <tbody>
            ${analytics.byDepartment
              .slice(0, 10)
              .map(
                d => `
              <tr>
                <td>${d.department}</td>
                <td>${d.totalEmployees}</td>
                <td>${d.activeEmployees}</td>
                <td>${d.syncCoverage}%</td>
              </tr>
            `
              )
              .join('')}
          </tbody>
        </table>
      </div>

      ${
        analytics.errorAnalysis.totalErrors > 0
          ? `
      <div class="section">
        <div class="section-title">Error Summary</div>
        <p>Total errors this period: <strong>${analytics.errorAnalysis.totalErrors}</strong></p>
        <p>Trend: <span class="${analytics.errorAnalysis.trend === 'decreasing' ? 'trend-up' : 'trend-down'}">
          ${analytics.errorAnalysis.trend === 'decreasing' ? '↓ Decreasing' : analytics.errorAnalysis.trend === 'increasing' ? '↑ Increasing' : '→ Stable'}
        </span></p>
      </div>
      `
          : ''
      }

      <div class="section">
        <div class="section-title">Recent Sync Operations</div>
        <table>
          <thead>
            <tr>
              <th>Date</th>
              <th>Type</th>
              <th>Status</th>
              <th>Added</th>
              <th>Updated</th>
            </tr>
          </thead>
          <tbody>
            ${analytics.recentSyncs
              .slice(0, 5)
              .map(
                s => `
              <tr>
                <td>${new Date(s.timestamp).toLocaleDateString()}</td>
                <td>${s.type}</td>
                <td>${s.status}</td>
                <td>${s.added}</td>
                <td>${s.updated}</td>
              </tr>
            `
              )
              .join('')}
          </tbody>
        </table>
      </div>
    </div>
  </div>
</body>
</html>`;
  }

  /**
   * Sends email using Microsoft Graph
   */
  private async sendEmail(
    recipients: INotificationRecipient[],
    subject: string,
    body: string
  ): Promise<void> {
    const message = {
      subject,
      body: {
        contentType: 'HTML',
        content: body
      },
      toRecipients: recipients.map(r => ({
        emailAddress: {
          address: r.email,
          name: r.name
        }
      }))
    };

    // Use raw fetch since PnPjs Graph doesn't expose sendMail directly
    const tokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
    const accessToken = await tokenProvider.getToken('https://graph.microsoft.com');

    const response = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ message, saveToSentItems: true })
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to send email: ${response.status} ${errorText}`);
    }
  }
}

export default SyncNotificationService;
