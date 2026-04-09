/**
 * EmailTemplateBuilder — Premium email notification templates
 * Compact gradient card design (Variation B) with colour-coded headers per notification type.
 *
 * Usage:
 *   import { EmailTemplateBuilder } from '../utils/EmailTemplateBuilder';
 *   const html = EmailTemplateBuilder.policyPublished({ recipientName, policyTitle, ... });
 *
 * All user-controlled content MUST be escaped via escapeHtml() before passing to these methods.
 */

import { escapeHtml } from './sanitizeHtml';

// ============================================================================
// TYPES
// ============================================================================

const F = "'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif";

export type EmailNotificationType =
  | 'policy-published' | 'policy-updated' | 'ack-required' | 'reminder-3day'
  | 'reminder-1day' | 'overdue' | 'ack-complete' | 'review-required'
  | 'approval-request' | 'approval-approved' | 'approval-rejected'
  | 'policy-expiring' | 'policy-retired' | 'sla-breach' | 'welcome';

interface IEmailTemplateConfig {
  gradientStart: string;
  gradientEnd: string;
  notificationLabel: string;
  ctaColor: string;
}

export interface IMetadataRow {
  label: string;
  value: string;
  valueColor?: string;
  valueBold?: boolean;
}

export interface IEmailParams {
  recipientName: string;
  headerTitle: string;
  bodyText: string;
  rows: IMetadataRow[];
  ctaText: string;
  ctaUrl: string;
}

// ============================================================================
// COLOUR CONFIGS PER NOTIFICATION TYPE
// ============================================================================

const TEMPLATE_CONFIGS: Record<EmailNotificationType, IEmailTemplateConfig> = {
  'policy-published':    { gradientStart: '#0d9488', gradientEnd: '#0f766e', notificationLabel: 'New Policy Published',    ctaColor: '#059669' },
  'policy-updated':      { gradientStart: '#2563eb', gradientEnd: '#1d4ed8', notificationLabel: 'Policy Updated',          ctaColor: '#2563eb' },
  'ack-required':        { gradientStart: '#0d9488', gradientEnd: '#0f766e', notificationLabel: 'Acknowledgement Required', ctaColor: '#0d9488' },
  'reminder-3day':       { gradientStart: '#d97706', gradientEnd: '#b45309', notificationLabel: 'Reminder',                ctaColor: '#d97706' },
  'reminder-1day':       { gradientStart: '#ea580c', gradientEnd: '#c2410c', notificationLabel: 'Final Reminder',          ctaColor: '#ea580c' },
  'overdue':             { gradientStart: '#dc2626', gradientEnd: '#991b1b', notificationLabel: 'OVERDUE',                 ctaColor: '#dc2626' },
  'ack-complete':        { gradientStart: '#059669', gradientEnd: '#047857', notificationLabel: 'Acknowledgement Complete', ctaColor: '#059669' },
  'review-required':     { gradientStart: '#0d9488', gradientEnd: '#0f766e', notificationLabel: 'Review Required',         ctaColor: '#0d9488' },
  'approval-request':    { gradientStart: '#d97706', gradientEnd: '#b45309', notificationLabel: 'Approval Request',        ctaColor: '#d97706' },
  'approval-approved':   { gradientStart: '#059669', gradientEnd: '#047857', notificationLabel: 'Approved',                ctaColor: '#059669' },
  'approval-rejected':   { gradientStart: '#dc2626', gradientEnd: '#b91c1c', notificationLabel: 'Revision Required',       ctaColor: '#dc2626' },
  'policy-expiring':     { gradientStart: '#d97706', gradientEnd: '#b45309', notificationLabel: 'Expiry Warning',          ctaColor: '#d97706' },
  'policy-retired':      { gradientStart: '#64748b', gradientEnd: '#475569', notificationLabel: 'Policy Retired',          ctaColor: '#64748b' },
  'sla-breach':          { gradientStart: '#991b1b', gradientEnd: '#7f1d1d', notificationLabel: 'SLA BREACH',              ctaColor: '#991b1b' },
  'welcome':             { gradientStart: '#0d9488', gradientEnd: '#0f766e', notificationLabel: 'Welcome',                 ctaColor: '#0d9488' },
};

// ============================================================================
// CORE BUILDER
// ============================================================================

export class EmailTemplateBuilder {

  /**
   * Build a premium email notification from type + params.
   * All string values in params should already be escaped via escapeHtml().
   */
  public static build(type: EmailNotificationType, params: IEmailParams): string {
    const config = TEMPLATE_CONFIGS[type];
    return EmailTemplateBuilder.renderTemplate(config, params);
  }

  // ============================================================================
  // CONVENIENCE METHODS — one per notification type
  // ============================================================================

  public static policyPublished(p: {
    recipientName: string; policyTitle: string; policyNumber: string;
    publishedBy: string; category: string; department: string;
    riskLevel: string; effectiveDate: string; ctaUrl: string;
  }): string {
    return this.build('policy-published', {
      recipientName: p.recipientName,
      headerTitle: escapeHtml(p.policyTitle),
      bodyText: 'A new policy has been published that requires your attention. Please review the details below and acknowledge within the specified timeframe.',
      rows: [
        { label: 'Policy Number', value: escapeHtml(p.policyNumber), valueBold: true },
        { label: 'Published By', value: escapeHtml(p.publishedBy) },
        { label: 'Category', value: escapeHtml(p.category) },
        { label: 'Department', value: escapeHtml(p.department) },
        { label: 'Risk Level', value: this.riskDisplay(p.riskLevel), valueColor: this.riskColor(p.riskLevel), valueBold: true },
        { label: 'Effective Date', value: escapeHtml(p.effectiveDate) },
        { label: 'Action Required', value: 'Read &amp; Acknowledge within 14 days', valueColor: '#0d9488', valueBold: true },
      ],
      ctaText: 'View Policy',
      ctaUrl: p.ctaUrl,
    });
  }

  public static policyUpdated(p: {
    recipientName: string; policyTitle: string; policyNumber: string;
    updatedBy: string; previousVersion: string; newVersion: string;
    keyChanges: string; ctaUrl: string;
  }): string {
    return this.build('policy-updated', {
      recipientName: p.recipientName,
      headerTitle: escapeHtml(p.policyTitle),
      bodyText: 'A policy you have previously acknowledged has been updated. Please review the changes and re-acknowledge the updated version.',
      rows: [
        { label: 'Policy Number', value: escapeHtml(p.policyNumber), valueBold: true },
        { label: 'Updated By', value: escapeHtml(p.updatedBy) },
        { label: 'Previous Version', value: escapeHtml(p.previousVersion) },
        { label: 'New Version', value: escapeHtml(p.newVersion), valueColor: '#2563eb', valueBold: true },
        { label: 'Key Changes', value: escapeHtml(p.keyChanges) },
        { label: 'Action Required', value: 'Review changes &amp; re-acknowledge', valueColor: '#2563eb', valueBold: true },
      ],
      ctaText: 'Review Changes',
      ctaUrl: p.ctaUrl,
    });
  }

  public static ackRequired(p: {
    recipientName: string; policyTitle: string; policyNumber: string;
    assignedBy: string; category: string; department: string;
    riskLevel: string; dueDate: string; quizRequired: boolean; ctaUrl: string;
  }): string {
    return this.build('ack-required', {
      recipientName: p.recipientName,
      headerTitle: escapeHtml(p.policyTitle),
      bodyText: 'You have been assigned a policy that requires your acknowledgement. Please read the policy carefully and confirm your understanding by the deadline.',
      rows: [
        { label: 'Policy Number', value: escapeHtml(p.policyNumber), valueBold: true },
        { label: 'Assigned By', value: escapeHtml(p.assignedBy) },
        { label: 'Category', value: escapeHtml(p.category) },
        { label: 'Department', value: escapeHtml(p.department) },
        { label: 'Risk Level', value: this.riskDisplay(p.riskLevel), valueColor: this.riskColor(p.riskLevel), valueBold: true },
        { label: 'Due Date', value: escapeHtml(p.dueDate), valueBold: true },
        { label: 'Quiz Required', value: p.quizRequired ? 'Yes — Complete after reading' : 'No' },
      ],
      ctaText: 'Read &amp; Acknowledge',
      ctaUrl: p.ctaUrl,
    });
  }

  public static reminder3Day(p: {
    recipientName: string; policyTitle: string; policyNumber: string;
    category: string; dueDate: string; ctaUrl: string;
  }): string {
    return this.build('reminder-3day', {
      recipientName: p.recipientName,
      headerTitle: escapeHtml(p.policyTitle),
      bodyText: 'This is a friendly reminder that you have 3 days remaining to acknowledge the policy below. Please take action before the deadline.',
      rows: [
        { label: 'Policy Number', value: escapeHtml(p.policyNumber), valueBold: true },
        { label: 'Policy', value: escapeHtml(p.policyTitle) },
        { label: 'Category', value: escapeHtml(p.category) },
        { label: 'Due Date', value: escapeHtml(p.dueDate), valueColor: '#d97706', valueBold: true },
        { label: 'Days Remaining', value: '3 days', valueColor: '#d97706', valueBold: true },
        { label: 'Status', value: 'Pending acknowledgement' },
      ],
      ctaText: 'Acknowledge Now',
      ctaUrl: p.ctaUrl,
    });
  }

  public static reminder1Day(p: {
    recipientName: string; policyTitle: string; policyNumber: string;
    dueDate: string; ctaUrl: string;
  }): string {
    return this.build('reminder-1day', {
      recipientName: p.recipientName,
      headerTitle: escapeHtml(p.policyTitle),
      bodyText: 'URGENT: You have less than 24 hours to acknowledge this policy. Failure to acknowledge may result in escalation to your manager.',
      rows: [
        { label: 'Policy Number', value: escapeHtml(p.policyNumber), valueBold: true },
        { label: 'Policy', value: escapeHtml(p.policyTitle) },
        { label: 'Due Date', value: 'Tomorrow', valueColor: '#ea580c', valueBold: true },
        { label: 'Days Remaining', value: '1 day', valueColor: '#ea580c', valueBold: true },
        { label: 'Escalation', value: 'Will notify your manager if not completed' },
        { label: 'Status', value: 'Action required immediately' },
      ],
      ctaText: 'Acknowledge Now',
      ctaUrl: p.ctaUrl,
    });
  }

  public static overdue(p: {
    recipientName: string; policyTitle: string; policyNumber: string;
    dueDate: string; daysOverdue: number; escalatedTo: string; ctaUrl: string;
  }): string {
    return this.build('overdue', {
      recipientName: p.recipientName,
      headerTitle: escapeHtml(p.policyTitle),
      bodyText: 'Your acknowledgement for this policy is now OVERDUE. This has been flagged to your manager. Please complete your acknowledgement immediately.',
      rows: [
        { label: 'Policy Number', value: escapeHtml(p.policyNumber), valueBold: true },
        { label: 'Policy', value: escapeHtml(p.policyTitle) },
        { label: 'Due Date', value: `${escapeHtml(p.dueDate)} — OVERDUE`, valueColor: '#dc2626', valueBold: true },
        { label: 'Days Overdue', value: `${p.daysOverdue} day${p.daysOverdue !== 1 ? 's' : ''} overdue`, valueColor: '#dc2626', valueBold: true },
        { label: 'Escalated To', value: escapeHtml(p.escalatedTo) },
        { label: 'Compliance Status', value: 'Non-compliant', valueColor: '#dc2626', valueBold: true },
      ],
      ctaText: 'Acknowledge Immediately',
      ctaUrl: p.ctaUrl,
    });
  }

  public static ackComplete(p: {
    recipientName: string; policyTitle: string; policyNumber: string;
    acknowledgedDate: string; category: string; quizResult?: string; ctaUrl: string;
  }): string {
    return this.build('ack-complete', {
      recipientName: p.recipientName,
      headerTitle: escapeHtml(p.policyTitle),
      bodyText: 'Thank you for acknowledging this policy. Your compliance has been recorded. No further action is required.',
      rows: [
        { label: 'Policy Number', value: escapeHtml(p.policyNumber), valueBold: true },
        { label: 'Acknowledged On', value: escapeHtml(p.acknowledgedDate) },
        { label: 'Policy', value: escapeHtml(p.policyTitle) },
        { label: 'Category', value: escapeHtml(p.category) },
        { label: 'Compliance Status', value: '&#10003; Compliant', valueColor: '#059669', valueBold: true },
        ...(p.quizResult ? [{ label: 'Quiz Result', value: escapeHtml(p.quizResult), valueColor: '#059669' }] : []),
      ],
      ctaText: 'View My Policies',
      ctaUrl: p.ctaUrl,
    });
  }

  public static reviewRequired(p: {
    recipientName: string; policyTitle: string; policyNumber: string;
    submittedBy: string; category: string; version: string;
    reviewDeadline: string; ctaUrl: string;
  }): string {
    return this.build('review-required', {
      recipientName: p.recipientName,
      headerTitle: escapeHtml(p.policyTitle),
      bodyText: 'A policy has been submitted for your review. Please review the content and provide your decision (approve, request changes, or reject).',
      rows: [
        { label: 'Policy Number', value: escapeHtml(p.policyNumber), valueBold: true },
        { label: 'Submitted By', value: escapeHtml(p.submittedBy) },
        { label: 'Policy', value: escapeHtml(p.policyTitle) },
        { label: 'Category', value: escapeHtml(p.category) },
        { label: 'Version', value: `${escapeHtml(p.version)} — Draft` },
        { label: 'Review Deadline', value: escapeHtml(p.reviewDeadline) },
      ],
      ctaText: 'Review Policy',
      ctaUrl: p.ctaUrl,
    });
  }

  public static approvalRequest(p: {
    recipientName: string; policyTitle: string; policyNumber: string;
    requestedBy: string; category: string; riskLevel: string;
    reviewerDecision: string; approvalDeadline: string; ctaUrl: string;
  }): string {
    return this.build('approval-request', {
      recipientName: p.recipientName,
      headerTitle: escapeHtml(p.policyTitle),
      bodyText: 'A policy requires your approval before it can be published. Please review the policy and the reviewer feedback, then make your decision.',
      rows: [
        { label: 'Policy Number', value: escapeHtml(p.policyNumber), valueBold: true },
        { label: 'Requested By', value: escapeHtml(p.requestedBy) },
        { label: 'Policy', value: escapeHtml(p.policyTitle) },
        { label: 'Category', value: escapeHtml(p.category) },
        { label: 'Risk Level', value: this.riskDisplay(p.riskLevel), valueColor: this.riskColor(p.riskLevel), valueBold: true },
        { label: 'Reviewer Decision', value: escapeHtml(p.reviewerDecision) },
        { label: 'Approval Deadline', value: escapeHtml(p.approvalDeadline) },
      ],
      ctaText: 'Review &amp; Approve',
      ctaUrl: p.ctaUrl,
    });
  }

  public static approvalApproved(p: {
    recipientName: string; policyTitle: string; policyNumber: string;
    approvedBy: string; decisionDate: string; comments: string; ctaUrl: string;
  }): string {
    return this.build('approval-approved', {
      recipientName: p.recipientName,
      headerTitle: escapeHtml(p.policyTitle),
      bodyText: 'Great news! Your policy has been approved and is now ready for publishing. You can publish it immediately or schedule it for a later date.',
      rows: [
        { label: 'Policy Number', value: escapeHtml(p.policyNumber), valueBold: true },
        { label: 'Approved By', value: escapeHtml(p.approvedBy) },
        { label: 'Decision Date', value: escapeHtml(p.decisionDate) },
        { label: 'Policy', value: escapeHtml(p.policyTitle) },
        { label: 'Comments', value: escapeHtml(p.comments) },
        { label: 'Next Step', value: 'Publish when ready', valueColor: '#059669', valueBold: true },
      ],
      ctaText: 'Go to Dashboard & Publish',
      ctaUrl: p.ctaUrl,
    });
  }

  public static approvalRejected(p: {
    recipientName: string; policyTitle: string; policyNumber: string;
    rejectedBy: string; decisionDate: string; reason: string; ctaUrl: string;
  }): string {
    return this.build('approval-rejected', {
      recipientName: p.recipientName,
      headerTitle: escapeHtml(p.policyTitle),
      bodyText: 'Your policy has been returned for revision. Please review the feedback below and make the necessary changes before resubmitting.',
      rows: [
        { label: 'Policy Number', value: escapeHtml(p.policyNumber), valueBold: true },
        { label: 'Rejected By', value: escapeHtml(p.rejectedBy) },
        { label: 'Decision Date', value: escapeHtml(p.decisionDate) },
        { label: 'Policy', value: escapeHtml(p.policyTitle) },
        { label: 'Reason', value: escapeHtml(p.reason), valueColor: '#dc2626' },
        { label: 'Action Required', value: 'Revise and resubmit for review', valueColor: '#dc2626', valueBold: true },
      ],
      ctaText: 'Edit Policy',
      ctaUrl: p.ctaUrl,
    });
  }

  public static policyExpiring(p: {
    recipientName: string; policyTitle: string; policyNumber: string;
    category: string; currentVersion: string; expiryDate: string;
    daysUntilExpiry: number; ctaUrl: string;
  }): string {
    return this.build('policy-expiring', {
      recipientName: p.recipientName,
      headerTitle: escapeHtml(p.policyTitle),
      bodyText: 'A policy you manage is approaching its expiry date. Please review and update the policy, or extend the expiry date if the content is still current.',
      rows: [
        { label: 'Policy Number', value: escapeHtml(p.policyNumber), valueBold: true },
        { label: 'Policy', value: escapeHtml(p.policyTitle) },
        { label: 'Category', value: escapeHtml(p.category) },
        { label: 'Current Version', value: escapeHtml(p.currentVersion) },
        { label: 'Expiry Date', value: escapeHtml(p.expiryDate), valueColor: '#d97706', valueBold: true },
        { label: 'Days Until Expiry', value: `${p.daysUntilExpiry} day${p.daysUntilExpiry !== 1 ? 's' : ''}`, valueColor: '#d97706', valueBold: true },
        { label: 'Action Required', value: 'Review, update, or extend' },
      ],
      ctaText: 'Review Policy',
      ctaUrl: p.ctaUrl,
    });
  }

  public static policyRetired(p: {
    recipientName: string; policyTitle: string; policyNumber: string;
    retiredBy: string; retirementDate: string; reason: string;
    userStatus: string; ctaUrl: string;
  }): string {
    return this.build('policy-retired', {
      recipientName: p.recipientName,
      headerTitle: escapeHtml(p.policyTitle),
      bodyText: 'A policy has been retired and is no longer in effect. Any outstanding acknowledgement requirements have been cancelled. No further action is required.',
      rows: [
        { label: 'Policy Number', value: escapeHtml(p.policyNumber), valueBold: true },
        { label: 'Retired By', value: escapeHtml(p.retiredBy) },
        { label: 'Retirement Date', value: escapeHtml(p.retirementDate) },
        { label: 'Policy', value: escapeHtml(p.policyTitle) },
        { label: 'Reason', value: escapeHtml(p.reason) },
        { label: 'Your Status', value: escapeHtml(p.userStatus) },
      ],
      ctaText: 'View Policies',
      ctaUrl: p.ctaUrl,
    });
  }

  public static slaBreach(p: {
    recipientName: string; policyTitle: string; policyNumber: string;
    slaTarget: string; actualRate: string; shortfall: string;
    affectedUsers: number; escalatedTo: string; ctaUrl: string;
  }): string {
    return this.build('sla-breach', {
      recipientName: p.recipientName,
      headerTitle: 'Acknowledgement SLA Exceeded',
      bodyText: 'ALERT: A Service Level Agreement has been breached. The acknowledgement completion target for the policy below was not met within the required timeframe.',
      rows: [
        { label: 'Policy Number', value: escapeHtml(p.policyNumber), valueBold: true },
        { label: 'Policy', value: escapeHtml(p.policyTitle) },
        { label: 'SLA Target', value: escapeHtml(p.slaTarget), valueColor: '#991b1b', valueBold: true },
        { label: 'Actual Rate', value: escapeHtml(p.actualRate), valueColor: '#dc2626', valueBold: true },
        { label: 'Shortfall', value: escapeHtml(p.shortfall), valueColor: '#dc2626', valueBold: true },
        { label: 'Affected Users', value: `${p.affectedUsers} employee${p.affectedUsers !== 1 ? 's' : ''}` },
        { label: 'Escalated To', value: escapeHtml(p.escalatedTo) },
      ],
      ctaText: 'View SLA Dashboard',
      ctaUrl: p.ctaUrl,
    });
  }

  public static welcome(p: {
    recipientName: string; role: string; policiesAssigned: number;
    dueDate: string; helpUrl: string; ctaUrl: string;
  }): string {
    return this.build('welcome', {
      recipientName: p.recipientName,
      headerTitle: 'Welcome to Policy Manager',
      bodyText: 'Welcome to the DWx Policy Manager. You have been set up with access to view and acknowledge organisational policies. Here\'s how to get started.',
      rows: [
        { label: 'Your Role', value: escapeHtml(p.role) },
        { label: 'Policies Assigned', value: `${p.policiesAssigned} polic${p.policiesAssigned !== 1 ? 'ies' : 'y'} awaiting acknowledgement` },
        { label: 'Due Date', value: escapeHtml(p.dueDate) },
        { label: 'Help &amp; Support', value: `<a href="${escapeHtml(p.helpUrl)}" style="color:#0d9488; text-decoration:underline;">Help Centre</a>` },
      ],
      ctaText: 'Go to My Policies',
      ctaUrl: p.ctaUrl,
    });
  }

  // ============================================================================
  // PRIVATE HELPERS
  // ============================================================================

  private static riskDisplay(risk: string): string {
    const r = escapeHtml(risk);
    if (['Critical', 'High'].includes(risk)) return `&#9650; ${r}`;
    if (risk === 'Medium') return `&#9679; ${r}`;
    return r;
  }

  private static riskColor(risk: string): string {
    switch (risk) {
      case 'Critical': return '#991b1b';
      case 'High': return '#dc2626';
      case 'Medium': return '#d97706';
      case 'Low': return '#059669';
      case 'Informational': return '#64748b';
      default: return '#334155';
    }
  }

  private static renderMetadataRow(row: IMetadataRow, index: number, isLast: boolean): string {
    const bg = index % 2 === 0 ? '#f8fafc' : '#ffffff';
    const borderBottom = isLast ? '' : 'border-bottom:1px solid #f1f5f9;';
    const valColor = row.valueColor || '#334155';
    const valWeight = row.valueBold ? '600' : '500';
    return `<tr>
  <td width="38%" style="background-color:${bg}; padding:12px 20px; font-family:${F}; font-size:11px; font-weight:600; text-transform:uppercase; letter-spacing:0.8px; color:#64748b; ${borderBottom}">${row.label}</td>
  <td width="62%" style="background-color:${bg}; padding:12px 20px; font-family:${F}; font-size:13px; font-weight:${valWeight}; color:${valColor}; ${borderBottom}">${row.value}</td>
</tr>`;
  }

  private static renderTemplate(config: IEmailTemplateConfig, params: IEmailParams): string {
    const metadataRows = params.rows.map((r, i) => this.renderMetadataRow(r, i, i === params.rows.length - 1)).join('\n');

    return `<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#f1f5f9;">
  <tr>
    <td align="center" style="padding:32px 16px;">
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="600" style="max-width:600px; width:100%; border-radius:12px; overflow:hidden; box-shadow:0 4px 24px rgba(0,0,0,0.08);">
        <tr>
          <td style="background-color:${config.gradientEnd}; background:linear-gradient(135deg, ${config.gradientStart} 0%, ${config.gradientEnd} 100%); padding:0;">
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
              <tr>
                <td style="padding:20px 40px 18px 40px;">
                  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
                    <tr>
                      <td valign="middle" style="font-family:${F};">
                        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="padding-bottom:10px;">
                          <tr>
                            <td style="font-size:11px; font-weight:600; letter-spacing:1.5px; text-transform:uppercase; color:rgba(255,255,255,0.6);">First Digital &bull; DWx Policy Manager</td>
                            <td align="right" style="font-size:11px; font-weight:600; letter-spacing:1px; text-transform:uppercase; color:rgba(255,255,255,0.5);">${config.notificationLabel}</td>
                          </tr>
                        </table>
                        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
                          <tr>
                            <td style="font-size:20px; font-weight:700; color:#ffffff; line-height:1.3; letter-spacing:-0.3px;">${params.headerTitle}</td>
                            <td width="44" valign="middle" align="right">
                              <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                                <tr><td style="width:44px; height:44px; border-radius:50%; background-color:rgba(255,255,255,0.1); font-size:1px; line-height:1px;">&nbsp;</td></tr>
                              </table>
                            </td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td style="background-color:#ffffff; padding:28px 40px 24px 40px;">
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
              <tr>
                <td style="font-family:${F}; font-size:14px; font-weight:400; line-height:1.65; color:#334155; padding-bottom:24px;">Hi <strong style="color:#0f172a;">${escapeHtml(params.recipientName)}</strong>,<br><br>${params.bodyText}</td>
              </tr>
            </table>
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border:1px solid #e2e8f0; border-radius:8px; overflow:hidden;">
${metadataRows}
            </table>
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
              <tr><td style="height:28px; font-size:1px; line-height:1px;">&nbsp;</td></tr>
            </table>
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" align="center">
              <tr>
                <td style="background-color:${config.ctaColor}; border-radius:8px;">
                  <a href="${params.ctaUrl}" target="_blank" style="display:inline-block; padding:14px 48px; font-family:${F}; font-size:14px; font-weight:600; color:#ffffff; text-decoration:none; letter-spacing:0.3px;">${params.ctaText}</a>
                </td>
              </tr>
            </table>
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
              <tr>
                <td align="center" style="padding-top:12px; font-family:${F}; font-size:11px; color:#94a3b8;">
                  Or copy this link: <a href="${params.ctaUrl}" style="color:#64748b; text-decoration:underline; word-break:break-all;">${params.ctaUrl}</a>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td style="background-color:#f8fafc; border-top:1px solid #e2e8f0; padding:20px 40px;">
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
              <tr>
                <td style="font-family:${F}; font-size:11px; color:#94a3b8; line-height:1.6;">First Digital &mdash; DWx Policy Manager<br><span style="color:#cbd5e1;">Policy Governance &amp; Compliance</span></td>
                <td align="right" style="font-family:${F}; font-size:11px;"><a href="#unsubscribe" style="color:#94a3b8; text-decoration:underline;">Unsubscribe</a></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>`;
  }
}
