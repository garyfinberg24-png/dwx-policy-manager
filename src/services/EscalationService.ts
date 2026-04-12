// @ts-nocheck
/**
 * EscalationService — SLA tracking and auto-escalation for policy approvals.
 *
 * Monitors pending approvals, detects SLA breaches, and executes escalation
 * actions (notify, reassign, auto-approve, reject) based on configured rules.
 *
 * Escalation flow:
 *   1. Approval created with SLA deadline (RequestedDate + SLA days)
 *   2. checkAndEscalate() called periodically (or on page load)
 *   3. For each overdue approval: execute escalation action
 *   4. Log escalation to PM_PolicyAuditLog + send notification
 */
import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { logger } from './LoggingService';

export interface IEscalationResult {
  approvalId: number;
  policyId: number;
  policyTitle: string;
  action: 'Notified' | 'Reassigned' | 'AutoApproved' | 'Rejected';
  escalatedTo?: string;
  daysOverdue: number;
}

export interface IEscalationConfig {
  enabled: boolean;
  warningDays: number;      // Days before SLA to send warning (e.g., 1 day before)
  defaultSLADays: number;   // Default SLA if not specified (e.g., 7 days)
  escalationAction: 'Notify' | 'Reassign' | 'AutoApprove' | 'Reject';
  escalateToRole: string;   // Role to escalate to (e.g., 'Admin')
  maxEscalations: number;   // Max number of escalations before auto-action
  notifyAuthor: boolean;    // Notify the policy author of escalation
}

const DEFAULT_CONFIG: IEscalationConfig = {
  enabled: true,
  warningDays: 1,
  defaultSLADays: 7,
  escalationAction: 'Notify',
  escalateToRole: 'Admin',
  maxEscalations: 3,
  notifyAuthor: true
};

export class EscalationService {
  private sp: SPFI;
  private config: IEscalationConfig;

  constructor(sp: SPFI, config?: Partial<IEscalationConfig>) {
    this.sp = sp;
    this.config = { ...DEFAULT_CONFIG, ...config };
  }

  /**
   * Load escalation config from PM_Configuration
   */
  public async loadConfig(): Promise<void> {
    try {
      const items = await this.sp.web.lists.getByTitle('PM_Configuration')
        .items.filter("substringof('Escalation', ConfigKey)")
        .select('ConfigKey', 'ConfigValue').top(20)();

      const configMap: Record<string, string> = {};
      items.forEach((item: any) => { configMap[item.ConfigKey] = item.ConfigValue; });

      if (configMap['Admin.Escalation.Enabled']) this.config.enabled = configMap['Admin.Escalation.Enabled'] === 'true';
      if (configMap['Admin.Escalation.WarningDays']) this.config.warningDays = Number(configMap['Admin.Escalation.WarningDays']) || 1;
      if (configMap['Admin.Escalation.DefaultSLADays']) this.config.defaultSLADays = Number(configMap['Admin.Escalation.DefaultSLADays']) || 7;
      if (configMap['Admin.Escalation.Action']) this.config.escalationAction = configMap['Admin.Escalation.Action'] as any || 'Notify';
      if (configMap['Admin.Escalation.EscalateToRole']) this.config.escalateToRole = configMap['Admin.Escalation.EscalateToRole'] || 'Admin';
      if (configMap['Admin.Escalation.MaxEscalations']) this.config.maxEscalations = Number(configMap['Admin.Escalation.MaxEscalations']) || 3;
    } catch {
      // Use defaults
    }
  }

  /** Cache of recently sent escalation keys to prevent duplicates within a session */
  private static _recentEscalations: Set<string> = new Set();
  private static _recentEscalationsExpiry: number = 0;

  /**
   * Check all pending approvals for SLA breaches and escalate as needed.
   * Call this on page load or periodically.
   * Includes deduplication: skips if an escalation for the same policy
   * was already queued in this session or exists pending in PM_NotificationQueue.
   */
  public async checkAndEscalate(): Promise<IEscalationResult[]> {
    if (!this.config.enabled) return [];

    const results: IEscalationResult[] = [];
    const now = new Date();

    // Reset session dedup cache every 24 hours
    if (Date.now() > EscalationService._recentEscalationsExpiry) {
      EscalationService._recentEscalations.clear();
      EscalationService._recentEscalationsExpiry = Date.now() + 24 * 60 * 60 * 1000;
    }

    // Pre-load existing pending escalations from queue to prevent duplicates
    let existingEscalationPolicyIds: Set<number> = new Set();
    try {
      const existing = await this.sp.web.lists.getByTitle('PM_NotificationQueue')
        .items.filter("NotificationType eq 'Escalation' and QueueStatus eq 'Pending'")
        .select('Id', 'PolicyId')
        .top(500)();
      existingEscalationPolicyIds = new Set(existing.map((e: any) => e.PolicyId).filter(Boolean));
    } catch { /* queue may not exist */ }

    try {
      // Load pending approvals from PM_Approvals
      const pendingApprovals = await this.sp.web.lists.getByTitle('PM_Approvals')
        .items.filter("Status eq 'Pending'")
        .select('Id', 'Title', 'ProcessID', 'RequestedDate', 'DueDate', 'ApproverId', 'ApprovalLevel', 'EscalationCount')
        .top(200)();

      for (const approval of pendingApprovals) {
        const requestedDate = approval.RequestedDate ? new Date(approval.RequestedDate) : null;
        const dueDate = approval.DueDate ? new Date(approval.DueDate) : null;

        // Calculate effective due date
        let effectiveDue = dueDate;
        if (!effectiveDue && requestedDate) {
          effectiveDue = new Date(requestedDate);
          effectiveDue.setDate(effectiveDue.getDate() + this.config.defaultSLADays);
        }
        if (!effectiveDue) continue;

        const daysOverdue = Math.floor((now.getTime() - effectiveDue.getTime()) / (1000 * 60 * 60 * 24));
        const escalationCount = approval.EscalationCount || 0;

        // Check if SLA is breached
        if (daysOverdue > 0) {
          // Dedup: skip if escalation already pending for this policy
          const policyId = approval.ProcessID || 0;
          const dedupKey = `approval-${policyId}`;
          if (existingEscalationPolicyIds.has(policyId) || EscalationService._recentEscalations.has(dedupKey)) {
            continue; // Already has a pending escalation — skip
          }
          // SLA breached — execute escalation
          const result = await this.executeEscalation(approval, daysOverdue, escalationCount);
          if (result) {
            results.push(result);
            EscalationService._recentEscalations.add(dedupKey);
          }
        } else if (daysOverdue >= -this.config.warningDays && escalationCount === 0) {
          // Approaching SLA — send warning
          await this.sendWarningNotification(approval, Math.abs(daysOverdue));
        }
      }

      // Also check PM_PolicyReviewers
      try {
        const pendingReviews = await this.sp.web.lists.getByTitle('PM_PolicyReviewers')
          .items.filter("ReviewStatus eq 'Pending'")
          .select('Id', 'PolicyId', 'ReviewerId', 'ReviewerType', 'AssignedDate')
          .top(200)();

        for (const review of pendingReviews) {
          const assignedDate = review.AssignedDate ? new Date(review.AssignedDate) : null;
          if (!assignedDate) continue;

          const effectiveDue = new Date(assignedDate);
          effectiveDue.setDate(effectiveDue.getDate() + this.config.defaultSLADays);
          const daysOverdue = Math.floor((now.getTime() - effectiveDue.getTime()) / (1000 * 60 * 60 * 24));

          if (daysOverdue > 0) {
            // Dedup: skip if escalation already pending for this policy
            const dedupKey = `review-${review.PolicyId}`;
            if (existingEscalationPolicyIds.has(review.PolicyId) || EscalationService._recentEscalations.has(dedupKey)) {
              continue; // Already has a pending escalation — skip
            }

            // Resolve reviewer email from ReviewerId
            let reviewerEmail = '';
            try {
              if (review.ReviewerId) {
                const user = await this.sp.web.siteUsers.getById(review.ReviewerId)();
                reviewerEmail = user.Email || user.LoginName || '';
              }
            } catch { /* user not found */ }

            await this.sendEscalationNotification(
              review.PolicyId,
              'Review overdue',
              `Review for policy ${review.PolicyId} by ${review.ReviewerType} is ${daysOverdue} days overdue`,
              reviewerEmail
            );
            EscalationService._recentEscalations.add(dedupKey);
          }
        }
      } catch { /* PM_PolicyReviewers may not exist */ }

    } catch (error) {
      logger.error('EscalationService', 'checkAndEscalate failed:', error);
    }

    return results;
  }

  /**
   * Execute escalation action for a specific approval
   */
  private async executeEscalation(approval: any, daysOverdue: number, escalationCount: number): Promise<IEscalationResult | null> {
    const policyId = approval.ProcessID || 0;
    const policyTitle = approval.Title || `Policy ${policyId}`;

    try {
      // Determine action based on escalation count
      let action: IEscalationResult['action'];

      if (escalationCount >= this.config.maxEscalations) {
        // Max escalations reached — take final action
        action = this.config.escalationAction === 'AutoApprove' ? 'AutoApproved' :
                 this.config.escalationAction === 'Reject' ? 'Rejected' : 'Reassigned';
      } else {
        // Intermediate escalation — notify and increment
        action = 'Notified';
      }

      // Execute the action
      switch (action) {
        case 'AutoApproved':
          await this.sp.web.lists.getByTitle('PM_Approvals').items.getById(approval.Id).update({
            Status: 'Approved',
            Comments: `Auto-approved after ${daysOverdue} days overdue (${escalationCount + 1} escalations)`,
            EscalationCount: escalationCount + 1
          });
          // Update policy status
          if (policyId > 0) {
            try {
              await this.sp.web.lists.getByTitle('PM_Policies').items.getById(policyId).update({ PolicyStatus: 'Approved' });
            } catch { /* best-effort */ }
          }
          break;

        case 'Rejected':
          await this.sp.web.lists.getByTitle('PM_Approvals').items.getById(approval.Id).update({
            Status: 'Rejected',
            Comments: `Auto-rejected after ${daysOverdue} days overdue (${escalationCount + 1} escalations, max reached)`,
            EscalationCount: escalationCount + 1
          });
          if (policyId > 0) {
            try {
              await this.sp.web.lists.getByTitle('PM_Policies').items.getById(policyId).update({ PolicyStatus: 'Draft' });
            } catch { /* best-effort */ }
          }
          break;

        case 'Reassigned':
          // Reassign to the escalation role (Admin by default)
          await this.sp.web.lists.getByTitle('PM_Approvals').items.getById(approval.Id).update({
            Comments: `Escalated: reassigned to ${this.config.escalateToRole} after ${daysOverdue} days overdue`,
            EscalationCount: escalationCount + 1
          });
          break;

        case 'Notified':
        default:
          // Just increment the count and send notification
          await this.sp.web.lists.getByTitle('PM_Approvals').items.getById(approval.Id).update({
            EscalationCount: escalationCount + 1
          });
          break;
      }

      // Resolve approver email for notification
      let approverEmail = '';
      try {
        if (approval.RequestedById) {
          const user = await this.sp.web.siteUsers.getById(approval.RequestedById)();
          approverEmail = user.Email || user.LoginName || '';
        }
      } catch { /* user not found */ }

      // Send escalation notification
      await this.sendEscalationNotification(policyId, policyTitle,
        `Approval for "${policyTitle}" is ${daysOverdue} days overdue. Action: ${action}. Escalation #${escalationCount + 1}.`,
        approverEmail);

      // Audit log
      try {
        await this.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
          Title: `Escalation: ${policyTitle}`,
          PolicyId: policyId,
          EntityType: 'Approval',
          EntityId: String(approval.Id),
          AuditAction: 'Escalation',
          ActionDescription: `SLA breach: ${daysOverdue}d overdue. Action: ${action}. Escalation #${escalationCount + 1}/${this.config.maxEscalations}.`,
          ComplianceRelevant: true
        });
      } catch { /* non-critical */ }

      return {
        approvalId: approval.Id,
        policyId,
        policyTitle,
        action,
        daysOverdue,
        escalatedTo: action === 'Reassigned' ? this.config.escalateToRole : undefined
      };
    } catch (error) {
      logger.error('EscalationService', `Failed to escalate approval ${approval.Id}:`, error);
      return null;
    }
  }

  /**
   * Send warning notification (approaching SLA deadline)
   */
  private async sendWarningNotification(approval: any, daysRemaining: number): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle('PM_NotificationQueue').items.add({
        Title: `SLA Warning: ${approval.Title || 'Policy approval'} — ${daysRemaining} day${daysRemaining !== 1 ? 's' : ''} remaining`,
        To: '', // Will be resolved from ApproverId
        Subject: `Reminder: Approval due in ${daysRemaining} day${daysRemaining !== 1 ? 's' : ''} — ${approval.Title}`,
        Message: `<p>The approval for <strong>${approval.Title}</strong> is due in <strong>${daysRemaining} day${daysRemaining !== 1 ? 's' : ''}</strong>.</p><p>Please take action before the SLA deadline.</p>`,
        QueueStatus: 'Pending',
        Priority: 'High',
        NotificationType: 'EscalationWarning',
        Channel: 'Email'
      });
    } catch { /* best-effort */ }
  }

  /**
   * Send escalation notification.
   * IMPORTANT: Only queues email if recipientEmail contains a valid email address.
   * Empty RecipientEmail causes the Logic App to crash with "To Field cannot be null".
   */
  private async sendEscalationNotification(policyId: number, policyTitle: string, message: string, recipientEmail?: string): Promise<void> {
    try {
      // Guard: never write to notification queue without a valid email
      if (!recipientEmail || !recipientEmail.includes('@')) {
        // Log to audit instead — don't crash the email pipeline
        try {
          await this.sp.web.lists.getByTitle('PM_PolicyAuditLog').items.add({
            AuditAction: 'EscalationSkipped',
            EntityType: 'Policy',
            EntityId: policyId,
            ActionDescription: `Escalation notification skipped — no recipient email resolved. ${message}`,
            PerformedBy: 'System',
            ComplianceRelevant: false
          });
        } catch { /* best-effort audit */ }
        return;
      }

      await this.sp.web.lists.getByTitle('PM_NotificationQueue').items.add({
        Title: `Escalation: ${policyTitle}`,
        Subject: `[ESCALATION] Approval overdue: ${policyTitle}`,
        RecipientEmail: recipientEmail,
        Message: `<p>${message}</p><p>Please review and take action immediately.</p>`,
        QueueStatus: 'Pending',
        Priority: 'Urgent',
        NotificationType: 'Escalation',
        Channel: 'Email'
      });
    } catch { /* best-effort */ }
  }

  /**
   * Get escalation summary for dashboard display
   */
  public async getEscalationSummary(): Promise<{
    totalOverdue: number;
    totalWarning: number;
    escalations: Array<{ policyId: number; title: string; daysOverdue: number; escalationCount: number; action: string }>;
  }> {
    try {
      const now = new Date();
      const pending = await this.sp.web.lists.getByTitle('PM_Approvals')
        .items.filter("Status eq 'Pending'")
        .select('Id', 'Title', 'ProcessID', 'RequestedDate', 'DueDate', 'EscalationCount')
        .top(200)();

      let totalOverdue = 0;
      let totalWarning = 0;
      const escalations: any[] = [];

      for (const approval of pending) {
        const requestedDate = approval.RequestedDate ? new Date(approval.RequestedDate) : null;
        const dueDate = approval.DueDate ? new Date(approval.DueDate) : null;
        let effectiveDue = dueDate;
        if (!effectiveDue && requestedDate) {
          effectiveDue = new Date(requestedDate);
          effectiveDue.setDate(effectiveDue.getDate() + this.config.defaultSLADays);
        }
        if (!effectiveDue) continue;

        const daysOverdue = Math.floor((now.getTime() - effectiveDue.getTime()) / (1000 * 60 * 60 * 24));

        if (daysOverdue > 0) {
          totalOverdue++;
          escalations.push({
            policyId: approval.ProcessID,
            title: approval.Title || `Policy ${approval.ProcessID}`,
            daysOverdue,
            escalationCount: approval.EscalationCount || 0,
            action: (approval.EscalationCount || 0) >= this.config.maxEscalations ? this.config.escalationAction : 'Pending'
          });
        } else if (daysOverdue >= -this.config.warningDays) {
          totalWarning++;
        }
      }

      return { totalOverdue, totalWarning, escalations };
    } catch {
      return { totalOverdue: 0, totalWarning: 0, escalations: [] };
    }
  }
}
