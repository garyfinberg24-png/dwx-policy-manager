// @ts-nocheck
/**
 * ReminderScheduleService
 * Manages scheduled reminders for policy revisions, acknowledgement deadlines,
 * review cycles, and expiry warnings. Creates reminders in PM_ReminderSchedule
 * and processes pending ones by queueing notification emails.
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { PM_LISTS } from '../constants/SharePointListNames';
import { escapeHtml } from '../utils/sanitizeHtml';
import { EmailTemplateBuilder } from '../utils/EmailTemplateBuilder';

// ============================================================================
// TYPES
// ============================================================================

export type ReminderType = 'RevisionDue' | 'AcknowledgementOverdue' | 'ReviewCycleDue' | 'ExpiryWarning' | 'CustomReminder';

export interface IReminder {
  Id?: number;
  Title: string;
  PolicyId: number;
  PolicyTitle: string;
  ReminderType: ReminderType;
  ScheduledDate: string;
  ReminderStatus: 'Pending' | 'Sent' | 'Skipped' | 'Failed';
  RecipientType: 'Author' | 'Reviewer' | 'Approver' | 'AllAssigned' | 'Custom';
  RecipientEmail?: string;
  RecipientId?: number;
  IsRecurring: boolean;
  RecurrenceInterval?: 'Daily' | 'Weekly' | 'Monthly' | 'Quarterly' | 'Annual';
  NextOccurrence?: string;
  SentDate?: string;
  FailureReason?: string;
  CreatedByEmail?: string;
}

// ============================================================================
// SERVICE
// ============================================================================

export class ReminderScheduleService {
  private sp: SPFI;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Schedule a revision reminder for a policy.
   * Called when a policy is published with a review frequency.
   */
  public async scheduleRevisionReminder(
    policyId: number,
    policyTitle: string,
    reviewFrequency: string,
    authorEmail: string
  ): Promise<void> {
    const nextDate = this.calculateNextDate(reviewFrequency);
    if (!nextDate) return;

    // Reminder 30 days before due
    const reminderDate = new Date(nextDate);
    reminderDate.setDate(reminderDate.getDate() - 30);

    await this.createReminder({
      Title: `Revision due: ${policyTitle}`,
      PolicyId: policyId,
      PolicyTitle: policyTitle,
      ReminderType: 'RevisionDue',
      ScheduledDate: reminderDate.toISOString(),
      ReminderStatus: 'Pending',
      RecipientType: 'Author',
      RecipientEmail: authorEmail,
      IsRecurring: true,
      RecurrenceInterval: this.frequencyToInterval(reviewFrequency),
      NextOccurrence: nextDate.toISOString(),
      CreatedByEmail: authorEmail
    });
  }

  /**
   * Schedule an expiry warning for a policy.
   * Called when a policy is published with an expiry date.
   */
  public async scheduleExpiryWarning(
    policyId: number,
    policyTitle: string,
    expiryDate: string,
    authorEmail: string
  ): Promise<void> {
    // Warn 30 days before expiry
    const warningDate = new Date(expiryDate);
    warningDate.setDate(warningDate.getDate() - 30);
    if (warningDate <= new Date()) return; // Already past

    await this.createReminder({
      Title: `Expiry warning: ${policyTitle}`,
      PolicyId: policyId,
      PolicyTitle: policyTitle,
      ReminderType: 'ExpiryWarning',
      ScheduledDate: warningDate.toISOString(),
      ReminderStatus: 'Pending',
      RecipientType: 'Author',
      RecipientEmail: authorEmail,
      IsRecurring: false,
      CreatedByEmail: authorEmail
    });
  }

  /**
   * Process all pending reminders that are due.
   * Called from Admin Centre or a scheduled Azure Function.
   */
  public async processPendingReminders(siteUrl: string): Promise<{ sent: number; failed: number }> {
    const now = new Date().toISOString();
    let sent = 0;
    let failed = 0;

    try {
      const pendingItems = await this.sp.web.lists
        .getByTitle(PM_LISTS.REMINDER_SCHEDULE)
        .items.filter(`ReminderStatus eq 'Pending' and ScheduledDate le '${now}'`)
        .select('Id', 'Title', 'PolicyId', 'PolicyTitle', 'ReminderType', 'RecipientEmail', 'RecipientId', 'RecipientType', 'IsRecurring', 'RecurrenceInterval', 'NextOccurrence')
        .orderBy('ScheduledDate')
        .top(50)();

      for (const item of pendingItems) {
        try {
          // Queue notification email
          if (item.RecipientEmail) {
            const policyUrl = `${siteUrl}/SitePages/PolicyDetails.aspx?policyId=${item.PolicyId}`;
            const emailHtml = EmailTemplateBuilder.policyExpiring({
              recipientName: item.RecipientEmail.split('@')[0] || 'Author',
              policyTitle: item.PolicyTitle,
              policyNumber: '',
              category: '',
              currentVersion: '',
              expiryDate: item.ReminderType === 'ExpiryWarning' ? 'Approaching' : 'Review due',
              daysUntilExpiry: 30,
              ctaUrl: policyUrl
            });

            // Write to notification queue (two-step QueueStatus pattern)
            const result = await this.sp.web.lists.getByTitle(PM_LISTS.NOTIFICATION_QUEUE).items.add({
              Title: item.Title,
              RecipientEmail: item.RecipientEmail,
              PolicyId: item.PolicyId,
              PolicyTitle: item.PolicyTitle,
              NotificationType: 'reminder',
              Channel: 'Email',
              Message: emailHtml,
              QueueStatus: 'Pending',
              Priority: 'Normal'
            });
            const newId = result?.data?.Id || result?.data?.id;
            if (newId) {
              try { await this.sp.web.lists.getByTitle(PM_LISTS.NOTIFICATION_QUEUE).items.getById(newId).update({ QueueStatus: 'Pending' }); } catch { /* best-effort */ }
            }
          }

          // Mark as sent
          const updateData: Record<string, unknown> = { ReminderStatus: 'Sent', SentDate: now };

          // Handle recurrence
          if (item.IsRecurring && item.RecurrenceInterval) {
            const nextDate = this.calculateNextFromInterval(item.RecurrenceInterval);
            if (nextDate) {
              // Create next occurrence
              await this.createReminder({
                Title: item.Title,
                PolicyId: item.PolicyId,
                PolicyTitle: item.PolicyTitle,
                ReminderType: item.ReminderType,
                ScheduledDate: nextDate.toISOString(),
                ReminderStatus: 'Pending',
                RecipientType: item.RecipientType || 'Author',
                RecipientEmail: item.RecipientEmail,
                RecipientId: item.RecipientId,
                IsRecurring: true,
                RecurrenceInterval: item.RecurrenceInterval,
                CreatedByEmail: item.RecipientEmail
              });
            }
          }

          await this.sp.web.lists.getByTitle(PM_LISTS.REMINDER_SCHEDULE).items.getById(item.Id).update(updateData);
          sent++;
        } catch (err) {
          failed++;
          try {
            await this.sp.web.lists.getByTitle(PM_LISTS.REMINDER_SCHEDULE).items.getById(item.Id).update({
              ReminderStatus: 'Failed',
              FailureReason: String(err)
            });
          } catch { /* best-effort */ }
        }
      }
    } catch (err) {
      console.error('[ReminderScheduleService] processPendingReminders failed:', err);
    }

    return { sent, failed };
  }

  /**
   * Get all reminders for a specific policy.
   */
  public async getRemindersForPolicy(policyId: number): Promise<IReminder[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(PM_LISTS.REMINDER_SCHEDULE)
        .items.filter(`PolicyId eq ${policyId}`)
        .select('Id', 'Title', 'PolicyId', 'PolicyTitle', 'ReminderType', 'ScheduledDate', 'ReminderStatus', 'RecipientType', 'RecipientEmail', 'IsRecurring', 'RecurrenceInterval', 'SentDate')
        .orderBy('ScheduledDate', false)
        .top(50)();
      return items as IReminder[];
    } catch {
      return [];
    }
  }

  // ============================================================================
  // PRIVATE HELPERS
  // ============================================================================

  private async createReminder(reminder: Omit<IReminder, 'Id'>): Promise<void> {
    await this.sp.web.lists.getByTitle(PM_LISTS.REMINDER_SCHEDULE).items.add(reminder);
  }

  private calculateNextDate(reviewFrequency: string): Date | null {
    const now = new Date();
    switch (reviewFrequency) {
      case 'Monthly': return new Date(now.getFullYear(), now.getMonth() + 1, now.getDate());
      case 'Quarterly': return new Date(now.getFullYear(), now.getMonth() + 3, now.getDate());
      case 'Biannual': return new Date(now.getFullYear(), now.getMonth() + 6, now.getDate());
      case 'Annual': return new Date(now.getFullYear() + 1, now.getMonth(), now.getDate());
      default: return null; // 'None' or unrecognised
    }
  }

  private frequencyToInterval(freq: string): 'Monthly' | 'Quarterly' | 'Annual' | undefined {
    switch (freq) {
      case 'Monthly': return 'Monthly';
      case 'Quarterly': return 'Quarterly';
      case 'Biannual': return 'Quarterly'; // approximate
      case 'Annual': return 'Annual';
      default: return undefined;
    }
  }

  private calculateNextFromInterval(interval: string): Date | null {
    const now = new Date();
    switch (interval) {
      case 'Daily': return new Date(now.getTime() + 86400000);
      case 'Weekly': return new Date(now.getTime() + 7 * 86400000);
      case 'Monthly': return new Date(now.getFullYear(), now.getMonth() + 1, now.getDate());
      case 'Quarterly': return new Date(now.getFullYear(), now.getMonth() + 3, now.getDate());
      case 'Annual': return new Date(now.getFullYear() + 1, now.getMonth(), now.getDate());
      default: return null;
    }
  }

  private getReminderTypeLabel(type: string): string {
    switch (type) {
      case 'RevisionDue': return 'Policy Revision Due';
      case 'AcknowledgementOverdue': return 'Acknowledgement Overdue';
      case 'ReviewCycleDue': return 'Review Cycle Due';
      case 'ExpiryWarning': return 'Policy Expiry Warning';
      case 'CustomReminder': return 'Policy Reminder';
      default: return 'Reminder';
    }
  }
}
