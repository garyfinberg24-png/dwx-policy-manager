// @ts-nocheck
/**
 * DistributionQueueViewService — Reads PM_DistributionQueue and PM_NotificationQueue
 * for visual display of distribution + email pipeline status.
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

// ============================================================================
// TYPES
// ============================================================================

export interface IDistributionJob {
  Id: number;
  Title: string;
  PolicyId: number;
  PolicyName: string;
  JobType: string;
  Status: string;
  TargetCount: number;
  ProcessedCount: number;
  QueuedBy: string;
  QueuedDate: string;
  CompletedDate?: string;
  ErrorMessage?: string;
}

export interface INotificationQueueItem {
  Id: number;
  Title: string;
  To: string;
  RecipientEmail: string;
  Subject: string;
  QueueStatus: string;
  Status: string;
  Priority: string;
  NotificationType: string;
  Channel: string;
  PolicyId: number;
  PolicyTitle: string;
  Created: string;
  RetryCount: number;
  LastError: string;
}

export interface IQueueSummary {
  distributionJobs: IDistributionJob[];
  notificationItems: INotificationQueueItem[];
  stats: {
    totalJobs: number;
    activeJobs: number;
    completedJobs: number;
    failedJobs: number;
    totalEmails: number;
    pendingEmails: number;
    sentEmails: number;
    failedEmails: number;
  };
  loadedAt: string;
}

// ============================================================================
// SERVICE
// ============================================================================

export class DistributionQueueViewService {
  private _sp: SPFI;

  constructor(sp: SPFI) {
    this._sp = sp;
  }

  /**
   * Load distribution queue and notification queue data for display.
   */
  public async loadQueueData(): Promise<IQueueSummary> {
    const distributionJobs = await this._loadDistributionJobs();
    const notificationItems = await this._loadNotificationQueue();

    const activeJobs = distributionJobs.filter(j => j.Status === 'Queued' || j.Status === 'Processing');
    const completedJobs = distributionJobs.filter(j => j.Status === 'Completed');
    const failedJobs = distributionJobs.filter(j => j.Status === 'Failed');

    const pendingEmails = notificationItems.filter(i => (i.QueueStatus || i.Status) === 'Pending');
    const sentEmails = notificationItems.filter(i => (i.QueueStatus || i.Status) === 'Sent');
    const failedEmails = notificationItems.filter(i => (i.QueueStatus || i.Status) === 'Failed');

    return {
      distributionJobs,
      notificationItems,
      stats: {
        totalJobs: distributionJobs.length,
        activeJobs: activeJobs.length,
        completedJobs: completedJobs.length,
        failedJobs: failedJobs.length,
        totalEmails: notificationItems.length,
        pendingEmails: pendingEmails.length,
        sentEmails: sentEmails.length,
        failedEmails: failedEmails.length,
      },
      loadedAt: new Date().toISOString(),
    };
  }

  private async _loadDistributionJobs(): Promise<IDistributionJob[]> {
    try {
      const items = await this._sp.web.lists.getByTitle('PM_DistributionQueue')
        .items
        .select('Id', 'Title', 'PolicyId', 'PolicyName', 'JobType', 'Status', 'TargetCount', 'ProcessedCount', 'QueuedBy', 'Created', 'CompletedDate', 'ErrorMessage')
        .orderBy('Created', false)
        .top(50)();

      return items.map((item: any) => ({
        Id: item.Id,
        Title: item.Title || '',
        PolicyId: item.PolicyId || 0,
        PolicyName: item.PolicyName || item.Title || '',
        JobType: item.JobType || 'Publish',
        Status: item.Status || 'Queued',
        TargetCount: item.TargetCount || 0,
        ProcessedCount: item.ProcessedCount || 0,
        QueuedBy: item.QueuedBy || '',
        QueuedDate: item.Created || '',
        CompletedDate: item.CompletedDate || '',
        ErrorMessage: item.ErrorMessage || '',
      }));
    } catch {
      return [];
    }
  }

  private async _loadNotificationQueue(): Promise<INotificationQueueItem[]> {
    try {
      const items = await this._sp.web.lists.getByTitle('PM_NotificationQueue')
        .items
        .select('Id', 'Title', 'To', 'RecipientEmail', 'Subject', 'QueueStatus', 'Status', 'Priority', 'NotificationType', 'Channel', 'PolicyId', 'PolicyTitle', 'Created', 'RetryCount', 'LastError')
        .orderBy('Created', false)
        .top(100)();

      return items.map((item: any) => ({
        Id: item.Id,
        Title: item.Title || '',
        To: item.To || item.RecipientEmail || '',
        RecipientEmail: item.RecipientEmail || item.To || '',
        Subject: item.Subject || item.Title || '',
        QueueStatus: item.QueueStatus || item.Status || 'Pending',
        Status: item.Status || item.QueueStatus || 'Pending',
        Priority: item.Priority || 'Normal',
        NotificationType: item.NotificationType || '',
        Channel: item.Channel || 'Email',
        PolicyId: item.PolicyId || 0,
        PolicyTitle: item.PolicyTitle || '',
        Created: item.Created || '',
        RetryCount: item.RetryCount || 0,
        LastError: item.LastError || '',
      }));
    } catch {
      return [];
    }
  }
}
