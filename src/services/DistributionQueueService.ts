// @ts-nocheck
/**
 * DistributionQueueService
 *
 * Client-side service for queueing bulk distribution jobs.
 * Jobs are written to PM_DistributionQueue (SharePoint list) and processed
 * by a server-side Azure Function / Logic App that polls the queue.
 *
 * Pattern: Same as EmailQueueService (PM_EmailQueue → Logic App → send emails)
 *
 * Client flow:
 *   1. queueDistribution() — writes ONE list item with all user IDs as JSON
 *   2. getJobStatus() — polls for progress updates
 *   3. getActiveJobs() — shows all in-flight jobs on Distribution page
 *
 * Server flow (Azure Function):
 *   1. Polls PM_DistributionQueue where QueueStatus = 'Queued'
 *   2. Sets QueueStatus = 'Processing', StartedDate = now
 *   3. Parses TargetUserIds JSON, processes in batches of 50
 *   4. Creates PM_PolicyAcknowledgements records
 *   5. Queues email notifications to PM_EmailQueue
 *   6. Updates ProcessedUsers after each batch
 *   7. Sets QueueStatus = 'Completed', CompletedDate = now
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { logger } from './LoggingService';

const QUEUE_LIST = 'PM_DistributionQueue';

export type QueueStatus = 'Queued' | 'Processing' | 'Completed' | 'Failed' | 'Cancelled';
export type JobType = 'Publish' | 'Redistribute' | 'Reminder' | 'Revoke';

export interface IDistributionJob {
  Id: number;
  Title: string;
  PolicyId: number;
  PolicyName: string;
  PolicyVersionNumber: string;
  TargetUserIds: string; // JSON array
  TotalUsers: number;
  ProcessedUsers: number;
  FailedUsers: number;
  QueueStatus: QueueStatus;
  JobType: JobType;
  DueDate?: string;
  SendNotifications: boolean;
  QueuedBy: string;
  QueuedByEmail: string;
  StartedDate?: string;
  CompletedDate?: string;
  ErrorLog?: string;
  Created: string;
  Modified: string;
}

export interface IQueueDistributionRequest {
  policyId: number;
  policyName: string;
  policyVersionNumber: string;
  targetUserIds: number[];
  jobType: JobType;
  dueDate?: Date;
  sendNotifications: boolean;
  queuedBy: string;
  queuedByEmail: string;
}

export interface IDistributionProgress {
  jobId: number;
  status: QueueStatus;
  totalUsers: number;
  processedUsers: number;
  failedUsers: number;
  progressPercent: number;
  startedDate?: Date;
  completedDate?: Date;
  estimatedTimeRemaining?: string;
  errors: string[];
}

export class DistributionQueueService {
  private sp: SPFI;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Queue a distribution job. This is FAST (single SP list item write).
   * The actual processing happens server-side via Azure Function.
   *
   * @returns The created job ID for status polling
   */
  public async queueDistribution(request: IQueueDistributionRequest): Promise<number> {
    try {
      // For very large user lists (>5000), split into chunks to avoid SP Note field limits
      // SP Note fields support up to ~256KB of text
      const userIdsJson = JSON.stringify(request.targetUserIds);

      if (userIdsJson.length > 200000) {
        // Split into multiple jobs
        const chunkSize = 2000;
        const jobIds: number[] = [];
        for (let i = 0; i < request.targetUserIds.length; i += chunkSize) {
          const chunk = request.targetUserIds.slice(i, i + chunkSize);
          const chunkNum = Math.floor(i / chunkSize) + 1;
          const totalChunks = Math.ceil(request.targetUserIds.length / chunkSize);
          const jobId = await this.createQueueItem({
            ...request,
            targetUserIds: chunk,
            policyName: `${request.policyName} (Part ${chunkNum}/${totalChunks})`
          });
          jobIds.push(jobId);
        }
        logger.info('DistributionQueueService', `Queued ${jobIds.length} jobs for ${request.targetUserIds.length} users (chunked)`);
        return jobIds[0]; // Return first job ID
      }

      const jobId = await this.createQueueItem(request);
      logger.info('DistributionQueueService', `Queued distribution job ${jobId}: ${request.policyName} → ${request.targetUserIds.length} users`);
      return jobId;
    } catch (error) {
      logger.error('DistributionQueueService', 'Failed to queue distribution:', error);
      throw new Error('Failed to queue distribution. Please try again.');
    }
  }

  private async createQueueItem(request: IQueueDistributionRequest): Promise<number> {
    const result = await this.sp.web.lists
      .getByTitle(QUEUE_LIST)
      .items.add({
        Title: `${request.jobType}: ${request.policyName}`,
        PolicyId: request.policyId,
        PolicyName: request.policyName,
        PolicyVersionNumber: request.policyVersionNumber,
        TargetUserIds: JSON.stringify(request.targetUserIds),
        TotalUsers: request.targetUserIds.length,
        ProcessedUsers: 0,
        FailedUsers: 0,
        QueueStatus: 'Queued',
        JobType: request.jobType,
        DueDate: request.dueDate ? request.dueDate.toISOString() : null,
        SendNotifications: request.sendNotifications,
        QueuedBy: request.queuedBy,
        QueuedByEmail: request.queuedByEmail,
        ErrorLog: '[]'
      });

    return result.data.Id;
  }

  /**
   * Get progress for a specific job. Use for polling.
   */
  public async getJobStatus(jobId: number): Promise<IDistributionProgress> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(QUEUE_LIST)
        .items.getById(jobId)
        .select('Id', 'QueueStatus', 'TotalUsers', 'ProcessedUsers', 'FailedUsers', 'StartedDate', 'CompletedDate', 'ErrorLog')() as any;

      const errors: string[] = [];
      try {
        const parsed = JSON.parse(item.ErrorLog || '[]');
        if (Array.isArray(parsed)) errors.push(...parsed);
      } catch { /* ignore */ }

      const total = item.TotalUsers || 1;
      const processed = item.ProcessedUsers || 0;
      const progressPercent = Math.round((processed / total) * 100);

      // Estimate time remaining based on processing speed
      let estimatedTimeRemaining: string | undefined;
      if (item.StartedDate && item.QueueStatus === 'Processing' && processed > 0) {
        const elapsed = Date.now() - new Date(item.StartedDate).getTime();
        const msPerUser = elapsed / processed;
        const remaining = (total - processed) * msPerUser;
        const minutesRemaining = Math.ceil(remaining / 60000);
        estimatedTimeRemaining = minutesRemaining <= 1
          ? 'Less than a minute'
          : `About ${minutesRemaining} minutes`;
      }

      return {
        jobId: item.Id,
        status: item.QueueStatus,
        totalUsers: total,
        processedUsers: processed,
        failedUsers: item.FailedUsers || 0,
        progressPercent,
        startedDate: item.StartedDate ? new Date(item.StartedDate) : undefined,
        completedDate: item.CompletedDate ? new Date(item.CompletedDate) : undefined,
        estimatedTimeRemaining,
        errors
      };
    } catch (error) {
      logger.error('DistributionQueueService', `Failed to get job status for ${jobId}:`, error);
      return {
        jobId,
        status: 'Failed',
        totalUsers: 0,
        processedUsers: 0,
        failedUsers: 0,
        progressPercent: 0,
        errors: ['Failed to retrieve job status']
      };
    }
  }

  /**
   * Get all active and recent jobs (for the Distribution dashboard)
   */
  public async getActiveJobs(): Promise<IDistributionJob[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(QUEUE_LIST)
        .items
        .select('Id', 'Title', 'PolicyId', 'PolicyName', 'PolicyVersionNumber', 'TotalUsers', 'ProcessedUsers', 'FailedUsers', 'QueueStatus', 'JobType', 'DueDate', 'SendNotifications', 'QueuedBy', 'QueuedByEmail', 'StartedDate', 'CompletedDate', 'ErrorLog', 'Created', 'Modified')
        .orderBy('Created', false)
        .top(50)();

      return items as IDistributionJob[];
    } catch (error) {
      logger.error('DistributionQueueService', 'Failed to get active jobs:', error);
      return [];
    }
  }

  /**
   * Get jobs for a specific policy
   */
  public async getJobsForPolicy(policyId: number): Promise<IDistributionJob[]> {
    try {
      const safeId = Number(policyId);
      if (!Number.isFinite(safeId) || safeId <= 0) return [];

      const items = await this.sp.web.lists
        .getByTitle(QUEUE_LIST)
        .items
        .filter(`PolicyId eq ${safeId}`)
        .select('Id', 'Title', 'PolicyId', 'PolicyName', 'TotalUsers', 'ProcessedUsers', 'FailedUsers', 'QueueStatus', 'JobType', 'StartedDate', 'CompletedDate', 'Created')
        .orderBy('Created', false)
        .top(10)();

      return items as IDistributionJob[];
    } catch (error) {
      logger.error('DistributionQueueService', `Failed to get jobs for policy ${policyId}:`, error);
      return [];
    }
  }

  /**
   * Cancel a queued job (only if still Queued, not yet Processing)
   */
  public async cancelJob(jobId: number): Promise<boolean> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(QUEUE_LIST)
        .items.getById(jobId)
        .select('QueueStatus')() as any;

      if (item.QueueStatus !== 'Queued') {
        logger.warn('DistributionQueueService', `Cannot cancel job ${jobId} — status is ${item.QueueStatus}`);
        return false;
      }

      await this.sp.web.lists
        .getByTitle(QUEUE_LIST)
        .items.getById(jobId)
        .update({ QueueStatus: 'Cancelled' });

      logger.info('DistributionQueueService', `Cancelled job ${jobId}`);
      return true;
    } catch (error) {
      logger.error('DistributionQueueService', `Failed to cancel job ${jobId}:`, error);
      return false;
    }
  }

  /**
   * Retry a failed job by resetting status to Queued
   */
  public async retryJob(jobId: number): Promise<boolean> {
    try {
      await this.sp.web.lists
        .getByTitle(QUEUE_LIST)
        .items.getById(jobId)
        .update({
          QueueStatus: 'Queued',
          ProcessedUsers: 0,
          FailedUsers: 0,
          StartedDate: null,
          CompletedDate: null,
          ErrorLog: '[]'
        });

      logger.info('DistributionQueueService', `Reset job ${jobId} for retry`);
      return true;
    } catch (error) {
      logger.error('DistributionQueueService', `Failed to retry job ${jobId}:`, error);
      return false;
    }
  }
}
