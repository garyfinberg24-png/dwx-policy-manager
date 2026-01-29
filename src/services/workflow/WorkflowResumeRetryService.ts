// @ts-nocheck
/* eslint-disable */
/**
 * WorkflowResumeRetryService
 * Automatic retry mechanism for failed workflow resume operations
 * CRITICAL: Ensures workflows are never permanently stuck due to transient failures
 */

import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import { logger } from '../LoggingService';
import { WorkflowInstanceStatus } from '../../models/IWorkflow';
import { WorkflowEngineService } from './WorkflowEngineService';
import { WorkflowInstanceService } from './WorkflowInstanceService';
import {
  retryOperation,
  IRetryOptions,
  workflowSyncDLQ,
  DeadLetterQueue
} from '../../utils/retryUtils';
import { PersistentDeadLetterQueueService } from './PersistentDeadLetterQueueService';

/**
 * Failed resume operation stored for retry
 */
export interface IFailedResumeOperation {
  id: string;
  workflowInstanceId: number;
  stepId?: string;
  taskId?: number;
  approvalId?: number;
  operationType: 'task' | 'approval' | 'manual';
  failedAt: Date;
  lastRetryAt?: Date;
  retryCount: number;
  lastError: string;
  completionData?: Record<string, unknown>;
}

/**
 * Retry configuration
 */
export interface IResumeRetryConfig {
  enabled: boolean;
  maxRetries: number;
  initialDelayMs: number;
  maxDelayMs: number;
  retryIntervalMs: number;
  maxConcurrentRetries: number;
  abandonAfterMs: number; // Auto-abandon after this time
}

/**
 * Retry result
 */
export interface IResumeRetryResult {
  operationId: string;
  success: boolean;
  retryCount: number;
  workflowInstanceId: number;
  newWorkflowStatus?: WorkflowInstanceStatus;
  error?: string;
  abandoned?: boolean;
}

/**
 * Batch retry result
 */
export interface IBatchRetryResult {
  processedAt: Date;
  totalOperations: number;
  succeeded: number;
  failed: number;
  abandoned: number;
  results: IResumeRetryResult[];
}

export class WorkflowResumeRetryService {
  private sp: SPFI;
  private context: WebPartContext;
  private workflowEngine: WorkflowEngineService;
  private instanceService: WorkflowInstanceService;
  private persistentDLQ: PersistentDeadLetterQueueService;
  private config: IResumeRetryConfig;
  private retryTimer: ReturnType<typeof setInterval> | null = null;
  private isProcessing: boolean = false;
  private readonly LIST_NAME = 'PM_WorkflowResumeRetry';

  constructor(
    sp: SPFI,
    context: WebPartContext,
    config?: Partial<IResumeRetryConfig>
  ) {
    this.sp = sp;
    this.context = context;
    this.workflowEngine = new WorkflowEngineService(sp, context);
    this.instanceService = new WorkflowInstanceService(sp);
    this.persistentDLQ = new PersistentDeadLetterQueueService(sp, workflowSyncDLQ);

    this.config = {
      enabled: true,
      maxRetries: 5,
      initialDelayMs: 5000,      // 5 seconds
      maxDelayMs: 300000,        // 5 minutes
      retryIntervalMs: 60000,    // 1 minute
      maxConcurrentRetries: 3,
      abandonAfterMs: 86400000,  // 24 hours
      ...config
    };
  }

  /**
   * Initialize the service
   */
  public async initialize(): Promise<void> {
    await this.persistentDLQ.initialize();
    logger.info('WorkflowResumeRetryService', 'Service initialized');
  }

  /**
   * Start automatic retry processing
   */
  public start(): void {
    if (!this.config.enabled) {
      logger.info('WorkflowResumeRetryService', 'Retry service is disabled');
      return;
    }

    if (this.retryTimer) {
      this.stop();
    }

    logger.info('WorkflowResumeRetryService',
      `Starting retry service (interval: ${this.config.retryIntervalMs}ms)`);

    // Run immediately
    this.processFailedResumes().catch(err => {
      logger.error('WorkflowResumeRetryService', 'Error in initial retry run', err);
    });

    // Schedule periodic runs
    this.retryTimer = setInterval(() => {
      this.processFailedResumes().catch(err => {
        logger.error('WorkflowResumeRetryService', 'Error in scheduled retry run', err);
      });
    }, this.config.retryIntervalMs);
  }

  /**
   * Stop automatic retry processing
   */
  public stop(): void {
    if (this.retryTimer) {
      clearInterval(this.retryTimer);
      this.retryTimer = null;
      logger.info('WorkflowResumeRetryService', 'Retry service stopped');
    }
  }

  /**
   * Queue a failed resume operation for retry
   */
  public async queueForRetry(
    workflowInstanceId: number,
    operationType: 'task' | 'approval' | 'manual',
    error: Error | string,
    options?: {
      stepId?: string;
      taskId?: number;
      approvalId?: number;
      completionData?: Record<string, unknown>;
    }
  ): Promise<string> {
    const errorMessage = typeof error === 'string' ? error : error.message;

    const operationId = await this.persistentDLQ.enqueue(
      `workflow-resume-${operationType}`,
      {
        workflowInstanceId,
        stepId: options?.stepId,
        taskId: options?.taskId,
        approvalId: options?.approvalId,
        completionData: options?.completionData
      },
      errorMessage,
      1, // Initial attempt count
      {
        operationType,
        workflowInstanceId,
        queuedAt: new Date().toISOString()
      }
    );

    logger.info('WorkflowResumeRetryService',
      `Queued workflow ${workflowInstanceId} resume for retry (ID: ${operationId})`);

    return operationId;
  }

  /**
   * Process all failed resume operations
   */
  public async processFailedResumes(): Promise<IBatchRetryResult> {
    if (this.isProcessing) {
      logger.info('WorkflowResumeRetryService', 'Already processing - skipping');
      return {
        processedAt: new Date(),
        totalOperations: 0,
        succeeded: 0,
        failed: 0,
        abandoned: 0,
        results: []
      };
    }

    this.isProcessing = true;
    const results: IResumeRetryResult[] = [];

    try {
      // Get pending items from DLQ
      const pendingItems = await this.persistentDLQ.getPendingItems();

      // Filter to only workflow resume operations
      const resumeOperations = pendingItems.filter(item =>
        item.operationType.startsWith('workflow-resume-')
      );

      if (resumeOperations.length === 0) {
        return {
          processedAt: new Date(),
          totalOperations: 0,
          succeeded: 0,
          failed: 0,
          abandoned: 0,
          results: []
        };
      }

      logger.info('WorkflowResumeRetryService',
        `Processing ${resumeOperations.length} failed resume operations`);

      // Process up to maxConcurrentRetries at a time
      const toProcess = resumeOperations.slice(0, this.config.maxConcurrentRetries);

      for (const item of toProcess) {
        // Check if should abandon
        const age = Date.now() - item.createdAt.getTime();
        if (age > this.config.abandonAfterMs) {
          await this.persistentDLQ.markAbandoned(item.id);
          results.push({
            operationId: item.id,
            success: false,
            retryCount: item.attempts,
            workflowInstanceId: (item.payload as any).workflowInstanceId,
            abandoned: true,
            error: 'Auto-abandoned after 24 hours'
          });
          continue;
        }

        // Check if max retries exceeded
        if (item.attempts >= this.config.maxRetries) {
          await this.persistentDLQ.markAbandoned(item.id);
          results.push({
            operationId: item.id,
            success: false,
            retryCount: item.attempts,
            workflowInstanceId: (item.payload as any).workflowInstanceId,
            abandoned: true,
            error: 'Max retries exceeded'
          });
          continue;
        }

        // Mark as processing
        await this.persistentDLQ.markProcessing(item.id);

        // Attempt retry
        const result = await this.retryResumeOperation(item);
        results.push(result);

        if (result.success) {
          await this.persistentDLQ.markResolved(item.id);
        } else {
          await this.persistentDLQ.updateAttempt(item.id, result.error);
        }
      }

      const succeeded = results.filter(r => r.success).length;
      const failed = results.filter(r => !r.success && !r.abandoned).length;
      const abandoned = results.filter(r => r.abandoned).length;

      logger.info('WorkflowResumeRetryService',
        `Retry batch complete: ${succeeded} succeeded, ${failed} failed, ${abandoned} abandoned`);

      return {
        processedAt: new Date(),
        totalOperations: results.length,
        succeeded,
        failed,
        abandoned,
        results
      };
    } finally {
      this.isProcessing = false;
    }
  }

  /**
   * Retry a single resume operation with exponential backoff
   */
  private async retryResumeOperation(
    item: { id: string; operationType: string; payload: unknown; attempts: number; context?: Record<string, unknown> }
  ): Promise<IResumeRetryResult> {
    const payload = item.payload as {
      workflowInstanceId: number;
      stepId?: string;
      taskId?: number;
      approvalId?: number;
      completionData?: Record<string, unknown>;
    };

    const retryOptions: IRetryOptions = {
      maxRetries: 2, // Per-operation retries (on top of queue retries)
      initialDelay: this.config.initialDelayMs,
      maxDelay: this.config.maxDelayMs,
      factor: 2,
      jitter: true,
      onRetry: (error, attempt) => {
        logger.info('WorkflowResumeRetryService',
          `Retry attempt ${attempt} for workflow ${payload.workflowInstanceId}: ${error?.message || error}`);
      }
    };

    try {
      // First, check if workflow still needs resume
      const instance = await this.instanceService.getById(payload.workflowInstanceId);

      // If workflow is already completed or failed, no need to retry
      if (instance.Status === WorkflowInstanceStatus.Completed ||
          instance.Status === WorkflowInstanceStatus.Failed ||
          instance.Status === WorkflowInstanceStatus.Cancelled) {
        return {
          operationId: item.id,
          success: true,
          retryCount: item.attempts,
          workflowInstanceId: payload.workflowInstanceId,
          newWorkflowStatus: instance.Status,
          error: `Workflow already in terminal state: ${instance.Status}`
        };
      }

      // If workflow is running, it may have been resumed already
      if (instance.Status === WorkflowInstanceStatus.Running) {
        return {
          operationId: item.id,
          success: true,
          retryCount: item.attempts,
          workflowInstanceId: payload.workflowInstanceId,
          newWorkflowStatus: instance.Status,
          error: 'Workflow already running - may have been resumed'
        };
      }

      // Attempt the resume with retry
      const result = await retryOperation(async () => {
        if (payload.stepId) {
          return await this.workflowEngine.completeWaitingStep(
            payload.workflowInstanceId,
            payload.stepId,
            {
              retryOperation: true,
              retryAttempt: item.attempts + 1,
              originalTaskId: payload.taskId,
              originalApprovalId: payload.approvalId,
              ...payload.completionData
            }
          );
        } else {
          return await this.workflowEngine.resumeWorkflow(
            payload.workflowInstanceId,
            {
              retryOperation: true,
              retryAttempt: item.attempts + 1,
              ...payload.completionData
            }
          );
        }
      }, retryOptions);

      return {
        operationId: item.id,
        success: result.success,
        retryCount: item.attempts + 1,
        workflowInstanceId: payload.workflowInstanceId,
        newWorkflowStatus: result.status,
        error: result.success ? undefined : result.message
      };
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      logger.error('WorkflowResumeRetryService',
        `Failed to retry workflow ${payload.workflowInstanceId}`, error);

      return {
        operationId: item.id,
        success: false,
        retryCount: item.attempts + 1,
        workflowInstanceId: payload.workflowInstanceId,
        error: errorMessage
      };
    }
  }

  /**
   * Force retry a specific operation
   */
  public async forceRetry(operationId: string): Promise<IResumeRetryResult> {
    const items = await this.persistentDLQ.getPendingItems();
    const item = items.find(i => i.id === operationId);

    if (!item) {
      return {
        operationId,
        success: false,
        retryCount: 0,
        workflowInstanceId: 0,
        error: 'Operation not found'
      };
    }

    await this.persistentDLQ.markProcessing(operationId);
    const result = await this.retryResumeOperation({
      id: item.id,
      operationType: item.operationType,
      payload: item.payload,
      attempts: item.attempts,
      context: item.context
    });

    if (result.success) {
      await this.persistentDLQ.markResolved(operationId);
    } else {
      await this.persistentDLQ.updateAttempt(operationId, result.error);
    }

    return result;
  }

  /**
   * Get statistics about retry queue
   */
  public async getStats(): Promise<{
    pending: number;
    processing: number;
    totalRetried: number;
    successRate: number;
    oldestPending?: Date;
    isRunning: boolean;
    config: IResumeRetryConfig;
  }> {
    const dlqStats = await this.persistentDLQ.getStats();

    return {
      pending: dlqStats.pending,
      processing: dlqStats.processing,
      totalRetried: dlqStats.resolved + dlqStats.abandoned,
      successRate: dlqStats.resolved / Math.max(1, dlqStats.resolved + dlqStats.abandoned),
      isRunning: this.retryTimer !== null,
      config: { ...this.config }
    };
  }

  /**
   * Get all pending retry operations
   */
  public async getPendingOperations(): Promise<IFailedResumeOperation[]> {
    const items = await this.persistentDLQ.getPendingItems();

    return items
      .filter(item => item.operationType.startsWith('workflow-resume-'))
      .map(item => {
        const payload = item.payload as any;
        return {
          id: item.id,
          workflowInstanceId: payload.workflowInstanceId,
          stepId: payload.stepId,
          taskId: payload.taskId,
          approvalId: payload.approvalId,
          operationType: item.operationType.replace('workflow-resume-', '') as 'task' | 'approval' | 'manual',
          failedAt: item.createdAt,
          lastRetryAt: item.lastAttemptAt,
          retryCount: item.attempts,
          lastError: item.error,
          completionData: payload.completionData
        };
      });
  }

  /**
   * Abandon an operation (mark it as not retryable)
   */
  public async abandonOperation(operationId: string, userId?: number): Promise<boolean> {
    return await this.persistentDLQ.markAbandoned(operationId, userId);
  }

  /**
   * Clear all pending retry operations
   * Use with caution - only for admin recovery scenarios
   */
  public async clearAllPending(): Promise<number> {
    const pending = await this.getPendingOperations();
    let cleared = 0;

    for (const op of pending) {
      if (await this.persistentDLQ.markAbandoned(op.id)) {
        cleared++;
      }
    }

    logger.warn('WorkflowResumeRetryService', `Cleared ${cleared} pending retry operations`);
    return cleared;
  }
}
