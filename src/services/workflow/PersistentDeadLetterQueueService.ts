// @ts-nocheck
/* eslint-disable */
/**
 * PersistentDeadLetterQueueService
 * Extends the in-memory DLQ with SharePoint list persistence
 * CRITICAL: Ensures failed operations are not lost on page refresh or session timeout
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { IDeadLetterItem, DeadLetterQueue } from '../../utils/retryUtils';
import { logger } from '../LoggingService';

/**
 * SharePoint list schema for DLQ persistence:
 * - Title: Operation Type
 * - OperationId: Unique identifier (dlq_xxx)
 * - Payload: JSON string of the operation payload
 * - ErrorMessage: Error details
 * - AttemptCount: Number of retry attempts
 * - Context: JSON string of additional context
 * - Status: Pending | Processing | Resolved | Abandoned
 * - CreatedDate: When first failed
 * - LastAttemptDate: When last retried
 * - ResolvedDate: When resolved (if applicable)
 * - ResolvedBy: User who resolved (if applicable)
 */

export interface IPersistentDLQItem extends IDeadLetterItem {
  spListItemId?: number;
  status: 'Pending' | 'Processing' | 'Resolved' | 'Abandoned';
  resolvedDate?: Date;
  resolvedById?: number;
}

export interface IDLQRetryResult {
  success: boolean;
  item?: IPersistentDLQItem;
  error?: string;
}

export class PersistentDeadLetterQueueService {
  private sp: SPFI;
  private inMemoryDLQ: DeadLetterQueue;
  private readonly LIST_NAME = 'PM_DeadLetterQueue';
  private initialized: boolean = false;

  constructor(sp: SPFI, inMemoryDLQ?: DeadLetterQueue) {
    this.sp = sp;
    this.inMemoryDLQ = inMemoryDLQ || new DeadLetterQueue(500);
  }

  /**
   * Initialize the service and sync in-memory items to SharePoint
   */
  public async initialize(): Promise<void> {
    if (this.initialized) return;

    try {
      // Check if list exists
      await this.ensureListExists();

      // Sync any in-memory items that might have accumulated
      await this.syncInMemoryToSharePoint();

      // Load pending items from SharePoint to in-memory
      await this.loadPendingItems();

      this.initialized = true;
      logger.info('PersistentDLQService', 'Dead Letter Queue service initialized');
    } catch (error) {
      logger.error('PersistentDLQService', 'Failed to initialize DLQ service', error);
      // Continue without persistence - in-memory DLQ will still work
    }
  }

  /**
   * Ensure the DLQ SharePoint list exists
   */
  private async ensureListExists(): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(this.LIST_NAME)();
    } catch (error) {
      // List doesn't exist - this is expected in new deployments
      logger.warn('PersistentDLQService',
        `DLQ list '${this.LIST_NAME}' not found. Persistence will be disabled until list is provisioned.`);
      throw error;
    }
  }

  /**
   * Add a failed operation to the persistent DLQ
   */
  public async enqueue(
    operationType: string,
    payload: unknown,
    error: Error | string,
    attempts: number,
    context?: Record<string, unknown>
  ): Promise<string> {
    // First add to in-memory queue
    const dlqId = this.inMemoryDLQ.enqueue(operationType, payload, error, attempts, context);

    // Then persist to SharePoint
    try {
      await this.persistToSharePoint({
        id: dlqId,
        operationType,
        payload,
        error: typeof error === 'string' ? error : error.message,
        attempts,
        createdAt: new Date(),
        lastAttemptAt: new Date(),
        context,
        status: 'Pending'
      });
    } catch (persistError) {
      logger.warn('PersistentDLQService',
        `Failed to persist DLQ item ${dlqId} to SharePoint. Item remains in memory.`, persistError);
    }

    return dlqId;
  }

  /**
   * Persist a DLQ item to SharePoint
   */
  private async persistToSharePoint(item: IPersistentDLQItem): Promise<number | undefined> {
    try {
      const result = await this.sp.web.lists.getByTitle(this.LIST_NAME).items.add({
        Title: item.operationType,
        OperationId: item.id,
        Payload: JSON.stringify(item.payload),
        ErrorMessage: item.error,
        AttemptCount: item.attempts,
        Context: item.context ? JSON.stringify(item.context) : null,
        Status: item.status,
        CreatedDate: item.createdAt.toISOString(),
        LastAttemptDate: item.lastAttemptAt.toISOString()
      });

      logger.info('PersistentDLQService', `Persisted DLQ item ${item.id} to SharePoint (ID: ${result.data.Id})`);
      return result.data.Id;
    } catch (error) {
      logger.error('PersistentDLQService', `Failed to persist DLQ item ${item.id}`, error);
      throw error;
    }
  }

  /**
   * Sync all in-memory items to SharePoint
   */
  private async syncInMemoryToSharePoint(): Promise<void> {
    const inMemoryItems = this.inMemoryDLQ.getAll();

    if (inMemoryItems.length === 0) return;

    logger.info('PersistentDLQService', `Syncing ${inMemoryItems.length} in-memory DLQ items to SharePoint`);

    for (const item of inMemoryItems) {
      try {
        // Check if already persisted
        const existing = await this.findByOperationId(item.id);
        if (!existing) {
          await this.persistToSharePoint({
            ...item,
            status: 'Pending'
          });
        }
      } catch (error) {
        logger.warn('PersistentDLQService', `Failed to sync item ${item.id}`, error);
      }
    }
  }

  /**
   * Load pending items from SharePoint to in-memory queue
   */
  private async loadPendingItems(): Promise<void> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.LIST_NAME)
        .items
        .filter("Status eq 'Pending' or Status eq 'Processing'")
        .select('*')
        .orderBy('CreatedDate', true)
        .top(500)();

      logger.info('PersistentDLQService', `Loaded ${items.length} pending DLQ items from SharePoint`);

      // Add to in-memory queue if not already there
      for (const item of items) {
        const existing = this.inMemoryDLQ.getAll().find(i => i.id === item.OperationId);
        if (!existing) {
          // Add directly to in-memory without re-persisting
          this.inMemoryDLQ.enqueue(
            item.Title,
            JSON.parse(item.Payload || '{}'),
            item.ErrorMessage,
            item.AttemptCount,
            item.Context ? JSON.parse(item.Context) : undefined
          );
        }
      }
    } catch (error) {
      logger.warn('PersistentDLQService', 'Failed to load pending items from SharePoint', error);
    }
  }

  /**
   * Find item by operation ID
   */
  private async findByOperationId(operationId: string): Promise<any | undefined> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.LIST_NAME)
        .items
        .filter(`OperationId eq '${operationId}'`)
        .top(1)();

      return items.length > 0 ? items[0] : undefined;
    } catch (error) {
      return undefined;
    }
  }

  /**
   * Mark an item as being processed (to prevent duplicate processing)
   */
  public async markProcessing(operationId: string): Promise<boolean> {
    try {
      const item = await this.findByOperationId(operationId);
      if (!item) return false;

      await this.sp.web.lists.getByTitle(this.LIST_NAME)
        .items.getById(item.Id)
        .update({
          Status: 'Processing',
          LastAttemptDate: new Date().toISOString()
        });

      return true;
    } catch (error) {
      logger.error('PersistentDLQService', `Failed to mark ${operationId} as processing`, error);
      return false;
    }
  }

  /**
   * Mark an item as resolved (successfully retried or manually resolved)
   */
  public async markResolved(operationId: string, userId?: number): Promise<boolean> {
    try {
      const item = await this.findByOperationId(operationId);
      if (!item) return false;

      await this.sp.web.lists.getByTitle(this.LIST_NAME)
        .items.getById(item.Id)
        .update({
          Status: 'Resolved',
          ResolvedDate: new Date().toISOString(),
          ResolvedById: userId
        });

      // Remove from in-memory queue
      this.inMemoryDLQ.remove(operationId);

      logger.info('PersistentDLQService', `Marked ${operationId} as resolved`);
      return true;
    } catch (error) {
      logger.error('PersistentDLQService', `Failed to mark ${operationId} as resolved`, error);
      return false;
    }
  }

  /**
   * Mark an item as abandoned (will not be retried)
   */
  public async markAbandoned(operationId: string, userId?: number): Promise<boolean> {
    try {
      const item = await this.findByOperationId(operationId);
      if (!item) return false;

      await this.sp.web.lists.getByTitle(this.LIST_NAME)
        .items.getById(item.Id)
        .update({
          Status: 'Abandoned',
          ResolvedDate: new Date().toISOString(),
          ResolvedById: userId
        });

      // Remove from in-memory queue
      this.inMemoryDLQ.remove(operationId);

      logger.info('PersistentDLQService', `Marked ${operationId} as abandoned`);
      return true;
    } catch (error) {
      logger.error('PersistentDLQService', `Failed to mark ${operationId} as abandoned`, error);
      return false;
    }
  }

  /**
   * Update retry attempt count
   */
  public async updateAttempt(operationId: string, newError?: string): Promise<void> {
    try {
      // Update in-memory
      this.inMemoryDLQ.updateAttempt(operationId);

      // Update in SharePoint
      const item = await this.findByOperationId(operationId);
      if (item) {
        const updateData: any = {
          AttemptCount: item.AttemptCount + 1,
          LastAttemptDate: new Date().toISOString(),
          Status: 'Pending' // Reset from Processing back to Pending if retry failed
        };

        if (newError) {
          updateData.ErrorMessage = newError;
        }

        await this.sp.web.lists.getByTitle(this.LIST_NAME)
          .items.getById(item.Id)
          .update(updateData);
      }
    } catch (error) {
      logger.warn('PersistentDLQService', `Failed to update attempt for ${operationId}`, error);
    }
  }

  /**
   * Get all pending items from SharePoint
   */
  public async getPendingItems(): Promise<IPersistentDLQItem[]> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.LIST_NAME)
        .items
        .filter("Status eq 'Pending'")
        .select('*')
        .orderBy('CreatedDate', true)();

      return items.map(item => ({
        id: item.OperationId,
        operationType: item.Title,
        payload: JSON.parse(item.Payload || '{}'),
        error: item.ErrorMessage,
        attempts: item.AttemptCount,
        createdAt: new Date(item.CreatedDate),
        lastAttemptAt: new Date(item.LastAttemptDate),
        context: item.Context ? JSON.parse(item.Context) : undefined,
        spListItemId: item.Id,
        status: item.Status
      }));
    } catch (error) {
      logger.error('PersistentDLQService', 'Failed to get pending items', error);
      return this.inMemoryDLQ.getAll().map(item => ({
        ...item,
        status: 'Pending' as const
      }));
    }
  }

  /**
   * Get statistics
   */
  public async getStats(): Promise<{
    total: number;
    pending: number;
    processing: number;
    resolved: number;
    abandoned: number;
    byType: Record<string, number>;
  }> {
    try {
      const items = await this.sp.web.lists.getByTitle(this.LIST_NAME)
        .items
        .select('Title', 'Status')();

      const stats = {
        total: items.length,
        pending: 0,
        processing: 0,
        resolved: 0,
        abandoned: 0,
        byType: {} as Record<string, number>
      };

      items.forEach(item => {
        // Count by status
        switch (item.Status) {
          case 'Pending': stats.pending++; break;
          case 'Processing': stats.processing++; break;
          case 'Resolved': stats.resolved++; break;
          case 'Abandoned': stats.abandoned++; break;
        }

        // Count by operation type (only pending/processing)
        if (item.Status === 'Pending' || item.Status === 'Processing') {
          stats.byType[item.Title] = (stats.byType[item.Title] || 0) + 1;
        }
      });

      return stats;
    } catch (error) {
      // Fallback to in-memory stats
      const inMemoryStats = this.inMemoryDLQ.getStats();
      return {
        total: inMemoryStats.total,
        pending: inMemoryStats.total,
        processing: 0,
        resolved: 0,
        abandoned: 0,
        byType: inMemoryStats.byType
      };
    }
  }

  /**
   * Get the underlying in-memory queue (for backward compatibility)
   */
  public getInMemoryQueue(): DeadLetterQueue {
    return this.inMemoryDLQ;
  }
}
