/**
 * Retry Utilities with Dead Letter Queue Support
 * Provides retry logic with exponential backoff and DLQ for failed operations
 */

import { logger } from '../services/LoggingService';

// ============================================================================
// INTERFACES
// ============================================================================

/**
 * Options for retry behavior
 */
export interface IRetryOptions {
  maxRetries: number;
  initialDelayMs: number;
  maxDelayMs: number;
  backoffMultiplier: number;
}

/**
 * Result from a retry operation
 */
export interface IRetryResult<T> {
  success: boolean;
  data?: T;
  error?: Error;
  attempts: number;
  totalDurationMs: number;
  deadLetterItemId?: string;
}

/**
 * Dead letter queue item
 */
export interface IDLQItem {
  id: string;
  operationType: string;
  payload: unknown;
  error?: Error;
  attempts: number;
  createdAt: Date;
  lastAttemptAt: Date;
  metadata?: Record<string, unknown>;
}

/**
 * Dead letter queue stats
 */
export interface IDLQStats {
  totalItems: number;
  byType: Record<string, number>;
  oldestItem?: Date;
  newestItem?: Date;
}

// ============================================================================
// IN-MEMORY DEAD LETTER QUEUE
// ============================================================================

/**
 * Simple in-memory dead letter queue for failed operations
 */
export class InMemoryDeadLetterQueue {
  private items: Map<string, IDLQItem> = new Map();

  /**
   * Add an item to the DLQ
   */
  public add(item: IDLQItem): void {
    this.items.set(item.id, item);
  }

  /**
   * Get all items from the DLQ
   */
  public getAll(): IDLQItem[] {
    return Array.from(this.items.values());
  }

  /**
   * Get items by operation type
   */
  public getByType(operationType: string): IDLQItem[] {
    return this.getAll().filter(item => item.operationType === operationType);
  }

  /**
   * Remove an item from the DLQ
   */
  public remove(id: string): boolean {
    return this.items.delete(id);
  }

  /**
   * Update the attempt count for an item
   */
  public updateAttempt(id: string): void {
    const item = this.items.get(id);
    if (item) {
      item.attempts += 1;
      item.lastAttemptAt = new Date();
    }
  }

  /**
   * Get DLQ statistics
   */
  public getStats(): IDLQStats {
    const items = this.getAll();
    const byType: Record<string, number> = {};

    items.forEach(item => {
      byType[item.operationType] = (byType[item.operationType] || 0) + 1;
    });

    const dates = items.map(i => i.createdAt.getTime());

    return {
      totalItems: items.length,
      byType,
      oldestItem: dates.length > 0 ? new Date(Math.min(...dates)) : undefined,
      newestItem: dates.length > 0 ? new Date(Math.max(...dates)) : undefined
    };
  }

  /**
   * Clear all items from the DLQ
   */
  public clear(): void {
    this.items.clear();
  }
}

// ============================================================================
// DEFAULT CONFIGURATION
// ============================================================================

/**
 * Default retry options for process sync operations
 */
export const PROCESS_SYNC_RETRY_OPTIONS: IRetryOptions = {
  maxRetries: 3,
  initialDelayMs: 1000,
  maxDelayMs: 10000,
  backoffMultiplier: 2
};

/**
 * Shared DLQ instance for workflow sync operations
 */
export const workflowSyncDLQ = new InMemoryDeadLetterQueue();

// ============================================================================
// RETRY WITH DLQ
// ============================================================================

/**
 * Generate a unique ID for DLQ items
 */
function generateDLQId(): string {
  return `dlq_${Date.now()}_${Math.random().toString(36).substring(2, 9)}`;
}

/**
 * Delay execution for a specified number of milliseconds
 */
function delay(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Execute an async operation with retry logic and dead letter queue support.
 * If all retries fail, the operation is added to the DLQ for later processing.
 *
 * @param operation - The async operation to execute
 * @param operationType - A string identifying the type of operation (for DLQ categorization)
 * @param payload - The payload associated with the operation (stored in DLQ on failure)
 * @param options - Retry configuration options
 * @param dlq - The dead letter queue to use for failed operations
 * @param metadata - Optional metadata to attach to DLQ items
 * @returns A result object indicating success/failure with attempt details
 */
export async function retryWithDLQ<T>(
  operation: () => Promise<T>,
  operationType: string,
  payload: unknown,
  options: IRetryOptions = PROCESS_SYNC_RETRY_OPTIONS,
  dlq: InMemoryDeadLetterQueue = workflowSyncDLQ,
  metadata?: Record<string, unknown>
): Promise<IRetryResult<T>> {
  const startTime = Date.now();
  let lastError: Error | undefined;

  for (let attempt = 1; attempt <= options.maxRetries + 1; attempt++) {
    try {
      const data = await operation();
      return {
        success: true,
        data,
        attempts: attempt,
        totalDurationMs: Date.now() - startTime
      };
    } catch (err) {
      lastError = err instanceof Error ? err : new Error(String(err));

      if (attempt <= options.maxRetries) {
        const delayMs = Math.min(
          options.initialDelayMs * Math.pow(options.backoffMultiplier, attempt - 1),
          options.maxDelayMs
        );
        logger.warn('retryUtils',
          `Retry ${attempt}/${options.maxRetries} for ${operationType} after ${delayMs}ms`,
          lastError
        );
        await delay(delayMs);
      }
    }
  }

  // All retries exhausted - add to DLQ
  const dlqId = generateDLQId();
  const dlqItem: IDLQItem = {
    id: dlqId,
    operationType,
    payload,
    error: lastError,
    attempts: options.maxRetries + 1,
    createdAt: new Date(),
    lastAttemptAt: new Date(),
    metadata
  };

  dlq.add(dlqItem);

  logger.error('retryUtils',
    `All ${options.maxRetries + 1} attempts failed for ${operationType}. Added to DLQ: ${dlqId}`,
    lastError
  );

  return {
    success: false,
    error: lastError,
    attempts: options.maxRetries + 1,
    totalDurationMs: Date.now() - startTime,
    deadLetterItemId: dlqId
  };
}
