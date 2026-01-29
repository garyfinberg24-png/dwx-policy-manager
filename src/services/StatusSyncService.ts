// @ts-nocheck
/**
 * StatusSyncService
 * Real-time status synchronization service using polling
 * Provides near-real-time updates for tasks, workflows, and approvals
 * CRITICAL: Ensures UI displays current state without manual refresh
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import { logger } from './LoggingService';
import { TaskStatus } from '../models/ICommon';
import { WorkflowInstanceStatus } from '../models/IWorkflow';

/**
 * Change event types
 */
export type ChangeType = 'created' | 'updated' | 'deleted' | 'statusChanged';

/**
 * Entity types that can be synchronized
 */
export type EntityType = 'task' | 'workflow' | 'approval' | 'process' | 'notification';

/**
 * Change event
 */
export interface IChangeEvent {
  entityType: EntityType;
  entityId: number;
  changeType: ChangeType;
  timestamp: Date;
  previousValue?: unknown;
  currentValue?: unknown;
  changedFields?: string[];
  changedBy?: {
    id: number;
    name: string;
  };
}

/**
 * Subscription to changes
 */
export interface IChangeSubscription {
  id: string;
  entityType: EntityType;
  entityId?: number; // Optional - if not set, subscribes to all of that type
  callback: (event: IChangeEvent) => void;
}

/**
 * Sync configuration
 */
export interface ISyncConfig {
  enabled: boolean;
  pollIntervalMs: number;
  entities: EntityType[];
  maxChangesPerPoll: number;
  enableNotifications: boolean;
}

/**
 * Entity state snapshot for change detection
 */
interface IEntitySnapshot {
  entityType: EntityType;
  entityId: number;
  status: string;
  modified: Date;
  modifiedById?: number;
  data: Record<string, unknown>;
}

export class StatusSyncService {
  private sp: SPFI;
  private config: ISyncConfig;
  private subscriptions: Map<string, IChangeSubscription> = new Map();
  private entitySnapshots: Map<string, IEntitySnapshot> = new Map();
  private pollTimer: ReturnType<typeof setInterval> | null = null;
  private isPolling: boolean = false;
  private lastPollTime: Date = new Date();

  constructor(sp: SPFI, config?: Partial<ISyncConfig>) {
    this.sp = sp;
    this.config = {
      enabled: true,
      pollIntervalMs: 10000, // 10 seconds
      entities: ['task', 'workflow', 'approval'],
      maxChangesPerPoll: 50,
      enableNotifications: true,
      ...config
    };
  }

  /**
   * Start the sync service
   */
  public start(): void {
    if (!this.config.enabled) {
      logger.info('StatusSyncService', 'Sync service is disabled');
      return;
    }

    if (this.pollTimer) {
      this.stop();
    }

    logger.info('StatusSyncService',
      `Starting sync service (interval: ${this.config.pollIntervalMs}ms)`);

    // Initial snapshot
    this.takeSnapshot().catch(err => {
      logger.error('StatusSyncService', 'Error taking initial snapshot', err);
    });

    // Start polling
    this.pollTimer = setInterval(() => {
      this.pollForChanges().catch(err => {
        logger.error('StatusSyncService', 'Error during poll', err);
      });
    }, this.config.pollIntervalMs);
  }

  /**
   * Stop the sync service
   */
  public stop(): void {
    if (this.pollTimer) {
      clearInterval(this.pollTimer);
      this.pollTimer = null;
      logger.info('StatusSyncService', 'Sync service stopped');
    }
  }

  /**
   * Subscribe to changes
   */
  public subscribe(
    entityType: EntityType,
    callback: (event: IChangeEvent) => void,
    entityId?: number
  ): string {
    const subscriptionId = `sub_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;

    this.subscriptions.set(subscriptionId, {
      id: subscriptionId,
      entityType,
      entityId,
      callback
    });

    logger.info('StatusSyncService',
      `New subscription ${subscriptionId} for ${entityType}${entityId ? ` (ID: ${entityId})` : ''}`);

    return subscriptionId;
  }

  /**
   * Unsubscribe from changes
   */
  public unsubscribe(subscriptionId: string): void {
    this.subscriptions.delete(subscriptionId);
    logger.info('StatusSyncService', `Unsubscribed ${subscriptionId}`);
  }

  /**
   * Force an immediate poll
   */
  public async forcePoll(): Promise<IChangeEvent[]> {
    return await this.pollForChanges();
  }

  /**
   * Take a snapshot of current entity states
   */
  private async takeSnapshot(): Promise<void> {
    try {
      if (this.config.entities.includes('task')) {
        await this.snapshotTasks();
      }

      if (this.config.entities.includes('workflow')) {
        await this.snapshotWorkflows();
      }

      if (this.config.entities.includes('approval')) {
        await this.snapshotApprovals();
      }

      this.lastPollTime = new Date();
    } catch (error) {
      logger.error('StatusSyncService', 'Error taking snapshot', error);
    }
  }

  /**
   * Poll for changes since last poll
   */
  private async pollForChanges(): Promise<IChangeEvent[]> {
    if (this.isPolling) {
      return [];
    }

    this.isPolling = true;
    const changes: IChangeEvent[] = [];

    try {
      // Poll each entity type
      if (this.config.entities.includes('task')) {
        const taskChanges = await this.pollTaskChanges();
        changes.push(...taskChanges);
      }

      if (this.config.entities.includes('workflow')) {
        const workflowChanges = await this.pollWorkflowChanges();
        changes.push(...workflowChanges);
      }

      if (this.config.entities.includes('approval')) {
        const approvalChanges = await this.pollApprovalChanges();
        changes.push(...approvalChanges);
      }

      // Notify subscribers
      for (const change of changes) {
        this.notifySubscribers(change);
      }

      this.lastPollTime = new Date();

      if (changes.length > 0) {
        logger.info('StatusSyncService', `Detected ${changes.length} changes`);
      }

      return changes;
    } finally {
      this.isPolling = false;
    }
  }

  /**
   * Poll for task changes
   */
  private async pollTaskChanges(): Promise<IChangeEvent[]> {
    const changes: IChangeEvent[] = [];

    try {
      const items = await this.sp.web.lists.getByTitle('PM_TaskAssignments')
        .items
        .filter(`Modified gt datetime'${this.lastPollTime.toISOString()}'`)
        .select('Id', 'Title', 'Status', 'Modified', 'EditorId')
        .orderBy('Modified', false)
        .top(this.config.maxChangesPerPoll)();

      for (const item of items) {
        const snapshotKey = `task_${item.Id}`;
        const previousSnapshot = this.entitySnapshots.get(snapshotKey);

        const currentSnapshot: IEntitySnapshot = {
          entityType: 'task',
          entityId: item.Id,
          status: item.Status,
          modified: new Date(item.Modified),
          modifiedById: item.EditorId,
          data: { Title: item.Title, Status: item.Status }
        };

        // Detect change type
        let changeType: ChangeType = 'updated';
        const changedFields: string[] = [];

        if (!previousSnapshot) {
          changeType = 'created';
        } else if (previousSnapshot.status !== currentSnapshot.status) {
          changeType = 'statusChanged';
          changedFields.push('Status');
        }

        changes.push({
          entityType: 'task',
          entityId: item.Id,
          changeType,
          timestamp: currentSnapshot.modified,
          previousValue: previousSnapshot?.data,
          currentValue: currentSnapshot.data,
          changedFields,
          changedBy: item.EditorId ? { id: item.EditorId, name: '' } : undefined
        });

        // Update snapshot
        this.entitySnapshots.set(snapshotKey, currentSnapshot);
      }
    } catch (error) {
      logger.warn('StatusSyncService', 'Error polling task changes', error);
    }

    return changes;
  }

  /**
   * Poll for workflow changes
   */
  private async pollWorkflowChanges(): Promise<IChangeEvent[]> {
    const changes: IChangeEvent[] = [];

    try {
      const items = await this.sp.web.lists.getByTitle('PM_WorkflowInstances')
        .items
        .filter(`Modified gt datetime'${this.lastPollTime.toISOString()}'`)
        .select('Id', 'Title', 'Status', 'CurrentStepId', 'CurrentStepName', 'Modified', 'EditorId')
        .orderBy('Modified', false)
        .top(this.config.maxChangesPerPoll)();

      for (const item of items) {
        const snapshotKey = `workflow_${item.Id}`;
        const previousSnapshot = this.entitySnapshots.get(snapshotKey);

        const currentSnapshot: IEntitySnapshot = {
          entityType: 'workflow',
          entityId: item.Id,
          status: item.Status,
          modified: new Date(item.Modified),
          modifiedById: item.EditorId,
          data: {
            Title: item.Title,
            Status: item.Status,
            CurrentStepId: item.CurrentStepId,
            CurrentStepName: item.CurrentStepName
          }
        };

        let changeType: ChangeType = 'updated';
        const changedFields: string[] = [];

        if (!previousSnapshot) {
          changeType = 'created';
        } else {
          if (previousSnapshot.status !== currentSnapshot.status) {
            changeType = 'statusChanged';
            changedFields.push('Status');
          }
          if (previousSnapshot.data.CurrentStepId !== currentSnapshot.data.CurrentStepId) {
            changedFields.push('CurrentStepId', 'CurrentStepName');
          }
        }

        changes.push({
          entityType: 'workflow',
          entityId: item.Id,
          changeType,
          timestamp: currentSnapshot.modified,
          previousValue: previousSnapshot?.data,
          currentValue: currentSnapshot.data,
          changedFields,
          changedBy: item.EditorId ? { id: item.EditorId, name: '' } : undefined
        });

        this.entitySnapshots.set(snapshotKey, currentSnapshot);
      }
    } catch (error) {
      logger.warn('StatusSyncService', 'Error polling workflow changes', error);
    }

    return changes;
  }

  /**
   * Poll for approval changes
   */
  private async pollApprovalChanges(): Promise<IChangeEvent[]> {
    const changes: IChangeEvent[] = [];

    try {
      const items = await this.sp.web.lists.getByTitle('PM_Approvals')
        .items
        .filter(`Modified gt datetime'${this.lastPollTime.toISOString()}'`)
        .select('Id', 'Title', 'Status', 'Modified', 'EditorId')
        .orderBy('Modified', false)
        .top(this.config.maxChangesPerPoll)();

      for (const item of items) {
        const snapshotKey = `approval_${item.Id}`;
        const previousSnapshot = this.entitySnapshots.get(snapshotKey);

        const currentSnapshot: IEntitySnapshot = {
          entityType: 'approval',
          entityId: item.Id,
          status: item.Status,
          modified: new Date(item.Modified),
          modifiedById: item.EditorId,
          data: { Title: item.Title, Status: item.Status }
        };

        let changeType: ChangeType = 'updated';
        const changedFields: string[] = [];

        if (!previousSnapshot) {
          changeType = 'created';
        } else if (previousSnapshot.status !== currentSnapshot.status) {
          changeType = 'statusChanged';
          changedFields.push('Status');
        }

        changes.push({
          entityType: 'approval',
          entityId: item.Id,
          changeType,
          timestamp: currentSnapshot.modified,
          previousValue: previousSnapshot?.data,
          currentValue: currentSnapshot.data,
          changedFields,
          changedBy: item.EditorId ? { id: item.EditorId, name: '' } : undefined
        });

        this.entitySnapshots.set(snapshotKey, currentSnapshot);
      }
    } catch (error) {
      logger.warn('StatusSyncService', 'Error polling approval changes', error);
    }

    return changes;
  }

  /**
   * Snapshot tasks for initial state
   */
  private async snapshotTasks(): Promise<void> {
    const items = await this.sp.web.lists.getByTitle('PM_TaskAssignments')
      .items
      .select('Id', 'Title', 'Status', 'Modified')
      .top(500)();

    for (const item of items) {
      this.entitySnapshots.set(`task_${item.Id}`, {
        entityType: 'task',
        entityId: item.Id,
        status: item.Status,
        modified: new Date(item.Modified),
        data: { Title: item.Title, Status: item.Status }
      });
    }
  }

  /**
   * Snapshot workflows for initial state
   */
  private async snapshotWorkflows(): Promise<void> {
    const items = await this.sp.web.lists.getByTitle('PM_WorkflowInstances')
      .items
      .select('Id', 'Title', 'Status', 'CurrentStepId', 'CurrentStepName', 'Modified')
      .top(500)();

    for (const item of items) {
      this.entitySnapshots.set(`workflow_${item.Id}`, {
        entityType: 'workflow',
        entityId: item.Id,
        status: item.Status,
        modified: new Date(item.Modified),
        data: {
          Title: item.Title,
          Status: item.Status,
          CurrentStepId: item.CurrentStepId,
          CurrentStepName: item.CurrentStepName
        }
      });
    }
  }

  /**
   * Snapshot approvals for initial state
   */
  private async snapshotApprovals(): Promise<void> {
    const items = await this.sp.web.lists.getByTitle('PM_Approvals')
      .items
      .select('Id', 'Title', 'Status', 'Modified')
      .top(500)();

    for (const item of items) {
      this.entitySnapshots.set(`approval_${item.Id}`, {
        entityType: 'approval',
        entityId: item.Id,
        status: item.Status,
        modified: new Date(item.Modified),
        data: { Title: item.Title, Status: item.Status }
      });
    }
  }

  /**
   * Notify relevant subscribers of a change
   */
  private notifySubscribers(event: IChangeEvent): void {
    for (const subscription of Array.from(this.subscriptions.values())) {
      // Check if subscription matches this event
      if (subscription.entityType !== event.entityType) {
        continue;
      }

      // If subscription is for specific entity, check ID
      if (subscription.entityId && subscription.entityId !== event.entityId) {
        continue;
      }

      try {
        subscription.callback(event);
      } catch (error) {
        logger.warn('StatusSyncService',
          `Error in subscription callback ${subscription.id}`, error);
      }
    }
  }

  /**
   * Get current status of the service
   */
  public getStatus(): {
    isRunning: boolean;
    subscriptionCount: number;
    entityCount: number;
    lastPollTime: Date;
    config: ISyncConfig;
  } {
    return {
      isRunning: this.pollTimer !== null,
      subscriptionCount: this.subscriptions.size,
      entityCount: this.entitySnapshots.size,
      lastPollTime: this.lastPollTime,
      config: { ...this.config }
    };
  }

  /**
   * Update configuration
   */
  public updateConfig(config: Partial<ISyncConfig>): void {
    this.config = { ...this.config, ...config };

    // Restart if running and interval changed
    if (this.pollTimer && config.pollIntervalMs) {
      this.stop();
      this.start();
    }
  }

  /**
   * Clear all snapshots (force re-detection)
   */
  public clearSnapshots(): void {
    this.entitySnapshots.clear();
    logger.info('StatusSyncService', 'Snapshots cleared');
  }

  /**
   * Subscribe to task status changes with typed callback
   */
  public subscribeToTaskChanges(
    callback: (event: IChangeEvent & { currentValue: { Status: TaskStatus } }) => void,
    taskId?: number
  ): string {
    return this.subscribe('task', callback as (event: IChangeEvent) => void, taskId);
  }

  /**
   * Subscribe to workflow status changes with typed callback
   */
  public subscribeToWorkflowChanges(
    callback: (event: IChangeEvent & { currentValue: { Status: WorkflowInstanceStatus } }) => void,
    workflowId?: number
  ): string {
    return this.subscribe('workflow', callback as (event: IChangeEvent) => void, workflowId);
  }

  /**
   * Subscribe to approval status changes
   */
  public subscribeToApprovalChanges(
    callback: (event: IChangeEvent) => void,
    approvalId?: number
  ): string {
    return this.subscribe('approval', callback, approvalId);
  }
}

/**
 * React hook for using StatusSyncService
 * Usage in components:
 *
 * const { subscribe, unsubscribe } = useStatusSync(statusSyncService);
 *
 * useEffect(() => {
 *   const subId = subscribe('task', (event) => {
 *     if (event.changeType === 'statusChanged') {
 *       // Refresh data
 *     }
 *   });
 *   return () => unsubscribe(subId);
 * }, []);
 */
export function createStatusSyncHooks(service: StatusSyncService) {
  return {
    subscribe: (
      entityType: EntityType,
      callback: (event: IChangeEvent) => void,
      entityId?: number
    ) => service.subscribe(entityType, callback, entityId),

    unsubscribe: (subscriptionId: string) => service.unsubscribe(subscriptionId),

    forcePoll: () => service.forcePoll(),

    getStatus: () => service.getStatus()
  };
}
