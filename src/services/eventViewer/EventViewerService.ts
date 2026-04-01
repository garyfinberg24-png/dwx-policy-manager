/**
 * EventViewerService — SharePoint CRUD for PM_EventLog list.
 * Also reads from PM_PolicyAuditLog and PM_DeadLetterQueue.
 *
 * Implements IPersistCallback so EventBuffer can auto-persist Error/Critical events.
 */

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { PM_LISTS } from '../../constants/SharePointListNames';
import { logger } from '../LoggingService';
import { IPersistCallback } from './EventBuffer';
import {
  IEventEntry,
  INetworkEvent,
  IPersistedEvent,
  IEventFilter,
  EventSeverity,
} from '../../models/IEventViewer';

// ============================================================================
// SERVICE
// ============================================================================

export class EventViewerService implements IPersistCallback {
  private sp: SPFI;
  private listName: string = PM_LISTS.EVENT_LOG;
  private auditListName: string = PM_LISTS.POLICY_AUDIT_LOG;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  // ==========================================================================
  // PERSIST — Write events to PM_EventLog
  // ==========================================================================

  /**
   * Persist a single event to PM_EventLog.
   * Returns the SP list item ID.
   */
  public async persistEvent(event: IEventEntry): Promise<number> {
    try {
      const item = await this.sp.web.lists
        .getByTitle(this.listName)
        .items.add(this._eventToListItem(event));

      event.persisted = true;
      event.persistedItemId = item.data.Id;
      return item.data.Id;
    } catch (error) {
      logger.error('EventViewerService', 'Failed to persist event:', error);
      throw error;
    }
  }

  /**
   * Persist a batch of events. Uses per-item try/catch — continues on failure.
   * Implements IPersistCallback for EventBuffer auto-persist.
   */
  public async persistBatch(events: IEventEntry[]): Promise<void> {
    let failCount = 0;
    for (let i = 0; i < events.length; i++) {
      try {
        await this.persistEvent(events[i]);
      } catch (err) {
        failCount++;
        logger.warn('EventViewerService', `Failed to persist event ${i + 1}/${events.length}:`, err);
      }
    }
    if (failCount > 0) {
      logger.error('EventViewerService', `${failCount}/${events.length} events failed to persist`);
    }
  }

  // ==========================================================================
  // READ — Query persisted events
  // ==========================================================================

  /**
   * Get persisted events from PM_EventLog with optional filters.
   */
  public async getPersistedEvents(filter?: IEventFilter, top: number = 100): Promise<IPersistedEvent[]> {
    try {
      let query = this.sp.web.lists
        .getByTitle(this.listName)
        .items.select(
          'Id', 'Title', 'EventCode', 'Severity', 'Channel', 'Source',
          'Message', 'StackTrace', 'CorrelationId', 'SessionId', 'UserLogin',
          'EventTimestamp', 'Duration', 'Url', 'HttpMethod', 'HttpStatus',
          'InvestigationNotes', 'Classification', 'IsInvestigated',
          'AutoPersisted', 'Metadata'
        )
        .orderBy('EventTimestamp', false)
        .top(top);

      // Apply OData filters
      const filters: string[] = [];
      if (filter?.severities && filter.severities.length === 1) {
        filters.push(`Severity eq '${this._severityToString(filter.severities[0])}'`);
      }
      if (filter?.channels && filter.channels.length === 1) {
        filters.push(`Channel eq '${filter.channels[0]}'`);
      }
      if (filter?.eventCodes && filter.eventCodes.length === 1) {
        filters.push(`EventCode eq '${filter.eventCodes[0]}'`);
      }
      if (filter?.startTime) {
        filters.push(`EventTimestamp ge datetime'${filter.startTime}'`);
      }
      if (filter?.endTime) {
        filters.push(`EventTimestamp le datetime'${filter.endTime}'`);
      }
      if (filter?.isInvestigated !== undefined) {
        filters.push(`IsInvestigated eq ${filter.isInvestigated ? 1 : 0}`);
      }

      if (filters.length > 0) {
        query = query.filter(filters.join(' and '));
      }

      return await query();
    } catch (error) {
      logger.error('EventViewerService', 'Failed to get persisted events:', error);
      return [];
    }
  }

  // ==========================================================================
  // INVESTIGATION — Update investigation notes + classification
  // ==========================================================================

  /**
   * Add investigation notes to a persisted event.
   */
  public async addInvestigationNote(eventId: number, notes: string): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.listName)
        .items.getById(eventId)
        .update({
          InvestigationNotes: notes,
          IsInvestigated: true,
        });
    } catch (error) {
      logger.error('EventViewerService', `Failed to update investigation notes for event ${eventId}:`, error);
      throw error;
    }
  }

  /**
   * Update the classification of a persisted event.
   */
  public async updateClassification(eventId: number, classification: string): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(this.listName)
        .items.getById(eventId)
        .update({ Classification: classification });
    } catch (error) {
      logger.error('EventViewerService', `Failed to update classification for event ${eventId}:`, error);
      throw error;
    }
  }

  // ==========================================================================
  // AUDIT LOG — Read from PM_PolicyAuditLog
  // ==========================================================================

  /**
   * Get recent audit log entries for the System Health / Audit channel.
   */
  public async getAuditLogEvents(top: number = 50): Promise<any[]> {
    try {
      return await this.sp.web.lists
        .getByTitle(this.auditListName)
        .items.select(
          'Id', 'Title', 'AuditAction', 'EntityType', 'EntityId',
          'EntityName', 'ActionDescription', 'PerformedByName',
          'Created', 'Severity'
        )
        .orderBy('Created', false)
        .top(top)();
    } catch (error) {
      logger.error('EventViewerService', 'Failed to get audit log events:', error);
      return [];
    }
  }

  // ==========================================================================
  // RETENTION — Delete old events
  // ==========================================================================

  /**
   * Delete events older than the specified number of days.
   * Returns the count of deleted items.
   */
  public async deleteOldEvents(olderThanDays: number): Promise<number> {
    try {
      const cutoffDate = new Date();
      cutoffDate.setDate(cutoffDate.getDate() - olderThanDays);
      const cutoffIso = cutoffDate.toISOString();

      const oldItems = await this.sp.web.lists
        .getByTitle(this.listName)
        .items.select('Id')
        .filter(`EventTimestamp lt datetime'${cutoffIso}'`)
        .top(500)();

      let deletedCount = 0;
      for (let i = 0; i < oldItems.length; i++) {
        try {
          await this.sp.web.lists
            .getByTitle(this.listName)
            .items.getById(oldItems[i].Id)
            .delete();
          deletedCount++;
        } catch (err) {
          logger.warn('EventViewerService', `Failed to delete event ${oldItems[i].Id}:`, err);
        }
      }

      if (deletedCount > 0) {
        logger.info('EventViewerService', `Retention cleanup: deleted ${deletedCount} events older than ${olderThanDays} days`);
      }

      return deletedCount;
    } catch (error) {
      logger.error('EventViewerService', 'Failed to run retention cleanup:', error);
      return 0;
    }
  }

  // ==========================================================================
  // STATS
  // ==========================================================================

  /**
   * Get aggregate statistics from persisted events.
   */
  public async getEventStats(): Promise<{
    total: number;
    byCode: Record<string, number>;
    bySeverity: Record<string, number>;
  }> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.listName)
        .items.select('EventCode', 'Severity')
        .top(5000)();

      const byCode: Record<string, number> = {};
      const bySeverity: Record<string, number> = {};

      for (let i = 0; i < items.length; i++) {
        const code = items[i].EventCode || 'Unknown';
        const sev = items[i].Severity || 'Unknown';
        byCode[code] = (byCode[code] || 0) + 1;
        bySeverity[sev] = (bySeverity[sev] || 0) + 1;
      }

      return { total: items.length, byCode, bySeverity };
    } catch (error) {
      logger.error('EventViewerService', 'Failed to get event stats:', error);
      return { total: 0, byCode: {}, bySeverity: {} };
    }
  }

  // ==========================================================================
  // PRIVATE HELPERS
  // ==========================================================================

  private _eventToListItem(event: IEventEntry): Record<string, any> {
    const networkEvent = event as INetworkEvent;
    return {
      Title: (event.message || '').substring(0, 255),
      EventCode: event.eventCode || '',
      Severity: this._severityToString(event.severity),
      Channel: event.channel,
      Source: event.source,
      Message: event.message,
      StackTrace: event.stackTrace || '',
      CorrelationId: event.id,
      SessionId: event.sessionId || '',
      UserLogin: '',  // Will be set by current user context
      EventTimestamp: event.timestamp,
      Duration: networkEvent.duration || null,
      Url: networkEvent.requestUrl || event.url || '',
      HttpMethod: networkEvent.httpMethod || '',
      HttpStatus: networkEvent.httpStatus || null,
      AutoPersisted: event.autoPersist || false,
      Metadata: event.metadata ? JSON.stringify(event.metadata) : '',
    };
  }

  private _severityToString(severity: EventSeverity): string {
    switch (severity) {
      case EventSeverity.Verbose: return 'Verbose';
      case EventSeverity.Information: return 'Information';
      case EventSeverity.Warning: return 'Warning';
      case EventSeverity.Error: return 'Error';
      case EventSeverity.Critical: return 'Critical';
      default: return 'Information';
    }
  }
}
