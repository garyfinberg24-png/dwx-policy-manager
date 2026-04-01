/**
 * EventTriageService — Client-side service that calls the Azure Function
 * with mode='event-triage' for AI-powered event analysis.
 *
 * Follows the same pattern as PolicyChatService.
 */

import { logger } from '../LoggingService';
import { EventBuffer } from './EventBuffer';
import {
  IEventEntry,
  INetworkEvent,
  IEventTriageRequest,
  IEventTriageResponse,
  ISessionInfo,
} from '../../models/IEventViewer';

// ============================================================================
// SERVICE
// ============================================================================

export class EventTriageService {
  private _functionUrl: string;
  private _eventBuffer: EventBuffer;

  constructor(functionUrl: string) {
    this._functionUrl = functionUrl;
    this._eventBuffer = EventBuffer.getInstance();
  }

  /**
   * Analyse all current session events.
   */
  public async triageSession(sessionInfo: ISessionInfo): Promise<IEventTriageResponse> {
    const events = this._eventBuffer.getAll();
    const networkEvents = this._eventBuffer.getNetworkEvents();

    // Build network stats
    const totalDuration = networkEvents.reduce((sum, e) => sum + (e.duration || 0), 0);
    const failedCount = networkEvents.filter(e => e.httpStatus && e.httpStatus >= 400).length;
    const throttledCount = networkEvents.filter(e => e.httpStatus === 429).length;

    const request: IEventTriageRequest = {
      mode: 'event-triage',
      message: 'Analyse this session. Identify all root causes, assess severity, suggest fixes, and provide a health summary.',
      eventContext: {
        events: this._compactEvents(events),
        sessionInfo: sessionInfo,
        networkStats: {
          totalRequests: networkEvents.length,
          avgLatency: networkEvents.length > 0 ? Math.round(totalDuration / networkEvents.length) : 0,
          errorRate: networkEvents.length > 0 ? Math.round((failedCount / networkEvents.length) * 100) : 0,
          throttledCount: throttledCount,
        },
      },
      conversationHistory: [],
      userRole: 'Admin',
    };

    return this._callFunction(request);
  }

  /**
   * Triage a single event with surrounding context.
   */
  public async triageEvent(event: IEventEntry): Promise<IEventTriageResponse> {
    // Include 5 events before and after for context
    const allEvents = this._eventBuffer.getAll();
    const idx = allEvents.findIndex(e => e.id === event.id);
    const contextStart = Math.max(0, idx - 5);
    const contextEnd = Math.min(allEvents.length, idx + 6);
    const contextEvents = allEvents.slice(contextStart, contextEnd);

    const request: IEventTriageRequest = {
      mode: 'event-triage',
      message: `Analyse this specific event in detail: [${event.eventCode || 'Unknown'}] ${event.message}\n\nProvide root cause, why it happened, and a specific code fix if applicable.`,
      eventContext: {
        events: this._compactEvents(contextEvents),
      },
      conversationHistory: [],
      userRole: 'Admin',
    };

    return this._callFunction(request);
  }

  /**
   * Ask a freeform question about the events.
   */
  public async askAI(
    question: string,
    conversationHistory: Array<{ role: string; content: string }>
  ): Promise<IEventTriageResponse> {
    // Send the most recent 30 events as context
    const recentEvents = this._eventBuffer.getAll().slice(0, 30);

    const request: IEventTriageRequest = {
      mode: 'event-triage',
      message: question,
      eventContext: {
        events: this._compactEvents(recentEvents),
      },
      conversationHistory: conversationHistory.slice(-8).map(m => ({
        role: m.role as 'user' | 'assistant',
        content: m.content,
      })),
      userRole: 'Admin',
    };

    return this._callFunction(request);
  }

  // ==========================================================================
  // PRIVATE
  // ==========================================================================

  private async _callFunction(request: IEventTriageRequest): Promise<IEventTriageResponse> {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 30000);

    try {
      const response = await fetch(this._functionUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(request),
        signal: controller.signal,
      });

      clearTimeout(timeoutId);

      if (!response.ok) {
        throw new Error(`AI Triage API returned ${response.status}`);
      }

      const data = await response.json();
      return {
        analysis: data.message || data.analysis || '',
        suggestedActions: data.suggestedActions?.map((a: any) => a.label) || [],
        confidence: 80,
      };
    } catch (error) {
      clearTimeout(timeoutId);
      if (error instanceof Error && error.name === 'AbortError') {
        logger.error('EventTriageService', 'AI triage request timed out after 30s');
        throw new Error('AI triage request timed out. Please try again.');
      }
      logger.error('EventTriageService', 'AI triage request failed:', error);
      throw error;
    }
  }

  /**
   * Compact events for the API payload — strip unnecessary fields and truncate.
   */
  private _compactEvents(events: IEventEntry[]): IEventTriageRequest['eventContext']['events'] {
    return events.map(e => {
      const net = e as INetworkEvent;
      return {
        id: e.id,
        timestamp: e.timestamp,
        severity: e.severity,
        channel: e.channel,
        source: e.source,
        message: e.message.substring(0, 300),
        eventCode: e.eventCode,
        stackTrace: e.stackTrace ? e.stackTrace.substring(0, 500) : undefined,
        httpMethod: net.httpMethod,
        httpStatus: net.httpStatus,
        duration: net.duration,
        requestUrl: net.requestUrl ? net.requestUrl.substring(0, 200) : undefined,
      };
    });
  }
}
