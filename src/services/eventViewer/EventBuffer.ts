/**
 * EventBuffer — Singleton ring buffer with pub-sub for the DWx Event Viewer.
 *
 * Maintains three fixed-size circular buffers (Application, Console, Network)
 * that capture events in-memory. Subscribers receive new events in real-time.
 *
 * Usage:
 *   const buffer = EventBuffer.getInstance();
 *   const unsub = buffer.subscribe(event => console.log(event));
 *   buffer.push({ id: '...', severity: EventSeverity.Error, ... });
 *   unsub(); // cleanup
 */

import {
  IEventEntry,
  INetworkEvent,
  IEventFilter,
  IEventBufferStats,
  EventSeverity,
  EventChannel,
  DEFAULT_EVENT_VIEWER_CONFIG,
} from '../../models/IEventViewer';
import { EventClassifier } from './EventClassifier';

// ============================================================================
// TYPES
// ============================================================================

/** Listener callback for new events */
export type EventListener = (event: IEventEntry) => void;

// ============================================================================
// RING BUFFER IMPLEMENTATION
// ============================================================================

/**
 * Fixed-size circular array — evicts oldest entries when full.
 */
class RingBuffer<T> {
  private _items: T[];
  private _maxSize: number;
  private _writeIndex: number = 0;
  private _count: number = 0;

  constructor(maxSize: number) {
    this._maxSize = maxSize;
    this._items = new Array<T>(maxSize);
  }

  /** Add an item — evicts oldest if full */
  public push(item: T): void {
    this._items[this._writeIndex] = item;
    this._writeIndex = (this._writeIndex + 1) % this._maxSize;
    if (this._count < this._maxSize) {
      this._count++;
    }
  }

  /** Get all items in chronological order (oldest first) */
  public getAll(): T[] {
    if (this._count === 0) return [];
    if (this._count < this._maxSize) {
      return this._items.slice(0, this._count);
    }
    // Buffer is full — read from writeIndex (oldest) to end, then start to writeIndex
    return [
      ...this._items.slice(this._writeIndex),
      ...this._items.slice(0, this._writeIndex),
    ];
  }

  /** Get all items in reverse chronological order (newest first) */
  public getAllReversed(): T[] {
    return this.getAll().reverse();
  }

  /** Current number of items in the buffer */
  public get count(): number {
    return this._count;
  }

  /** Maximum buffer capacity */
  public get capacity(): number {
    return this._maxSize;
  }

  /** Clear all items */
  public clear(): void {
    this._items = new Array<T>(this._maxSize);
    this._writeIndex = 0;
    this._count = 0;
  }

  /** Resize buffer (preserves existing items up to new size) */
  public resize(newSize: number): void {
    const existing = this.getAllReversed().slice(0, newSize);
    this._maxSize = newSize;
    this._items = new Array<T>(newSize);
    this._writeIndex = 0;
    this._count = 0;
    // Re-insert in chronological order
    existing.reverse().forEach(item => this.push(item));
  }
}

// ============================================================================
// EVENT BUFFER SINGLETON
// ============================================================================

export class EventBuffer {
  private static _instance: EventBuffer;

  // Ring buffers by channel
  private _appBuffer: RingBuffer<IEventEntry>;
  private _consoleBuffer: RingBuffer<IEventEntry>;
  private _networkBuffer: RingBuffer<INetworkEvent>;

  // Pub-sub listeners
  private _listeners: EventListener[] = [];

  // Session ID for correlation
  private _sessionId: string;

  // Optional persistence service reference (set in Phase 3)
  private _persistService: IPersistCallback | undefined;

  // Debounced persist queue
  private _persistQueue: IEventEntry[] = [];
  private _persistTimer: ReturnType<typeof setTimeout> | null = null;
  private static readonly PERSIST_DEBOUNCE_MS = 500;

  private constructor() {
    const config = DEFAULT_EVENT_VIEWER_CONFIG;
    this._appBuffer = new RingBuffer<IEventEntry>(config.appBufferSize);
    this._consoleBuffer = new RingBuffer<IEventEntry>(config.consoleBufferSize);
    this._networkBuffer = new RingBuffer<INetworkEvent>(config.networkBufferSize);
    this._sessionId = `sess_${Date.now().toString(36)}_${Math.random().toString(36).substring(2, 6)}`;
  }

  public static getInstance(): EventBuffer {
    if (!EventBuffer._instance) {
      EventBuffer._instance = new EventBuffer();
    }
    return EventBuffer._instance;
  }

  /** For testing — reset singleton */
  public static resetInstance(): void {
    if (EventBuffer._instance) {
      EventBuffer._instance.dispose();
    }
    EventBuffer._instance = undefined as unknown as EventBuffer;
  }

  // ==========================================================================
  // PUSH — Add an event to the appropriate buffer
  // ==========================================================================

  /**
   * Push an event into the buffer. Automatically:
   * - Assigns sessionId if not set
   * - Routes to the correct buffer by channel
   * - Notifies all subscribers
   * - Queues for persistence if severity >= Error
   */
  public push(event: IEventEntry): void {
    // Assign session ID
    if (!event.sessionId) {
      event.sessionId = this._sessionId;
    }

    // Assign timestamp if not set
    if (!event.timestamp) {
      event.timestamp = new Date().toISOString();
    }

    // Auto-classify: assign event code if not already set
    if (!event.eventCode || event.channel === EventChannel.Console) {
      const classification = event.channel === EventChannel.Console
        ? EventClassifier.reclassifyConsoleEvent(event)
        : EventClassifier.classify(event);
      if (classification) {
        event.eventCode = classification.eventCode;
      }
    }

    // Route to appropriate buffer
    switch (event.channel) {
      case EventChannel.Console:
        this._consoleBuffer.push(event);
        break;
      case EventChannel.Network:
        this._networkBuffer.push(event as INetworkEvent);
        break;
      case EventChannel.Application:
      case EventChannel.Audit:
      case EventChannel.DLQ:
      case EventChannel.System:
      default:
        this._appBuffer.push(event);
        break;
    }

    // Flag auto-persist for Error and Critical
    if (event.severity >= EventSeverity.Error) {
      event.autoPersist = true;
      this._queueForPersist(event);
    }

    // Notify subscribers
    this._notifyListeners(event);
  }

  // ==========================================================================
  // SUBSCRIBE / UNSUBSCRIBE — Pub-sub pattern
  // ==========================================================================

  /**
   * Subscribe to new events. Returns an unsubscribe function.
   */
  public subscribe(listener: EventListener): () => void {
    this._listeners.push(listener);
    return () => {
      const idx = this._listeners.indexOf(listener);
      if (idx !== -1) {
        this._listeners.splice(idx, 1);
      }
    };
  }

  // ==========================================================================
  // QUERY — Get events from buffers
  // ==========================================================================

  /**
   * Get all events across all buffers (newest first).
   */
  public getAll(channel?: EventChannel): IEventEntry[] {
    if (channel) {
      return this._getBufferForChannel(channel).getAllReversed();
    }

    // Merge all buffers, sort by timestamp descending
    const all: IEventEntry[] = [
      ...this._appBuffer.getAll(),
      ...this._consoleBuffer.getAll(),
      ...this._networkBuffer.getAll(),
    ];

    return all.sort((a, b) =>
      new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime()
    );
  }

  /**
   * Get events matching a filter.
   */
  public getByFilter(filter: IEventFilter): IEventEntry[] {
    let events = this.getAll();

    if (filter.channels && filter.channels.length > 0) {
      events = events.filter(e => filter.channels!.indexOf(e.channel) !== -1);
    }

    if (filter.severities && filter.severities.length > 0) {
      events = events.filter(e => filter.severities!.indexOf(e.severity) !== -1);
    }

    if (filter.minSeverity !== undefined) {
      events = events.filter(e => e.severity >= filter.minSeverity!);
    }

    if (filter.source) {
      const src = filter.source.toLowerCase();
      events = events.filter(e => e.source.toLowerCase().indexOf(src) !== -1);
    }

    if (filter.searchText) {
      const q = filter.searchText.toLowerCase();
      events = events.filter(e =>
        e.message.toLowerCase().indexOf(q) !== -1 ||
        e.source.toLowerCase().indexOf(q) !== -1 ||
        (e.eventCode && e.eventCode.toLowerCase().indexOf(q) !== -1) ||
        (e.url && e.url.toLowerCase().indexOf(q) !== -1)
      );
    }

    if (filter.eventCodes && filter.eventCodes.length > 0) {
      events = events.filter(e =>
        e.eventCode && filter.eventCodes!.indexOf(e.eventCode) !== -1
      );
    }

    if (filter.startTime) {
      const start = new Date(filter.startTime).getTime();
      events = events.filter(e => new Date(e.timestamp).getTime() >= start);
    }

    if (filter.endTime) {
      const end = new Date(filter.endTime).getTime();
      events = events.filter(e => new Date(e.timestamp).getTime() <= end);
    }

    if (filter.spListName) {
      const listName = filter.spListName.toLowerCase();
      events = events.filter(e =>
        (e as INetworkEvent).spListName &&
        (e as INetworkEvent).spListName!.toLowerCase().indexOf(listName) !== -1
      );
    }

    if (filter.minDuration !== undefined) {
      events = events.filter(e =>
        (e as INetworkEvent).duration !== undefined &&
        (e as INetworkEvent).duration! >= filter.minDuration!
      );
    }

    if (filter.includeAssets === false) {
      events = events.filter(e => !(e as INetworkEvent).isAssetRequest);
    }

    return events;
  }

  /**
   * Get events grouped by event code.
   */
  public getByEventCode(code: string): IEventEntry[] {
    return this.getAll().filter(e => e.eventCode === code);
  }

  /**
   * Get network events only (typed).
   */
  public getNetworkEvents(): INetworkEvent[] {
    return this._networkBuffer.getAllReversed();
  }

  // ==========================================================================
  // STATS
  // ==========================================================================

  /**
   * Get buffer statistics.
   */
  public getStats(): IEventBufferStats {
    const all = this.getAll();
    return {
      appCount: this._appBuffer.count,
      consoleCount: this._consoleBuffer.count,
      networkCount: this._networkBuffer.count,
      totalCount: this._appBuffer.count + this._consoleBuffer.count + this._networkBuffer.count,
      errorCount: all.filter(e => e.severity === EventSeverity.Error).length,
      warningCount: all.filter(e => e.severity === EventSeverity.Warning).length,
      criticalCount: all.filter(e => e.severity === EventSeverity.Critical).length,
      capacity: {
        app: this._appBuffer.capacity,
        console: this._consoleBuffer.capacity,
        network: this._networkBuffer.capacity,
      },
    };
  }

  // ==========================================================================
  // SESSION
  // ==========================================================================

  /** Get the current session ID */
  public get sessionId(): string {
    return this._sessionId;
  }

  // ==========================================================================
  // CLEAR / RESIZE / DISPOSE
  // ==========================================================================

  /** Clear one or all buffers */
  public clear(channel?: EventChannel): void {
    if (channel) {
      this._getBufferForChannel(channel).clear();
    } else {
      this._appBuffer.clear();
      this._consoleBuffer.clear();
      this._networkBuffer.clear();
    }
  }

  /** Resize buffers (e.g. from admin config) */
  public resizeBuffers(appSize: number, consoleSize: number, networkSize: number): void {
    this._appBuffer.resize(appSize);
    this._consoleBuffer.resize(consoleSize);
    this._networkBuffer.resize(networkSize);
  }

  /** Clean up resources */
  public dispose(): void {
    this._listeners = [];
    if (this._persistTimer) {
      clearTimeout(this._persistTimer);
      this._persistTimer = null;
    }
    this._persistQueue = [];
  }

  // ==========================================================================
  // PERSISTENCE INTEGRATION (wired in Phase 3)
  // ==========================================================================

  /**
   * Set the persistence callback for auto-persisting Error/Critical events.
   * Called by EventViewer component when EventViewerService is available.
   */
  public setPersistCallback(callback: IPersistCallback): void {
    this._persistService = callback;
  }

  /** Clear the persistence callback (on unmount) */
  public clearPersistCallback(): void {
    this._persistService = undefined;
  }

  // ==========================================================================
  // PRIVATE HELPERS
  // ==========================================================================

  private _getBufferForChannel(channel: EventChannel): RingBuffer<IEventEntry> {
    switch (channel) {
      case EventChannel.Console:
        return this._consoleBuffer;
      case EventChannel.Network:
        return this._networkBuffer as RingBuffer<IEventEntry>;
      default:
        return this._appBuffer;
    }
  }

  private _notifyListeners(event: IEventEntry): void {
    for (let i = 0; i < this._listeners.length; i++) {
      try {
        this._listeners[i](event);
      } catch (_) {
        // Never let a listener error break event capture
      }
    }
  }

  private _queueForPersist(event: IEventEntry): void {
    if (!this._persistService) return;

    this._persistQueue.push(event);

    // Debounce: batch persist after 500ms of no new events
    if (this._persistTimer) {
      clearTimeout(this._persistTimer);
    }
    this._persistTimer = setTimeout(() => {
      this._flushPersistQueue();
    }, EventBuffer.PERSIST_DEBOUNCE_MS);
  }

  private _flushPersistQueue(): void {
    if (!this._persistService || this._persistQueue.length === 0) return;

    const batch = this._persistQueue.splice(0);
    // Fire-and-forget — persistence failure must not lose in-memory events
    try {
      this._persistService.persistBatch(batch);
    } catch (_) {
      // Swallow — events remain in ring buffer regardless
    }
  }
}

// ============================================================================
// PERSISTENCE CALLBACK INTERFACE (implemented by EventViewerService in Phase 3)
// ============================================================================

export interface IPersistCallback {
  persistBatch(events: IEventEntry[]): Promise<void>;
}
