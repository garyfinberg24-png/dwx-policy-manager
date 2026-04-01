/**
 * ConsoleInterceptor — Monkey-patches console.log/warn/error/debug to capture
 * output into the EventBuffer while preserving original console behaviour.
 *
 * Usage:
 *   ConsoleInterceptor.getInstance().install();
 *   // ... console output is now captured ...
 *   ConsoleInterceptor.getInstance().uninstall(); // restore originals
 */

import { EventBuffer } from './EventBuffer';
import {
  IEventEntry,
  EventSeverity,
  EventChannel,
  ConsoleOrigin,
} from '../../models/IEventViewer';

// ============================================================================
// TYPES
// ============================================================================

type ConsoleMethod = 'log' | 'warn' | 'error' | 'debug';

interface IOriginalMethods {
  log: typeof console.log;
  warn: typeof console.warn;
  error: typeof console.error;
  debug: typeof console.debug;
}

/** Map console method → severity + default event code */
const METHOD_MAP: Record<ConsoleMethod, { severity: EventSeverity; code: string }> = {
  error: { severity: EventSeverity.Error, code: 'CON-001' },
  warn: { severity: EventSeverity.Warning, code: 'CON-002' },
  log: { severity: EventSeverity.Information, code: 'CON-003' },
  debug: { severity: EventSeverity.Verbose, code: 'CON-004' },
};

/** Patterns to detect console sub-origin */
const ORIGIN_PATTERNS: Array<{ pattern: RegExp; origin: ConsoleOrigin }> = [
  { pattern: /^\[SPLoader|^\[SPPageContext|^\[SPPropertyBag|^\[SPModule/i, origin: ConsoleOrigin.Framework },
  { pattern: /^PnPLogging|^\[@pnp/i, origin: ConsoleOrigin.Library },
  { pattern: /^Warning:.*Did not expect|^Warning:.*setState|^Warning:.*Each child|^Warning:.*Failed prop/i, origin: ConsoleOrigin.React },
  { pattern: /^\[(INFO|WARN|ERROR|VERBOSE)\]\s*\[/, origin: ConsoleOrigin.App },
];

// ============================================================================
// SINGLETON
// ============================================================================

export class ConsoleInterceptor {
  private static _instance: ConsoleInterceptor;

  private _installed: boolean = false;
  private _originals: IOriginalMethods | null = null;
  private _eventBuffer: EventBuffer;

  private constructor() {
    this._eventBuffer = EventBuffer.getInstance();
  }

  public static getInstance(): ConsoleInterceptor {
    if (!ConsoleInterceptor._instance) {
      ConsoleInterceptor._instance = new ConsoleInterceptor();
    }
    return ConsoleInterceptor._instance;
  }

  // ==========================================================================
  // INSTALL / UNINSTALL
  // ==========================================================================

  /**
   * Install console interceptors. Stores original methods for restoration.
   */
  public install(): void {
    if (this._installed) return;

    // Store originals
    this._originals = {
      log: console.log.bind(console),
      warn: console.warn.bind(console),
      error: console.error.bind(console),
      debug: console.debug.bind(console),
    };

    // Patch each method
    const methods: ConsoleMethod[] = ['log', 'warn', 'error', 'debug'];
    for (let i = 0; i < methods.length; i++) {
      const method = methods[i];
      this._patchMethod(method);
    }

    // Capture unhandled promise rejections
    if (typeof window !== 'undefined') {
      window.addEventListener('unhandledrejection', this._onUnhandledRejection);
    }

    this._installed = true;
  }

  /**
   * Uninstall interceptors — restore original console methods.
   */
  public uninstall(): void {
    if (!this._installed || !this._originals) return;

    console.log = this._originals.log;
    console.warn = this._originals.warn;
    console.error = this._originals.error;
    console.debug = this._originals.debug;

    if (typeof window !== 'undefined') {
      window.removeEventListener('unhandledrejection', this._onUnhandledRejection);
    }

    this._installed = false;
  }

  /** Whether interceptors are currently active */
  public get isInstalled(): boolean {
    return this._installed;
  }

  // ==========================================================================
  // PRIVATE — Patching
  // ==========================================================================

  private _patchMethod(method: ConsoleMethod): void {
    const originals = this._originals!;
    const interceptor = this;

    (console as any)[method] = function (...args: any[]): void {
      // Always call original first — never swallow output
      originals[method].apply(console, args);

      // Capture into EventBuffer
      try {
        interceptor._captureConsoleCall(method, args);
      } catch (_) {
        // Never let capture break the console
      }
    };
  }

  private _captureConsoleCall(method: ConsoleMethod, args: any[]): void {
    // Build message from args
    const message = this._argsToString(args);

    // Prevent recursion — skip our own EventViewer/EventBuffer output
    if (message.indexOf('[EventViewer]') !== -1) return;

    const mapping = METHOD_MAP[method];
    const origin = this._detectOrigin(message);

    // Extract stack trace for errors
    let stackTrace: string | undefined;
    if (method === 'error') {
      for (let i = 0; i < args.length; i++) {
        if (args[i] instanceof Error && args[i].stack) {
          stackTrace = args[i].stack;
          break;
        }
      }
    }

    // Extract source from logger format: [LEVEL] [Source] message
    const source = this._extractSource(message) || 'Console';

    const event: IEventEntry = {
      id: `evt_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
      timestamp: new Date().toISOString(),
      severity: mapping.severity,
      channel: EventChannel.Console,
      source: source,
      message: message,
      eventCode: mapping.code,
      stackTrace: stackTrace,
      consoleOrigin: origin,
      url: typeof window !== 'undefined' ? window.location.pathname : undefined,
    };

    this._eventBuffer.push(event);
  }

  // ==========================================================================
  // PRIVATE — Unhandled rejection handler
  // ==========================================================================

  private _onUnhandledRejection = (e: PromiseRejectionEvent): void => {
    try {
      const reason = e.reason;
      const message = reason instanceof Error
        ? reason.message
        : String(reason);

      const event: IEventEntry = {
        id: `evt_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
        timestamp: new Date().toISOString(),
        severity: EventSeverity.Critical,
        channel: EventChannel.Console,
        source: 'UnhandledRejection',
        message: `Unhandled promise rejection: ${message}`,
        eventCode: 'CON-005',
        stackTrace: reason instanceof Error ? reason.stack : undefined,
        consoleOrigin: ConsoleOrigin.App,
        url: typeof window !== 'undefined' ? window.location.pathname : undefined,
      };

      this._eventBuffer.push(event);
    } catch (_) {
      // Never break on capture failure
    }
  };

  // ==========================================================================
  // PRIVATE — Helpers
  // ==========================================================================

  /**
   * Convert console.log arguments to a single string.
   */
  private _argsToString(args: any[]): string {
    const parts: string[] = [];
    for (let i = 0; i < args.length; i++) {
      const arg = args[i];
      if (arg === null) {
        parts.push('null');
      } else if (arg === undefined) {
        parts.push('undefined');
      } else if (typeof arg === 'string') {
        parts.push(arg);
      } else if (arg instanceof Error) {
        parts.push(arg.message);
      } else if (typeof arg === 'object') {
        try {
          parts.push(JSON.stringify(arg, null, 0).substring(0, 500));
        } catch (_) {
          parts.push('[Object]');
        }
      } else {
        parts.push(String(arg));
      }
    }
    return parts.join(' ');
  }

  /**
   * Detect the console sub-origin from the message content.
   */
  private _detectOrigin(message: string): ConsoleOrigin {
    for (let i = 0; i < ORIGIN_PATTERNS.length; i++) {
      if (ORIGIN_PATTERNS[i].pattern.test(message)) {
        return ORIGIN_PATTERNS[i].origin;
      }
    }
    return ConsoleOrigin.Browser;
  }

  /**
   * Extract source component name from logger format: [LEVEL] [Source] message
   */
  private _extractSource(message: string): string | undefined {
    // Match: [INFO] [PolicyService] ... or [ERROR] [ApprovalService] ...
    const match = message.match(/^\[(INFO|WARN|ERROR|VERBOSE)\]\s*\[([^\]]+)\]/);
    if (match) {
      return match[2];
    }
    return undefined;
  }
}
