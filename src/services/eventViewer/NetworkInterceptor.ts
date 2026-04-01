/**
 * NetworkInterceptor — Wraps window.fetch and XMLHttpRequest to capture
 * all HTTP traffic into the EventBuffer with timing, status, and SP list tagging.
 *
 * Usage:
 *   NetworkInterceptor.getInstance().install();
 *   // ... all fetch/XHR calls are now captured ...
 *   NetworkInterceptor.getInstance().uninstall();
 */

import { EventBuffer } from './EventBuffer';
import {
  INetworkEvent,
  EventSeverity,
  EventChannel,
} from '../../models/IEventViewer';
import {
  TELEMETRY_EXCLUSION_PATTERNS,
  ASSET_EXTENSIONS,
  SLOW_REQUEST_THRESHOLD_MS,
} from '../../constants/EventCodes';

// ============================================================================
// SINGLETON
// ============================================================================

export class NetworkInterceptor {
  private static _instance: NetworkInterceptor;

  private _installed: boolean = false;
  private _originalFetch: typeof window.fetch | null = null;
  private _originalXhrOpen: typeof XMLHttpRequest.prototype.open | null = null;
  private _originalXhrSend: typeof XMLHttpRequest.prototype.send | null = null;
  private _eventBuffer: EventBuffer;

  private constructor() {
    this._eventBuffer = EventBuffer.getInstance();
  }

  public static getInstance(): NetworkInterceptor {
    if (!NetworkInterceptor._instance) {
      NetworkInterceptor._instance = new NetworkInterceptor();
    }
    return NetworkInterceptor._instance;
  }

  // ==========================================================================
  // INSTALL / UNINSTALL
  // ==========================================================================

  public install(): void {
    if (this._installed) return;
    if (typeof window === 'undefined') return;

    this._patchFetch();
    this._patchXhr();
    this._installed = true;
  }

  public uninstall(): void {
    if (!this._installed) return;

    // Restore fetch
    if (this._originalFetch) {
      window.fetch = this._originalFetch;
      this._originalFetch = null;
    }

    // Restore XHR
    if (this._originalXhrOpen) {
      XMLHttpRequest.prototype.open = this._originalXhrOpen;
      this._originalXhrOpen = null;
    }
    if (this._originalXhrSend) {
      XMLHttpRequest.prototype.send = this._originalXhrSend;
      this._originalXhrSend = null;
    }

    this._installed = false;
  }

  public get isInstalled(): boolean {
    return this._installed;
  }

  // ==========================================================================
  // FETCH INTERCEPTION
  // ==========================================================================

  private _patchFetch(): void {
    this._originalFetch = window.fetch.bind(window);
    const interceptor = this;
    const originalFetch: typeof window.fetch = this._originalFetch!;

    window.fetch = function (input: RequestInfo | URL, init?: RequestInit): Promise<Response> {
      const url = interceptor._extractUrl(input);
      const method = (init?.method || 'GET').toUpperCase();

      // Skip telemetry endpoints
      if (interceptor._isExcluded(url)) {
        return originalFetch(input, init);
      }

      const startTime = Date.now();

      // Capture request body and headers for inspection
      const reqBody = interceptor._truncateBody(init?.body);
      const reqHeaders = interceptor._extractHeaders(init?.headers);

      return originalFetch(input, init).then(
        (response: Response) => {
          try {
            const duration = Date.now() - startTime;
            const resHeaders = interceptor._extractResponseHeaders(response.headers);
            interceptor._captureNetworkEvent(url, method, response.status, duration, response, undefined, reqBody, reqHeaders, resHeaders);
          } catch (_) {
            // Never break fetch on capture failure
          }
          return response;
        },
        (error: Error) => {
          try {
            const duration = Date.now() - startTime;
            interceptor._captureNetworkError(url, method, duration, error);
          } catch (_) {
            // Never break fetch on capture failure
          }
          throw error;
        }
      );
    };
  }

  // ==========================================================================
  // XHR INTERCEPTION
  // ==========================================================================

  private _patchXhr(): void {
    this._originalXhrOpen = XMLHttpRequest.prototype.open;
    this._originalXhrSend = XMLHttpRequest.prototype.send;
    const interceptor = this;
    const originalOpen = this._originalXhrOpen;
    const originalSend = this._originalXhrSend;

    // Patch open() to capture URL and method
    XMLHttpRequest.prototype.open = function (
      method: string,
      url: string | URL,
      async?: boolean,
      username?: string | null,
      password?: string | null
    ): void {
      (this as any)._evUrl = String(url);
      (this as any)._evMethod = method.toUpperCase();
      return originalOpen.apply(this, [method, url, async !== false, username, password] as any);
    };

    // Patch send() to capture timing and response
    XMLHttpRequest.prototype.send = function (body?: Document | XMLHttpRequestBodyInit | null): void {
      const url: string = (this as any)._evUrl || '';
      const method: string = (this as any)._evMethod || 'GET';

      // Skip telemetry endpoints
      if (interceptor._isExcluded(url)) {
        return originalSend.apply(this, [body] as any);
      }

      const startTime = Date.now();
      const xhr = this;

      const onLoadEnd = (): void => {
        try {
          const duration = Date.now() - startTime;
          interceptor._captureNetworkEvent(url, method, xhr.status, duration, undefined, xhr);
        } catch (_) {
          // Never break XHR on capture failure
        }
        xhr.removeEventListener('loadend', onLoadEnd);
        xhr.removeEventListener('error', onError);
      };

      const onError = (): void => {
        try {
          const duration = Date.now() - startTime;
          interceptor._captureNetworkError(url, method, duration, new Error('XHR network error'));
        } catch (_) {
          // Never break XHR on capture failure
        }
        xhr.removeEventListener('loadend', onLoadEnd);
        xhr.removeEventListener('error', onError);
      };

      xhr.addEventListener('loadend', onLoadEnd);
      xhr.addEventListener('error', onError);

      return originalSend.apply(this, [body] as any);
    };
  }

  // ==========================================================================
  // EVENT CREATION
  // ==========================================================================

  private _captureNetworkEvent(
    url: string,
    method: string,
    status: number,
    duration: number,
    response?: Response,
    xhr?: XMLHttpRequest,
    reqBody?: string,
    reqHeaders?: Record<string, string>,
    resHeaders?: Record<string, string>
  ): void {
    const spListName = this._extractSpListName(url);
    const isAsset = ASSET_EXTENSIONS.test(url);
    const severity = this._statusToSeverity(status, duration);
    const eventCode = this._statusToEventCode(status, duration);

    // Try to get response size
    let responseSize: number | undefined;
    if (response) {
      const cl = response.headers.get('content-length');
      if (cl) responseSize = parseInt(cl, 10);
    } else if (xhr) {
      const cl = xhr.getResponseHeader('content-length');
      if (cl) responseSize = parseInt(cl, 10);
    }

    const statusText = status >= 400
      ? `${method} ${this._truncateUrl(url)} — ${status} ${this._statusLabel(status)}`
      : `${method} ${this._truncateUrl(url)}`;

    const event: INetworkEvent = {
      id: `evt_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
      timestamp: new Date().toISOString(),
      severity: severity,
      channel: EventChannel.Network,
      source: spListName || this._extractHost(url),
      message: statusText,
      eventCode: eventCode,
      url: typeof window !== 'undefined' ? window.location.pathname : undefined,
      requestUrl: url,
      httpMethod: method,
      httpStatus: status,
      duration: duration,
      responseSize: responseSize,
      spListName: spListName,
      isAssetRequest: isAsset,
      requestBody: reqBody,
      requestHeaders: reqHeaders || (xhr ? this._extractXhrResponseHeaders(xhr) : undefined),
      responseHeaders: resHeaders || (xhr ? this._extractXhrResponseHeaders(xhr) : undefined),
    };

    this._eventBuffer.push(event);
  }

  private _captureNetworkError(
    url: string,
    method: string,
    duration: number,
    error: Error
  ): void {
    const spListName = this._extractSpListName(url);
    const isAsset = ASSET_EXTENSIONS.test(url);

    // Determine specific error code
    let eventCode = 'NET-021'; // Generic network failure
    const msg = error.message.toLowerCase();
    if (msg.indexOf('abort') !== -1 || msg.indexOf('timeout') !== -1) {
      eventCode = 'NET-020';
    } else if (msg.indexOf('cors') !== -1 || msg.indexOf('cross-origin') !== -1) {
      eventCode = 'NET-022';
    }

    const event: INetworkEvent = {
      id: `evt_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
      timestamp: new Date().toISOString(),
      severity: EventSeverity.Error,
      channel: EventChannel.Network,
      source: spListName || this._extractHost(url),
      message: `${method} ${this._truncateUrl(url)} — ${error.message}`,
      eventCode: eventCode,
      stackTrace: error.stack,
      url: typeof window !== 'undefined' ? window.location.pathname : undefined,
      requestUrl: url,
      httpMethod: method,
      httpStatus: 0,
      duration: duration,
      spListName: spListName,
      isAssetRequest: isAsset,
    };

    this._eventBuffer.push(event);
  }

  // ==========================================================================
  // PRIVATE HELPERS
  // ==========================================================================

  private _extractUrl(input: RequestInfo | URL): string {
    if (typeof input === 'string') return input;
    if (input instanceof URL) return input.href;
    if (input instanceof Request) return input.url;
    return String(input);
  }

  private _isExcluded(url: string): boolean {
    for (let i = 0; i < TELEMETRY_EXCLUSION_PATTERNS.length; i++) {
      if (TELEMETRY_EXCLUSION_PATTERNS[i].test(url)) return true;
    }
    return false;
  }

  /**
   * Extract SharePoint list name from API URLs.
   * Matches: /_api/web/lists/getbytitle('PM_Policies')/items
   */
  private _extractSpListName(url: string): string | undefined {
    const match = url.match(/getbytitle\('([^']+)'\)/i);
    if (match) return match[1];

    // Also match getById pattern with list name in path
    const match2 = url.match(/lists\/([^/]+)\/items/i);
    if (match2 && match2[1].indexOf('PM_') === 0) return match2[1];

    return undefined;
  }

  private _extractHost(url: string): string {
    try {
      // Handle relative URLs
      if (url.startsWith('/')) return 'SharePoint';
      if (url.startsWith('_api/')) return 'SharePoint';
      const parsed = new URL(url);
      if (parsed.hostname.indexOf('sharepoint.com') !== -1) return 'SharePoint';
      if (parsed.hostname.indexOf('azurewebsites.net') !== -1) return 'Azure Function';
      if (parsed.hostname.indexOf('graph.microsoft.com') !== -1) return 'MS Graph';
      return parsed.hostname;
    } catch (_) {
      return 'Unknown';
    }
  }

  private _truncateUrl(url: string): string {
    // For SP API calls, show the meaningful part
    const apiIdx = url.indexOf('/_api/');
    if (apiIdx !== -1) return url.substring(apiIdx);

    // For Azure Functions, show path only
    const funcIdx = url.indexOf('/api/');
    if (funcIdx !== -1) {
      const path = url.substring(funcIdx);
      // Hide function key
      const codeIdx = path.indexOf('?code=');
      if (codeIdx !== -1) return path.substring(0, codeIdx) + '?code=***';
      return path;
    }

    // Truncate long URLs
    if (url.length > 120) return url.substring(0, 117) + '...';
    return url;
  }

  private _statusToSeverity(status: number, duration: number): EventSeverity {
    if (status === 0) return EventSeverity.Error;
    if (status === 429) return EventSeverity.Error;
    if (status >= 500) return EventSeverity.Error;
    if (status >= 400) return EventSeverity.Warning;
    if (duration > SLOW_REQUEST_THRESHOLD_MS) return EventSeverity.Warning;
    return EventSeverity.Verbose;
  }

  private _statusToEventCode(status: number, duration: number): string {
    if (status === 429) return 'NET-010';
    if (status === 400) return 'NET-002';
    if (status === 401) return 'NET-003';
    if (status === 403) return 'NET-004';
    if (status === 404) return 'NET-005';
    if (status === 409) return 'NET-006';
    if (status === 500) return 'NET-011';
    if (status === 502 || status === 503) return 'NET-012';
    if (status >= 400) return 'NET-002';
    if (duration > SLOW_REQUEST_THRESHOLD_MS) return 'NET-001';
    return 'NET-000';
  }

  private _statusLabel(status: number): string {
    switch (status) {
      case 400: return 'Bad Request';
      case 401: return 'Unauthorized';
      case 403: return 'Forbidden';
      case 404: return 'Not Found';
      case 409: return 'Conflict';
      case 429: return 'Too Many Requests';
      case 500: return 'Internal Server Error';
      case 502: return 'Bad Gateway';
      case 503: return 'Service Unavailable';
      default: return '';
    }
  }

  // ==========================================================================
  // REQUEST/RESPONSE INSPECTION HELPERS
  // ==========================================================================

  private static readonly MAX_BODY_SIZE = 4096; // 4KB truncation limit

  /** Truncate request body to safe size for memory */
  private _truncateBody(body: RequestInit['body']): string | undefined {
    if (!body) return undefined;
    try {
      let text: string;
      if (typeof body === 'string') {
        text = body;
      } else if (body instanceof URLSearchParams) {
        text = body.toString();
      } else {
        // FormData, Blob, ArrayBuffer — skip (too large / binary)
        return '[Binary or FormData body]';
      }
      if (text.length > NetworkInterceptor.MAX_BODY_SIZE) {
        return text.substring(0, NetworkInterceptor.MAX_BODY_SIZE) + '... [truncated]';
      }
      return text;
    } catch (_) {
      return undefined;
    }
  }

  /** Extract headers from RequestInit.headers */
  private _extractHeaders(headers?: HeadersInit): Record<string, string> | undefined {
    if (!headers) return undefined;
    try {
      const result: Record<string, string> = {};
      if (headers instanceof Headers) {
        headers.forEach((value, key) => { result[key] = value; });
      } else if (Array.isArray(headers)) {
        for (const [key, value] of headers) { result[key] = value; }
      } else {
        for (const key of Object.keys(headers)) { result[key] = (headers as Record<string, string>)[key]; }
      }
      return Object.keys(result).length > 0 ? result : undefined;
    } catch (_) {
      return undefined;
    }
  }

  /** Extract response headers from fetch Response */
  private _extractResponseHeaders(headers: Headers): Record<string, string> | undefined {
    try {
      const result: Record<string, string> = {};
      headers.forEach((value, key) => { result[key] = value; });
      return Object.keys(result).length > 0 ? result : undefined;
    } catch (_) {
      return undefined;
    }
  }

  /** Extract response headers from XHR */
  private _extractXhrResponseHeaders(xhr: XMLHttpRequest): Record<string, string> | undefined {
    try {
      const raw = xhr.getAllResponseHeaders();
      if (!raw) return undefined;
      const result: Record<string, string> = {};
      const lines = raw.trim().split(/[\r\n]+/);
      for (let i = 0; i < lines.length; i++) {
        const parts = lines[i].split(': ');
        if (parts.length >= 2) {
          result[parts[0]] = parts.slice(1).join(': ');
        }
      }
      return Object.keys(result).length > 0 ? result : undefined;
    } catch (_) {
      return undefined;
    }
  }
}
