/**
 * LoggingService - Telemetry & Logging for Policy Manager
 *
 * Supports two modes:
 * 1. Console-only (default) — all telemetry goes to console.log/warn/error
 * 2. Application Insights — when initialized with a connection string, sends
 *    telemetry to Azure Application Insights via the Beacon API (no npm dependency)
 *
 * Usage:
 *   const logging = LoggingService.getInstance();
 *   logging.initializeAppInsights('InstrumentationKey=xxx;IngestionEndpoint=https://...');
 *   logging.trackPageView('PolicyHub');
 *   logging.trackEvent('PolicyAcknowledged', { policyId: '42' });
 */

// Enum for log severity (compatible with Application Insights)
export enum SeverityLevel {
  Verbose = 0,
  Information = 1,
  Warning = 2,
  Error = 3,
  Critical = 4,
}

// Lightweight Application Insights envelope structure
interface IAppInsightsEnvelope {
  name: string;
  time: string;
  iKey: string;
  tags: Record<string, string>;
  data: {
    baseType: string;
    baseData: Record<string, unknown>;
  };
}

export class LoggingService {
  private static instance: LoggingService;
  private isProduction: boolean = false;
  private isDevelopment: boolean = true;

  // Application Insights state
  private instrumentationKey: string = '';
  private ingestionEndpoint: string = '';
  private appInsightsEnabled: boolean = false;
  private userId: string = '';
  private sessionId: string = '';
  private pendingEnvelopes: IAppInsightsEnvelope[] = [];
  // @ts-ignore TS6133 — assigned in init() for interval cleanup, not read elsewhere yet
  private flushTimer: ReturnType<typeof setTimeout> | null = null;

  private constructor() {
    this.isProduction =
      window.location.hostname !== 'localhost' &&
      window.location.hostname.includes('.sharepoint.com');
    this.isDevelopment = !this.isProduction;
    this.sessionId = this.generateSessionId();
  }

  public static getInstance(): LoggingService {
    if (!LoggingService.instance) {
      LoggingService.instance = new LoggingService();
    }
    return LoggingService.instance;
  }

  /** Clean up interval timer when no longer needed */
  public dispose(): void {
    if (this.flushTimer) {
      clearInterval(this.flushTimer);
      this.flushTimer = null;
    }
  }

  /**
   * Initialize Application Insights telemetry.
   * Accepts either:
   *   - A connection string: "InstrumentationKey=xxx;IngestionEndpoint=https://..."
   *   - Just an instrumentation key (GUID)
   *
   * If the input is empty or invalid, App Insights stays disabled (console-only mode).
   */
  public initializeAppInsights(connectionStringOrKey: string): void {
    if (!connectionStringOrKey) {
      console.log('[LoggingService] No connection string — console-only mode');
      return;
    }

    try {
      if (connectionStringOrKey.includes('InstrumentationKey=')) {
        // Parse connection string
        const parts: Record<string, string> = {};
        connectionStringOrKey.split(';').forEach((segment) => {
          const [key, ...valueParts] = segment.split('=');
          if (key && valueParts.length > 0) {
            parts[key.trim()] = valueParts.join('=').trim();
          }
        });
        this.instrumentationKey = parts['InstrumentationKey'] || '';
        this.ingestionEndpoint =
          parts['IngestionEndpoint'] || 'https://dc.services.visualstudio.com';
      } else {
        // Assume bare instrumentation key
        this.instrumentationKey = connectionStringOrKey;
        this.ingestionEndpoint = 'https://dc.services.visualstudio.com';
      }

      if (this.instrumentationKey) {
        this.appInsightsEnabled = true;
        // Auto-flush every 15 seconds in production
        if (this.isProduction) {
          this.flushTimer = setInterval(() => this.flush(), 15000);
        }
        console.log(
          `[LoggingService] App Insights initialized (iKey: ${this.instrumentationKey.substring(0, 8)}...)`
        );
      }
    } catch (err) {
      console.warn('[LoggingService] Failed to parse App Insights connection string:', err);
    }
  }

  // ---------------------------------------------------------------------------
  // Public API — unchanged interface, now with App Insights support
  // ---------------------------------------------------------------------------

  public info(message: string, properties?: Record<string, unknown>): void {
    if (this.isDevelopment) {
      console.log(`[INFO] ${message}`, properties || '');
    }
    this.sendTrace(message, SeverityLevel.Information, properties);
  }

  public warn(source: string, message: string, properties?: Record<string, unknown>): void {
    console.warn(`[WARN] [${source}] ${message}`, properties || '');
    this.sendTrace(`[${source}] ${message}`, SeverityLevel.Warning, properties);
  }

  public error(
    source: string,
    message: string,
    error?: Error,
    properties?: Record<string, unknown>
  ): void {
    console.error(`[ERROR] [${source}] ${message}`, error || '', properties || '');
    if (error) {
      this.trackException(error, SeverityLevel.Error, {
        source,
        message,
        ...properties,
      });
    } else {
      this.sendTrace(`[${source}] ${message}`, SeverityLevel.Error, properties);
    }
  }

  public verbose(message: string, properties?: Record<string, unknown>): void {
    if (this.isDevelopment) {
      console.debug(`[VERBOSE] ${message}`, properties || '');
    }
    this.sendTrace(message, SeverityLevel.Verbose, properties);
  }

  public trackEvent(
    name: string,
    properties?: Record<string, unknown>,
    measurements?: Record<string, number>
  ): void {
    if (this.isDevelopment) {
      console.log(`[EVENT] ${name}`, properties || '', measurements || '');
    }
    if (this.appInsightsEnabled) {
      this.enqueue(
        'Microsoft.ApplicationInsights.Event',
        'EventData',
        {
          ver: 2,
          name,
          properties: properties || {},
          measurements: measurements || {},
        }
      );
    }
  }

  public trackMetric(
    name: string,
    average: number,
    properties?: Record<string, unknown>
  ): void {
    if (this.isDevelopment) {
      console.log(`[METRIC] ${name}: ${average}`, properties || '');
    }
    if (this.appInsightsEnabled) {
      this.enqueue(
        'Microsoft.ApplicationInsights.Metric',
        'MetricData',
        {
          ver: 2,
          metrics: [{ name, value: average, count: 1 }],
          properties: properties || {},
        }
      );
    }
  }

  public trackException(
    error: Error,
    severityLevel?: SeverityLevel,
    properties?: Record<string, unknown>
  ): void {
    console.error(`[EXCEPTION] ${error.message}`, error, severityLevel, properties || '');
    if (this.appInsightsEnabled) {
      this.enqueue(
        'Microsoft.ApplicationInsights.Exception',
        'ExceptionData',
        {
          ver: 2,
          exceptions: [
            {
              typeName: error.name || 'Error',
              message: error.message,
              stack: error.stack || '',
              hasFullStack: !!error.stack,
            },
          ],
          severityLevel: severityLevel ?? SeverityLevel.Error,
          properties: properties || {},
        }
      );
    }
  }

  public trackPageView(
    name: string,
    url?: string,
    properties?: Record<string, unknown>
  ): void {
    if (this.isDevelopment) {
      console.log(`[PAGEVIEW] ${name}`, url || '', properties || '');
    }
    if (this.appInsightsEnabled) {
      this.enqueue(
        'Microsoft.ApplicationInsights.PageView',
        'PageviewData',
        {
          ver: 2,
          name,
          url: url || window.location.href,
          properties: properties || {},
        }
      );
    }
  }

  public trackDependency(
    name: string,
    duration: number,
    success: boolean,
    resultCode?: number,
    properties?: Record<string, unknown>
  ): void {
    if (this.isDevelopment) {
      console.log(`[DEPENDENCY] ${name}`, { duration, success, resultCode }, properties || '');
    }
    if (this.appInsightsEnabled) {
      this.enqueue(
        'Microsoft.ApplicationInsights.RemoteDependency',
        'RemoteDependencyData',
        {
          ver: 2,
          name,
          duration: this.formatDuration(duration),
          success,
          resultCode: resultCode ?? (success ? 200 : 500),
          type: 'HTTP',
          properties: properties || {},
        }
      );
    }
  }

  public setUserId(userId: string): void {
    this.userId = userId;
    if (this.isDevelopment) {
      console.log(`[USER] Set user ID: ${userId}`);
    }
  }

  /**
   * Flush all pending telemetry to Application Insights.
   * Uses navigator.sendBeacon for reliable delivery on page unload.
   */
  public flush(): void {
    if (!this.appInsightsEnabled || this.pendingEnvelopes.length === 0) return;

    const envelopes = [...this.pendingEnvelopes];
    this.pendingEnvelopes = [];

    const endpoint = `${this.ingestionEndpoint}/v2/track`;
    const payload = envelopes.map((e) => JSON.stringify(e)).join('\n');

    try {
      if (typeof navigator.sendBeacon === 'function') {
        navigator.sendBeacon(endpoint, payload);
      } else {
        // Fallback for older browsers
        const xhr = new XMLHttpRequest();
        xhr.open('POST', endpoint, true);
        xhr.setRequestHeader('Content-Type', 'application/json');
        xhr.send(payload);
      }
    } catch (err) {
      // Silently ignore telemetry failures — never break the app
      if (this.isDevelopment) {
        console.warn('[LoggingService] Failed to send telemetry:', err);
      }
    }
  }

  // ---------------------------------------------------------------------------
  // Private helpers
  // ---------------------------------------------------------------------------

  private sendTrace(
    message: string,
    severityLevel: SeverityLevel,
    properties?: Record<string, unknown>
  ): void {
    if (!this.appInsightsEnabled) return;
    this.enqueue(
      'Microsoft.ApplicationInsights.Message',
      'MessageData',
      {
        ver: 2,
        message,
        severityLevel,
        properties: properties || {},
      }
    );
  }

  private enqueue(name: string, baseType: string, baseData: Record<string, unknown>): void {
    const envelope: IAppInsightsEnvelope = {
      name,
      time: new Date().toISOString(),
      iKey: this.instrumentationKey,
      tags: {
        'ai.session.id': this.sessionId,
        'ai.user.id': this.userId || 'anonymous',
        'ai.application.ver': '1.2.3',
        'ai.cloud.roleName': 'PolicyManager-SPFx',
        'ai.device.type': 'Browser',
        'ai.operation.name': window.location.pathname,
      },
      data: { baseType, baseData },
    };

    this.pendingEnvelopes.push(envelope);

    // Auto-flush when batch reaches 25 items
    if (this.pendingEnvelopes.length >= 25) {
      this.flush();
    }
  }

  private generateSessionId(): string {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
      const r = (Math.random() * 16) | 0;
      const v = c === 'x' ? r : (r & 0x3) | 0x8;
      return v.toString(16);
    });
  }

  /**
   * Format milliseconds into Application Insights duration format: "HH:MM:SS.mmm"
   */
  private formatDuration(ms: number): string {
    const hours = Math.floor(ms / 3600000);
    const minutes = Math.floor((ms % 3600000) / 60000);
    const seconds = Math.floor((ms % 60000) / 1000);
    const millis = Math.floor(ms % 1000);
    return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}.${String(millis).padStart(3, '0')}`;
  }
}

// Export logger object for compatibility
export const logger = {
  info: (source: string, message: string, data?: unknown): void => {
    LoggingService.getInstance().info(`[${source}] ${message}`, data as Record<string, unknown>);
  },
  warn: (source: string, message: string, data?: unknown): void => {
    LoggingService.getInstance().warn(source, message, data as Record<string, unknown>);
  },
  error: (source: string, message: string, error?: unknown): void => {
    LoggingService.getInstance().error(source, message, error as Error);
  },
  debug: (source: string, message: string, data?: unknown): void => {
    LoggingService.getInstance().verbose(`[${source}] ${message}`, data as Record<string, unknown>);
  },
};

export default LoggingService;
