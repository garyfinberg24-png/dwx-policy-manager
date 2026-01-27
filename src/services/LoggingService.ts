// @ts-nocheck
/**
 * LoggingService - Standalone console-based logging for Policy Manager
 * No Application Insights dependency
 */

// Enum for log severity (compatible with original interface)
export enum SeverityLevel {
  Verbose = 0,
  Information = 1,
  Warning = 2,
  Error = 3,
  Critical = 4
}

export class LoggingService {
  private static instance: LoggingService;
  private isProduction: boolean = false;
  private isDevelopment: boolean = true;

  private constructor() {
    this.isProduction = window.location.hostname !== 'localhost' &&
                        !window.location.hostname.includes('.sharepoint.com');
    this.isDevelopment = !this.isProduction;
  }

  public static getInstance(): LoggingService {
    if (!LoggingService.instance) {
      LoggingService.instance = new LoggingService();
    }
    return LoggingService.instance;
  }

  public initializeAppInsights(_instrumentationKey: string): void {
    // No-op for standalone version
    console.log('[LoggingService] App Insights not available in standalone version');
  }

  public info(message: string, properties?: Record<string, unknown>): void {
    if (this.isDevelopment) {
      console.log(`[INFO] ${message}`, properties || '');
    }
  }

  public warn(source: string, message: string, properties?: Record<string, unknown>): void {
    console.warn(`[WARN] [${source}] ${message}`, properties || '');
  }

  public error(source: string, message: string, error?: Error, properties?: Record<string, unknown>): void {
    console.error(`[ERROR] [${source}] ${message}`, error || '', properties || '');
  }

  public verbose(message: string, properties?: Record<string, unknown>): void {
    if (this.isDevelopment) {
      console.debug(`[VERBOSE] ${message}`, properties || '');
    }
  }

  public trackEvent(name: string, properties?: Record<string, unknown>, measurements?: Record<string, number>): void {
    if (this.isDevelopment) {
      console.log(`[EVENT] ${name}`, properties || '', measurements || '');
    }
  }

  public trackMetric(name: string, average: number, properties?: Record<string, unknown>): void {
    if (this.isDevelopment) {
      console.log(`[METRIC] ${name}: ${average}`, properties || '');
    }
  }

  public trackException(error: Error, severityLevel?: SeverityLevel, properties?: Record<string, unknown>): void {
    console.error(`[EXCEPTION] ${error.message}`, error, severityLevel, properties || '');
  }

  public trackPageView(name: string, url?: string, properties?: Record<string, unknown>): void {
    if (this.isDevelopment) {
      console.log(`[PAGEVIEW] ${name}`, url || '', properties || '');
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
  }

  public setUserId(userId: string): void {
    if (this.isDevelopment) {
      console.log(`[USER] Set user ID: ${userId}`);
    }
  }

  public flush(): void {
    // No-op for console-based logging
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
  }
};

export default LoggingService;
