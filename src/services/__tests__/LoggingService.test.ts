/**
 * LoggingService Unit Tests
 *
 * Tests the singleton pattern, log level methods, App Insights envelope creation,
 * flush behavior, and the exported logger compatibility object.
 */

import { LoggingService, SeverityLevel, logger } from '../LoggingService';

// ---------------------------------------------------------------------------
// Helpers — reset singleton between tests
// ---------------------------------------------------------------------------

/**
 * Force-reset the private static `instance` so each test gets a fresh
 * LoggingService.  We use bracket-notation to access the private field.
 */
function resetSingleton(): void {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (LoggingService as any).instance = undefined;
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe('LoggingService', () => {
  let service: LoggingService;

  beforeEach(() => {
    resetSingleton();
    service = LoggingService.getInstance();
    jest.restoreAllMocks();
  });

  afterEach(() => {
    service.dispose();
    resetSingleton();
  });

  // ===== Singleton =====

  describe('Singleton pattern', () => {
    it('should return the same instance on repeated calls', () => {
      const a = LoggingService.getInstance();
      const b = LoggingService.getInstance();
      expect(a).toBe(b);
    });

    it('should return a new instance after reset', () => {
      const a = LoggingService.getInstance();
      resetSingleton();
      const b = LoggingService.getInstance();
      expect(a).not.toBe(b);
    });
  });

  // ===== Log level methods (console-only mode) =====

  describe('Console-only logging', () => {
    it('info() should log to console.log in dev mode', () => {
      const spy = jest.spyOn(console, 'log').mockImplementation();
      service.info('test message', { key: 'value' });
      expect(spy).toHaveBeenCalledWith(
        '[INFO] test message',
        { key: 'value' }
      );
    });

    it('warn() should log to console.warn', () => {
      const spy = jest.spyOn(console, 'warn').mockImplementation();
      service.warn('Source', 'warning message');
      expect(spy).toHaveBeenCalledWith(
        '[WARN] [Source] warning message',
        ''
      );
    });

    it('error() should log to console.error', () => {
      const spy = jest.spyOn(console, 'error').mockImplementation();
      service.error('Source', 'error message');
      expect(spy).toHaveBeenCalledWith(
        '[ERROR] [Source] error message',
        '',
        ''
      );
    });

    it('error() with an Error object should log the error', () => {
      const spy = jest.spyOn(console, 'error').mockImplementation();
      const err = new Error('boom');
      service.error('Source', 'error message', err);
      expect(spy).toHaveBeenCalled();
      // The first console.error call is from error()
      const firstCallArgs = spy.mock.calls[0];
      expect(firstCallArgs[0]).toContain('[ERROR]');
    });

    it('verbose() should log to console.debug in dev mode', () => {
      const spy = jest.spyOn(console, 'debug').mockImplementation();
      service.verbose('verbose message');
      expect(spy).toHaveBeenCalledWith(
        '[VERBOSE] verbose message',
        ''
      );
    });
  });

  // ===== App Insights initialization =====

  describe('initializeAppInsights', () => {
    it('should stay in console-only mode with empty string', () => {
      const spy = jest.spyOn(console, 'log').mockImplementation();
      service.initializeAppInsights('');
      expect(spy).toHaveBeenCalledWith(
        '[LoggingService] No connection string — console-only mode'
      );
    });

    it('should parse a full connection string', () => {
      jest.spyOn(console, 'log').mockImplementation();
      service.initializeAppInsights(
        'InstrumentationKey=abc-123;IngestionEndpoint=https://custom.endpoint.com'
      );
      // Verify appInsightsEnabled is true by checking that trackEvent enqueues
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      expect((service as any).appInsightsEnabled).toBe(true);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      expect((service as any).instrumentationKey).toBe('abc-123');
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      expect((service as any).ingestionEndpoint).toBe('https://custom.endpoint.com');
    });

    it('should accept a bare instrumentation key', () => {
      jest.spyOn(console, 'log').mockImplementation();
      service.initializeAppInsights('abcd-1234-efgh-5678');
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      expect((service as any).appInsightsEnabled).toBe(true);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      expect((service as any).instrumentationKey).toBe('abcd-1234-efgh-5678');
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      expect((service as any).ingestionEndpoint).toBe('https://dc.services.visualstudio.com');
    });
  });

  // ===== Envelope creation & queuing =====

  describe('Envelope creation (App Insights enabled)', () => {
    beforeEach(() => {
      jest.spyOn(console, 'log').mockImplementation();
      jest.spyOn(console, 'debug').mockImplementation();
      service.initializeAppInsights('test-ikey-1234');
    });

    it('trackEvent should enqueue an EventData envelope', () => {
      service.trackEvent('PolicyAcknowledged', { policyId: '42' });
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const pending: any[] = (service as any).pendingEnvelopes;
      expect(pending.length).toBe(1);
      expect(pending[0].name).toBe('Microsoft.ApplicationInsights.Event');
      expect(pending[0].data.baseType).toBe('EventData');
      expect(pending[0].data.baseData.name).toBe('PolicyAcknowledged');
    });

    it('trackPageView should enqueue a PageviewData envelope', () => {
      service.trackPageView('PolicyHub', 'https://test.sharepoint.com/PolicyHub');
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const pending: any[] = (service as any).pendingEnvelopes;
      expect(pending.length).toBe(1);
      expect(pending[0].data.baseType).toBe('PageviewData');
      expect(pending[0].data.baseData.name).toBe('PolicyHub');
      expect(pending[0].data.baseData.url).toBe('https://test.sharepoint.com/PolicyHub');
    });

    it('trackMetric should enqueue a MetricData envelope', () => {
      service.trackMetric('LoadTime', 1500);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const pending: any[] = (service as any).pendingEnvelopes;
      expect(pending.length).toBe(1);
      expect(pending[0].data.baseType).toBe('MetricData');
      expect(pending[0].data.baseData.metrics[0].value).toBe(1500);
    });

    it('trackException should enqueue an ExceptionData envelope', () => {
      jest.spyOn(console, 'error').mockImplementation();
      const error = new Error('Test error');
      service.trackException(error, SeverityLevel.Error);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const pending: any[] = (service as any).pendingEnvelopes;
      expect(pending.length).toBe(1);
      expect(pending[0].data.baseType).toBe('ExceptionData');
      expect(pending[0].data.baseData.exceptions[0].message).toBe('Test error');
    });

    it('trackDependency should enqueue a RemoteDependencyData envelope', () => {
      service.trackDependency('GET /api/policies', 350, true, 200);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const pending: any[] = (service as any).pendingEnvelopes;
      expect(pending.length).toBe(1);
      expect(pending[0].data.baseType).toBe('RemoteDependencyData');
      expect(pending[0].data.baseData.success).toBe(true);
    });

    it('envelope should include session and user tags', () => {
      service.setUserId('user@company.com');
      service.trackEvent('TestEvent');
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const pending: any[] = (service as any).pendingEnvelopes;
      expect(pending[0].tags['ai.user.id']).toBe('user@company.com');
      expect(pending[0].tags['ai.session.id']).toBeDefined();
      expect(pending[0].tags['ai.cloud.roleName']).toBe('PolicyManager-SPFx');
    });
  });

  // ===== Flush behavior =====

  describe('Flush behavior', () => {
    beforeEach(() => {
      jest.spyOn(console, 'log').mockImplementation();
      jest.spyOn(console, 'debug').mockImplementation();
      service.initializeAppInsights('test-ikey-1234');
    });

    it('flush should call navigator.sendBeacon with enqueued envelopes', () => {
      const beaconSpy = jest.spyOn(navigator, 'sendBeacon').mockReturnValue(true);
      service.trackEvent('E1');
      service.trackEvent('E2');
      service.flush();

      expect(beaconSpy).toHaveBeenCalledTimes(1);
      const payload = beaconSpy.mock.calls[0][1] as string;
      expect(payload).toContain('E1');
      expect(payload).toContain('E2');
    });

    it('flush should clear the pending queue', () => {
      jest.spyOn(navigator, 'sendBeacon').mockReturnValue(true);
      service.trackEvent('E1');
      service.flush();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      expect((service as any).pendingEnvelopes.length).toBe(0);
    });

    it('flush should do nothing when no envelopes are pending', () => {
      // Reset singleton so there are zero pending envelopes
      service.dispose();
      resetSingleton();
      service = LoggingService.getInstance();
      jest.spyOn(console, 'log').mockImplementation();
      service.initializeAppInsights('test-ikey-1234');

      // Clear any prior sendBeacon calls and spy fresh
      jest.restoreAllMocks();
      jest.spyOn(console, 'log').mockImplementation();
      jest.spyOn(console, 'debug').mockImplementation();
      const beaconSpy = jest.spyOn(navigator, 'sendBeacon').mockReturnValue(true);

      service.flush();
      expect(beaconSpy).not.toHaveBeenCalled();
    });

    it('should auto-flush when reaching 25 envelopes', () => {
      // Reset singleton for a clean count
      service.dispose();
      resetSingleton();
      service = LoggingService.getInstance();
      jest.spyOn(console, 'log').mockImplementation();
      service.initializeAppInsights('test-ikey-1234');

      jest.restoreAllMocks();
      jest.spyOn(console, 'log').mockImplementation();
      jest.spyOn(console, 'debug').mockImplementation();
      const beaconSpy = jest.spyOn(navigator, 'sendBeacon').mockReturnValue(true);

      for (let i = 0; i < 25; i++) {
        service.trackEvent(`Event${i}`);
      }
      // Auto-flush happens at 25
      expect(beaconSpy).toHaveBeenCalledTimes(1);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      expect((service as any).pendingEnvelopes.length).toBe(0);
    });
  });

  // ===== SeverityLevel enum =====

  describe('SeverityLevel enum', () => {
    it('should have correct numeric values', () => {
      expect(SeverityLevel.Verbose).toBe(0);
      expect(SeverityLevel.Information).toBe(1);
      expect(SeverityLevel.Warning).toBe(2);
      expect(SeverityLevel.Error).toBe(3);
      expect(SeverityLevel.Critical).toBe(4);
    });
  });

  // ===== logger compatibility object =====

  describe('logger export', () => {
    it('logger.info should call LoggingService.info', () => {
      const spy = jest.spyOn(console, 'log').mockImplementation();
      logger.info('TestSource', 'hello info');
      expect(spy).toHaveBeenCalled();
    });

    it('logger.warn should call LoggingService.warn', () => {
      const spy = jest.spyOn(console, 'warn').mockImplementation();
      logger.warn('TestSource', 'hello warn');
      expect(spy).toHaveBeenCalled();
    });

    it('logger.error should call LoggingService.error', () => {
      const spy = jest.spyOn(console, 'error').mockImplementation();
      logger.error('TestSource', 'hello error');
      expect(spy).toHaveBeenCalled();
    });

    it('logger.debug should call LoggingService.verbose', () => {
      const spy = jest.spyOn(console, 'debug').mockImplementation();
      logger.debug('TestSource', 'hello debug');
      expect(spy).toHaveBeenCalled();
    });
  });

  // ===== setUserId =====

  describe('setUserId', () => {
    it('should set the userId on the service', () => {
      jest.spyOn(console, 'log').mockImplementation();
      service.setUserId('john@company.com');
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      expect((service as any).userId).toBe('john@company.com');
    });
  });

  // ===== dispose =====

  describe('dispose', () => {
    it('should not throw when called multiple times', () => {
      expect(() => {
        service.dispose();
        service.dispose();
      }).not.toThrow();
    });
  });
});
