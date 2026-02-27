// Global setup for Jest tests
// Provides minimal browser globals that SPFx code may reference

// Mock localStorage
const localStorageMock = (() => {
  let store = {};
  return {
    getItem: (key) => (key in store ? store[key] : null),
    setItem: (key, value) => { store[key] = String(value); },
    removeItem: (key) => { delete store[key]; },
    clear: () => { store = {}; },
    get length() { return Object.keys(store).length; },
    key: (index) => Object.keys(store)[index] || null,
  };
})();

Object.defineProperty(window, 'localStorage', { value: localStorageMock });

// Mock navigator.sendBeacon — use a plain function (not jest.fn) so that
// jest.spyOn(navigator, 'sendBeacon') in tests can wrap it cleanly.
if (!navigator.sendBeacon) {
  Object.defineProperty(navigator, 'sendBeacon', {
    value: function sendBeacon() { return true; },
    writable: true,
    configurable: true,
  });
}

// Ensure window.location is available (jsdom provides it, but ensure hostname)
if (!window.location.hostname) {
  Object.defineProperty(window, 'location', {
    value: {
      hostname: 'localhost',
      href: 'http://localhost',
      pathname: '/',
      search: '',
      hash: '',
    },
    writable: true,
  });
}
