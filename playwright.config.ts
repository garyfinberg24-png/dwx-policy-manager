import { defineConfig, devices } from '@playwright/test';

/**
 * Playwright configuration for Policy Manager SPFx E2E tests
 * Tests run against SharePoint Online
 */
export default defineConfig({
  testDir: './e2e',
  fullyParallel: false, // Run tests sequentially for SharePoint
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 2 : 0,
  workers: 1, // Single worker for authenticated SharePoint sessions
  reporter: [
    ['html', { outputFolder: 'playwright-report' }],
    ['list']
  ],

  use: {
    // SharePoint Online base URL
    baseURL: 'https://mf7m.sharepoint.com/sites/PolicyManager',

    // Capture traces and screenshots on failure
    trace: 'on-first-retry',
    screenshot: 'only-on-failure',
    video: 'on-first-retry',

    // Longer timeouts for SharePoint
    actionTimeout: 30000,
    navigationTimeout: 60000,

    // Use stored authentication state
    storageState: './e2e/.auth/user.json',
  },

  // Global setup for authentication
  globalSetup: './e2e/global-setup.ts',

  projects: [
    // Setup project for authentication
    {
      name: 'setup',
      testMatch: /.*\.setup\.ts/,
      use: {
        storageState: undefined, // Don't use stored state during setup
      },
    },

    // Main test project
    {
      name: 'chromium',
      use: {
        ...devices['Desktop Chrome'],
        viewport: { width: 1920, height: 1080 },
      },
      dependencies: ['setup'],
    },
  ],

  // Timeout for each test
  timeout: 120000,

  // Expect timeout
  expect: {
    timeout: 30000,
  },
});
