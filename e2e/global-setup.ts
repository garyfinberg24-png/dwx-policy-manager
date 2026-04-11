import { chromium, FullConfig } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';

/**
 * Global setup for Playwright tests
 * Handles M365 authentication and saves state for reuse.
 *
 * Launches a VISIBLE browser for manual login (MFA, password, etc.)
 * Then saves session cookies/storage for test reuse.
 */
async function globalSetup(config: FullConfig): Promise<void> {
  const authDir = path.join(__dirname, '.auth');
  const authFile = path.join(authDir, 'user.json');

  // Ensure auth directory exists
  if (!fs.existsSync(authDir)) {
    fs.mkdirSync(authDir, { recursive: true });
  }

  // Check if we have valid stored auth
  if (fs.existsSync(authFile)) {
    const stats = fs.statSync(authFile);
    const ageInHours = (Date.now() - stats.mtimeMs) / (1000 * 60 * 60);

    // Reuse auth if less than 2 hours old
    if (ageInHours < 2) {
      console.log(`Using cached authentication state (${Math.round(ageInHours * 60)}min old)`);
      return;
    }
  }

  console.log('========================================');
  console.log('  M365 Authentication Required');
  console.log('  A browser window will open.');
  console.log('  Please log in manually (including MFA).');
  console.log('  The browser will close automatically');
  console.log('  once SharePoint loads.');
  console.log('========================================');

  const browser = await chromium.launch({
    headless: false,
    slowMo: 100, // Slow down slightly for login reliability
  });
  const context = await browser.newContext({
    viewport: { width: 1280, height: 720 },
  });
  const page = await context.newPage();

  try {
    // Navigate to SharePoint site - will redirect to M365 login
    await page.goto('https://mf7m.sharepoint.com/sites/PolicyManager', { timeout: 30000 });

    // Give the user up to 3 minutes to complete login (MFA, password, etc.)
    console.log('Waiting for login to complete (up to 3 minutes)...');
    await page.waitForURL('**/sites/PolicyManager**', { timeout: 180000 });

    // Additional wait for page to fully load
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(3000);

    console.log('✅ Authentication successful!');

    // Save authentication state
    await context.storageState({ path: authFile });
    console.log(`Auth state saved to ${authFile}`);

  } catch (error) {
    console.error('❌ Authentication failed or timed out:', error);
    console.error('Please ensure you can log in to SharePoint manually.');
    throw error;
  } finally {
    await browser.close();
  }
}

export default globalSetup;
