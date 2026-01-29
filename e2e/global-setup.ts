import { chromium, FullConfig } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';

/**
 * Global setup for Playwright tests
 * Handles M365 authentication and saves state for reuse
 */
async function globalSetup(config: FullConfig): Promise<void> {
  const authFile = path.join(__dirname, '.auth', 'user.json');

  // Check if we have valid stored auth
  if (fs.existsSync(authFile)) {
    const stats = fs.statSync(authFile);
    const ageInHours = (Date.now() - stats.mtimeMs) / (1000 * 60 * 60);

    // Reuse auth if less than 1 hour old
    if (ageInHours < 1) {
      console.log('Using cached authentication state');
      return;
    }
  }

  console.log('Performing fresh M365 authentication...');

  const browser = await chromium.launch({ headless: false }); // Visible for MFA if needed
  const context = await browser.newContext();
  const page = await context.newPage();

  try {
    // Navigate to SharePoint site - will redirect to M365 login
    await page.goto('https://mf7m.sharepoint.com/sites/PolicyManager');

    // Wait for login page
    await page.waitForSelector('input[type="email"]', { timeout: 30000 });

    // Enter email
    await page.fill('input[type="email"]', process.env.M365_USERNAME || 'gf_admin@mf7m.onmicrosoft.com');
    await page.click('input[type="submit"]');

    // Wait for password field
    await page.waitForSelector('input[type="password"]', { timeout: 30000 });

    // Enter password
    await page.fill('input[type="password"]', process.env.M365_PASSWORD || '');
    await page.click('input[type="submit"]');

    // Handle "Stay signed in?" prompt if it appears
    try {
      await page.waitForSelector('input[value="Yes"]', { timeout: 5000 });
      await page.click('input[value="Yes"]');
    } catch {
      // Prompt didn't appear, continue
    }

    // Wait for SharePoint to load (indicates successful auth)
    await page.waitForURL('**/sites/PolicyManager**', { timeout: 60000 });

    // Additional wait for page to fully load
    await page.waitForLoadState('networkidle');

    console.log('Authentication successful!');

    // Save authentication state
    await context.storageState({ path: authFile });
    console.log(`Auth state saved to ${authFile}`);

  } catch (error) {
    console.error('Authentication failed:', error);
    throw error;
  } finally {
    await browser.close();
  }
}

export default globalSetup;
