import { test as setup, expect } from '@playwright/test';
import * as path from 'path';

const authFile = path.join(__dirname, '.auth', 'user.json');

/**
 * Authentication setup test
 * Runs before all other tests to establish M365 session
 */
setup('authenticate with Microsoft 365', async ({ page }) => {
  // Navigate to SharePoint - triggers M365 login
  await page.goto('https://mf7m.sharepoint.com/sites/PolicyManager');

  // Check if already authenticated (storage state loaded)
  const isLoggedIn = await page.url().includes('sharepoint.com/sites/PolicyManager');

  if (!isLoggedIn || page.url().includes('login.microsoftonline.com')) {
    console.log('Performing M365 login...');

    // Wait for and fill email
    await page.waitForSelector('input[type="email"]', { timeout: 30000 });
    await page.fill('input[type="email"]', process.env.M365_USERNAME || 'gf_admin@mf7m.onmicrosoft.com');
    await page.click('input[type="submit"]');

    // Wait for and fill password
    await page.waitForSelector('input[type="password"]', { timeout: 30000 });
    await page.fill('input[type="password"]', process.env.M365_PASSWORD || '');
    await page.click('input[type="submit"]');

    // Handle "Stay signed in?" prompt
    try {
      const staySignedIn = page.locator('input[value="Yes"]');
      await staySignedIn.waitFor({ timeout: 10000 });
      await staySignedIn.click();
    } catch {
      // Prompt may not appear
    }

    // Wait for SharePoint to load
    await page.waitForURL('**/sites/PolicyManager**', { timeout: 60000 });
  }

  // Verify we're on SharePoint
  await expect(page).toHaveURL(/sharepoint\.com\/sites\/PolicyManager/);

  // Save authentication state
  await page.context().storageState({ path: authFile });
  console.log('Authentication state saved');
});
