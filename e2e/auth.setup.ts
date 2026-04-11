import { test as setup, expect } from '@playwright/test';
import * as path from 'path';
import * as fs from 'fs';

const authFile = path.join(__dirname, '.auth', 'user.json');

/**
 * Authentication setup test
 * Verifies auth state is valid and refreshes if needed.
 * The heavy lifting is done in global-setup.ts (interactive login).
 * This just validates the saved state works.
 */
setup('verify M365 authentication', async ({ page }) => {
  // Check if auth file exists (should have been created by global-setup)
  if (!fs.existsSync(authFile)) {
    console.log('No auth state found — global-setup should have created it');
    setup.skip();
    return;
  }

  // Navigate to SharePoint with saved auth state
  await page.goto('https://mf7m.sharepoint.com/sites/PolicyManager', { timeout: 60000 });

  // If redirected to login, auth state is stale
  const currentUrl = page.url();
  if (currentUrl.includes('login.microsoftonline.com') || currentUrl.includes('login.live.com')) {
    console.log('Auth state expired — need fresh login');
    // Delete stale auth file so global-setup re-runs next time
    fs.unlinkSync(authFile);
    setup.skip();
    return;
  }

  // Wait for SharePoint to load
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});

  // Verify we're on SharePoint
  await expect(page).toHaveURL(/sharepoint\.com\/sites\/PolicyManager/);

  // Re-save fresh auth state
  await page.context().storageState({ path: authFile });
  console.log('✅ Authentication verified and state saved');
});
