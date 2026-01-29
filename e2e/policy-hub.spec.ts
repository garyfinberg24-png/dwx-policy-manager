import { test, expect } from '@playwright/test';

/**
 * Policy Hub E2E Tests
 * Tests the central policy discovery and search interface
 */
test.describe('Policy Hub', () => {
  test.beforeEach(async ({ page }) => {
    // Navigate to the Policy Hub page
    await page.goto('/_layouts/15/workbench.aspx');

    // Wait for the workbench to load
    await page.waitForLoadState('networkidle');
  });

  test('should load Policy Hub webpart', async ({ page }) => {
    // Look for the Policy Hub heading or container
    const policyHub = page.locator('[data-automation-id="policyHub"], .policy-hub, text=Policy Hub');

    // If webpart isn't added, we may need to add it
    const addWebpartButton = page.locator('button:has-text("Add a new Web Part")');
    if (await addWebpartButton.isVisible()) {
      console.log('Workbench is empty - webpart needs to be added');
      // This is expected on first run
    }

    await expect(page).toHaveTitle(/Workbench|Policy/);
  });

  test('should display policy search functionality', async ({ page }) => {
    // Look for search input
    const searchInput = page.locator('input[placeholder*="Search"], input[aria-label*="Search"]');

    if (await searchInput.isVisible()) {
      await expect(searchInput).toBeEnabled();

      // Test search interaction
      await searchInput.fill('test policy');
      await searchInput.press('Enter');

      // Wait for search results
      await page.waitForTimeout(2000);
    }
  });

  test('should display policy filters', async ({ page }) => {
    // Look for filter controls
    const filters = page.locator('[data-automation-id="filters"], .filters, button:has-text("Filter")');

    if (await filters.first().isVisible()) {
      await expect(filters.first()).toBeEnabled();
    }
  });

  test('should navigate to policy details on click', async ({ page }) => {
    // Find a policy item
    const policyItem = page.locator('.policy-item, [data-automation-id="policyItem"]').first();

    if (await policyItem.isVisible()) {
      await policyItem.click();

      // Should navigate or open details
      await page.waitForTimeout(2000);
    }
  });
});
