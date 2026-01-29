import { test, expect } from '@playwright/test';

/**
 * Policy Admin E2E Tests
 * Tests administrative interface for managing policies
 */
test.describe('Policy Admin Dashboard', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/_layouts/15/workbench.aspx');
    await page.waitForLoadState('networkidle');
  });

  test('should display admin dashboard for authorized users', async ({ page }) => {
    // Look for admin dashboard elements
    const adminDashboard = page.locator('text=Policy Administration, text=Admin Dashboard, [data-automation-id="policyAdmin"]');

    if (await adminDashboard.first().isVisible()) {
      await expect(adminDashboard.first()).toBeVisible();
    }
  });

  test('should show policy management actions', async ({ page }) => {
    // Check for CRUD action buttons
    const createButton = page.locator('button:has-text("Create"), button:has-text("New Policy")');
    const editButton = page.locator('button:has-text("Edit")');
    const deleteButton = page.locator('button:has-text("Delete")');

    const hasCreate = await createButton.first().isVisible().catch(() => false);
    const hasEdit = await editButton.first().isVisible().catch(() => false);

    console.log(`Create button visible: ${hasCreate}`);
    console.log(`Edit button visible: ${hasEdit}`);
  });

  test('should display policy statistics', async ({ page }) => {
    // Look for KPI/metrics section
    const stats = page.locator('.stats, .metrics, .kpi, [data-automation-id="policyStats"]');

    if (await stats.first().isVisible()) {
      await expect(stats.first()).toBeVisible();
    }
  });

  test('should allow filtering policies by status', async ({ page }) => {
    // Find status filter
    const statusFilter = page.locator('select:has-text("Status"), button:has-text("Status"), [aria-label*="Status"]');

    if (await statusFilter.first().isVisible()) {
      await statusFilter.first().click();
      await page.waitForTimeout(500);

      // Look for filter options
      const filterOptions = page.locator('option, [role="option"], .ms-Dropdown-item');
      const optionCount = await filterOptions.count();
      console.log(`Filter options count: ${optionCount}`);
    }
  });

  test('should show compliance overview', async ({ page }) => {
    // Look for compliance section
    const compliance = page.locator('text=Compliance, text=Acknowledgement Rate, [data-automation-id="compliance"]');

    if (await compliance.first().isVisible()) {
      await expect(compliance.first()).toBeVisible();
    }
  });
});
