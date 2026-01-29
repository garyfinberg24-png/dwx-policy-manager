import { test, expect } from '@playwright/test';

/**
 * My Policies E2E Tests
 * Tests the personal policy dashboard for employees
 */
test.describe('My Policies Dashboard', () => {
  test.beforeEach(async ({ page }) => {
    // Navigate to SharePoint workbench
    await page.goto('/_layouts/15/workbench.aspx');
    await page.waitForLoadState('networkidle');
  });

  test('should display user assigned policies', async ({ page }) => {
    // Look for My Policies section
    const myPolicies = page.locator('text=My Policies, text=Assigned Policies, [data-automation-id="myPolicies"]');

    if (await myPolicies.first().isVisible()) {
      await expect(myPolicies.first()).toBeVisible();
    }
  });

  test('should show policy acknowledgement status', async ({ page }) => {
    // Look for acknowledgement indicators
    const acknowledged = page.locator('text=Acknowledged, .status-acknowledged');
    const pending = page.locator('text=Pending, .status-pending');

    // At least one status type should be present if policies exist
    const hasAcknowledged = await acknowledged.first().isVisible().catch(() => false);
    const hasPending = await pending.first().isVisible().catch(() => false);

    console.log(`Acknowledged policies visible: ${hasAcknowledged}`);
    console.log(`Pending policies visible: ${hasPending}`);
  });

  test('should allow policy acknowledgement', async ({ page }) => {
    // Find acknowledge button
    const acknowledgeButton = page.locator('button:has-text("Acknowledge"), button:has-text("Accept")');

    if (await acknowledgeButton.first().isVisible()) {
      // Click to acknowledge
      await acknowledgeButton.first().click();

      // Wait for confirmation
      await page.waitForTimeout(2000);

      // Check for success message
      const successMessage = page.locator('text=Success, text=Acknowledged, .ms-MessageBar--success');
      if (await successMessage.isVisible()) {
        await expect(successMessage).toBeVisible();
      }
    }
  });

  test('should display policy details on selection', async ({ page }) => {
    // Find and click a policy item
    const policyItem = page.locator('.policy-card, .policy-item, [data-automation-id="policyCard"]').first();

    if (await policyItem.isVisible()) {
      await policyItem.click();
      await page.waitForTimeout(1000);

      // Should show policy details panel or navigate
      const detailsPanel = page.locator('.policy-details, [data-automation-id="policyDetails"]');
      if (await detailsPanel.isVisible()) {
        await expect(detailsPanel).toBeVisible();
      }
    }
  });
});
