import { test } from '@playwright/test';

test('Find Manager nav button', async ({ page }) => {
  // Go to Manager View page (guaranteed to have the full nav)
  await page.goto('https://mf7m.sharepoint.com/sites/PolicyManager/SitePages/PolicyManagerView.aspx', { timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  // Get ALL buttons and log them
  const allButtons = await page.locator('button').allTextContents();
  const buttonTexts = allButtons.map(t => t.trim()).filter(t => t.length > 0 && t.length < 30);
  console.log(`All buttons (${buttonTexts.length}):`);
  for (const t of buttonTexts) console.log(`  "${t}"`);

  // Find Manager button specifically
  const managerBtns = page.locator('button').filter({ hasText: /^Manager$/ });
  const count = await managerBtns.count();
  console.log(`\nManager buttons (exact): ${count}`);

  // Try clicking it
  if (count > 0) {
    await managerBtns.first().click();
    await page.waitForTimeout(1500);

    // Log what appeared (dropdown items)
    const newElements = await page.locator('a, [role="menuitem"], [role="option"]').allTextContents();
    const items = newElements.map(t => t.trim()).filter(t => t.length > 0 && t.length < 40);
    console.log(`\nDropdown items after click (${items.length}):`);
    for (const t of items) console.log(`  "${t}"`);
  }
});
