import { test } from '@playwright/test';
import * as path from 'path';

test('Click Manager and screenshot dropdown', async ({ page }) => {
  await page.goto('https://mf7m.sharepoint.com/sites/PolicyManager/SitePages/PolicyManagerView.aspx', { timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  // Click Manager button
  const managerBtn = page.locator('button').filter({ hasText: /^Manager$/ }).first();
  await managerBtn.click();
  await page.waitForTimeout(2000);

  // Take screenshot to see the dropdown
  await page.screenshot({ path: path.join(process.cwd(), 'e2e-manager-dropdown.png') });
  console.log('📸 e2e-manager-dropdown.png');

  // Now scan for ANY new elements that appeared (custom dropdown/flyout)
  const allElements = await page.evaluate(() => {
    const results: string[] = [];
    // Look for any visible dropdown, flyout, menu, or popover
    const selectors = [
      '[class*="dropdown"]', '[class*="flyout"]', '[class*="menu"]',
      '[class*="popover"]', '[class*="callout"]', '[class*="panel"]',
      '[class*="Dropdown"]', '[class*="Callout"]', '[class*="Flyout"]'
    ];
    for (const sel of selectors) {
      const els = document.querySelectorAll(sel);
      for (const el of els) {
        const rect = el.getBoundingClientRect();
        if (rect.width > 0 && rect.height > 0) {
          const text = (el.textContent || '').trim().slice(0, 80);
          results.push(`${sel}: ${el.tagName} ${rect.width}x${rect.height} at y=${rect.y} "${text}"`);
        }
      }
    }
    return results;
  });

  console.log(`\nVisible dropdown/flyout/menu elements:`);
  for (const el of allElements) {
    console.log(`  ${el}`);
  }

  // Also check for elements with "Request" in text that just appeared
  const requestElements = page.locator('*').filter({ hasText: /Request Policy/i });
  const reqCount = await requestElements.count();
  console.log(`\n"Request Policy" elements visible: ${reqCount}`);
  for (let i = 0; i < Math.min(reqCount, 5); i++) {
    const text = (await requestElements.nth(i).textContent() || '').trim().slice(0, 60);
    const tag = await requestElements.nth(i).evaluate(el => el.tagName);
    console.log(`  [${i}] <${tag}> "${text}"`);
  }
});
