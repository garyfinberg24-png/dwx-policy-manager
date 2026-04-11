import { test, Page } from '@playwright/test';
import * as path from 'path';

const BASE = 'https://mf7m.sharepoint.com/sites/PolicyManager/SitePages';

async function snap(page: Page, name: string): Promise<void> {
  await page.screenshot({ path: path.join(process.cwd(), `e2e-reqdbg-${name}.png`), fullPage: false });
  console.log(`📸 e2e-reqdbg-${name}.png`);
}

test('Debug: Find Manager nav and Request Policy', async ({ page }) => {
  console.log('\n=== DEBUG: FINDING MANAGER NAV ===');

  // Navigate to Policy Hub (where nav is visible)
  await page.goto(`${BASE}/PolicyHub.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  // Dump ALL text content that contains "Manager" or "Request"
  const elements = await page.evaluate(() => {
    const results: string[] = [];
    const all = document.querySelectorAll('*');
    for (const el of all) {
      const text = (el.textContent || '').trim();
      if (text.length > 0 && text.length < 100) {
        if (text.includes('Manager') || text.includes('Request')) {
          const tag = el.tagName.toLowerCase();
          const role = el.getAttribute('role') || '';
          const cls = el.className?.toString().slice(0, 50) || '';
          const href = el.getAttribute('href') || '';
          results.push(`<${tag} role="${role}" class="${cls}" href="${href}"> "${text.slice(0, 60)}"`);
        }
      }
    }
    return results.slice(0, 30);
  });

  console.log(`  Elements containing "Manager" or "Request":`);
  for (const el of elements) {
    console.log(`    ${el}`);
  }

  await snap(page, 'debug-hub');

  // Try clicking "Manager" text directly on the page
  const managerElements = page.getByText('Manager');
  const managerCount = await managerElements.count();
  console.log(`\n  getByText('Manager') count: ${managerCount}`);

  for (let i = 0; i < Math.min(managerCount, 5); i++) {
    const bbox = await managerElements.nth(i).boundingBox().catch(() => null);
    const text = (await managerElements.nth(i).textContent() || '').trim().slice(0, 40);
    console.log(`    [${i}] "${text}" at ${bbox ? `x=${bbox.x},y=${bbox.y},w=${bbox.width}` : 'no bbox'}`);
  }

  // Click the first "Manager" that's in the top area of the page (nav bar)
  if (managerCount > 0) {
    for (let i = 0; i < managerCount; i++) {
      const bbox = await managerElements.nth(i).boundingBox().catch(() => null);
      if (bbox && bbox.y < 100) { // Nav bar is typically in the top 100px
        console.log(`\n  Clicking Manager nav at y=${bbox.y}`);
        await managerElements.nth(i).click();
        await page.waitForTimeout(1500);

        await snap(page, 'debug-manager-clicked');

        // Now look for "Request Policy" in any dropdown that appeared
        const requestItems = page.getByText('Request Policy');
        const reqCount = await requestItems.count();
        console.log(`  "Request Policy" items after click: ${reqCount}`);

        if (reqCount > 0) {
          await requestItems.first().click();
          await page.waitForTimeout(2000);
          await snap(page, 'debug-request-wizard');

          const bodyText = await page.textContent('body') || '';
          console.log(`  Wizard opened: ${bodyText.includes('Policy Details') || bodyText.includes('What policy')}`);
        }
        break;
      }
    }
  }
});
