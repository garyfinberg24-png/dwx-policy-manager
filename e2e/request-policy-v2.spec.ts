import { test, Page } from '@playwright/test';
import * as path from 'path';

const BASE = 'https://mf7m.sharepoint.com/sites/PolicyManager/SitePages';

async function snap(page: Page, name: string): Promise<void> {
  await page.screenshot({ path: path.join(process.cwd(), `e2e-reqv2-${name}.png`), fullPage: false });
  console.log(`📸 e2e-reqv2-${name}.png`);
}

test('1 — Find and click Manager > Request Policy nav', async ({ page }) => {
  console.log('\n=== FINDING MANAGER NAV ===');

  // Try the Manager View page directly where the nav is guaranteed to be visible
  await page.goto(`${BASE}/PolicyManagerView.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  await snap(page, '1-manager-page');

  // Dump all elements with "Manager" or "Request" in their text
  const navInfo = await page.evaluate(() => {
    const results: string[] = [];
    const all = document.querySelectorAll('*');
    for (const el of all) {
      // Only check direct text content (not children)
      const directText = Array.from(el.childNodes)
        .filter(n => n.nodeType === Node.TEXT_NODE)
        .map(n => n.textContent?.trim())
        .filter(t => t && t.length > 0)
        .join(' ');

      if (directText && (directText.includes('Manager') || directText.includes('Request') || directText.includes('request'))) {
        const tag = el.tagName.toLowerCase();
        const role = el.getAttribute('role') || '';
        const cls = (el.className?.toString() || '').slice(0, 40);
        const href = el.getAttribute('href') || '';
        const onclick = el.getAttribute('onclick') ? 'has onclick' : '';
        results.push(`<${tag}${role ? ` role="${role}"` : ''}${href ? ` href="${href}"` : ''}${onclick ? ' onclick' : ''}> "${directText.slice(0, 40)}" cls="${cls}"`);
      }
    }
    return results;
  });

  console.log('  Direct text "Manager"/"Request" elements:');
  for (const item of navInfo) {
    console.log(`    ${item}`);
  }

  // Also try finding via the header component class names
  const headerNavItems = await page.evaluate(() => {
    const results: string[] = [];
    // Look for elements with nav-related class names
    const selectors = [
      '[class*="navItem"]',
      '[class*="headerNav"]',
      '[class*="navLink"]',
      '[class*="dropdownTrigger"]',
      '[class*="menuItem"]',
      '[class*="navGroup"]',
      '[data-key]',
      'a[href*="#"]'
    ];
    for (const selector of selectors) {
      const els = document.querySelectorAll(selector);
      for (const el of els) {
        const text = (el.textContent || '').trim().slice(0, 50);
        if (text) {
          results.push(`${selector}: "${text}" href="${el.getAttribute('href') || ''}" key="${el.getAttribute('data-key') || ''}"`);
        }
      }
    }
    return results.slice(0, 30);
  });

  console.log('\n  Nav-related elements:');
  for (const item of headerNavItems) {
    console.log(`    ${item}`);
  }
});

test('2 — Navigate directly to Manager View and find Request Policy', async ({ page }) => {
  console.log('\n=== MANAGER VIEW PAGE ===');

  await page.goto(`${BASE}/PolicyManagerView.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  const bodyText = await page.textContent('body') || '';
  console.log(`  Contains "Request Policy": ${bodyText.includes('Request Policy')}`);
  console.log(`  Contains "Request": ${bodyText.includes('Request')}`);
  console.log(`  Contains "Team Compliance": ${bodyText.includes('Team Compliance')}`);
  console.log(`  Contains "Approval": ${bodyText.includes('Approval')}`);

  await snap(page, '2-manager-view');

  // Try finding "Request Policy" button or nav item directly
  const requestBtn = page.locator('button, a, [role="button"], [role="menuitem"]').filter({ hasText: /Request Policy/i });
  const reqCount = await requestBtn.count();
  console.log(`  "Request Policy" elements: ${reqCount}`);

  // Also check the Author page for the Requests tab
  console.log('\n  Checking Author page...');
  await page.goto(`${BASE}/PolicyAuthor.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  await snap(page, '2-author-page');

  // Find all tab-like elements
  const tabs = page.locator('button, [role="tab"]').filter({ hasText: /Request|request/i });
  const tabCount = await tabs.count();
  console.log(`  "Request" tabs: ${tabCount}`);

  for (let i = 0; i < tabCount; i++) {
    const text = (await tabs.nth(i).textContent() || '').trim().slice(0, 40);
    console.log(`    tab[${i}]: "${text}"`);
  }

  // Click Policy Requests tab
  const reqTab = page.locator('button, [role="tab"]').filter({ hasText: /Policy Requests/i }).first();
  if (await reqTab.isVisible().catch(() => false)) {
    await reqTab.click();
    await page.waitForTimeout(3000);
    await snap(page, '2-requests-tab');
    console.log('  ✅ Policy Requests tab clicked');
  }
});
