import { test, Page } from '@playwright/test';
import * as path from 'path';

/**
 * Check PM_NotificationQueue for bad records and diagnose email pipeline
 */

const BASE = 'https://mf7m.sharepoint.com/sites/PolicyManager';

async function snap(page: Page, name: string): Promise<void> {
  await page.screenshot({ path: path.join(process.cwd(), `e2e-email-${name}.png`), fullPage: false });
  console.log(`📸 e2e-email-${name}.png`);
}

test('Check PM_NotificationQueue for empty RecipientEmail', async ({ page }) => {
  console.log('=== CHECKING PM_NotificationQueue ===\n');

  // Navigate to the SP list
  await page.goto(`${BASE}/Lists/PM_NotificationQueue/AllItems.aspx`, {
    waitUntil: 'domcontentloaded', timeout: 60000
  });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  await snap(page, 'queue-list');

  const bodyText = await page.textContent('body') || '';
  console.log('Page loaded:', bodyText.length > 100);
  console.log('Contains "Pending":', bodyText.includes('Pending'));
  console.log('Contains "Sent":', bodyText.includes('Sent'));
  console.log('Contains "Failed":', bodyText.includes('Failed'));
  console.log('Contains "Processing":', bodyText.includes('Processing'));

  // Count items visible
  const rows = page.locator('[role="row"], tr').filter({ hasText: /Pending|Sent|Failed|Processing/ });
  const rowCount = await rows.count();
  console.log(`\nQueue items visible: ${rowCount}`);

  // Try to find empty RecipientEmail cells
  // SP list view shows columns — look for empty cells in the RecipientEmail column
  const allCells = page.locator('td, [role="gridcell"]');
  const cellCount = await allCells.count();
  console.log(`Total cells: ${cellCount}`);
});

test('Check PM_NotificationQueue via REST API', async ({ page }) => {
  console.log('=== CHECKING QUEUE VIA REST API ===\n');

  // Use SP REST API to get queue items with empty RecipientEmail
  await page.goto(`${BASE}/_api/web/lists/getByTitle('PM_NotificationQueue')/items?$select=Id,Title,RecipientEmail,QueueStatus,NotificationType,AttemptCount&$orderby=Id desc&$top=20`, {
    waitUntil: 'domcontentloaded', timeout: 30000
  });
  await page.waitForTimeout(3000);

  const bodyText = await page.textContent('body') || '';

  // SP REST returns XML by default — check if we got data
  if (bodyText.includes('RecipientEmail') || bodyText.includes('d:RecipientEmail')) {
    console.log('REST API returned data');

    // Count items with empty RecipientEmail
    const emptyEmailCount = (bodyText.match(/RecipientEmail><\/d:|RecipientEmail":\s*""/g) || []).length;
    console.log(`Items with EMPTY RecipientEmail: ${emptyEmailCount}`);

    // Count by status
    const pendingCount = (bodyText.match(/Pending/g) || []).length;
    const sentCount = (bodyText.match(/>Sent</g) || []).length;
    const failedCount = (bodyText.match(/Failed/g) || []).length;
    console.log(`Pending: ${pendingCount}, Sent: ${sentCount}, Failed: ${failedCount}`);
  } else {
    console.log('REST API response (first 500 chars):');
    console.log(bodyText.slice(0, 500));
  }

  await snap(page, 'queue-rest-api');
});

test('List all PM_ lists and their item counts', async ({ page }) => {
  console.log('=== LISTING ALL PM_ LISTS ===\n');

  await page.goto(`${BASE}`, { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForTimeout(3000);

  const lists = await page.evaluate(async (siteUrl: string) => {
    try {
      const resp = await fetch(
        `${siteUrl}/_api/web/lists?$select=Title,ItemCount,Hidden&$filter=startswith(Title,'PM_')&$orderby=Title`,
        { headers: { 'Accept': 'application/json;odata=verbose' } }
      );
      const data = await resp.json();
      return (data.d?.results || []).map((l: any) => ({ title: l.Title, count: l.ItemCount }));
    } catch (e) { return { error: (e as Error).message }; }
  }, BASE);

  if (Array.isArray(lists)) {
    console.log(`Found ${lists.length} PM_ lists:\n`);
    for (const l of lists) {
      const flag = l.title.includes('Notification') || l.title.includes('Email') || l.title.includes('Queue') ? ' ⬅️ EMAIL' : '';
      console.log(`  ${l.title.padEnd(40)} ${String(l.count).padStart(5)} items${flag}`);
    }
  } else {
    console.log('Error:', JSON.stringify(lists));
  }
});

test('Check queue via JSON API', async ({ page }) => {
  console.log('=== CHECKING QUEUE VIA JSON API ===\n');

  await page.goto(`${BASE}`, { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForTimeout(3000);

  // Get Pending items first (these are the ones the Logic App will try to send)
  const result = await page.evaluate(async (siteUrl: string) => {
    try {
      const response = await fetch(
        `${siteUrl}/_api/web/lists/getByTitle('PM_NotificationQueue')/items?$top=5&$orderby=Id desc`,
        {
          headers: {
            'Accept': 'application/json;odata=verbose'
          }
        }
      );
      const data = await response.json();
      return data.d?.results || [];
    } catch (e) {
      return { error: (e as Error).message };
    }
  }, BASE);

  // Dump raw first item for field inspection
  if (Array.isArray(result) && result.length > 0) {
    console.log('--- RAW FIRST ITEM FIELDS ---');
    const firstItem = result[0];
    for (const [key, value] of Object.entries(firstItem)) {
      if (!key.startsWith('__') && !key.startsWith('odata')) {
        console.log(`  ${key}: ${JSON.stringify(value).slice(0, 80)}`);
      }
    }
    console.log('--- END RAW ---\n');
  }

  if (Array.isArray(result)) {
    console.log(`Total items returned: ${result.length}\n`);

    // Categorise
    let emptyRecipients = 0;
    let pending = 0;
    let sent = 0;
    let failed = 0;
    let processing = 0;

    for (const item of result) {
      const status = item.QueueStatus || 'Unknown';
      const email = item.RecipientEmail || '';
      const title = (item.Title || '').slice(0, 60);

      if (!email || email.trim() === '') emptyRecipients++;
      if (status === 'Pending') pending++;
      else if (status === 'Sent') sent++;
      else if (status === 'Failed') failed++;
      else if (status === 'Processing') processing++;

      // Log each item
      const emailDisplay = email ? email.slice(0, 30) : '⚠️ EMPTY';
      console.log(`  [${item.Id}] ${status.padEnd(12)} | ${emailDisplay.padEnd(32)} | ${title}`);
    }

    console.log(`\n--- Summary ---`);
    console.log(`Total: ${result.length}`);
    console.log(`Pending: ${pending}`);
    console.log(`Sent: ${sent}`);
    console.log(`Failed: ${failed}`);
    console.log(`Processing: ${processing}`);
    console.log(`⚠️ Empty RecipientEmail: ${emptyRecipients}`);

    if (emptyRecipients > 0) {
      console.log(`\n🔴 FOUND ${emptyRecipients} ITEMS WITH EMPTY RECIPIENT — these are causing the Logic App failure!`);
      console.log(`Fix: Delete these items from PM_NotificationQueue or update RecipientEmail`);
    } else {
      console.log(`\n✅ All items have RecipientEmail — Logic App failure may be from a different cause`);
    }
  } else {
    console.log('API result:', JSON.stringify(result).slice(0, 300));
  }
});
