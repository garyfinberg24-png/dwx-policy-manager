import { test, Page } from '@playwright/test';

/**
 * Fix PM_NotificationQueue — delete items with empty or invalid RecipientEmail
 */

const BASE = 'https://mf7m.sharepoint.com/sites/PolicyManager';

test('Delete bad notification queue items', async ({ page }) => {
  console.log('=== FIXING PM_NotificationQueue ===\n');

  await page.goto(`${BASE}`, { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForTimeout(3000);

  // Get request digest for write operations
  const digest = await page.evaluate(async (siteUrl: string) => {
    const resp = await fetch(`${siteUrl}/_api/contextinfo`, {
      method: 'POST',
      headers: { 'Accept': 'application/json;odata=verbose' }
    });
    const data = await resp.json();
    return data.d.GetContextWebInformation.FormDigestValue;
  }, BASE);

  console.log('Got request digest');

  // Find all Pending items with empty or invalid RecipientEmail
  const badItems = await page.evaluate(async (siteUrl: string) => {
    const resp = await fetch(
      `${siteUrl}/_api/web/lists/getByTitle('PM_NotificationQueue')/items?$select=Id,Title,RecipientEmail,QueueStatus&$filter=QueueStatus eq 'Pending'&$top=100&$orderby=Id desc`,
      { headers: { 'Accept': 'application/json;odata=verbose' } }
    );
    const data = await resp.json();
    const items = data.d?.results || [];

    // Find items where RecipientEmail is empty, null, or doesn't contain @
    return items.filter((item: any) => {
      const email = item.RecipientEmail || '';
      return !email || !email.includes('@');
    }).map((item: any) => ({
      id: item.Id,
      title: item.Title,
      recipientEmail: item.RecipientEmail || '(empty)',
      status: item.QueueStatus
    }));
  }, BASE);

  console.log(`Found ${badItems.length} items with empty/invalid RecipientEmail:\n`);
  for (const item of badItems) {
    console.log(`  [${item.id}] "${item.title}" → RecipientEmail: "${item.recipientEmail}"`);
  }

  if (badItems.length === 0) {
    console.log('\n✅ No bad items to fix!');
    return;
  }

  // Delete each bad item
  console.log(`\nDeleting ${badItems.length} bad items...`);

  for (const item of badItems) {
    const deleted = await page.evaluate(async (args: { siteUrl: string; id: number; digest: string }) => {
      try {
        const resp = await fetch(
          `${args.siteUrl}/_api/web/lists/getByTitle('PM_NotificationQueue')/items(${args.id})`,
          {
            method: 'POST',
            headers: {
              'Accept': 'application/json;odata=verbose',
              'X-RequestDigest': args.digest,
              'IF-MATCH': '*',
              'X-HTTP-Method': 'DELETE'
            }
          }
        );
        return resp.ok || resp.status === 200 || resp.status === 204;
      } catch (e) {
        return false;
      }
    }, { siteUrl: BASE, id: item.id, digest });

    console.log(`  [${item.id}] ${deleted ? '✅ Deleted' : '❌ Failed to delete'}`);
  }

  // Verify
  const remaining = await page.evaluate(async (siteUrl: string) => {
    const resp = await fetch(
      `${siteUrl}/_api/web/lists/getByTitle('PM_NotificationQueue')/items?$select=Id,RecipientEmail,QueueStatus&$filter=QueueStatus eq 'Pending'&$top=10&$orderby=Id desc`,
      { headers: { 'Accept': 'application/json;odata=verbose' } }
    );
    const data = await resp.json();
    const items = data.d?.results || [];
    return items.filter((item: any) => {
      const email = item.RecipientEmail || '';
      return !email || !email.includes('@');
    }).length;
  }, BASE);

  console.log(`\nRemaining bad items: ${remaining}`);
  if (remaining === 0) {
    console.log('✅ All bad items cleaned up — Logic App should work now');
  }
});
