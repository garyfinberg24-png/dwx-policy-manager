import { test, Page } from '@playwright/test';
import * as path from 'path';

/**
 * LIFECYCLE v2 — Fixed selectors based on diagnostic findings
 *
 * Key selector patterns discovered:
 * - My Policies: rows are table rows, eye icon for View (last column)
 * - Pipeline actions: IconButton with ariaLabel="Verb PolicyTitle"
 *   e.g. ariaLabel="Publish HTML Editor Test", ariaLabel="Retire HTML Editor Test"
 * - Acknowledge: button with text "Acknowledge" on PolicyDetails page
 */

const BASE = 'https://mf7m.sharepoint.com/sites/PolicyManager/SitePages';

let shotCount = 0;
async function snap(page: Page, name: string): Promise<void> {
  if (shotCount >= 15) return;
  await page.screenshot({ path: path.join(process.cwd(), `e2e-lv2-${name}.png`), fullPage: false });
  shotCount++;
  console.log(`📸 [${shotCount}/15] e2e-lv2-${name}.png`);
}

// ============================================================
// TEST 1: Acknowledgement — click eye icon on My Policies row
// ============================================================
test('1 — Acknowledgement: open policy from My Policies', async ({ page }) => {
  console.log('\n=== ACKNOWLEDGEMENT FLOW ===');

  await page.goto(`${BASE}/MyPolicies.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  // My Policies has rows — find the eye/view icon in the last column
  const viewIcons = page.locator('button[aria-label*="View" i], [class*="iconButton"], svg').filter({ hasText: '' });

  // Better: find clickable rows — they might be div[role="row"] or tr
  const policyRows = page.locator('[role="row"], tr').filter({ hasText: /POL-|Policy/i });
  const rowCount = await policyRows.count();
  console.log(`  Policy rows: ${rowCount}`);

  if (rowCount > 0) {
    // Click on the first policy row
    await policyRows.first().click();
    await page.waitForTimeout(3000);

    // Check if it navigated to PolicyDetails
    const url = page.url();
    console.log(`  Navigated to: ${url}`);

    if (url.includes('PolicyDetails')) {
      await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
      await page.waitForTimeout(3000);
      await snap(page, '1-policy-detail');

      const bodyText = await page.textContent('body') || '';
      console.log(`  Acknowledge button: ${bodyText.includes('Acknowledge')}`);
      console.log(`  Already acknowledged: ${bodyText.includes('Acknowledged')}`);

      // Check viewer mode
      const hasPdf = await page.locator('object[type="application/pdf"]').isVisible().catch(() => false);
      const hasOffice = await page.locator('iframe[src*="WopiFrame"]').isVisible().catch(() => false);
      const hasHtml = await page.locator('h1, h2, h3').first().isVisible().catch(() => false);
      const viewer = hasPdf ? 'PDF' : hasOffice ? 'Office Online' : hasHtml ? 'HTML' : 'Other';
      console.log(`  Viewer: ${viewer}`);
    } else {
      // Maybe it opened a panel instead of navigating
      console.log(`  Didn't navigate — checking for panel/detail view`);
      await snap(page, '1-click-result');
    }
  } else {
    // Try clicking the first policy name text directly
    const policyName = page.locator('text=/Data Privacy|Information Security|Code of Conduct|Health|Whistleblower/i').first();
    if (await policyName.isVisible().catch(() => false)) {
      console.log(`  Clicking policy name directly`);
      await policyName.click();
      await page.waitForTimeout(3000);
      await snap(page, '1-name-click');
    }
  }
});


// ============================================================
// TEST 2: Pipeline — Publish an Approved policy
// ============================================================
test('2 — Publish: click Publish icon on Approved policy', async ({ page }) => {
  console.log('\n=== PUBLISH APPROVED POLICY ===');

  await page.goto(`${BASE}/PolicyAuthor.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  // Click "Approved" filter
  const approvedFilter = page.locator('button, [role="button"]').filter({ hasText: /Approved/i }).first();
  if (await approvedFilter.isVisible().catch(() => false)) {
    await approvedFilter.click();
    await page.waitForTimeout(2000);
  }

  // Find a Publish icon — ariaLabel pattern: "Publish {PolicyTitle}"
  const publishBtns = page.locator('button[aria-label*="Publish"]');
  const publishCount = await publishBtns.count();
  console.log(`  Publish buttons: ${publishCount}`);

  for (let i = 0; i < Math.min(publishCount, 5); i++) {
    const label = await publishBtns.nth(i).getAttribute('aria-label') || '';
    const disabled = await publishBtns.nth(i).isDisabled().catch(() => true);
    console.log(`    [${i}]: "${label}" disabled=${disabled}`);
  }

  // Click first enabled Publish button
  for (let i = 0; i < publishCount; i++) {
    const disabled = await publishBtns.nth(i).isDisabled().catch(() => true);
    if (!disabled) {
      const label = await publishBtns.nth(i).getAttribute('aria-label') || '';
      console.log(`  Clicking: "${label}"`);
      await publishBtns.nth(i).click();
      await page.waitForTimeout(5000);
      await snap(page, '2-publish-dialog');

      // Handle confirmation dialog
      const confirmBtn = page.locator('button').filter({ hasText: /Publish|Confirm|Yes/i }).last();
      if (await confirmBtn.isVisible({ timeout: 3000 }).catch(() => false)) {
        await confirmBtn.click();
        await page.waitForTimeout(5000);

        const ok = page.locator('button').filter({ hasText: /^OK$/i }).first();
        if (await ok.isVisible({ timeout: 5000 }).catch(() => false)) {
          console.log(`  ✅ PUBLISHED — success dialog`);
          await ok.click();
        }
        await snap(page, '2-publish-result');
      }
      break;
    }
  }
});


// ============================================================
// TEST 3: Pipeline — Retire a Published policy
// ============================================================
test('3 — Retire: click Retire icon on Published policy', async ({ page }) => {
  console.log('\n=== RETIRE PUBLISHED POLICY ===');

  await page.goto(`${BASE}/PolicyAuthor.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  // Filter to Published
  const publishedFilter = page.locator('button, [role="button"]').filter({ hasText: /Published/i }).first();
  if (await publishedFilter.isVisible().catch(() => false)) {
    await publishedFilter.click();
    await page.waitForTimeout(2000);
  }

  // Find Retire icon — ariaLabel: "Retire {PolicyTitle}"
  const retireBtns = page.locator('button[aria-label*="Retire"]');
  const retireCount = await retireBtns.count();
  console.log(`  Retire buttons: ${retireCount}`);

  // Click first enabled Retire button
  for (let i = 0; i < retireCount; i++) {
    const disabled = await retireBtns.nth(i).isDisabled().catch(() => true);
    if (!disabled) {
      const label = await retireBtns.nth(i).getAttribute('aria-label') || '';
      console.log(`  Clicking: "${label}"`);
      await retireBtns.nth(i).click();
      await page.waitForTimeout(3000);
      await snap(page, '3-retire-dialog');

      // Handle confirmation dialog — might ask for reason
      const bodyText = await page.textContent('body') || '';
      console.log(`  Dialog: ${bodyText.includes('Retire') || bodyText.includes('retire')}`);

      const confirmBtn = page.locator('button').filter({ hasText: /Retire|Confirm|Yes/i }).last();
      if (await confirmBtn.isVisible({ timeout: 3000 }).catch(() => false)) {
        await confirmBtn.click();
        await page.waitForTimeout(5000);

        const ok = page.locator('button').filter({ hasText: /^OK$/i }).first();
        if (await ok.isVisible({ timeout: 5000 }).catch(() => false)) {
          console.log(`  ✅ RETIRED — success dialog`);
          await ok.click();
        }
        await snap(page, '3-retire-result');
      }
      break;
    }
  }
});


// ============================================================
// TEST 4: Pipeline — Edit rejected policy (Revise & Resubmit)
// ============================================================
test('4 — Edit Rejected: find Revise & Resubmit button', async ({ page }) => {
  console.log('\n=== EDIT REJECTED POLICY ===');

  await page.goto(`${BASE}/PolicyAuthor.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  // Filter to Rejected
  const rejectedFilter = page.locator('button, [role="button"]').filter({ hasText: /Rejected/i }).first();
  if (await rejectedFilter.isVisible().catch(() => false)) {
    await rejectedFilter.click();
    await page.waitForTimeout(2000);
  }

  await snap(page, '4-rejected-list');

  const bodyText = await page.textContent('body') || '';
  console.log(`  Contains "Rejected": ${bodyText.includes('Rejected')}`);

  // Look for "Revise & Resubmit" DefaultButton or any Edit icon
  const reviseBtn = page.locator('button').filter({ hasText: /Revise.*Resubmit/i }).first();
  const hasRevise = await reviseBtn.isVisible().catch(() => false);
  console.log(`  "Revise & Resubmit" button: ${hasRevise}`);

  // Also check for Edit icon with policy aria-label
  const editBtns = page.locator('button[aria-label*="Edit"]');
  const editCount = await editBtns.count();
  console.log(`  Edit buttons: ${editCount}`);
  for (let i = 0; i < Math.min(editCount, 5); i++) {
    const label = await editBtns.nth(i).getAttribute('aria-label') || '';
    if (label.includes('Edit') && !label.includes('site')) {
      console.log(`    "${label}"`);
    }
  }

  if (hasRevise) {
    await reviseBtn.click();
    await page.waitForTimeout(5000);
    await snap(page, '4-revise-wizard');
    console.log(`  ✅ Revise & Resubmit clicked — wizard should open`);
  }
});


// ============================================================
// TEST 5: Direct acknowledgement test via PolicyDetails
// ============================================================
test('5 — Acknowledge policy directly via PolicyDetails', async ({ page }) => {
  console.log('\n=== DIRECT ACKNOWLEDGEMENT ===');

  // Navigate to a published policy that we haven't acknowledged yet
  // Try policy IDs 2, 3, 4 (likely published and unacknowledged)
  for (const policyId of [2, 3, 4, 5]) {
    await page.goto(`${BASE}/PolicyDetails.aspx?policyId=${policyId}`, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(5000);

    const bodyText = await page.textContent('body') || '';
    const hasAckBtn = await page.locator('button').filter({ hasText: /Acknowledge/i }).first().isVisible().catch(() => false);
    const isAlreadyAcked = bodyText.includes('Already Acknowledged') || bodyText.includes('You have acknowledged');

    console.log(`  Policy ${policyId}: ackBtn=${hasAckBtn}, alreadyAcked=${isAlreadyAcked}`);

    if (hasAckBtn && !isAlreadyAcked) {
      await snap(page, `5-policy-${policyId}-before-ack`);

      // Scroll to bottom to fulfil read requirement
      await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
      await page.waitForTimeout(2000);

      const ackBtn = page.locator('button').filter({ hasText: /Acknowledge/i }).first();
      if (await ackBtn.isEnabled().catch(() => false)) {
        await ackBtn.click();
        await page.waitForTimeout(5000);

        await snap(page, `5-policy-${policyId}-after-ack`);

        const resultBody = await page.textContent('body') || '';
        console.log(`  ✅ Policy ${policyId}: ${resultBody.includes('Acknowledged') || resultBody.includes('success') ? 'ACKNOWLEDGED' : 'check result'}`);

        // Dismiss dialog
        const ok = page.locator('button').filter({ hasText: /^OK$/i }).first();
        if (await ok.isVisible({ timeout: 3000 }).catch(() => false)) await ok.click();
      }
      break; // Only need to acknowledge one
    }
  }
});
