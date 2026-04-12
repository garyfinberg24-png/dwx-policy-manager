import { test, Page } from '@playwright/test';
import * as path from 'path';

/**
 * REMAINING LIFECYCLE TESTS
 * 1. Acknowledgement flow
 * 2. Edit after Rejection → Resubmit
 * 3. Revise published policy
 * 4. Retire published policy
 * 5. Distribution campaign creation
 */

const BASE = 'https://mf7m.sharepoint.com/sites/PolicyManager/SitePages';

let shotCount = 0;
async function snap(page: Page, name: string): Promise<void> {
  if (shotCount >= 20) return;
  await page.screenshot({ path: path.join(process.cwd(), `e2e-life-${name}.png`), fullPage: false });
  shotCount++;
  console.log(`📸 [${shotCount}/20] e2e-life-${name}.png`);
}

// ============================================================
// TEST 1: Acknowledgement Flow
// ============================================================
test.describe.serial('1 — Acknowledgement Flow', () => {

  test('1.1 — My Policies shows pending policies', async ({ page }) => {
    console.log('\n=== MY POLICIES — PENDING ACKNOWLEDGEMENTS ===');

    await page.goto(`${BASE}/MyPolicies.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(5000);

    await snap(page, '1-1-my-policies');

    const bodyText = await page.textContent('body') || '';
    console.log(`  Contains "Pending": ${bodyText.includes('Pending')}`);
    console.log(`  Contains "Overdue": ${bodyText.includes('Overdue')}`);
    console.log(`  Contains "Completed": ${bodyText.includes('Completed')}`);
    console.log(`  Contains "Acknowledged": ${bodyText.includes('Acknowledged')}`);

    // Count policy rows
    const policyLinks = page.locator('a[href*="PolicyDetails"]');
    const linkCount = await policyLinks.count();
    console.log(`  Policy links: ${linkCount}`);

    // Check for pending policies specifically
    const pendingFilter = page.locator('button, [role="tab"]').filter({ hasText: /Pending/i }).first();
    if (await pendingFilter.isVisible().catch(() => false)) {
      await pendingFilter.click();
      await page.waitForTimeout(2000);
      console.log(`  Filtered to Pending`);
    }

    await snap(page, '1-1-pending-filtered');
  });

  test('1.2 — Open a policy and read it', async ({ page }) => {
    console.log('\n=== OPEN AND READ POLICY ===');

    await page.goto(`${BASE}/MyPolicies.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(5000);

    // Click a pending policy
    const policyLink = page.locator('a[href*="PolicyDetails"]').first();
    if (await policyLink.isVisible().catch(() => false)) {
      const href = await policyLink.getAttribute('href') || '';
      console.log(`  Clicking policy: ${href}`);
      await policyLink.click();
      await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
      await page.waitForTimeout(5000);

      await snap(page, '1-2-policy-reader');

      const bodyText = await page.textContent('body') || '';

      // Check for scroll progress bar
      const progressBar = page.locator('[class*="scrollProgress"], [class*="progressBar"], [class*="progress"]');
      const hasProgress = await progressBar.first().isVisible().catch(() => false);
      console.log(`  Scroll progress bar: ${hasProgress}`);

      // Check for acknowledge button
      const ackButton = page.locator('button').filter({ hasText: /Acknowledge|Accept|Confirm.*Read/i }).first();
      const hasAck = await ackButton.isVisible().catch(() => false);
      console.log(`  Acknowledge button: ${hasAck}`);

      // Check viewer mode
      const hasPdf = await page.locator('object[type="application/pdf"]').isVisible().catch(() => false);
      const hasOffice = await page.locator('iframe[src*="WopiFrame"]').isVisible().catch(() => false);
      const hasHtml = await page.locator('h1, h2, h3').first().isVisible().catch(() => false);
      const viewer = hasPdf ? 'PDF' : hasOffice ? 'Office Online' : hasHtml ? 'HTML' : 'Unknown';
      console.log(`  Viewer: ${viewer}`);

      // Try scrolling to bottom (to trigger read progress)
      await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
      await page.waitForTimeout(2000);

      // Check if acknowledge button is now enabled
      if (hasAck) {
        const isEnabled = await ackButton.isEnabled().catch(() => false);
        console.log(`  Acknowledge enabled after scroll: ${isEnabled}`);
      }
    } else {
      console.log('  ⚠️ No policy links found');
    }
  });

  test('1.3 — Click Acknowledge button', async ({ page }) => {
    console.log('\n=== ACKNOWLEDGE POLICY ===');

    // Navigate to a specific published policy
    await page.goto(`${BASE}/PolicyDetails.aspx?policyId=1`, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(5000);

    await snap(page, '1-3-before-ack');

    const bodyText = await page.textContent('body') || '';

    // Look for acknowledge button
    const ackButton = page.locator('button').filter({ hasText: /Acknowledge|Accept|Confirm/i }).first();
    const hasAck = await ackButton.isVisible().catch(() => false);
    console.log(`  Acknowledge button: ${hasAck}`);

    // Look for "Already Acknowledged" indicator
    const alreadyAcked = bodyText.includes('Acknowledged') && !bodyText.includes('Not Acknowledged');
    console.log(`  Already acknowledged: ${alreadyAcked}`);

    // Scroll to bottom first
    await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
    await page.waitForTimeout(1000);

    if (hasAck) {
      const isEnabled = await ackButton.isEnabled().catch(() => false);
      console.log(`  Button enabled: ${isEnabled}`);

      if (isEnabled) {
        await ackButton.click();
        await page.waitForTimeout(5000);

        await snap(page, '1-3-after-ack');

        const resultBody = await page.textContent('body') || '';
        const success = resultBody.includes('Acknowledged') || resultBody.includes('success') || resultBody.includes('Thank');
        console.log(`  ✅ Acknowledgement result: ${success ? 'SUCCESS' : 'check screenshot'}`);

        // Dismiss any dialog
        const okBtn = page.locator('button').filter({ hasText: /^OK$|^Close$|^Done$/i }).first();
        if (await okBtn.isVisible({ timeout: 3000 }).catch(() => false)) {
          await okBtn.click();
        }
      }
    }
  });
});


// ============================================================
// TEST 2: Edit After Rejection → Resubmit
// ============================================================
test.describe('2 — Edit After Rejection', () => {

  test('2.1 — Find rejected policy in pipeline', async ({ page }) => {
    console.log('\n=== FIND REJECTED POLICY ===');

    await page.goto(`${BASE}/PolicyAuthor.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(5000);

    // Click Rejected filter
    const rejectedFilter = page.locator('button, [role="tab"], [role="button"]').filter({ hasText: /Rejected/i }).first();
    if (await rejectedFilter.isVisible().catch(() => false)) {
      await rejectedFilter.click();
      await page.waitForTimeout(2000);
      console.log(`  Filtered to Rejected`);
    }

    await snap(page, '2-1-rejected-pipeline');

    const bodyText = await page.textContent('body') || '';
    console.log(`  Contains "Rejected": ${bodyText.includes('Rejected')}`);

    // Look for Edit & Resubmit action
    const editBtn = page.locator('button[aria-label*="Edit" i], button[title*="Edit" i], button[aria-label*="Resubmit" i]').first();
    const hasEdit = await editBtn.isVisible().catch(() => false);
    console.log(`  Edit/Resubmit button: ${hasEdit}`);

    if (hasEdit) {
      await editBtn.click();
      await page.waitForTimeout(5000);

      await snap(page, '2-1-edit-rejected');

      const editBody = await page.textContent('body') || '';
      console.log(`  Wizard opened: ${editBody.includes('Creation Method') || editBody.includes('Basic Info') || editBody.includes('Policy Builder')}`);
    }
  });
});


// ============================================================
// TEST 3: Revise Published Policy
// ============================================================
test.describe('3 — Revise Published Policy', () => {

  test('3.1 — Find published policy and click Revise', async ({ page }) => {
    console.log('\n=== REVISE PUBLISHED POLICY ===');

    await page.goto(`${BASE}/PolicyAuthor.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(5000);

    // Click Published filter
    const publishedFilter = page.locator('button, [role="tab"], [role="button"]').filter({ hasText: /Published/i }).first();
    if (await publishedFilter.isVisible().catch(() => false)) {
      await publishedFilter.click();
      await page.waitForTimeout(2000);
      console.log(`  Filtered to Published`);
    }

    await snap(page, '3-1-published-pipeline');

    const bodyText = await page.textContent('body') || '';
    console.log(`  Contains "Published": ${bodyText.includes('Published')}`);

    // Look for Edit button on a published policy (Edit = Revise for published)
    // The pipeline uses IconButtons with ariaLabel like "Edit Policy Name"
    const editBtns = page.locator('button[aria-label*="Edit"]').filter({ hasText: '' });
    const editCount = await editBtns.count();
    console.log(`  Edit buttons (revise): ${editCount}`);

    // Also look for "Revise & Resubmit" text button
    const reviseTextBtn = page.locator('button').filter({ hasText: /Revise/i }).first();
    const hasReviseText = await reviseTextBtn.isVisible().catch(() => false);
    console.log(`  "Revise" text button: ${hasReviseText}`);

    // Log first 10 icon buttons in the pipeline area
    const iconBtns = page.locator('button[aria-label]');
    const iconCount = await iconBtns.count();
    const pipelineBtns: string[] = [];
    for (let i = 0; i < Math.min(iconCount, 30); i++) {
      const label = await iconBtns.nth(i).getAttribute('aria-label') || '';
      if (label && (label.includes('Edit') || label.includes('View') || label.includes('Publish') || label.includes('Retire') || label.includes('Duplicate') || label.includes('Delete'))) {
        pipelineBtns.push(label);
      }
    }
    console.log(`  Pipeline action buttons:`);
    for (const btn of pipelineBtns.slice(0, 10)) {
      console.log(`    "${btn}"`);
    }
  });
});


// ============================================================
// TEST 4: Retire Published Policy
// ============================================================
test.describe('4 — Retire Published Policy', () => {

  test('4.1 — Find published policy and click Retire', async ({ page }) => {
    console.log('\n=== RETIRE PUBLISHED POLICY ===');

    await page.goto(`${BASE}/PolicyAuthor.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(5000);

    const publishedFilter = page.locator('button, [role="tab"], [role="button"]').filter({ hasText: /Published/i }).first();
    if (await publishedFilter.isVisible().catch(() => false)) {
      await publishedFilter.click();
      await page.waitForTimeout(2000);
    }

    // Look for Retire action button
    const retireBtn = page.locator('button[aria-label*="Retire" i], button[title*="Retire" i]').first();
    const hasRetire = await retireBtn.isVisible().catch(() => false);
    console.log(`  Retire button: ${hasRetire}`);

    if (hasRetire) {
      await retireBtn.click();
      await page.waitForTimeout(3000);

      await snap(page, '4-1-retire-dialog');

      // Check for confirmation dialog with reason
      const bodyText = await page.textContent('body') || '';
      console.log(`  Retire dialog: ${bodyText.includes('Retire') || bodyText.includes('retire') || bodyText.includes('reason')}`);

      // Fill reason if dialog appeared
      const reasonField = page.locator('textarea').first();
      if (await reasonField.isVisible().catch(() => false)) {
        await reasonField.fill('E2E test — policy superseded by updated version. All outstanding acknowledgements should be cancelled.');
        console.log(`  ✅ Retirement reason filled`);

        // Click confirm
        const confirmBtn = page.locator('button').filter({ hasText: /Confirm|Retire|Yes/i }).last();
        if (await confirmBtn.isVisible().catch(() => false)) {
          await confirmBtn.click();
          await page.waitForTimeout(5000);

          await snap(page, '4-1-after-retire');

          const resultBody = await page.textContent('body') || '';
          console.log(`  ✅ Retirement result: ${resultBody.includes('Retired') || resultBody.includes('retired') || resultBody.includes('success')}`);

          // Dismiss dialog
          const okBtn = page.locator('button').filter({ hasText: /^OK$/i }).first();
          if (await okBtn.isVisible({ timeout: 3000 }).catch(() => false)) await okBtn.click();
        }
      }
    } else {
      console.log('  ⚠️ No Retire button found');
    }
  });
});


// ============================================================
// TEST 5: Distribution Campaign Creation
// ============================================================
test.describe('5 — Distribution Campaign', () => {

  test('5.1 — Distribution page with existing campaigns', async ({ page }) => {
    console.log('\n=== DISTRIBUTION CAMPAIGNS ===');

    await page.goto(`${BASE}/PolicyDistribution.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(5000);

    await snap(page, '5-1-distribution');

    const bodyText = await page.textContent('body') || '';
    console.log(`  Contains "Distribution": ${bodyText.includes('Distribution')}`);
    console.log(`  Contains "Campaign": ${bodyText.includes('Campaign')}`);
    console.log(`  Contains "Create": ${bodyText.includes('Create')}`);

    // Check KPIs
    const kpis = ['Total', 'Active', 'Completed', 'Acknowledged', 'Overdue'];
    for (const kpi of kpis) {
      console.log(`    KPI "${kpi}": ${bodyText.includes(kpi)}`);
    }

    // Look for Create Campaign button
    const createBtn = page.locator('button').filter({ hasText: /Create.*Campaign|New.*Campaign|Create/i }).first();
    const hasCreate = await createBtn.isVisible().catch(() => false);
    console.log(`  Create Campaign button: ${hasCreate}`);

    if (hasCreate) {
      await createBtn.click();
      await page.waitForTimeout(3000);

      await snap(page, '5-1-create-campaign');

      const panelBody = await page.textContent('body') || '';
      console.log(`  Campaign panel opened: ${panelBody.includes('Campaign') || panelBody.includes('Name') || panelBody.includes('Select')}`);

      // Fill campaign name if visible
      const nameInput = page.locator('input[type="text"]:not([readonly]):not([disabled])').first();
      if (await nameInput.isVisible().catch(() => false)) {
        await nameInput.clear();
        await nameInput.fill('E2E Test Campaign — Q3 2026 Policy Rollout');
        console.log(`  ✅ Campaign name filled`);
      }

      // Check for policy selector
      const policySelector = page.locator('.ms-Dropdown, select, [class*="selector"]').first();
      console.log(`  Policy selector: ${await policySelector.isVisible().catch(() => false)}`);
    }
  });

  test('5.2 — View existing campaign details', async ({ page }) => {
    console.log('\n=== VIEW CAMPAIGN DETAILS ===');

    await page.goto(`${BASE}/PolicyDistribution.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(5000);

    // Click on first campaign card
    const campaignCards = page.locator('[style*="cursor: pointer"], [style*="cursor:pointer"]').filter({ hasText: /Campaign|Policy|Update/i });
    const cardCount = await campaignCards.count();
    console.log(`  Campaign cards: ${cardCount}`);

    if (cardCount > 0) {
      await campaignCards.first().click();
      await page.waitForTimeout(2000);

      await snap(page, '5-2-campaign-detail');

      const bodyText = await page.textContent('body') || '';
      console.log(`  Contains "Recipients": ${bodyText.includes('Recipients') || bodyText.includes('recipients')}`);
      console.log(`  Contains progress %: ${bodyText.includes('%')}`);
      console.log(`  Contains "Acknowledged": ${bodyText.includes('Acknowledged')}`);
    }
  });
});
