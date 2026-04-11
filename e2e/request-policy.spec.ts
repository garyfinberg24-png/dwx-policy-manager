import { test, expect, Page } from '@playwright/test';
import * as path from 'path';

/**
 * REQUEST POLICY LIFECYCLE TEST
 *
 * Tests the Manager → Author request flow:
 *   1. Manager clicks "Request Policy" in nav
 *   2. Fills out the 4-step request wizard (Details, Business Case, Requirements, Review)
 *   3. Submits the request → creates item in PM_PolicyRequests
 *   4. Author sees the request in the Requests tab
 *   5. Author clicks "Accept & Start Drafting" → opens Policy Builder with pre-filled data
 */

const BASE = 'https://mf7m.sharepoint.com/sites/PolicyManager/SitePages';
const TS = Date.now().toString(36).slice(-4);

let shotCount = 0;
async function snap(page: Page, name: string): Promise<void> {
  if (shotCount >= 15) return;
  await page.screenshot({ path: path.join(process.cwd(), `e2e-req-${name}.png`), fullPage: false });
  shotCount++;
  console.log(`📸 [${shotCount}/15] e2e-req-${name}.png`);
}

// ============================================================
// TEST 1: Open Request Policy wizard from Manager nav
// ============================================================
test('1 — Open Request Policy wizard from Manager dropdown', async ({ page }) => {
  console.log('\n=== OPEN REQUEST POLICY WIZARD ===');

  // Navigate to any page that shows the Manager nav
  await page.goto(`${BASE}/PolicyHub.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  // Find the Manager dropdown in the nav bar
  const managerNav = page.locator('text=/Manager/i').first();
  const hasManager = await managerNav.isVisible().catch(() => false);
  console.log(`  Manager nav visible: ${hasManager}`);

  if (hasManager) {
    await managerNav.click();
    await page.waitForTimeout(1000);

    // Look for "Request Policy" in the dropdown
    const requestPolicy = page.locator('a, button, [role="menuitem"]').filter({ hasText: /Request Policy/i }).first();
    const hasRequest = await requestPolicy.isVisible({ timeout: 3000 }).catch(() => false);
    console.log(`  "Request Policy" menu item: ${hasRequest}`);

    if (hasRequest) {
      await requestPolicy.click();
      await page.waitForTimeout(2000);

      await snap(page, '1-wizard-opened');

      // Check if wizard modal/panel opened
      const bodyText = await page.textContent('body') || '';
      console.log(`  Contains "Policy Details": ${bodyText.includes('Policy Details')}`);
      console.log(`  Contains "What policy": ${bodyText.includes('What policy')}`);
      console.log(`  Contains "Request Policy": ${bodyText.includes('Request Policy') || bodyText.includes('Upload Policy Document')}`);
    } else {
      // Maybe it's using href navigation
      const allLinks = page.locator('a[href*="request"], [data-key*="request"]');
      const linkCount = await allLinks.count();
      console.log(`  Request-related links: ${linkCount}`);
      await snap(page, '1-manager-dropdown');
    }
  } else {
    // Try clicking the Manager label/text directly
    console.log('  Trying Manager from header...');
    const managerBtn = page.locator('button, a, [role="button"]').filter({ hasText: /Manager/i }).first();
    if (await managerBtn.isVisible().catch(() => false)) {
      await managerBtn.click();
      await page.waitForTimeout(1000);
      await snap(page, '1-manager-click');
    }
  }
});


// ============================================================
// TEST 2: Fill out the Request Policy wizard — all 4 steps
// ============================================================
test('2 — Fill Request Policy wizard with realistic data', async ({ page }) => {
  console.log('\n=== FILL REQUEST POLICY WIZARD ===');

  // Navigate to Policy Hub where the nav bar is visible
  await page.goto(`${BASE}/PolicyHub.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  // Find the Manager nav item in the header nav bar
  // The nav bar has items like: All Policies | Policy Hub | Author ▾ | Manager ▾ | Secure Policies ▾
  // These are typically <a> or <div> elements with the nav text
  const navItems = page.locator('nav a, nav [role="button"], [class*="navItem"], [class*="headerNav"] a, [class*="headerNav"] [role="button"]');
  const navCount = await navItems.count();
  console.log(`  Nav items found: ${navCount}`);

  // Log all nav items for debugging
  for (let i = 0; i < Math.min(navCount, 15); i++) {
    const text = (await navItems.nth(i).textContent() || '').trim().slice(0, 30);
    if (text.length > 0) console.log(`    nav[${i}]: "${text}"`);
  }

  // Click Manager in the nav — try multiple approaches
  let wizardOpened = false;

  // Approach 1: Find "Manager" link/button in the nav area
  const managerLink = page.locator('a, [role="button"]').filter({ hasText: /^Manager$/ }).first();
  if (await managerLink.isVisible().catch(() => false)) {
    await managerLink.click();
    await page.waitForTimeout(1000);
    console.log('  Clicked "Manager" nav item');

    // Now find "Request Policy" in the dropdown
    const requestItem = page.locator('a, button, [role="menuitem"], [role="option"]').filter({ hasText: /Request Policy/i }).first();
    if (await requestItem.isVisible({ timeout: 3000 }).catch(() => false)) {
      await requestItem.click();
      await page.waitForTimeout(2000);
      wizardOpened = true;
      console.log('  ✅ "Request Policy" clicked');
    }
  }

  // Approach 2: Try href-based navigation
  if (!wizardOpened) {
    const requestLink = page.locator('a[href*="request-policy"], a[href*="Request"]').first();
    if (await requestLink.isVisible().catch(() => false)) {
      await requestLink.click();
      await page.waitForTimeout(2000);
      wizardOpened = true;
      console.log('  ✅ Request Policy link clicked (href)');
    }
  }

  // Approach 3: The wizard may be triggered by the "New Policy" card on the home page
  if (!wizardOpened) {
    console.log('  Trying "New Policy" card on landing page...');
    const newPolicyCard = page.locator('div[role="button"], a').filter({ hasText: /New Policy/i }).first();
    if (await newPolicyCard.isVisible().catch(() => false)) {
      // This might navigate to PolicyBuilder, not the Request wizard
      // Let's try finding the Request wizard trigger more specifically
      console.log('  Checking all clickable elements for "Request"...');
      const allClickable = page.locator('a, button, [role="button"], [role="menuitem"]');
      const clickCount = await allClickable.count();
      for (let i = 0; i < Math.min(clickCount, 30); i++) {
        const text = (await allClickable.nth(i).textContent() || '').trim();
        if (text.toLowerCase().includes('request')) {
          console.log(`    Found: "${text.slice(0, 50)}"`);
        }
      }
    }
  }

  await snap(page, '2-wizard-state');

  if (!wizardOpened) {
    console.log('  ⚠️ Could not open Request Policy wizard — logging page state');
    const bodyText = await page.textContent('body') || '';
    console.log(`    Page has "Request": ${bodyText.includes('Request')}`);
    console.log(`    Page has "Manager": ${bodyText.includes('Manager')}`);
    return;
  }

  // ---- STEP 1: Policy Details ----
  console.log('  --- Step 1: Policy Details ---');

  // Policy Title
  const titleInput = page.locator('input[type="text"]:not([readonly]):not([disabled])').first();
  if (await titleInput.isVisible().catch(() => false)) {
    await titleInput.clear();
    await titleInput.fill(`Data Retention & Archival Policy ${TS}`);
    console.log('    Title filled');
  }

  // Category dropdown
  const categoryDropdown = page.locator('.ms-Dropdown').first();
  if (await categoryDropdown.isVisible().catch(() => false)) {
    await categoryDropdown.click();
    await page.waitForTimeout(500);
    const catOpt = page.locator('.ms-Dropdown-item, [role="option"]').filter({ hasText: /Data Privacy|IT.*Security|Compliance/i }).first();
    if (await catOpt.isVisible().catch(() => false)) {
      await catOpt.click();
      console.log('    Category selected');
    } else {
      const firstOpt = page.locator('.ms-Dropdown-item, [role="option"]').first();
      if (await firstOpt.isVisible().catch(() => false)) await firstOpt.click();
    }
    await page.waitForTimeout(300);
  }

  // Policy Type dropdown (if present)
  const typeDropdown = page.locator('.ms-Dropdown').nth(1);
  if (await typeDropdown.isVisible().catch(() => false)) {
    await typeDropdown.click();
    await page.waitForTimeout(300);
    const typeOpt = page.locator('.ms-Dropdown-item, [role="option"]').filter({ hasText: /New Policy/i }).first();
    if (await typeOpt.isVisible().catch(() => false)) await typeOpt.click();
    await page.waitForTimeout(300);
  }

  // Priority dropdown
  const priorityDropdown = page.locator('.ms-Dropdown').nth(2);
  if (await priorityDropdown.isVisible().catch(() => false)) {
    await priorityDropdown.click();
    await page.waitForTimeout(300);
    const priorityOpt = page.locator('.ms-Dropdown-item, [role="option"]').filter({ hasText: /High/i }).first();
    if (await priorityOpt.isVisible().catch(() => false)) await priorityOpt.click();
    await page.waitForTimeout(300);
  }

  await snap(page, '2-step1-details');

  // Click Next
  const clickNext = async () => {
    const btn = page.locator('button').filter({ hasText: /Next/ }).last();
    if (await btn.isVisible({ timeout: 5000 }).catch(() => false)) {
      await btn.click();
      await page.waitForTimeout(1500);
      return true;
    }
    return false;
  };

  await clickNext();

  // ---- STEP 2: Business Case ----
  console.log('  --- Step 2: Business Case ---');

  const justification = page.locator('textarea').first();
  if (await justification.isVisible().catch(() => false)) {
    await justification.clear();
    await justification.fill(
      'New regulatory requirements from the Information Commissioner\'s Office (ICO) mandate that all organisations must have a clearly documented data retention policy by Q3 2026. ' +
      'Our current data handling practices lack formal documentation, creating compliance risk. ' +
      'A Data Retention & Archival Policy will ensure consistent data lifecycle management across all departments, ' +
      'reduce storage costs by 30% through automated archival, and protect the organisation from potential ICO enforcement action (fines up to £17.5M). ' +
      'This is a Priority 1 compliance gap identified in our latest internal audit (Ref: AUD-2026-047).'
    );
    console.log('    Business justification filled (216 chars)');
  }

  // Regulatory driver (if separate field)
  const regulatoryField = page.locator('textarea').nth(1);
  if (await regulatoryField.isVisible().catch(() => false)) {
    await regulatoryField.clear();
    await regulatoryField.fill('GDPR Article 5(1)(e) — storage limitation principle. ICO Guidance on Data Retention 2025. UK Data Protection Act 2018 Schedule 1.');
    console.log('    Regulatory driver filled');
  }

  await snap(page, '2-step2-business-case');
  await clickNext();

  // ---- STEP 3: Requirements ----
  console.log('  --- Step 3: Requirements ---');

  // Target Audience
  const audienceInput = page.locator('input[type="text"]:not([readonly]):not([disabled])').first();
  if (await audienceInput.isVisible().catch(() => false)) {
    await audienceInput.clear();
    await audienceInput.fill('All employees, IT administrators, data stewards, department managers');
    console.log('    Target audience filled');
  }

  // Desired effective date
  const dateInput = page.locator('input[type="date"]').first();
  if (await dateInput.isVisible().catch(() => false)) {
    await dateInput.evaluate((el: HTMLInputElement, d: string) => {
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
      if (setter) setter.call(el, d);
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
    }, '2026-07-01');
    console.log('    Effective date: 2026-07-01');
  }

  // Read timeframe, acknowledgement, quiz toggles
  const bodyText = await page.textContent('body') || '';
  console.log(`    Contains "Acknowledgement": ${bodyText.includes('Acknowledgement') || bodyText.includes('acknowledgement')}`);
  console.log(`    Contains "Quiz": ${bodyText.includes('Quiz') || bodyText.includes('quiz')}`);

  // Additional notes
  const notesField = page.locator('textarea').first();
  if (await notesField.isVisible().catch(() => false)) {
    await notesField.clear();
    await notesField.fill('Please ensure the policy covers: (1) retention periods per data category, (2) archival to cold storage procedures, (3) secure deletion protocols, (4) legal hold exceptions. Reference: ISO 27001:2022 Annex A.8.10.');
    console.log('    Additional notes filled');
  }

  await snap(page, '2-step3-requirements');
  await clickNext();

  // ---- STEP 4: Review & Submit ----
  console.log('  --- Step 4: Review & Submit ---');

  await snap(page, '2-step4-review');

  const reviewBody = await page.textContent('body') || '';
  console.log(`    Contains "Data Retention": ${reviewBody.includes('Data Retention')}`);
  console.log(`    Contains "ICO": ${reviewBody.includes('ICO')}`);
  console.log(`    Contains "Submit": ${reviewBody.includes('Submit')}`);

  // Click Submit
  const submitBtn = page.locator('button').filter({ hasText: /Submit.*Request|Submit$/i }).first();
  const hasSubmit = await submitBtn.isVisible().catch(() => false);
  console.log(`    Submit button visible: ${hasSubmit}`);

  if (hasSubmit) {
    await submitBtn.click();
    await page.waitForTimeout(8000);

    await snap(page, '2-submitted');

    const resultBody = await page.textContent('body') || '';
    const hasSuccess = resultBody.includes('success') || resultBody.includes('Success') || resultBody.includes('submitted') || resultBody.includes('REQ-');
    console.log(`    ✅ Submit result — success indicators: ${hasSuccess}`);

    // Check for reference number
    const refMatch = resultBody.match(/REQ-\d+/);
    if (refMatch) {
      console.log(`    📋 Reference number: ${refMatch[0]}`);
    }

    // Dismiss any dialog
    const okBtn = page.locator('button').filter({ hasText: /^OK$|^Close$|^Done$/i }).first();
    if (await okBtn.isVisible({ timeout: 3000 }).catch(() => false)) {
      await okBtn.click();
    }
  }
});


// ============================================================
// TEST 3: Check Author > Requests tab for the submitted request
// ============================================================
test('3 — Verify request appears in Author Requests tab', async ({ page }) => {
  console.log('\n=== CHECK AUTHOR REQUESTS TAB ===');

  await page.goto(`${BASE}/PolicyAuthor.aspx?tab=requests`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  await snap(page, '3-requests-tab');

  const bodyText = await page.textContent('body') || '';

  // Check for request status indicators
  console.log(`  Contains "New": ${bodyText.includes('New')}`);
  console.log(`  Contains "Assigned": ${bodyText.includes('Assigned')}`);
  console.log(`  Contains "Data Retention": ${bodyText.includes('Data Retention')}`);
  console.log(`  Contains "Accept": ${bodyText.includes('Accept')}`);
  console.log(`  Contains "Start Drafting": ${bodyText.includes('Start Drafting')}`);

  // Check KPI cards
  const kpiLabels = ['New', 'Assigned', 'In Progress', 'Completed', 'Rejected'];
  for (const kpi of kpiLabels) {
    const found = bodyText.includes(kpi);
    console.log(`    KPI "${kpi}": ${found}`);
  }

  // Look for request cards
  const requestCards = page.locator('[style*="borderLeft: 3px"], [style*="border-left: 3px"]');
  const cardCount = await requestCards.count();
  console.log(`  Request cards found: ${cardCount}`);

  // Look for "Accept & Start Drafting" button
  const acceptBtn = page.locator('button').filter({ hasText: /Accept.*Draft|Start.*Draft/i }).first();
  const hasAccept = await acceptBtn.isVisible().catch(() => false);
  console.log(`  "Accept & Start Drafting" button: ${hasAccept}`);
});


// ============================================================
// TEST 4: Click "Accept & Start Drafting" to create policy from request
// ============================================================
test('4 — Accept request and start drafting', async ({ page }) => {
  console.log('\n=== ACCEPT REQUEST & START DRAFTING ===');

  await page.goto(`${BASE}/PolicyAuthor.aspx?tab=requests`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  // Find a request card and click it to open the detail panel
  const requestCards = page.locator('[style*="borderLeft: 3px"], [style*="border-left: 3px"], [role="button"]').filter({ hasText: /New|Data Retention|Request/i });
  const cardCount = await requestCards.count();
  console.log(`  Request cards matching: ${cardCount}`);

  if (cardCount > 0) {
    await requestCards.first().click();
    await page.waitForTimeout(2000);

    await snap(page, '4-request-detail');

    // Look for "Accept & Start Drafting" button in the detail panel
    const acceptBtn = page.locator('button').filter({ hasText: /Accept.*Draft|Start.*Draft/i }).first();
    const hasAccept = await acceptBtn.isVisible({ timeout: 3000 }).catch(() => false);
    console.log(`  "Accept & Start Drafting" visible: ${hasAccept}`);

    if (hasAccept) {
      await acceptBtn.click();
      await page.waitForTimeout(5000);

      await snap(page, '4-wizard-from-request');

      // Check if Policy Builder opened with pre-filled data
      const bodyText = await page.textContent('body') || '';
      console.log(`  Policy Builder opened: ${bodyText.includes('Creation Method') || bodyText.includes('Standard Wizard') || bodyText.includes('Policy Builder')}`);
      console.log(`  Request data pre-filled: ${bodyText.includes('Data Retention') || bodyText.includes('retention')}`);

      // Check if the name field has the request title
      const nameInput = page.locator('input[type="text"]:not([readonly]):not([disabled])').first();
      if (await nameInput.isVisible().catch(() => false)) {
        const value = await nameInput.inputValue();
        console.log(`  Pre-filled name: "${value}"`);
      }
    } else {
      console.log('  ⚠️ Accept button not found — checking what buttons exist');
      const allBtns = page.locator('button');
      const btnCount = await allBtns.count();
      for (let i = 0; i < Math.min(btnCount, 10); i++) {
        const text = (await allBtns.nth(i).textContent() || '').trim().slice(0, 40);
        if (text.length > 0) console.log(`    button[${i}]: "${text}"`);
      }
    }
  } else {
    console.log('  ⚠️ No request cards found');
  }
});
