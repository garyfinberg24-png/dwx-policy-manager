import { test, Page } from '@playwright/test';
import * as path from 'path';

const BASE = 'https://mf7m.sharepoint.com/sites/PolicyManager/SitePages';
const TS = Date.now().toString(36).slice(-4);

let shotCount = 0;
async function snap(page: Page, name: string): Promise<void> {
  if (shotCount >= 15) return;
  await page.screenshot({ path: path.join(process.cwd(), `e2e-reqfin-${name}.png`), fullPage: false });
  shotCount++;
  console.log(`📸 [${shotCount}/15] e2e-reqfin-${name}.png`);
}

test('1 — Open Request Policy wizard from Manager nav', async ({ page }) => {
  console.log('\n=== OPEN REQUEST POLICY WIZARD ===');

  await page.goto(`${BASE}/PolicyHub.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  // Click the "Manager" nav button (class contains "navItem")
  const managerBtn = page.locator('button[class*="navItem"]').filter({ hasText: 'Manager' }).first();
  const hasManager = await managerBtn.isVisible().catch(() => false);
  console.log(`  Manager nav button: ${hasManager}`);

  if (hasManager) {
    await managerBtn.click();
    await page.waitForTimeout(1500);
    await snap(page, '1-manager-dropdown');

    // Find "Request Policy" in the dropdown
    const dropdownItems = page.locator('a, button, [role="menuitem"], [role="option"], [class*="dropdownItem"], [class*="menuItem"]');
    const itemCount = await dropdownItems.count();
    console.log(`  Dropdown items: ${itemCount}`);

    // Log dropdown items
    for (let i = 0; i < Math.min(itemCount, 15); i++) {
      const text = (await dropdownItems.nth(i).textContent() || '').trim().slice(0, 40);
      if (text.length > 0 && text.length < 40) console.log(`    [${i}]: "${text}"`);
    }

    // Click "Request Policy"
    const requestItem = page.locator('a, button, [role="menuitem"]').filter({ hasText: /Request Policy/i }).first();
    const hasRequest = await requestItem.isVisible({ timeout: 3000 }).catch(() => false);
    console.log(`  "Request Policy" item: ${hasRequest}`);

    if (hasRequest) {
      await requestItem.click();
      await page.waitForTimeout(2000);
      await snap(page, '1-request-wizard-opened');

      const bodyText = await page.textContent('body') || '';
      console.log(`  Wizard opened: ${bodyText.includes('Policy Details') || bodyText.includes('What policy') || bodyText.includes('Upload Policy')}`);
    }
  }
});

test('2 — Fill Request Policy wizard and submit', async ({ page }) => {
  console.log('\n=== FILL & SUBMIT REQUEST POLICY ===');

  await page.goto(`${BASE}/PolicyHub.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  // Open Manager → Request Policy
  const managerBtn = page.locator('button[class*="navItem"]').filter({ hasText: 'Manager' }).first();
  if (await managerBtn.isVisible().catch(() => false)) {
    await managerBtn.click();
    await page.waitForTimeout(1500);

    const requestItem = page.locator('a, button, [role="menuitem"]').filter({ hasText: /Request Policy/i }).first();
    if (await requestItem.isVisible({ timeout: 3000 }).catch(() => false)) {
      await requestItem.click();
      await page.waitForTimeout(2000);
    }
  }

  // ---- STEP 1: Policy Details ----
  console.log('  --- Step 1: Policy Details ---');

  // Find all inputs and textareas in the wizard
  const allInputs = page.locator('input[type="text"]:not([readonly]):not([disabled])');
  const inputCount = await allInputs.count();
  console.log(`    Editable inputs: ${inputCount}`);

  // Fill policy title (first text input)
  if (inputCount > 0) {
    await allInputs.first().clear();
    await allInputs.first().fill(`Data Retention & Archival Policy ${TS}`);
    console.log('    ✅ Title filled');
  }

  // Fill category dropdown
  const dropdowns = page.locator('.ms-Dropdown');
  const ddCount = await dropdowns.count();
  console.log(`    Dropdowns: ${ddCount}`);

  if (ddCount > 0) {
    // First dropdown = Category
    await dropdowns.first().click();
    await page.waitForTimeout(500);
    const catOpt = page.locator('.ms-Dropdown-item, [role="option"]').first();
    if (await catOpt.isVisible().catch(() => false)) {
      await catOpt.click();
      console.log('    ✅ Category selected');
    }
    await page.waitForTimeout(300);
  }

  // Priority dropdown (if visible)
  if (ddCount > 2) {
    await dropdowns.nth(2).click();
    await page.waitForTimeout(300);
    const highOpt = page.locator('.ms-Dropdown-item, [role="option"]').filter({ hasText: /High/i }).first();
    if (await highOpt.isVisible().catch(() => false)) {
      await highOpt.click();
      console.log('    ✅ Priority: High');
    } else {
      await page.keyboard.press('Escape');
    }
  }

  await snap(page, '2-step1');

  // Click Next
  const clickNext = async (): Promise<boolean> => {
    const btn = page.locator('button').filter({ hasText: /Next/ }).last();
    if (await btn.isVisible({ timeout: 3000 }).catch(() => false)) {
      await btn.click();
      await page.waitForTimeout(1500);
      return true;
    }
    // Try any button that looks like "Next" or "Continue"
    const altBtn = page.locator('button').filter({ hasText: /Continue|next|→/i }).first();
    if (await altBtn.isVisible({ timeout: 2000 }).catch(() => false)) {
      await altBtn.click();
      await page.waitForTimeout(1500);
      return true;
    }
    return false;
  };

  const step1Next = await clickNext();
  console.log(`    Next clicked: ${step1Next}`);

  // ---- STEP 2: Business Case ----
  console.log('  --- Step 2: Business Case ---');

  const textarea = page.locator('textarea').first();
  if (await textarea.isVisible().catch(() => false)) {
    await textarea.clear();
    await textarea.fill(
      'New ICO guidance mandates documented data retention policies for all UK organisations by Q3 2026. ' +
      'Our internal audit (AUD-2026-047) identified this as a Priority 1 compliance gap. ' +
      'Without this policy, we face potential fines up to £17.5M under GDPR Article 83. ' +
      'The policy will ensure consistent data lifecycle management, reduce storage costs by 30%, ' +
      'and provide defensible deletion schedules for all data categories.'
    );
    console.log('    ✅ Business justification filled');
  }

  // Regulatory driver (second textarea if present)
  const textarea2 = page.locator('textarea').nth(1);
  if (await textarea2.isVisible().catch(() => false)) {
    await textarea2.clear();
    await textarea2.fill('GDPR Article 5(1)(e), ICO Retention Guidance 2025, UK DPA 2018 Schedule 1, ISO 27001:2022 A.8.10');
    console.log('    ✅ Regulatory driver filled');
  }

  await snap(page, '2-step2');
  await clickNext();

  // ---- STEP 3: Requirements ----
  console.log('  --- Step 3: Requirements ---');

  // Target audience (first editable input or textarea)
  const audienceInput = page.locator('input[type="text"]:not([readonly]):not([disabled])').first();
  if (await audienceInput.isVisible().catch(() => false)) {
    await audienceInput.clear();
    await audienceInput.fill('All employees, IT administrators, data stewards, department managers');
    console.log('    ✅ Target audience filled');
  }

  // Date
  const dateInput = page.locator('input[type="date"]').first();
  if (await dateInput.isVisible().catch(() => false)) {
    await dateInput.evaluate((el: HTMLInputElement, d: string) => {
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
      if (setter) setter.call(el, d);
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
    }, '2026-07-01');
    console.log('    ✅ Date: 2026-07-01');
  }

  // Notes textarea
  const notesArea = page.locator('textarea').first();
  if (await notesArea.isVisible().catch(() => false)) {
    await notesArea.clear();
    await notesArea.fill('Please cover: (1) retention periods per data category, (2) archival to cold storage, (3) secure deletion protocols, (4) legal hold exceptions.');
    console.log('    ✅ Additional notes filled');
  }

  await snap(page, '2-step3');
  await clickNext();

  // ---- STEP 4: Review & Submit ----
  console.log('  --- Step 4: Review & Submit ---');
  await snap(page, '2-step4-review');

  const submitBtn = page.locator('button').filter({ hasText: /Submit/i }).last();
  const hasSubmit = await submitBtn.isVisible().catch(() => false);
  console.log(`    Submit button: ${hasSubmit}`);

  if (hasSubmit) {
    const btnText = await submitBtn.textContent();
    console.log(`    Button text: "${btnText}"`);
    await submitBtn.click();
    await page.waitForTimeout(8000);

    await snap(page, '2-submitted');

    const resultBody = await page.textContent('body') || '';
    console.log(`    Success: ${resultBody.includes('success') || resultBody.includes('Success') || resultBody.includes('submitted') || resultBody.includes('REQ-')}`);

    // Look for reference number
    const refMatch = resultBody.match(/REQ-\w+/);
    if (refMatch) console.log(`    📋 Reference: ${refMatch[0]}`);

    // Dismiss dialog
    const okBtn = page.locator('button').filter({ hasText: /^OK$|^Close$|^Done$/i }).first();
    if (await okBtn.isVisible({ timeout: 3000 }).catch(() => false)) {
      await okBtn.click();
    }
  }
});

test('3 — Verify request in Author Requests tab + Accept', async ({ page }) => {
  console.log('\n=== AUTHOR REQUESTS TAB ===');

  await page.goto(`${BASE}/PolicyAuthor.aspx?tab=requests`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  await snap(page, '3-requests-tab');

  const bodyText = await page.textContent('body') || '';
  console.log(`  Contains "Data Retention": ${bodyText.includes('Data Retention')}`);
  console.log(`  Contains "New": ${bodyText.includes('New')}`);
  console.log(`  Contains "Accept": ${bodyText.includes('Accept')}`);

  // Click on a request card to open detail panel
  const cards = page.locator('[style*="cursor: pointer"], [style*="cursor:pointer"]').filter({ hasText: /Policy|Request/i });
  const cardCount = await cards.count();
  console.log(`  Clickable request cards: ${cardCount}`);

  if (cardCount > 0) {
    await cards.first().click();
    await page.waitForTimeout(2000);
    await snap(page, '3-request-detail');

    // Look for "Accept & Start Drafting"
    const acceptBtn = page.locator('button').filter({ hasText: /Accept|Start.*Draft/i }).first();
    const hasAccept = await acceptBtn.isVisible({ timeout: 3000 }).catch(() => false);
    console.log(`  "Accept" button: ${hasAccept}`);

    if (hasAccept) {
      console.log('  ✅ Accept & Start Drafting button found!');
      // Don't click it yet — just confirm it exists
    }
  }
});
