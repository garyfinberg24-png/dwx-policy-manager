import { test, Page } from '@playwright/test';
import * as path from 'path';

const BASE = 'https://mf7m.sharepoint.com/sites/PolicyManager/SitePages';
const TS = Date.now().toString(36).slice(-4);

let shotCount = 0;
async function snap(page: Page, name: string): Promise<void> {
  if (shotCount >= 15) return;
  await page.screenshot({ path: path.join(process.cwd(), `e2e-reqw-${name}.png`), fullPage: false });
  shotCount++;
  console.log(`📸 [${shotCount}/15] e2e-reqw-${name}.png`);
}

async function openRequestWizard(page: Page): Promise<boolean> {
  // Click Manager button in nav
  const managerBtn = page.locator('button').filter({ hasText: /^Manager$/ }).first();
  if (!await managerBtn.isVisible({ timeout: 5000 }).catch(() => false)) return false;

  await managerBtn.click();
  await page.waitForTimeout(1500);

  // Click "Request Policy" in the dropdown — it's a text element inside the flyout
  const requestItem = page.getByText('Request Policy', { exact: true }).first();
  if (await requestItem.isVisible({ timeout: 3000 }).catch(() => false)) {
    await requestItem.click();
    await page.waitForTimeout(2000);
    return true;
  }
  return false;
}

test('1 — Open Request Policy wizard', async ({ page }) => {
  console.log('\n=== OPEN REQUEST POLICY WIZARD ===');

  await page.goto(`${BASE}/PolicyManagerView.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  const opened = await openRequestWizard(page);
  console.log(`  Wizard opened: ${opened}`);

  await snap(page, '1-wizard');

  if (opened) {
    const bodyText = await page.textContent('body') || '';
    console.log(`  Contains "Policy Details": ${bodyText.includes('Policy Details')}`);
    console.log(`  Contains "What policy": ${bodyText.includes('What policy')}`);
    console.log(`  Contains "policy title": ${bodyText.toLowerCase().includes('policy title') || bodyText.toLowerCase().includes('title')}`);
  }
});

test('2 — Fill and submit Request Policy wizard', async ({ page }) => {
  console.log('\n=== FILL & SUBMIT REQUEST ===');

  await page.goto(`${BASE}/PolicyManagerView.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  const opened = await openRequestWizard(page);
  if (!opened) {
    console.log('  ❌ Could not open wizard');
    return;
  }

  await snap(page, '2-step1-empty');

  // ---- STEP 1: Policy Details ----
  console.log('  --- Step 1: Policy Details ---');

  // Fill title
  const titleInput = page.locator('input[type="text"]:not([readonly]):not([disabled])').first();
  if (await titleInput.isVisible().catch(() => false)) {
    await titleInput.clear();
    await titleInput.fill(`Data Retention & Archival Policy ${TS}`);
    console.log('    ✅ Title');
  }

  // Category dropdown — may be native <select> or custom dropdown
  // Try native select first
  const nativeSelect = page.locator('select').first();
  const hasNativeSelect = await nativeSelect.isVisible().catch(() => false);
  console.log(`    Native <select>: ${hasNativeSelect}`);

  if (hasNativeSelect) {
    await nativeSelect.selectOption({ label: 'Data Privacy' });
    console.log('    ✅ Category (native select)');
  } else {
    // Try Fluent Dropdown
    const fluentDropdown = page.locator('.ms-Dropdown, [class*="dropdown"], [class*="Dropdown"]').first();
    if (await fluentDropdown.isVisible().catch(() => false)) {
      await fluentDropdown.click();
      await page.waitForTimeout(500);
      const opt = page.locator('.ms-Dropdown-item, [role="option"]').first();
      if (await opt.isVisible().catch(() => false)) {
        await opt.click();
        console.log('    ✅ Category (Fluent dropdown)');
      }
    } else {
      // Try clicking on "Select category..." text
      const selectCat = page.getByText('Select category').first();
      if (await selectCat.isVisible().catch(() => false)) {
        await selectCat.click();
        await page.waitForTimeout(500);
        const opt = page.locator('[role="option"], [role="listbox"] *').filter({ hasText: /Data Privacy|Compliance/i }).first();
        if (await opt.isVisible().catch(() => false)) {
          await opt.click();
          console.log('    ✅ Category (text click)');
        }
      }
    }
  }

  await snap(page, '2-step1-filled');

  // Next
  const clickNext = async () => {
    const btn = page.locator('button').filter({ hasText: /Next/ }).last();
    if (await btn.isVisible({ timeout: 3000 }).catch(() => false)) {
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
      'New ICO guidance mandates documented data retention policies by Q3 2026. Internal audit (AUD-2026-047) flagged this as Priority 1. ' +
      'Without formal policy, risk of GDPR fines up to £17.5M. Policy will reduce storage costs 30% and provide defensible deletion schedules.'
    );
    console.log('    ✅ Justification');
  }

  const regField = page.locator('textarea').nth(1);
  if (await regField.isVisible().catch(() => false)) {
    await regField.fill('GDPR Art 5(1)(e), ICO Retention Guidance 2025, UK DPA 2018, ISO 27001:2022 A.8.10');
    console.log('    ✅ Regulatory driver');
  }

  await snap(page, '2-step2');
  await clickNext();

  // ---- STEP 3: Requirements ----
  console.log('  --- Step 3: Requirements ---');

  // Target audience — could be input or textarea
  const audience = page.locator('input[type="text"]:not([readonly]):not([disabled]), textarea').first();
  if (await audience.isVisible().catch(() => false)) {
    await audience.clear();
    await audience.fill('All employees, IT admins, data stewards, department managers');
    console.log('    ✅ Audience');
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
    console.log('    ✅ Date');
  }

  await snap(page, '2-step3');
  await clickNext();

  // ---- STEP 4: Review & Submit ----
  console.log('  --- Step 4: Review & Submit ---');
  await snap(page, '2-step4');

  // Submit
  const submitBtn = page.locator('button').filter({ hasText: /Submit/i }).last();
  if (await submitBtn.isVisible().catch(() => false)) {
    console.log(`    Submit button text: "${await submitBtn.textContent()}"`);
    await submitBtn.click();
    await page.waitForTimeout(8000);

    await snap(page, '2-submitted');

    const resultBody = await page.textContent('body') || '';
    const hasSuccess = resultBody.includes('success') || resultBody.includes('Success') || resultBody.includes('submitted') || resultBody.includes('REQ-');
    console.log(`    ✅ Submitted — success: ${hasSuccess}`);

    const refMatch = resultBody.match(/REQ-\w+/);
    if (refMatch) console.log(`    📋 Reference: ${refMatch[0]}`);

    // Dismiss
    const okBtn = page.locator('button').filter({ hasText: /^OK$|^Close$|^Done$/i }).first();
    if (await okBtn.isVisible({ timeout: 3000 }).catch(() => false)) await okBtn.click();
  } else {
    console.log('    ⚠️ No Submit button found');
  }
});

test('3 — Verify request in Author Requests tab', async ({ page }) => {
  console.log('\n=== AUTHOR REQUESTS TAB ===');

  await page.goto(`${BASE}/PolicyAuthor.aspx?tab=requests`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  await snap(page, '3-requests');

  const bodyText = await page.textContent('body') || '';
  console.log(`  Contains "Data Retention": ${bodyText.includes('Data Retention')}`);
  console.log(`  Contains "Testing Policy": ${bodyText.includes('Testing Policy')}`);
  console.log(`  Contains "Policy Requests": ${bodyText.includes('Policy Requests')}`);

  // Click on any request card
  const cards = page.locator('div[style*="cursor"]').filter({ hasText: /Policy|Request|Testing/i });
  const cardCount = await cards.count();
  console.log(`  Request cards: ${cardCount}`);

  if (cardCount > 0) {
    await cards.first().click();
    await page.waitForTimeout(2000);
    await snap(page, '3-detail-panel');

    // Check for Accept button
    const acceptBtn = page.locator('button').filter({ hasText: /Accept|Start|Draft/i });
    const acceptCount = await acceptBtn.count();
    console.log(`  Accept/Draft buttons: ${acceptCount}`);
    for (let i = 0; i < acceptCount; i++) {
      console.log(`    [${i}]: "${(await acceptBtn.nth(i).textContent() || '').trim().slice(0, 40)}"`);
    }
  }
});
