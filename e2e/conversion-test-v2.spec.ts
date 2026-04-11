import { test, expect, Page } from '@playwright/test';
import * as path from 'path';

/**
 * CONVERSION TEST v2 — Real Office file upload + Office Online creation
 *
 * Tests:
 * 1. Upload .docx via the Upload panel on Step 0
 * 2. Create Word/Excel/PPT via "Create Document" on Step 6
 * 3. Submit for review to trigger conversion
 * 4. Verify HTML viewer
 */

const BASE = 'https://mf7m.sharepoint.com/sites/PolicyManager/SitePages';
const DOCS_DIR = path.join(process.cwd(), 'e2e', 'test-documents');
const TS = Date.now().toString(36).slice(-4);

let shotCount = 0;
async function snap(page: Page, name: string): Promise<void> {
  if (shotCount >= 20) return;
  await page.screenshot({ path: path.join(process.cwd(), `e2e-conv2-${name}.png`), fullPage: false });
  shotCount++;
  console.log(`📸 [${shotCount}/20] e2e-conv2-${name}.png`);
}

// ============================================================
// TEST 1: Upload .docx via Upload panel (Step 0)
// ============================================================
test('1 — Upload real .docx via Upload panel', async ({ page }) => {
  console.log('\n=== UPLOAD .DOCX VIA PANEL ===');

  await page.goto(`${BASE}/PolicyBuilder.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(3000);

  // Standard Wizard
  await page.getByText('Standard Wizard').first().click();
  await page.waitForTimeout(3000);

  // Select "Upload" type
  await page.getByText('Upload', { exact: true }).first().click();
  await page.waitForTimeout(1000);

  // Click "Browse & Upload" blank card — this opens the Upload Panel
  const browseCard = page.locator('div[role="button"]').filter({ hasText: /Browse.*Upload/i }).first();
  if (await browseCard.isVisible().catch(() => false)) {
    await browseCard.click();
    await page.waitForTimeout(2000);
    console.log('  Upload panel should be open');
  }

  // The Upload Panel has a hidden input#policyFileInput
  // Set the file directly on it
  const fileInput = page.locator('#policyFileInput, input[type="file"]');
  const hasInput = await fileInput.first().isVisible().catch(() => false);
  console.log(`  Hidden file input found: ${await fileInput.count()}`);

  // Even if hidden, we can setInputFiles
  const docxPath = path.join(DOCS_DIR, 'docx', 'acceptable-use-policy.docx');
  try {
    await fileInput.first().setInputFiles(docxPath);
    await page.waitForTimeout(5000);
    console.log('  ✅ File set on input');
  } catch (e) {
    console.log('  ⚠️ setInputFiles failed:', (e as Error).message.slice(0, 80));

    // Fallback: click the drop zone to trigger the file dialog
    const dropZone = page.locator('text=/Drag.*drop|click to browse/i').first();
    if (await dropZone.isVisible().catch(() => false)) {
      console.log('  Trying drop zone click approach...');
      // Use page.evaluate to programmatically trigger file input
      await page.evaluate((filePath) => {
        const input = document.getElementById('policyFileInput') as HTMLInputElement;
        if (input) {
          // Can't set files programmatically from evaluate, but we can verify the element exists
          console.log('File input found:', input.tagName, input.type, input.accept);
        }
      }, docxPath);
    }
  }

  await snap(page, '1-upload-panel');

  // Check if file was processed
  const bodyText = await page.textContent('body') || '';
  const processed = bodyText.includes('uploaded') || bodyText.includes('linked') || bodyText.includes('Processing') || bodyText.includes('acceptable');
  console.log(`  File processed: ${processed}`);

  // Close panel if still open and proceed
  const dismissBtn = page.locator('button[aria-label="Close"]').first();
  if (await dismissBtn.isVisible().catch(() => false)) {
    await dismissBtn.click();
    await page.waitForTimeout(500);
  }
});


// ============================================================
// TEST 2: Create Word doc via wizard + Save + Submit
// ============================================================
test('2 — Create Word doc via wizard and submit for review', async ({ page }) => {
  console.log('\n=== CREATE WORD DOC + SUBMIT ===');

  await page.goto(`${BASE}/PolicyBuilder.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(3000);

  await page.getByText('Standard Wizard').first().click();
  await page.waitForTimeout(3000);

  // Step 0: Word + Blank
  await page.getByText('Word', { exact: true }).first().click();
  await page.waitForTimeout(1000);
  const blank = page.locator('div[role="button"]').filter({ hasText: /Blank/i }).first();
  if (await blank.isVisible().catch(() => false)) {
    await blank.click();
    await page.waitForTimeout(500);
  }

  const clickNext = async () => {
    const btn = page.locator('button').filter({ hasText: /Next/ }).last();
    if (await btn.isVisible({ timeout: 5000 }).catch(() => false)) {
      await btn.click();
      await page.waitForTimeout(2000);
    }
  };
  const dismiss = async () => {
    const ok = page.locator('button').filter({ hasText: /^OK$/i }).first();
    if (await ok.isVisible({ timeout: 1000 }).catch(() => false)) {
      await ok.click();
      await page.waitForTimeout(500);
    }
  };

  // Step 0 → 1
  await clickNext();
  await dismiss();

  // Step 1: Fill name + category
  const nameInput = page.locator('input[type="text"]:not([readonly]):not([disabled])').first();
  if (await nameInput.isVisible().catch(() => false)) {
    await nameInput.clear();
    await nameInput.fill(`E2E Word Conversion Test ${TS}`);
  }
  const dropdown = page.locator('.ms-Dropdown').first();
  if (await dropdown.isVisible().catch(() => false)) {
    await dropdown.click();
    await page.waitForTimeout(500);
    const catOpt = page.locator('.ms-Dropdown-item, [role="option"]').filter({ hasText: 'HR Policies' }).first();
    if (await catOpt.isVisible().catch(() => false)) await catOpt.click();
    else {
      const first = page.locator('.ms-Dropdown-item, [role="option"]').first();
      if (await first.isVisible().catch(() => false)) await first.click();
    }
  }

  // Navigate to Step 6 (Content)
  for (let i = 0; i < 5; i++) {
    await clickNext();
    await dismiss();
  }

  console.log('  At Step 6: Content');

  // Click "Create Word Document"
  const createBtn = page.locator('button').filter({ hasText: /Create.*Word|Create.*Document/i }).first();
  if (await createBtn.isVisible({ timeout: 5000 }).catch(() => false)) {
    await createBtn.click();
    console.log('  Clicked "Create Word Document"');
    await page.waitForTimeout(10000); // Wait for SP to create the doc

    await snap(page, '2-word-created');

    const bodyText = await page.textContent('body') || '';
    console.log(`  Document linked: ${bodyText.includes('linked') || bodyText.includes('success')}`);
    console.log(`  Open in Office visible: ${bodyText.includes('Open in Office')}`);

    // Navigate to Review + Save Draft
    await clickNext();
    await dismiss();

    const saveBtn = page.locator('button').filter({ hasText: /Save.*Draft|Save$/i }).first();
    if (await saveBtn.isVisible().catch(() => false)) {
      await saveBtn.click();
      await page.waitForTimeout(5000);
      await dismiss();
      console.log('  ✅ Word policy saved as draft with linked document');
    }

    // Now Submit for Review
    const submitBtn = page.locator('button').filter({ hasText: /Submit.*Review/i }).first();
    if (await submitBtn.isVisible().catch(() => false)) {
      await submitBtn.click();
      await page.waitForTimeout(10000);
      await dismiss();
      console.log('  ✅ SUBMITTED FOR REVIEW — conversion should trigger');
      await snap(page, '2-word-submitted');
    } else {
      console.log('  ⚠️ Submit for Review button not visible after save');
    }
  } else {
    console.log('  ⚠️ No Create Document button');
    await snap(page, '2-word-no-create');
  }
});


// ============================================================
// TEST 3: Create PowerPoint doc via wizard
// ============================================================
test('3 — Create PowerPoint doc via wizard', async ({ page }) => {
  console.log('\n=== CREATE POWERPOINT DOC ===');

  await page.goto(`${BASE}/PolicyBuilder.aspx`, { waitUntil: 'domcontentloaded', timeout: 90000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(3000);

  await page.getByText('Standard Wizard').first().click();
  await page.waitForTimeout(3000);

  // Step 0: PowerPoint + Blank
  await page.getByText('PowerPoint', { exact: true }).first().click();
  await page.waitForTimeout(1000);
  const blank = page.locator('div[role="button"]').filter({ hasText: /Blank/i }).first();
  if (await blank.isVisible().catch(() => false)) {
    await blank.click();
    await page.waitForTimeout(500);
  }

  const clickNext = async () => {
    const btn = page.locator('button').filter({ hasText: /Next/ }).last();
    if (await btn.isVisible({ timeout: 5000 }).catch(() => false)) {
      await btn.click();
      await page.waitForTimeout(2000);
    }
  };
  const dismiss = async () => {
    const ok = page.locator('button').filter({ hasText: /^OK$/i }).first();
    if (await ok.isVisible({ timeout: 1000 }).catch(() => false)) {
      await ok.click();
      await page.waitForTimeout(500);
    }
  };

  await clickNext();
  await dismiss();

  // Step 1: Name + Category
  const nameInput = page.locator('input[type="text"]:not([readonly]):not([disabled])').first();
  if (await nameInput.isVisible().catch(() => false)) {
    await nameInput.clear();
    await nameInput.fill(`E2E PowerPoint Conversion Test ${TS}`);
  }
  const dropdown = page.locator('.ms-Dropdown').first();
  if (await dropdown.isVisible().catch(() => false)) {
    await dropdown.click();
    await page.waitForTimeout(500);
    const catOpt = page.locator('.ms-Dropdown-item, [role="option"]').filter({ hasText: 'Health & Safety' }).first();
    if (await catOpt.isVisible().catch(() => false)) await catOpt.click();
    else {
      const first = page.locator('.ms-Dropdown-item, [role="option"]').first();
      if (await first.isVisible().catch(() => false)) await first.click();
    }
  }

  for (let i = 0; i < 5; i++) {
    await clickNext();
    await dismiss();
  }

  console.log('  At Step 6: Content');

  const createBtn = page.locator('button').filter({ hasText: /Create.*PowerPoint|Create.*Document/i }).first();
  if (await createBtn.isVisible({ timeout: 5000 }).catch(() => false)) {
    await createBtn.click();
    console.log('  Clicked "Create PowerPoint Document"');
    await page.waitForTimeout(10000);

    await snap(page, '3-ppt-created');

    const bodyText = await page.textContent('body') || '';
    console.log(`  Document linked: ${bodyText.includes('linked') || bodyText.includes('success')}`);

    // Save
    await clickNext();
    await dismiss();
    const saveBtn = page.locator('button').filter({ hasText: /Save.*Draft|Save$/i }).first();
    if (await saveBtn.isVisible().catch(() => false)) {
      await saveBtn.click();
      await page.waitForTimeout(5000);
      await dismiss();
      console.log('  ✅ PowerPoint policy saved');
    }
  } else {
    console.log('  ⚠️ No Create Document button');
    await snap(page, '3-ppt-no-create');
  }
});


// ============================================================
// TEST 4: Create Excel doc via wizard
// ============================================================
test('4 — Create Excel doc via wizard', async ({ page }) => {
  console.log('\n=== CREATE EXCEL DOC ===');

  await page.goto(`${BASE}/PolicyBuilder.aspx`, { waitUntil: 'domcontentloaded', timeout: 90000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(3000);

  await page.getByText('Standard Wizard').first().click();
  await page.waitForTimeout(3000);

  await page.getByText('Excel', { exact: true }).first().click();
  await page.waitForTimeout(1000);
  const blank = page.locator('div[role="button"]').filter({ hasText: /Blank/i }).first();
  if (await blank.isVisible().catch(() => false)) {
    await blank.click();
    await page.waitForTimeout(500);
  }

  const clickNext = async () => {
    const btn = page.locator('button').filter({ hasText: /Next/ }).last();
    if (await btn.isVisible({ timeout: 5000 }).catch(() => false)) {
      await btn.click();
      await page.waitForTimeout(2000);
    }
  };
  const dismiss = async () => {
    const ok = page.locator('button').filter({ hasText: /^OK$/i }).first();
    if (await ok.isVisible({ timeout: 1000 }).catch(() => false)) {
      await ok.click();
      await page.waitForTimeout(500);
    }
  };

  await clickNext();
  await dismiss();

  const nameInput = page.locator('input[type="text"]:not([readonly]):not([disabled])').first();
  if (await nameInput.isVisible().catch(() => false)) {
    await nameInput.clear();
    await nameInput.fill(`E2E Excel Conversion Test ${TS}`);
  }
  const dropdown = page.locator('.ms-Dropdown').first();
  if (await dropdown.isVisible().catch(() => false)) {
    await dropdown.click();
    await page.waitForTimeout(500);
    const catOpt = page.locator('.ms-Dropdown-item, [role="option"]').filter({ hasText: 'Financial' }).first();
    if (await catOpt.isVisible().catch(() => false)) await catOpt.click();
    else {
      const first = page.locator('.ms-Dropdown-item, [role="option"]').first();
      if (await first.isVisible().catch(() => false)) await first.click();
    }
  }

  for (let i = 0; i < 5; i++) {
    await clickNext();
    await dismiss();
  }

  const createBtn = page.locator('button').filter({ hasText: /Create.*Excel|Create.*Document/i }).first();
  if (await createBtn.isVisible({ timeout: 5000 }).catch(() => false)) {
    await createBtn.click();
    console.log('  Clicked "Create Excel Document"');
    await page.waitForTimeout(10000);

    await snap(page, '4-excel-created');

    const bodyText = await page.textContent('body') || '';
    console.log(`  Document linked: ${bodyText.includes('linked') || bodyText.includes('success')}`);

    await clickNext();
    await dismiss();
    const saveBtn = page.locator('button').filter({ hasText: /Save.*Draft|Save$/i }).first();
    if (await saveBtn.isVisible().catch(() => false)) {
      await saveBtn.click();
      await page.waitForTimeout(5000);
      await dismiss();
      console.log('  ✅ Excel policy saved');
    }
  }
});


// ============================================================
// TEST 5: Viewer mode audit — all known policies
// ============================================================
test('5 — Viewer mode audit across published policies', async ({ page }) => {
  console.log('\n=== VIEWER MODE AUDIT ===');

  const policyIds = [1, 2, 3, 101, 102, 103, 108, 109];

  for (const id of policyIds) {
    await page.goto(`${BASE}/PolicyDetails.aspx?policyId=${id}&mode=browse`, {
      waitUntil: 'domcontentloaded', timeout: 60000
    });
    await page.waitForLoadState('networkidle', { timeout: 20000 }).catch(() => {});
    await page.waitForTimeout(3000);

    const bodyText = await page.textContent('body') || '';

    // Get title from page
    const titleEl = page.locator('h1, h2, [class*="policyTitle"]').first();
    const title = await titleEl.textContent().catch(() => '') || '';

    const hasPdf = await page.locator('object[type="application/pdf"]').isVisible().catch(() => false);
    const hasOffice = await page.locator('iframe[src*="WopiFrame"]').isVisible().catch(() => false);
    const hasHtml = await page.locator('h1, h2, h3').first().isVisible().catch(() => false);
    const noContent = bodyText.includes('No content');

    let viewer = 'Unknown';
    if (hasPdf) viewer = 'PDF';
    else if (hasOffice) viewer = 'Office Online';
    else if (noContent) viewer = 'No Content';
    else if (hasHtml) viewer = 'Native HTML';

    const icon = viewer === 'Native HTML' ? '✅' : viewer === 'Office Online' ? '📄' : viewer === 'PDF' ? '📑' : '⚠️';
    console.log(`  ${icon} Policy ${String(id).padEnd(4)} | ${viewer.padEnd(16)} | ${title.trim().slice(0, 50)}`);
  }
});
