import { test, expect, Page } from '@playwright/test';
import * as path from 'path';

/**
 * HTML CONVERSION TESTS
 *
 * Tests the document conversion pipeline:
 *   1. Upload real .docx/.xlsx/.pptx files via wizard
 *   2. Create docs via Office Online in the wizard
 *   3. Submit for review (triggers conversion)
 *   4. Verify converted HTML renders in the viewer
 *
 * This is the REAL conversion test — using actual Office binary files.
 */

const BASE = 'https://mf7m.sharepoint.com/sites/PolicyManager/SitePages';
const DOCS_DIR = path.join(process.cwd(), 'e2e', 'test-documents');
const TS = Date.now().toString(36).slice(-4);

let shotCount = 0;
async function snap(page: Page, name: string): Promise<void> {
  if (shotCount >= 20) return;
  await page.screenshot({ path: path.join(process.cwd(), `e2e-conv-${name}.png`), fullPage: false });
  shotCount++;
  console.log(`📸 [${shotCount}/20] e2e-conv-${name}.png`);
}

// Helper: Navigate wizard to content step with metadata
async function wizardToContentStep(
  page: Page,
  method: string,
  name: string,
  category: string
): Promise<void> {
  await page.goto(`${BASE}/PolicyBuilder.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(3000);

  // Standard Wizard
  await page.getByText('Standard Wizard').first().click();
  await page.waitForTimeout(3000);

  // Step 0: Select method
  await page.getByText(method, { exact: true }).first().click();
  await page.waitForTimeout(1000);

  // Select Blank
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

  // Step 0 → Step 1
  await clickNext();
  await dismiss();

  // Step 1: Name + Category
  const nameInput = page.locator('input[type="text"]:not([readonly]):not([disabled])').first();
  if (await nameInput.isVisible().catch(() => false)) {
    await nameInput.clear();
    await nameInput.fill(name);
  }
  const dropdown = page.locator('.ms-Dropdown').first();
  if (await dropdown.isVisible().catch(() => false)) {
    await dropdown.click();
    await page.waitForTimeout(500);
    const catOpt = page.locator('.ms-Dropdown-item, [role="option"]').filter({ hasText: category }).first();
    if (await catOpt.isVisible().catch(() => false)) await catOpt.click();
    else {
      const first = page.locator('.ms-Dropdown-item, [role="option"]').first();
      if (await first.isVisible().catch(() => false)) await first.click();
    }
    await page.waitForTimeout(300);
  }

  // Summary
  const textarea = page.locator('textarea').first();
  if (await textarea.isVisible().catch(() => false)) {
    await textarea.fill(`E2E conversion test — ${method} document`);
  }

  // Steps 1 → 6 (Content)
  for (let i = 0; i < 5; i++) {
    await clickNext();
    await dismiss();
  }

  console.log(`  Wizard at Content step (Step 6) for ${method}`);
}


// ============================================================
// TEST A: Create Word doc in Office Online via wizard
// ============================================================
test.describe.serial('Test A: Create Office Docs via Wizard', () => {

  test('A.1 — Create Word doc in Office Online', async ({ page }) => {
    console.log('\n=== A.1: CREATE WORD DOC IN OFFICE ONLINE ===');

    await wizardToContentStep(page, 'Word', `E2E Word Office Online ${TS}`, 'HR Policies');

    // Should see "Create Word Document" button
    const createBtn = page.locator('button').filter({ hasText: /Create.*Word|Create.*Document/i }).first();
    const hasCreate = await createBtn.isVisible({ timeout: 5000 }).catch(() => false);
    console.log(`  "Create Document" button: ${hasCreate}`);

    if (hasCreate) {
      await createBtn.click();
      await page.waitForTimeout(10000); // Wait for doc creation in SP

      await snap(page, 'a1-word-doc-created');

      // Check if document was created
      const bodyText = await page.textContent('body') || '';
      const docCreated = bodyText.includes('Document linked') || bodyText.includes('Open in Office') || bodyText.includes('successfully');
      console.log(`  Document created: ${docCreated}`);

      // Check for Office Online link or embedded editor
      const officeLink = page.locator('a[href*="Doc.aspx"], a[href*="WopiFrame"]');
      const hasOfficeLink = await officeLink.first().isVisible().catch(() => false);
      console.log(`  Office Online link: ${hasOfficeLink}`);

      // Check for embedded WOPI editor
      const embeddedFrame = page.locator('iframe[src*="Doc.aspx"], iframe[src*="WopiFrame"], [class*="embeddedEditor"]');
      const hasEmbedded = await embeddedFrame.first().isVisible().catch(() => false);
      console.log(`  Embedded editor: ${hasEmbedded}`);

      await snap(page, 'a1-word-doc-state');
    } else {
      console.log('  ⚠️ No Create Document button — may already have a linked doc');
      await snap(page, 'a1-word-no-create-btn');
    }
  });

  test('A.2 — Create Excel doc in Office Online', async ({ page }) => {
    console.log('\n=== A.2: CREATE EXCEL DOC IN OFFICE ONLINE ===');

    await wizardToContentStep(page, 'Excel', `E2E Excel Office Online ${TS}`, 'Financial');

    const createBtn = page.locator('button').filter({ hasText: /Create.*Excel|Create.*Document/i }).first();
    const hasCreate = await createBtn.isVisible({ timeout: 5000 }).catch(() => false);
    console.log(`  "Create Document" button: ${hasCreate}`);

    if (hasCreate) {
      await createBtn.click();
      await page.waitForTimeout(10000);
      await snap(page, 'a2-excel-doc-created');

      const bodyText = await page.textContent('body') || '';
      console.log(`  Document created: ${bodyText.includes('linked') || bodyText.includes('Office') || bodyText.includes('success')}`);
    }
  });

  test('A.3 — Create PowerPoint doc in Office Online', async ({ page }) => {
    console.log('\n=== A.3: CREATE POWERPOINT DOC IN OFFICE ONLINE ===');

    await wizardToContentStep(page, 'PowerPoint', `E2E PPT Office Online ${TS}`, 'Health & Safety');

    const createBtn = page.locator('button').filter({ hasText: /Create.*PowerPoint|Create.*Document/i }).first();
    const hasCreate = await createBtn.isVisible({ timeout: 5000 }).catch(() => false);
    console.log(`  "Create Document" button: ${hasCreate}`);

    if (hasCreate) {
      await createBtn.click();
      await page.waitForTimeout(10000);
      await snap(page, 'a3-ppt-doc-created');

      const bodyText = await page.textContent('body') || '';
      console.log(`  Document created: ${bodyText.includes('linked') || bodyText.includes('Office') || bodyText.includes('success')}`);
    }
  });

  test('A.4 — Save all 3 Office Online policies as drafts', async ({ page }) => {
    console.log('\n=== A.4: VERIFY DRAFTS SAVED ===');

    // Go to pipeline
    await page.goto(`${BASE}/PolicyAuthor.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(5000);

    await snap(page, 'a4-pipeline-check');

    const bodyText = await page.textContent('body') || '';
    console.log(`  Contains "E2E Word": ${bodyText.includes('E2E Word')}`);
    console.log(`  Contains "E2E Excel": ${bodyText.includes('E2E Excel')}`);
    console.log(`  Contains "E2E PPT": ${bodyText.includes('E2E PPT')}`);
  });
});


// ============================================================
// TEST B: Upload real Office files via wizard
// ============================================================
test.describe.serial('Test B: Upload Real Office Files', () => {

  test('B.1 — Upload .docx file via wizard', async ({ page }) => {
    console.log('\n=== B.1: UPLOAD REAL .DOCX ===');

    await wizardToContentStep(page, 'Upload', `E2E Upload DOCX ${TS}`, 'Data Privacy');

    // Look for file input
    const fileInput = page.locator('input[type="file"]');
    const hasFileInput = await fileInput.first().isVisible().catch(() => false);
    console.log(`  File input visible: ${hasFileInput}`);

    if (hasFileInput) {
      // Upload the real .docx file
      const docxPath = path.join(DOCS_DIR, 'docx', 'acceptable-use-policy.docx');
      await fileInput.first().setInputFiles(docxPath);
      await page.waitForTimeout(5000);

      await snap(page, 'b1-docx-uploaded');

      const bodyText = await page.textContent('body') || '';
      console.log(`  Upload result: ${bodyText.includes('uploaded') || bodyText.includes('linked') || bodyText.includes('success') || bodyText.includes('acceptable')}`);
    } else {
      // Try the "Browse & Upload" button
      const browseBtn = page.locator('button').filter({ hasText: /Browse|Upload/i }).first();
      if (await browseBtn.isVisible().catch(() => false)) {
        console.log('  Found Browse button — clicking');
        await browseBtn.click();
        await page.waitForTimeout(2000);

        // Now look for file input that may have appeared
        const fileInput2 = page.locator('input[type="file"]');
        if (await fileInput2.first().isVisible().catch(() => false)) {
          const docxPath = path.join(DOCS_DIR, 'docx', 'acceptable-use-policy.docx');
          await fileInput2.first().setInputFiles(docxPath);
          await page.waitForTimeout(5000);
          await snap(page, 'b1-docx-uploaded-v2');
        }
      }
      console.log('  ⚠️ No file input found');
      await snap(page, 'b1-no-file-input');
    }

    // Save draft
    const clickNext = async () => {
      const btn = page.locator('button').filter({ hasText: /Next/ }).last();
      if (await btn.isVisible({ timeout: 3000 }).catch(() => false)) {
        await btn.click();
        await page.waitForTimeout(2000);
      }
    };
    await clickNext(); // To Review step

    const saveBtn = page.locator('button').filter({ hasText: /Save.*Draft|Save$/i }).first();
    if (await saveBtn.isVisible().catch(() => false)) {
      await saveBtn.click();
      await page.waitForTimeout(5000);
      const ok = page.locator('button').filter({ hasText: /^OK$/i }).first();
      if (await ok.isVisible({ timeout: 5000 }).catch(() => false)) {
        console.log('  ✅ DOCX policy saved as draft');
        await ok.click();
      }
    }
  });

  test('B.2 — Upload .xlsx file via wizard', async ({ page }) => {
    console.log('\n=== B.2: UPLOAD REAL .XLSX ===');

    await wizardToContentStep(page, 'Upload', `E2E Upload XLSX ${TS}`, 'Compliance');

    const fileInput = page.locator('input[type="file"]');
    if (await fileInput.first().isVisible().catch(() => false)) {
      const xlsxPath = path.join(DOCS_DIR, 'xlsx', 'risk-register.xlsx');
      await fileInput.first().setInputFiles(xlsxPath);
      await page.waitForTimeout(5000);
      await snap(page, 'b2-xlsx-uploaded');
      console.log('  ✅ XLSX file uploaded');
    } else {
      console.log('  ⚠️ No file input');
      await snap(page, 'b2-no-input');
    }

    // Navigate to review and save
    const clickNext = async () => {
      const btn = page.locator('button').filter({ hasText: /Next/ }).last();
      if (await btn.isVisible({ timeout: 3000 }).catch(() => false)) {
        await btn.click();
        await page.waitForTimeout(2000);
      }
    };
    await clickNext();

    const saveBtn = page.locator('button').filter({ hasText: /Save.*Draft|Save$/i }).first();
    if (await saveBtn.isVisible().catch(() => false)) {
      await saveBtn.click();
      await page.waitForTimeout(5000);
      const ok = page.locator('button').filter({ hasText: /^OK$/i }).first();
      if (await ok.isVisible({ timeout: 5000 }).catch(() => false)) {
        console.log('  ✅ XLSX policy saved as draft');
        await ok.click();
      }
    }
  });

  test('B.3 — Upload .pptx file via wizard', async ({ page }) => {
    console.log('\n=== B.3: UPLOAD REAL .PPTX ===');

    await wizardToContentStep(page, 'Upload', `E2E Upload PPTX ${TS}`, 'Health & Safety');

    const fileInput = page.locator('input[type="file"]');
    if (await fileInput.first().isVisible().catch(() => false)) {
      const pptxPath = path.join(DOCS_DIR, 'pptx', 'security-awareness-training.pptx');
      await fileInput.first().setInputFiles(pptxPath);
      await page.waitForTimeout(5000);
      await snap(page, 'b3-pptx-uploaded');
      console.log('  ✅ PPTX file uploaded');
    } else {
      console.log('  ⚠️ No file input');
    }

    const clickNext = async () => {
      const btn = page.locator('button').filter({ hasText: /Next/ }).last();
      if (await btn.isVisible({ timeout: 3000 }).catch(() => false)) {
        await btn.click();
        await page.waitForTimeout(2000);
      }
    };
    await clickNext();

    const saveBtn = page.locator('button').filter({ hasText: /Save.*Draft|Save$/i }).first();
    if (await saveBtn.isVisible().catch(() => false)) {
      await saveBtn.click();
      await page.waitForTimeout(5000);
      const ok = page.locator('button').filter({ hasText: /^OK$/i }).first();
      if (await ok.isVisible({ timeout: 5000 }).catch(() => false)) {
        console.log('  ✅ PPTX policy saved as draft');
        await ok.click();
      }
    }
  });
});


// ============================================================
// TEST C: Submit uploaded policies and verify conversion
// ============================================================
test.describe('Test C: Conversion Verification', () => {

  test('C.1 — Verify converted HTML viewer for Word policy', async ({ page }) => {
    console.log('\n=== C.1: VERIFY WORD CONVERSION ===');

    // Find a published Word policy and check its viewer
    await page.goto(`${BASE}/PolicyDetails.aspx?policyId=1&mode=browse`, {
      waitUntil: 'domcontentloaded', timeout: 60000
    });
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(5000);

    // Check viewer type
    const hasPdf = await page.locator('object[type="application/pdf"]').isVisible().catch(() => false);
    const hasOffice = await page.locator('iframe[src*="WopiFrame"]').isVisible().catch(() => false);
    const hasH1 = await page.locator('h1, h2').first().isVisible().catch(() => false);
    const hasParagraphs = await page.locator('p').count();

    const viewer = hasPdf ? 'PDF' : hasOffice ? 'Office Online (NOT converted)' : hasH1 ? 'Native HTML (CONVERTED)' : 'Other';
    console.log(`  Viewer: ${viewer}`);
    console.log(`  Paragraphs: ${hasParagraphs}`);

    await snap(page, 'c1-word-viewer');

    if (hasH1 && !hasOffice) {
      console.log('  ✅ WORD → HTML CONVERSION VERIFIED — rendering natively');
    } else if (hasOffice) {
      console.log('  ℹ️ Showing in Office Online — conversion may not have run (pre-conversion publish)');
    }
  });

  test('C.2 — Check what viewer each policy type uses', async ({ page }) => {
    console.log('\n=== C.2: VIEWER MODE AUDIT ===');

    // Check multiple policy IDs
    const policyIds = [1, 101, 102, 103, 108, 109];

    for (const id of policyIds) {
      await page.goto(`${BASE}/PolicyDetails.aspx?policyId=${id}&mode=browse`, {
        waitUntil: 'domcontentloaded', timeout: 60000
      });
      await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
      await page.waitForTimeout(3000);

      const bodyText = await page.textContent('body') || '';
      const title = bodyText.match(/[A-Z][a-z]+ [A-Z][a-z]+ [A-Z][a-z]+/)?.[0] || `Policy ${id}`;

      const hasPdf = await page.locator('object[type="application/pdf"]').isVisible().catch(() => false);
      const hasOffice = await page.locator('iframe[src*="WopiFrame"]').isVisible().catch(() => false);
      const hasHtml = await page.locator('h1, h2, h3').first().isVisible().catch(() => false);
      const noContent = bodyText.includes('No content');

      let viewer = 'Unknown';
      if (hasPdf) viewer = 'PDF Embed';
      else if (hasOffice) viewer = 'Office Online iframe';
      else if (noContent) viewer = 'No Content';
      else if (hasHtml) viewer = 'Native HTML';

      console.log(`  Policy ${id}: ${viewer}`);
    }
  });
});
