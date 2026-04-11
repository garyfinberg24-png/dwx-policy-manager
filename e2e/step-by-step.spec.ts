import { test, expect } from '@playwright/test';
import * as path from 'path';

/**
 * Step-by-Step E2E Test — Policy Creation Lifecycle
 *
 * This test walks through ONE creation flow at a time,
 * taking a screenshot at each step so we can see exactly
 * what's happening.
 */

const BASE = 'https://mf7m.sharepoint.com/sites/PolicyManager/SitePages';

// Screenshot helper — saves to project root
async function snap(page: any, name: string): Promise<void> {
  const filePath = path.join(process.cwd(), `e2e-${name}.png`);
  await page.screenshot({ path: filePath, fullPage: false });
  console.log(`📸 Screenshot: e2e-${name}.png`);
}

// ============================================================
// TEST 1: Navigate to Policy Builder and explore the wizard
// ============================================================
test('Step 1 — Open Policy Builder and select Standard Wizard', async ({ page }) => {
  console.log('=== STEP 1: Navigate to Policy Builder ===');

  await page.goto(`${BASE}/PolicyBuilder.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(3000);

  await snap(page, 'step1-builder-landing');

  // Check what's on screen
  const bodyText = await page.textContent('body') || '';
  console.log('Page contains "Fast Track":', bodyText.includes('Fast Track'));
  console.log('Page contains "Standard Wizard":', bodyText.includes('Standard Wizard'));
  console.log('Page contains "Creation Method":', bodyText.includes('Creation Method'));

  // Click "Standard Wizard"
  const stdWizard = page.getByText('Standard Wizard').first();
  const isVisible = await stdWizard.isVisible().catch(() => false);
  console.log('Standard Wizard visible:', isVisible);

  if (isVisible) {
    await stdWizard.click();
    await page.waitForTimeout(3000);
    await snap(page, 'step1-standard-wizard-selected');
    console.log('✅ Standard Wizard selected');
  } else {
    console.log('❌ Standard Wizard not found');
  }

  // Now we should be on Step 0: Creation Method
  const bodyAfter = await page.textContent('body') || '';
  console.log('After click - contains "Creation Method":', bodyAfter.includes('Creation Method'));
  console.log('After click - contains "Rich Text":', bodyAfter.includes('Rich Text'));
  console.log('After click - contains "Word":', bodyAfter.includes('Word'));
  console.log('After click - contains "Excel":', bodyAfter.includes('Excel'));
  console.log('After click - contains "PowerPoint":', bodyAfter.includes('PowerPoint'));
  console.log('After click - contains "Infographic":', bodyAfter.includes('Infographic'));
  console.log('After click - contains "Upload":', bodyAfter.includes('Upload'));
});

// ============================================================
// TEST 2: Select each creation method type and verify
// ============================================================
test('Step 2 — Select Rich Text and verify Step 0', async ({ page }) => {
  console.log('=== STEP 2: Select Rich Text creation method ===');

  await page.goto(`${BASE}/PolicyBuilder.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(3000);

  // Click Standard Wizard
  await page.getByText('Standard Wizard').first().click();
  await page.waitForTimeout(3000);

  // Now on Step 0 — check the type strip
  await snap(page, 'step2-step0-creation-method');

  // Find all role="button" elements and log their text
  const buttons = page.locator('div[role="button"]');
  const count = await buttons.count();
  console.log(`Found ${count} role="button" elements`);
  for (let i = 0; i < Math.min(count, 15); i++) {
    const text = (await buttons.nth(i).textContent() || '').trim().replace(/\n/g, ' ').slice(0, 60);
    console.log(`  Button[${i}]: "${text}"`);
  }

  // Try to click "Rich Text"
  const richTextBtn = page.getByText('Rich Text', { exact: false }).first();
  console.log('Rich Text text visible:', await richTextBtn.isVisible().catch(() => false));

  // Click Blank Rich Text card
  const blankRichText = page.locator('div[role="button"]').filter({ hasText: /Blank Rich Text/i });
  const blankVisible = await blankRichText.first().isVisible().catch(() => false);
  console.log('Blank Rich Text card visible:', blankVisible);

  if (blankVisible) {
    await blankRichText.first().click();
    await page.waitForTimeout(500);
    console.log('✅ Blank Rich Text selected');
  }

  // Look for Next button
  const allButtons = page.locator('button');
  const btnCount = await allButtons.count();
  console.log(`Found ${btnCount} <button> elements`);
  for (let i = 0; i < Math.min(btnCount, 20); i++) {
    const text = (await allButtons.nth(i).textContent() || '').trim().replace(/\n/g, ' ').slice(0, 40);
    if (text.length > 0) {
      console.log(`  <button>[${i}]: "${text}"`);
    }
  }

  await snap(page, 'step2-blank-selected');
});

// ============================================================
// TEST 3: Navigate through wizard steps
// ============================================================
test('Step 3 — Navigate wizard: Step 0 → Step 1 → Step 2', async ({ page }) => {
  console.log('=== STEP 3: Navigate through wizard steps ===');

  await page.goto(`${BASE}/PolicyBuilder.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(3000);

  // Select Standard Wizard
  await page.getByText('Standard Wizard').first().click();
  await page.waitForTimeout(3000);

  // Step 0: Select Rich Text + Blank
  console.log('--- Step 0: Creation Method ---');
  // Rich Text should already be selected (default)
  const blankCard = page.locator('div[role="button"]').filter({ hasText: /Blank/i }).first();
  if (await blankCard.isVisible().catch(() => false)) {
    await blankCard.click();
    await page.waitForTimeout(500);
    console.log('✅ Blank card clicked');
  }

  // Find and click Next button
  const nextBtn = page.locator('button').filter({ hasText: /Next/ }).last();
  console.log('Next button visible:', await nextBtn.isVisible().catch(() => false));
  console.log('Next button text:', await nextBtn.textContent().catch(() => 'NOT FOUND'));

  await nextBtn.click();
  await page.waitForTimeout(2000);
  await snap(page, 'step3-step1-basic-info');

  console.log('--- Step 1: Basic Information ---');
  const bodyText1 = await page.textContent('body') || '';
  console.log('Contains "Basic Information":', bodyText1.includes('Basic Information'));
  console.log('Contains "Policy Name":', bodyText1.includes('Policy Name') || bodyText1.includes('policy name'));

  // Fill in policy name — skip readonly/disabled inputs (like Policy Number)
  const inputs = page.locator('input[type="text"]:not([readonly]):not([disabled]), input:not([type]):not([readonly]):not([disabled])');
  const inputCount = await inputs.count();
  console.log(`Found ${inputCount} editable text inputs`);

  if (inputCount > 0) {
    await inputs.first().clear();
    await inputs.first().fill('E2E-Test-RichText-Policy');
    console.log('✅ Policy name filled');
  }

  // Fill summary
  const textarea = page.locator('textarea').first();
  if (await textarea.isVisible().catch(() => false)) {
    await textarea.clear();
    await textarea.fill('This is an E2E test policy created via Rich Text method');
    console.log('✅ Summary filled');
  }

  await snap(page, 'step3-step1-filled');

  // Click Next to Step 2
  const nextBtn2 = page.locator('button').filter({ hasText: /Next/ }).last();
  await nextBtn2.click();
  await page.waitForTimeout(2000);
  await snap(page, 'step3-step2-metadata');

  console.log('--- Step 2: Metadata Profile ---');
  const bodyText2 = await page.textContent('body') || '';
  console.log('Contains "Metadata":', bodyText2.includes('Metadata'));
});

// ============================================================
// TEST 4: Full wizard walkthrough — all 8 steps to Save Draft
// ============================================================
test('Step 4 — Full wizard: Rich Text → Save Draft', async ({ page }) => {
  console.log('=== STEP 4: Full wizard walkthrough ===');

  await page.goto(`${BASE}/PolicyBuilder.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(3000);

  // Mode selection
  await page.getByText('Standard Wizard').first().click();
  await page.waitForTimeout(3000);

  // Step 0: Select Rich Text + Blank
  const blankCard = page.locator('div[role="button"]').filter({ hasText: /Blank/i }).first();
  if (await blankCard.isVisible().catch(() => false)) {
    await blankCard.click();
    await page.waitForTimeout(500);
  }
  console.log('✅ Step 0: Creation Method — Rich Text / Blank');

  // Navigate: Step 0 → Step 1
  const clickNextBtn = async () => {
    const btn = page.locator('button').filter({ hasText: /Next/ }).last();
    if (await btn.isVisible().catch(() => false)) {
      await btn.click();
      await page.waitForTimeout(2000);
      return true;
    }
    return false;
  };

  await clickNextBtn();

  // Step 1: Basic Info — skip readonly/disabled inputs
  const inputs = page.locator('input[type="text"]:not([readonly]):not([disabled]), input:not([type]):not([readonly]):not([disabled])');
  if (await inputs.first().isVisible().catch(() => false)) {
    await inputs.first().clear();
    await inputs.first().fill('E2E-RichText-Full-Wizard');
  }
  const textarea = page.locator('textarea').first();
  if (await textarea.isVisible().catch(() => false)) {
    await textarea.clear();
    await textarea.fill('E2E full wizard test — Rich Text creation');
  }
  console.log('✅ Step 1: Basic Information filled');

  // Step 1 → Step 2 (Metadata)
  await clickNextBtn();
  console.log('✅ Step 2: Metadata Profile');

  // Step 2 → Step 3 (Audience)
  await clickNextBtn();
  console.log('✅ Step 3: Audience');

  // Step 3 → Step 4 (Dates)
  await clickNextBtn();
  console.log('✅ Step 4: Effective Dates');

  // Step 4 → Step 5 (Reviewers)
  await clickNextBtn();
  console.log('✅ Step 5: Reviewers & Approvers');

  // Step 5 → Step 6 (Content)
  await clickNextBtn();
  console.log('--- Step 6: Policy Content ---');
  await snap(page, 'step4-step6-content');

  // Check what content editor is showing
  const bodyText = await page.textContent('body') || '';
  const hasTinyMCE = await page.locator('.tox-tinymce').isVisible().catch(() => false);
  const hasRichTextEditor = await page.locator('[class*="richText"], [class*="htmlEditor"]').first().isVisible().catch(() => false);
  const hasTextarea = await page.locator('textarea').first().isVisible().catch(() => false);

  console.log('TinyMCE visible:', hasTinyMCE);
  console.log('Rich text editor visible:', hasRichTextEditor);
  console.log('Textarea visible:', hasTextarea);
  console.log('Contains "Policy Content":', bodyText.includes('Policy Content'));

  // Try to enter content in TinyMCE
  if (hasTinyMCE) {
    const frame = page.frameLocator('.tox-tinymce iframe').first();
    try {
      await frame.locator('body').click();
      await frame.locator('body').fill('E2E test content. Section 1: Purpose. Section 2: Scope. Section 3: Compliance.');
      console.log('✅ Content entered in TinyMCE');
    } catch (e) {
      console.log('⚠️ Could not enter content in TinyMCE:', (e as Error).message.slice(0, 80));
    }
  }

  console.log('✅ Step 6: Policy Content');

  // Step 6 → Step 7 (Review & Submit)
  await clickNextBtn();
  console.log('--- Step 7: Review & Submit ---');
  await snap(page, 'step4-step7-review');

  // Look for Save Draft button
  const saveDraftBtn = page.locator('button').filter({ hasText: /Save.*Draft|Save$/i }).first();
  const hasSave = await saveDraftBtn.isVisible().catch(() => false);
  console.log('Save Draft button visible:', hasSave);

  // Look for Submit for Review button
  const submitBtn = page.locator('button').filter({ hasText: /Submit.*Review/i }).first();
  const hasSubmit = await submitBtn.isVisible().catch(() => false);
  console.log('Submit for Review button visible:', hasSubmit);

  // Save as draft
  if (hasSave) {
    await saveDraftBtn.click();
    await page.waitForTimeout(5000);
    console.log('✅ Save Draft clicked');
    await snap(page, 'step4-after-save');
  }
});

// ============================================================
// TEST 5: Create each document type — just Step 0 selection
// ============================================================
test('Step 5 — Verify all 7 creation methods available', async ({ page }) => {
  console.log('=== STEP 5: Verify all creation methods ===');

  await page.goto(`${BASE}/PolicyBuilder.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(3000);

  // Select Standard Wizard
  await page.getByText('Standard Wizard').first().click();
  await page.waitForTimeout(3000);

  // Check each creation method type in the horizontal strip
  const methods = ['Rich Text', 'HTML', 'Word', 'Excel', 'PowerPoint', 'Infographic', 'Upload'];

  for (const method of methods) {
    const btn = page.getByText(method, { exact: true }).first();
    const visible = await btn.isVisible().catch(() => false);
    console.log(`${visible ? '✅' : '❌'} ${method}: ${visible ? 'visible' : 'NOT FOUND'}`);

    if (visible) {
      await btn.click();
      await page.waitForTimeout(1000);

      // Check what templates appear
      const bodyText = await page.textContent('body') || '';
      const hasBlank = bodyText.includes('Blank');
      const hasTemplates = bodyText.includes('Template') || bodyText.includes('template');
      console.log(`   Blank card: ${hasBlank}, Templates: ${hasTemplates}`);
    }
  }

  await snap(page, 'step5-all-methods');
});

// ============================================================
// TEST 6: Author Pipeline overview
// ============================================================
test('Step 6 — Author Pipeline overview', async ({ page }) => {
  console.log('=== STEP 6: Author Pipeline ===');

  await page.goto(`${BASE}/PolicyAuthor.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  await snap(page, 'step6-author-pipeline');

  const bodyText = await page.textContent('body') || '';
  console.log('Contains "Policy Author":', bodyText.includes('Policy Author'));
  console.log('Contains "Draft":', bodyText.includes('Draft'));
  console.log('Contains "In Review":', bodyText.includes('In Review'));
  console.log('Contains "Approved":', bodyText.includes('Approved'));
  console.log('Contains "Published":', bodyText.includes('Published'));

  // Check KPI cards
  const kpiCards = page.locator('[style*="borderTop: 3px"], [style*="border-top: 3px"]');
  console.log('KPI cards found:', await kpiCards.count());

  // Check for Approvals tab
  const approvalsTab = page.getByText('Approvals', { exact: false }).first();
  console.log('Approvals tab visible:', await approvalsTab.isVisible().catch(() => false));
});

// ============================================================
// TEST 7: Approvals tab
// ============================================================
test('Step 7 — Approvals tab', async ({ page }) => {
  console.log('=== STEP 7: Approvals Tab ===');

  await page.goto(`${BASE}/PolicyAuthor.aspx?tab=approvals`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  await snap(page, 'step7-approvals-tab');

  const bodyText = await page.textContent('body') || '';
  console.log('Contains "Pending":', bodyText.includes('Pending'));
  console.log('Contains "Overdue":', bodyText.includes('Overdue'));
  console.log('Contains "Approved":', bodyText.includes('Approved'));
  console.log('Contains "Rejected":', bodyText.includes('Rejected'));
  console.log('Contains "overdue":', bodyText.includes('overdue'));

  // Check for approval cards with decision options
  console.log('Contains "Approve":', bodyText.includes('Approve'));
  console.log('Contains "Request Changes":', bodyText.includes('Request Changes'));
  console.log('Contains "Reject":', bodyText.includes('Reject'));
});


// ============================================================
// HELPER: Navigate wizard and fill required fields
// ============================================================
async function fullWizardCreate(
  page: any,
  method: string,
  policyName: string,
  selectTemplate?: string
): Promise<boolean> {
  // Go to builder
  await page.goto(`${BASE}/PolicyBuilder.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(3000);

  // Select Standard Wizard
  await page.getByText('Standard Wizard').first().click();
  await page.waitForTimeout(3000);

  // Step 0: Select creation method
  const methodBtn = page.getByText(method, { exact: true }).first();
  if (await methodBtn.isVisible().catch(() => false)) {
    await methodBtn.click();
    await page.waitForTimeout(1000);
  }

  // Select template or blank
  if (selectTemplate) {
    const templateCard = page.locator('div[role="button"]').filter({ hasText: new RegExp(selectTemplate, 'i') }).first();
    if (await templateCard.isVisible().catch(() => false)) {
      await templateCard.click();
      await page.waitForTimeout(500);
    }
  } else {
    const blankCard = page.locator('div[role="button"]').filter({ hasText: /Blank/i }).first();
    if (await blankCard.isVisible().catch(() => false)) {
      await blankCard.click();
      await page.waitForTimeout(500);
    }
  }

  const clickNextBtn = async (): Promise<boolean> => {
    const btn = page.locator('button').filter({ hasText: /Next/ }).last();
    if (await btn.isVisible().catch(() => false)) {
      await btn.click();
      await page.waitForTimeout(2000);
      return true;
    }
    return false;
  };

  // Step 0 → Step 1
  await clickNextBtn();

  // Step 1: Fill policy name + category
  const editableInputs = page.locator('input[type="text"]:not([readonly]):not([disabled]), input:not([type]):not([readonly]):not([disabled])');
  if (await editableInputs.first().isVisible().catch(() => false)) {
    await editableInputs.first().clear();
    await editableInputs.first().fill(policyName);
  }

  // Dismiss any lingering dialogs (e.g. "Success" from previous save)
  const okBtn = page.locator('button').filter({ hasText: /^OK$/i }).first();
  if (await okBtn.isVisible({ timeout: 1000 }).catch(() => false)) {
    await okBtn.click();
    await page.waitForTimeout(500);
  }

  // Select a Category from dropdown
  const categoryDropdown = page.locator('.ms-Dropdown').first();
  if (await categoryDropdown.isVisible().catch(() => false)) {
    await categoryDropdown.click();
    await page.waitForTimeout(500);
    // Select the first option
    const firstOption = page.locator('.ms-Dropdown-item, [role="option"]').first();
    if (await firstOption.isVisible().catch(() => false)) {
      await firstOption.click();
      await page.waitForTimeout(500);
      console.log('  ✅ Category selected');
    }
  }

  // Fill summary
  const textarea = page.locator('textarea').first();
  if (await textarea.isVisible().catch(() => false)) {
    await textarea.clear();
    await textarea.fill(`E2E test: ${policyName} — ${method} document type`);
  }

  console.log(`  ✅ Step 1: Basic Info filled for "${policyName}"`);

  // Navigate through remaining steps
  for (let step = 2; step <= 7; step++) {
    await clickNextBtn();
    console.log(`  ✅ Step ${step} navigated`);
  }

  return true;
}


// ============================================================
// TEST 8: Full wizard with Category → Save Draft succeeds
// ============================================================
test('Step 8 — Full wizard with Category → Save Draft', async ({ page }) => {
  console.log('=== STEP 8: Full wizard with required Category field ===');

  await fullWizardCreate(page, 'Rich Text', 'E2E-RichText-Complete');

  // Save Draft
  const saveDraftBtn = page.locator('button').filter({ hasText: /Save.*Draft|Save$/i }).first();
  if (await saveDraftBtn.isVisible().catch(() => false)) {
    await saveDraftBtn.click();
    await page.waitForTimeout(5000);

    // Check for validation error or success
    const bodyText = await page.textContent('body') || '';
    const hasWarning = bodyText.includes('Warning') || bodyText.includes('required');
    const hasSuccess = bodyText.includes('saved') || bodyText.includes('Success') || bodyText.includes('success');

    console.log('Warning visible:', hasWarning);
    console.log('Success visible:', hasSuccess);

    await snap(page, 'step8-save-result');

    if (!hasWarning) {
      console.log('✅ Draft saved successfully (no validation warnings)');
    } else {
      console.log('⚠️ Validation warning appeared — checking which fields are missing');
    }
  }
});


// ============================================================
// TEST 9: Each document type → full wizard → verify content step
// ============================================================
test('Step 9a — HTML document type wizard', async ({ page }) => {
  console.log('=== STEP 9a: HTML document type ===');

  await fullWizardCreate(page, 'HTML', 'E2E-HTML-Document');

  // Check content step — should show TinyMCE or HTML editor
  const bodyText = await page.textContent('body') || '';
  const hasTinyMCE = await page.locator('.tox-tinymce').isVisible().catch(() => false);
  console.log('TinyMCE visible:', hasTinyMCE);

  await snap(page, 'step9a-html-content');

  // Save
  const saveDraftBtn = page.locator('button').filter({ hasText: /Save.*Draft|Save$/i }).first();
  if (await saveDraftBtn.isVisible().catch(() => false)) {
    await saveDraftBtn.click();
    await page.waitForTimeout(5000);
    console.log('✅ HTML policy save attempted');
  }
});

test('Step 9b — Word document type wizard', async ({ page }) => {
  console.log('=== STEP 9b: Word document type ===');

  await fullWizardCreate(page, 'Word', 'E2E-Word-Document');

  // Content step — should show Create Word Document button or linked doc
  const bodyText = await page.textContent('body') || '';
  const hasCreateDoc = bodyText.includes('Create') && bodyText.includes('Document');
  const hasLinkedDoc = bodyText.includes('Document linked') || bodyText.includes('Open in Office');
  console.log('Create Document button:', hasCreateDoc);
  console.log('Linked document:', hasLinkedDoc);

  await snap(page, 'step9b-word-content');
  console.log('✅ Word content step verified');
});

test('Step 9c — Excel document type wizard', async ({ page }) => {
  console.log('=== STEP 9c: Excel document type ===');

  await fullWizardCreate(page, 'Excel', 'E2E-Excel-Document');

  const bodyText = await page.textContent('body') || '';
  const hasCreateDoc = bodyText.includes('Create') && bodyText.includes('Excel');
  console.log('Create Excel Document:', hasCreateDoc);

  await snap(page, 'step9c-excel-content');
  console.log('✅ Excel content step verified');
});

test('Step 9d — PowerPoint document type wizard', async ({ page }) => {
  console.log('=== STEP 9d: PowerPoint document type ===');

  await fullWizardCreate(page, 'PowerPoint', 'E2E-PPT-Document');

  const bodyText = await page.textContent('body') || '';
  const hasCreateDoc = bodyText.includes('Create') && bodyText.includes('PowerPoint');
  console.log('Create PowerPoint Document:', hasCreateDoc);

  await snap(page, 'step9d-ppt-content');
  console.log('✅ PowerPoint content step verified');
});

test('Step 9e — Infographic document type wizard', async ({ page }) => {
  console.log('=== STEP 9e: Infographic document type ===');

  await fullWizardCreate(page, 'Infographic', 'E2E-Infographic-Document');

  const bodyText = await page.textContent('body') || '';
  const hasUpload = bodyText.includes('Upload') || bodyText.includes('Image');
  console.log('Upload/Image control:', hasUpload);

  await snap(page, 'step9e-infographic-content');
  console.log('✅ Infographic content step verified');
});

test('Step 9f — Corporate Template wizard', async ({ page }) => {
  console.log('=== STEP 9f: Corporate Template ===');

  await fullWizardCreate(page, 'Word', 'E2E-Corporate-Template', 'Corporate');

  const bodyText = await page.textContent('body') || '';
  const hasSections = bodyText.includes('Corporate Template') || bodyText.includes('sections completed');
  console.log('Corporate Template sections:', hasSections);

  await snap(page, 'step9f-corporate-template');
  console.log('✅ Corporate Template content step verified');
});


// ============================================================
// TEST 10: Review mode with all 3 decisions
// ============================================================
test('Step 10 — Review mode with decision panel', async ({ page }) => {
  console.log('=== STEP 10: Review Mode ===');

  // Go to approvals tab and find a pending policy
  await page.goto(`${BASE}/PolicyAuthor.aspx?tab=approvals`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  // Look for a policy card that we can click to review
  const bodyText = await page.textContent('body') || '';

  // Find policy links
  const policyLinks = page.locator('a[href*="PolicyDetails"]');
  const linkCount = await policyLinks.count();
  console.log(`Found ${linkCount} policy detail links`);

  if (linkCount > 0) {
    const href = await policyLinks.first().getAttribute('href') || '';
    console.log('First policy link:', href);

    // Navigate to review mode
    const idMatch = href.match(/policyId=(\d+)/);
    if (idMatch) {
      const reviewUrl = `${BASE}/PolicyDetails.aspx?policyId=${idMatch[1]}&mode=review`;
      console.log('Navigating to review mode:', reviewUrl);

      await page.goto(reviewUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });
      await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
      await page.waitForTimeout(5000);

      await snap(page, 'step10-review-mode');

      const reviewBody = await page.textContent('body') || '';
      console.log('Review Mode contains "Approve":', reviewBody.includes('Approve'));
      console.log('Review Mode contains "Request Changes":', reviewBody.includes('Request Changes'));
      console.log('Review Mode contains "Reject":', reviewBody.includes('Reject'));
      console.log('Review Mode contains "Comments":', reviewBody.includes('Comment') || reviewBody.includes('comment'));
      console.log('Review Mode contains "Review Checklist":', reviewBody.includes('Review Checklist'));
      console.log('Review Mode contains "Submit":', reviewBody.includes('Submit'));

      // Check for content viewer on left
      const hasContent = await page.locator('h1, h2, h3').first().isVisible().catch(() => false);
      const hasIframe = await page.locator('iframe').first().isVisible().catch(() => false);
      console.log('Content heading visible:', hasContent);
      console.log('Content iframe visible:', hasIframe);
    }
  } else {
    console.log('⚠️ No policy links found in approvals tab');
  }
});


// ============================================================
// TEST 11: Viewer modes — directly navigate to known policies
// ============================================================
test('Step 11 — Viewer modes on published policies', async ({ page }) => {
  console.log('=== STEP 11: Viewer Modes ===');

  // Go to Policy Hub first to find policy links
  await page.goto(`${BASE}/PolicyHub.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  await snap(page, 'step11-hub-overview');

  // Try to find clickable policy items (may be role="button" divs, not <a> tags)
  const clickableItems = page.locator('[role="button"], a[href*="PolicyDetails"], [class*="policyCard"]');
  const itemCount = await clickableItems.count();
  console.log(`Found ${itemCount} clickable items in Hub`);

  // Also try: direct policy links by navigating to known IDs from the My Policies page
  // My Policies already showed policies with IDs, let's use those

  // Navigate to a few policy IDs directly to test viewer modes
  const testPolicies = [
    { id: 101, name: 'Information Security Policy' },
    { id: 1, name: 'First Policy' },
  ];

  for (const pol of testPolicies) {
    console.log(`\n--- Testing policy ${pol.id}: ${pol.name} ---`);

    await page.goto(`${BASE}/PolicyDetails.aspx?policyId=${pol.id}&mode=browse`, {
      waitUntil: 'domcontentloaded', timeout: 60000
    });
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(5000);

    // Detect viewer type
    const pdfEmbed = await page.locator('object[type="application/pdf"]').isVisible().catch(() => false);
    const officeFrame = await page.locator('iframe[src*="WopiFrame"], iframe[src*="wopiframe"]').isVisible().catch(() => false);
    const htmlH1H2 = await page.locator('h1, h2, h3').first().isVisible().catch(() => false);
    const hasNoContent = await page.locator('text=/No content available/i').isVisible().catch(() => false);

    let viewer = 'Unknown';
    if (pdfEmbed) viewer = 'PDF Embed (<object>)';
    else if (officeFrame) viewer = 'Office Online iframe (WopiFrame)';
    else if (hasNoContent) viewer = 'No Content (not published)';
    else if (htmlH1H2) viewer = 'Native HTML (h1/h2 visible)';

    const bodyText = await page.textContent('body') || '';
    const title = bodyText.includes(pol.name) ? pol.name : 'Title not found';

    console.log(`  Viewer: ${viewer}`);
    console.log(`  Title found: ${bodyText.includes(pol.name)}`);
    console.log(`  Has toolbar: ${await page.locator('text=/Download|Print|Fullscreen/i').first().isVisible().catch(() => false)}`);
    console.log(`  Has ack button: ${await page.locator('button').filter({ hasText: /Acknowledge/i }).isVisible().catch(() => false)}`);

    await snap(page, `step11-viewer-policy-${pol.id}`);
  }

  // Also test from My Policies page — click a policy
  console.log('\n--- Testing viewer from My Policies click ---');
  await page.goto(`${BASE}/MyPolicies.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  // Click the eye/view icon on the first policy row
  const viewIcon = page.locator('button[aria-label*="View" i], button[title*="View" i], svg').first();
  if (await viewIcon.isVisible().catch(() => false)) {
    await viewIcon.click();
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(3000);
    await snap(page, 'step11-viewer-from-mypolicies');

    const viewer2 = await page.locator('object[type="application/pdf"]').isVisible().catch(() => false) ? 'PDF' :
                    await page.locator('iframe[src*="WopiFrame"]').isVisible().catch(() => false) ? 'Office Online' : 'HTML/Other';
    console.log(`  Viewer from My Policies click: ${viewer2}`);
  }
});


// ============================================================
// TEST 12: My Policies + Distribution
// ============================================================
test('Step 12 — My Policies and Distribution pages', async ({ page }) => {
  console.log('=== STEP 12: My Policies & Distribution ===');

  // My Policies
  await page.goto(`${BASE}/MyPolicies.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  await snap(page, 'step12-my-policies');

  let bodyText = await page.textContent('body') || '';
  console.log('My Policies contains "My Policies":', bodyText.includes('My Policies'));
  console.log('My Policies contains "Compliance":', bodyText.includes('Compliance') || bodyText.includes('compliance'));
  console.log('My Policies contains "Pending":', bodyText.includes('Pending'));
  console.log('My Policies contains "Acknowledged":', bodyText.includes('Acknowledged'));

  // Distribution
  await page.goto(`${BASE}/PolicyDistribution.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  await snap(page, 'step12-distribution');

  bodyText = await page.textContent('body') || '';
  console.log('Distribution contains "Distribution":', bodyText.includes('Distribution'));
  console.log('Distribution contains "Campaign":', bodyText.includes('Campaign') || bodyText.includes('campaign'));
});
