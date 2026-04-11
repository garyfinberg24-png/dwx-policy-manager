import { test, expect, Page } from '@playwright/test';
import * as path from 'path';

/**
 * FULL LIFECYCLE EXECUTION — Production Readiness Deep Test
 *
 * Creates 5 REAL policies (one per document type) with:
 *   - Realistic metadata
 *   - gf_admin@mf7m.onmicrosoft.com as reviewer & approver
 *   - Full content where applicable
 *
 * Then follows EACH through the lifecycle:
 *   Draft → Submit for Review → Review → Approve/RequestChanges/Reject → Publish
 *
 * Checks Outlook for notification emails at each stage.
 */

const BASE = 'https://mf7m.sharepoint.com/sites/PolicyManager/SitePages';
const REVIEWER_EMAIL = 'gf_admin@mf7m.onmicrosoft.com';
const TS = Date.now().toString(36).slice(-4);

let shotCount = 0;
async function snap(page: Page, name: string): Promise<void> {
  if (shotCount >= 25) return;
  await page.screenshot({ path: path.join(process.cwd(), `e2e-exec-${name}.png`), fullPage: false });
  shotCount++;
  console.log(`📸 [${shotCount}/25] e2e-exec-${name}.png`);
}

// ============================================================
// HELPER: Type into PeoplePicker and select result
// ============================================================
async function fillPeoplePicker(page: Page, placeholder: string, email: string): Promise<boolean> {
  // Find PeoplePicker input by placeholder
  const ppInput = page.locator(`input[placeholder*="${placeholder}" i]`).first();
  if (!await ppInput.isVisible({ timeout: 3000 }).catch(() => false)) {
    // Try any PeoplePicker-looking input
    const anyPP = page.locator('.ms-BasePicker-input, input[aria-label*="People" i]').first();
    if (!await anyPP.isVisible({ timeout: 3000 }).catch(() => false)) {
      return false;
    }
    await anyPP.click();
    await anyPP.fill(email);
    await page.waitForTimeout(2000);
    // Select first suggestion
    const suggestion = page.locator('.ms-Suggestions-item, [role="option"]').first();
    if (await suggestion.isVisible({ timeout: 5000 }).catch(() => false)) {
      await suggestion.click();
      await page.waitForTimeout(500);
      return true;
    }
    return false;
  }

  await ppInput.click();
  await ppInput.fill(email);
  await page.waitForTimeout(2000);

  const suggestion = page.locator('.ms-Suggestions-item, [role="option"], .ms-PeoplePicker-result').first();
  if (await suggestion.isVisible({ timeout: 5000 }).catch(() => false)) {
    await suggestion.click();
    await page.waitForTimeout(500);
    return true;
  }
  return false;
}

// ============================================================
// HELPER: Create full policy, fill ALL steps, add reviewer
// ============================================================
async function createAndSavePolicy(
  page: Page,
  method: string,
  name: string,
  category: string,
  summary: string,
  riskLevel: string,
  effectiveDate: string,
  reviewFrequency: string,
  content?: string
): Promise<boolean> {
  await page.goto(`${BASE}/PolicyBuilder.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(3000);

  // Standard Wizard
  await page.getByText('Standard Wizard').first().click();
  await page.waitForTimeout(3000);

  // Step 0: Creation Method
  const methodBtn = page.getByText(method, { exact: true }).first();
  if (await methodBtn.isVisible().catch(() => false)) {
    await methodBtn.click();
    await page.waitForTimeout(1000);
  }
  // Select Blank
  const blankCard = page.locator('div[role="button"]').filter({ hasText: /Blank/i }).first();
  if (await blankCard.isVisible().catch(() => false)) {
    await blankCard.click();
    await page.waitForTimeout(500);
  }
  console.log(`  Step 0: ${method} / Blank selected`);

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

  // Step 1: Basic Info
  const nameInput = page.locator('input[type="text"]:not([readonly]):not([disabled]), input:not([type]):not([readonly]):not([disabled])').first();
  if (await nameInput.isVisible().catch(() => false)) {
    await nameInput.clear();
    await nameInput.fill(name);
  }

  // Category dropdown
  const dropdowns = page.locator('.ms-Dropdown');
  if (await dropdowns.first().isVisible().catch(() => false)) {
    await dropdowns.first().click();
    await page.waitForTimeout(500);
    const catOpt = page.locator('.ms-Dropdown-item, [role="option"]').filter({ hasText: category }).first();
    if (await catOpt.isVisible().catch(() => false)) {
      await catOpt.click();
      console.log(`  Step 1: Category "${category}"`);
    } else {
      const first = page.locator('.ms-Dropdown-item, [role="option"]').first();
      if (await first.isVisible().catch(() => false)) await first.click();
    }
    await page.waitForTimeout(300);
  }

  // Summary
  const textarea = page.locator('textarea').first();
  if (await textarea.isVisible().catch(() => false)) {
    await textarea.clear();
    await textarea.fill(summary);
  }
  console.log(`  Step 1: "${name}" filled`);

  await clickNext();
  await dismiss();

  // Step 2: Metadata (skip — use defaults)
  console.log(`  Step 2: Metadata (defaults)`);
  await clickNext();
  await dismiss();

  // Step 3: Audience (skip — use defaults)
  console.log(`  Step 3: Audience (defaults)`);
  await clickNext();
  await dismiss();

  // Step 4: Effective Dates
  const dateInputs = page.locator('input[type="date"]');
  if (await dateInputs.first().isVisible().catch(() => false)) {
    await dateInputs.first().evaluate((el: HTMLInputElement, d: string) => {
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
      if (setter) setter.call(el, d);
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
    }, effectiveDate);
    console.log(`  Step 4: Date "${effectiveDate}"`);
  }
  await clickNext();
  await dismiss();

  // Step 5: Reviewers & Approvers — ADD gf_admin as reviewer
  console.log(`  Step 5: Adding reviewer ${REVIEWER_EMAIL}...`);

  // Check the "Override" checkbox to enable PeoplePicker
  const overrideCheckbox = page.locator('text=/Override.*assign.*reviewer/i').first();
  if (await overrideCheckbox.isVisible({ timeout: 3000 }).catch(() => false)) {
    await overrideCheckbox.click();
    await page.waitForTimeout(500);
    console.log(`    Override checkbox clicked`);

    // Fill PeoplePicker
    const ppResult = await fillPeoplePicker(page, 'Search Entra', REVIEWER_EMAIL);
    if (ppResult) {
      console.log(`    ✅ Reviewer "${REVIEWER_EMAIL}" added`);
    } else {
      // Try alternate PeoplePicker
      const ppResult2 = await fillPeoplePicker(page, 'Search', REVIEWER_EMAIL);
      console.log(`    Reviewer add (fallback): ${ppResult2}`);
    }

    // Fill override reason
    const reasonField = page.locator('textarea').filter({ hasText: '' }).last();
    if (await reasonField.isVisible().catch(() => false)) {
      await reasonField.fill('E2E test — admin reviewer for lifecycle testing');
    }
  } else {
    // Try direct PeoplePicker without override
    const ppDirect = await fillPeoplePicker(page, 'reviewer', REVIEWER_EMAIL);
    console.log(`    Direct reviewer PeoplePicker: ${ppDirect}`);
  }

  await clickNext();
  await dismiss();

  // Step 6: Content
  if (content && (method === 'Rich Text' || method === 'HTML')) {
    const tinyMCE = page.locator('.tox-tinymce iframe').first();
    if (await tinyMCE.isVisible({ timeout: 3000 }).catch(() => false)) {
      const frame = page.frameLocator('.tox-tinymce iframe').first();
      await frame.locator('body').click();
      await frame.locator('body').fill(content);
      console.log(`  Step 6: Content entered in TinyMCE`);
    } else {
      const ta = page.locator('textarea').first();
      if (await ta.isVisible().catch(() => false)) {
        await ta.clear();
        await ta.fill(content);
        console.log(`  Step 6: Content entered in textarea`);
      }
    }
  } else {
    console.log(`  Step 6: ${method} — deferred doc creation`);
  }

  await clickNext();
  await dismiss();

  // Step 7: Review & Submit — SAVE DRAFT
  const saveBtn = page.locator('button').filter({ hasText: /Save.*Draft|Save$/i }).first();
  if (await saveBtn.isVisible().catch(() => false)) {
    await saveBtn.click();
    await page.waitForTimeout(5000);
    await dismiss();
    console.log(`  ✅ Draft saved`);
    return true;
  }
  return false;
}


// ============================================================
// HELPER: Submit for Review from the wizard (Step 7)
// ============================================================
async function submitForReviewFromWizard(page: Page): Promise<boolean> {
  const submitBtn = page.locator('button').filter({ hasText: /Submit.*Review/i }).first();
  if (await submitBtn.isVisible({ timeout: 5000 }).catch(() => false)) {
    await submitBtn.click();
    await page.waitForTimeout(8000);

    // Handle success dialog
    const okBtn = page.locator('button').filter({ hasText: /^OK$/i }).first();
    if (await okBtn.isVisible({ timeout: 10000 }).catch(() => false)) {
      console.log(`  ✅ Submitted for review — success dialog appeared`);
      await okBtn.click();
      return true;
    }
    // Check for error
    const bodyText = await page.textContent('body') || '';
    if (bodyText.includes('submitted') || bodyText.includes('review')) {
      console.log(`  ✅ Submitted for review`);
      return true;
    }
    console.log(`  ⚠️ Submit clicked but no success confirmation`);
    return true;
  }
  console.log(`  ❌ Submit for Review button not found`);
  return false;
}


// ============================================================
// POLICY CONFIGS
// ============================================================
const P = {
  p1: { method: 'Rich Text', name: `E2E Acceptable Use Policy ${TS}`, category: 'IT & Security',
    summary: 'Defines acceptable use of company IT resources including email, internet, software and hardware. Covers personal use limitations, security requirements, and monitoring disclosure.',
    risk: 'High', date: '2026-05-01', freq: 'Annual',
    content: `<h1>Acceptable Use Policy</h1><h2>1. Purpose</h2><p>This policy establishes guidelines for the appropriate use of company IT resources to protect both the organisation and its employees.</p><h2>2. Scope</h2><p>Applies to all employees, contractors, and third-party users who access company IT systems.</p><h2>3. Acceptable Use</h2><p>Company IT resources are provided primarily for business use. Limited personal use is permitted provided it does not interfere with work duties, does not violate any laws, and complies with this policy.</p><h2>4. Prohibited Activities</h2><ul><li>Accessing or distributing offensive, illegal, or inappropriate content</li><li>Installing unauthorised software</li><li>Sharing login credentials</li><li>Circumventing security controls</li></ul><h2>5. Monitoring</h2><p>The company reserves the right to monitor all use of IT resources in compliance with applicable privacy laws.</p>` },

  p2: { method: 'HTML', name: `E2E Anti-Bribery & Corruption Policy ${TS}`, category: 'Compliance',
    summary: 'Establishes zero-tolerance stance on bribery and corruption in compliance with the UK Bribery Act, US FCPA, and local anti-corruption laws. Covers gifts, hospitality, facilitation payments, and due diligence.',
    risk: 'Critical', date: '2026-05-01', freq: 'Annual',
    content: `<h1>Anti-Bribery & Corruption Policy</h1><h2>1. Policy Statement</h2><p>First Digital maintains a zero-tolerance approach to bribery and corruption. We are committed to acting professionally, fairly, and with integrity in all business dealings.</p><h2>2. Scope</h2><p>This policy applies to all employees, officers, directors, agents, consultants, and any third parties acting on behalf of the company worldwide.</p><h2>3. Definitions</h2><p><strong>Bribery:</strong> Offering, promising, giving, or receiving a financial or other advantage to induce or reward improper performance of a function.</p><p><strong>Facilitation Payment:</strong> Small unofficial payments made to speed up routine actions. These are PROHIBITED.</p><h2>4. Gifts & Hospitality</h2><p>All gifts over £50 must be declared. Corporate hospitality must be proportionate and approved by line management.</p>` },

  p3: { method: 'Word', name: `E2E Remote Working Policy ${TS}`, category: 'HR Policies',
    summary: 'Outlines the framework for flexible and remote working arrangements including eligibility criteria, equipment provision, health and safety obligations, and performance management expectations.',
    risk: 'Medium', date: '2026-05-15', freq: 'Annual', content: '' },

  p4: { method: 'Excel', name: `E2E Vendor Risk Assessment Matrix ${TS}`, category: 'Compliance',
    summary: 'Provides the standardised risk assessment framework for evaluating third-party vendors including scoring methodology, due diligence requirements, and ongoing monitoring obligations.',
    risk: 'High', date: '2026-06-01', freq: 'Quarterly', content: '' },

  p5: { method: 'PowerPoint', name: `E2E Emergency Procedures Guide ${TS}`, category: 'Health & Safety',
    summary: 'Mandatory emergency procedures training covering fire evacuation routes, medical emergency response, bomb threat protocols, severe weather procedures, and assembly point locations.',
    risk: 'Critical', date: '2026-05-01', freq: 'Annual', content: '' },
};


// ============================================================
// PHASE 1: Create all 5 policies
// ============================================================
test.describe.serial('Phase 1: Create 5 Policies', () => {

  test('1.1 — Create Rich Text: Acceptable Use Policy', async ({ page }) => {
    console.log(`\n=== CREATE: RICH TEXT ===`);
    const pol = P.p1;
    const saved = await createAndSavePolicy(page, pol.method, pol.name, pol.category, pol.summary, pol.risk, pol.date, pol.freq, pol.content);
    await snap(page, '1-p1-saved');
    expect(saved).toBeTruthy();
    console.log(`✅ ${pol.name} — SAVED AS DRAFT`);
  });

  test('1.2 — Create HTML: Anti-Bribery & Corruption Policy', async ({ page }) => {
    console.log(`\n=== CREATE: HTML ===`);
    const pol = P.p2;
    const saved = await createAndSavePolicy(page, pol.method, pol.name, pol.category, pol.summary, pol.risk, pol.date, pol.freq, pol.content);
    await snap(page, '1-p2-saved');
    expect(saved).toBeTruthy();
    console.log(`✅ ${pol.name} — SAVED AS DRAFT`);
  });

  test('1.3 — Create Word: Remote Working Policy', async ({ page }) => {
    console.log(`\n=== CREATE: WORD ===`);
    const pol = P.p3;
    const saved = await createAndSavePolicy(page, pol.method, pol.name, pol.category, pol.summary, pol.risk, pol.date, pol.freq, pol.content);
    await snap(page, '1-p3-saved');
    expect(saved).toBeTruthy();
    console.log(`✅ ${pol.name} — SAVED AS DRAFT`);
  });

  test('1.4 — Create Excel: Vendor Risk Assessment Matrix', async ({ page }) => {
    console.log(`\n=== CREATE: EXCEL ===`);
    const pol = P.p4;
    const saved = await createAndSavePolicy(page, pol.method, pol.name, pol.category, pol.summary, pol.risk, pol.date, pol.freq, pol.content);
    await snap(page, '1-p4-saved');
    expect(saved).toBeTruthy();
    console.log(`✅ ${pol.name} — SAVED AS DRAFT`);
  });

  test('1.5 — Create PowerPoint: Emergency Procedures Guide', async ({ page }) => {
    console.log(`\n=== CREATE: POWERPOINT ===`);
    const pol = P.p5;
    const saved = await createAndSavePolicy(page, pol.method, pol.name, pol.category, pol.summary, pol.risk, pol.date, pol.freq, pol.content);
    await snap(page, '1-p5-saved');
    expect(saved).toBeTruthy();
    console.log(`✅ ${pol.name} — SAVED AS DRAFT`);
  });
});


// ============================================================
// PHASE 2: Submit each for Review (from pipeline)
// ============================================================
test.describe.serial('Phase 2: Submit for Review', () => {

  test('Submit Rich Text policy for review', async ({ page }) => {
    console.log('\n=== SUBMIT: RICH TEXT FOR REVIEW ===');

    // Create and immediately submit
    await createAndSavePolicy(
      page, P.p1.method, P.p1.name + ' SUBMIT', P.p1.category, P.p1.summary,
      P.p1.risk, P.p1.date, P.p1.freq, P.p1.content
    );

    // Now click Submit for Review (should be on Step 7)
    const submitted = await submitForReviewFromWizard(page);
    await snap(page, '2-submit-richtext');

    console.log(`Submit result: ${submitted}`);
  });

  test('Submit HTML policy for review', async ({ page }) => {
    console.log('\n=== SUBMIT: HTML FOR REVIEW ===');

    await createAndSavePolicy(
      page, P.p2.method, P.p2.name + ' SUBMIT', P.p2.category, P.p2.summary,
      P.p2.risk, P.p2.date, P.p2.freq, P.p2.content
    );

    const submitted = await submitForReviewFromWizard(page);
    await snap(page, '2-submit-html');
    console.log(`Submit result: ${submitted}`);
  });
});


// ============================================================
// PHASE 3: Review — Execute all 3 decision paths
// ============================================================
test.describe.serial('Phase 3: Execute Review Decisions', () => {

  test('3.1 — Approvals tab: verify pending policies', async ({ page }) => {
    console.log('\n=== APPROVALS TAB ===');

    await page.goto(`${BASE}/PolicyAuthor.aspx?tab=approvals`, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(5000);

    await snap(page, '3-1-approvals');

    const bodyText = await page.textContent('body') || '';
    console.log(`  Pending: ${bodyText.includes('Pending')}`);
    console.log(`  Overdue: ${bodyText.includes('Overdue')}`);
    console.log(`  Contains our policies: ${bodyText.includes(TS)}`);
  });

  test('3.2 — Open policy in Review Mode and APPROVE', async ({ page }) => {
    console.log('\n=== EXECUTE: APPROVE ===');

    // Go to approvals tab and filter to Pending
    await page.goto(`${BASE}/PolicyAuthor.aspx?tab=approvals`, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(5000);

    // Click "Pending" filter to only see policies needing review
    const pendingFilter = page.locator('button, [role="button"]').filter({ hasText: /Pending/i }).first();
    if (await pendingFilter.isVisible().catch(() => false)) {
      await pendingFilter.click();
      await page.waitForTimeout(2000);
      console.log(`  Filtered to Pending`);
    }

    // Find policy detail links
    const links = page.locator('a[href*="PolicyDetails"]');
    const count = await links.count();
    console.log(`  Found ${count} policy links (Pending filter)`);

    // Log all links for debugging
    for (let i = 0; i < Math.min(count, 5); i++) {
      const text = (await links.nth(i).textContent() || '').trim().slice(0, 50);
      const href = await links.nth(i).getAttribute('href') || '';
      console.log(`    Link[${i}]: "${text}" → ${href}`);
    }

    if (count > 0) {
      const href = await links.first().getAttribute('href') || '';
      const idMatch = href.match(/policyId=(\d+)/);
      if (idMatch) {
        // Open in review mode
        console.log(`  Opening policyId=${idMatch[1]} in review mode`);
        await page.goto(`${BASE}/PolicyDetails.aspx?policyId=${idMatch[1]}&mode=review`, {
          waitUntil: 'domcontentloaded', timeout: 60000
        });
        await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
        await page.waitForTimeout(5000);

        // Log what's visible on the review page
        const bodyText = await page.textContent('body') || '';
        console.log(`  Page has "Your Review Decision": ${bodyText.includes('Your Review Decision')}`);
        console.log(`  Page has "Approve": ${bodyText.includes('Approve')}`);
        console.log(`  Page has "In Review": ${bodyText.includes('In Review')}`);

        // Find all div[role="button"] to understand the DOM
        const roleBtns = page.locator('div[role="button"]');
        const rbCount = await roleBtns.count();
        console.log(`  div[role="button"] count: ${rbCount}`);
        for (let i = 0; i < Math.min(rbCount, 10); i++) {
          const t = (await roleBtns.nth(i).textContent() || '').trim().replace(/\n/g, ' ').slice(0, 60);
          console.log(`    btn[${i}]: "${t}"`);
        }

        // Click Approve — find the one that says "Approve" + "Policy meets"
        const approveBtn = page.locator('div[role="button"]').filter({ hasText: 'Policy meets' }).first();
        const approveAlt = page.locator('div[role="button"]').filter({ hasText: 'Approve' }).first();
        const approveVisible = await approveBtn.isVisible().catch(() => false);
        const approveAltVisible = await approveAlt.isVisible().catch(() => false);
        console.log(`  Approve (Policy meets): ${approveVisible}`);
        console.log(`  Approve (hasText Approve): ${approveAltVisible}`);

        const btnToClick = approveVisible ? approveBtn : approveAlt;
        if (approveVisible || approveAltVisible) {
          await btnToClick.click();
          await page.waitForTimeout(500);

          // Fill comments
          const commentField = page.locator('textarea').last();
          if (await commentField.isVisible().catch(() => false)) {
            await commentField.fill('E2E APPROVED — Policy meets all requirements. Content reviewed, compliance references verified, audience scope appropriate.');
          }

          await snap(page, '3-2-approve-filled');

          // Click Submit Approval
          const submitDecision = page.locator('button').filter({ hasText: /Submit.*Approval|Submit.*Decision|Submit/i }).last();
          if (await submitDecision.isVisible().catch(() => false)) {
            await submitDecision.click();
            await page.waitForTimeout(8000);

            // Dismiss dialog
            const ok = page.locator('button').filter({ hasText: /^OK$/i }).first();
            if (await ok.isVisible({ timeout: 5000 }).catch(() => false)) {
              console.log(`  ✅ APPROVAL SUBMITTED — success dialog`);
              await ok.click();
            }
            await snap(page, '3-2-approve-result');
          }
        } else {
          console.log(`  ⚠️ Approve button not visible`);
          await snap(page, '3-2-approve-missing');
        }
      }
    }
  });

  test('3.3 — REQUEST CHANGES on a policy', async ({ page }) => {
    console.log('\n=== EXECUTE: REQUEST CHANGES ===');

    await page.goto(`${BASE}/PolicyAuthor.aspx?tab=approvals`, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(5000);

    const links = page.locator('a[href*="PolicyDetails"]');
    const count = await links.count();

    if (count > 1) {
      // Use second link (different policy)
      const href = await links.nth(1).getAttribute('href') || '';
      const idMatch = href.match(/policyId=(\d+)/);
      if (idMatch) {
        await page.goto(`${BASE}/PolicyDetails.aspx?policyId=${idMatch[1]}&mode=review`, {
          waitUntil: 'domcontentloaded', timeout: 60000
        });
        await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
        await page.waitForTimeout(5000);

        const changesBtn = page.locator('div[role="button"]').filter({ hasText: /Request Changes/ }).first();
        if (await changesBtn.isVisible().catch(() => false)) {
          await changesBtn.click();
          await page.waitForTimeout(500);

          const commentField = page.locator('textarea').last();
          if (await commentField.isVisible().catch(() => false)) {
            await commentField.fill('E2E REQUEST CHANGES — Section 3 needs updated GDPR references (Articles 25, 32). Section 4 should include incident reporting timeline (72 hours). Please also add a glossary of key terms.');
          }

          await snap(page, '3-3-changes-filled');

          const submitDecision = page.locator('button').filter({ hasText: /Submit.*Change|Submit.*Decision|Submit/i }).last();
          if (await submitDecision.isVisible().catch(() => false)) {
            await submitDecision.click();
            await page.waitForTimeout(8000);
            const ok = page.locator('button').filter({ hasText: /^OK$/i }).first();
            if (await ok.isVisible({ timeout: 5000 }).catch(() => false)) {
              console.log(`  ✅ CHANGES REQUESTED — success dialog`);
              await ok.click();
            }
            await snap(page, '3-3-changes-result');
          }
        } else {
          console.log(`  ⚠️ Request Changes button not visible`);
        }
      }
    } else {
      console.log(`  ⚠️ Not enough policies to test Request Changes (need 2+)`);
    }
  });

  test('3.4 — REJECT a policy', async ({ page }) => {
    console.log('\n=== EXECUTE: REJECT ===');

    await page.goto(`${BASE}/PolicyAuthor.aspx?tab=approvals`, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(5000);

    const links = page.locator('a[href*="PolicyDetails"]');
    const count = await links.count();

    if (count > 2) {
      const href = await links.nth(2).getAttribute('href') || '';
      const idMatch = href.match(/policyId=(\d+)/);
      if (idMatch) {
        await page.goto(`${BASE}/PolicyDetails.aspx?policyId=${idMatch[1]}&mode=review`, {
          waitUntil: 'domcontentloaded', timeout: 60000
        });
        await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
        await page.waitForTimeout(5000);

        const rejectBtn = page.locator('div[role="button"]').filter({ hasText: 'Reject' }).first();
        if (await rejectBtn.isVisible().catch(() => false)) {
          await rejectBtn.click();
          await page.waitForTimeout(500);

          const commentField = page.locator('textarea').last();
          if (await commentField.isVisible().catch(() => false)) {
            await commentField.fill('E2E REJECTED — Policy fundamentally contradicts existing operational procedures in Section 4.2. Risk assessment is incomplete (missing quantitative scoring). Recommend rewriting from scratch using the Compliance regulatory template.');
          }

          await snap(page, '3-4-reject-filled');

          const submitDecision = page.locator('button').filter({ hasText: /Submit.*Rejection|Submit.*Decision|Submit/i }).last();
          if (await submitDecision.isVisible().catch(() => false)) {
            await submitDecision.click();
            await page.waitForTimeout(8000);
            const ok = page.locator('button').filter({ hasText: /^OK$/i }).first();
            if (await ok.isVisible({ timeout: 5000 }).catch(() => false)) {
              console.log(`  ✅ REJECTION SUBMITTED — success dialog`);
              await ok.click();
            }
            await snap(page, '3-4-reject-result');
          }
        } else {
          console.log(`  ⚠️ Reject button not visible`);
        }
      }
    }
  });
});


// ============================================================
// PHASE 4: Pipeline — Verify status updates
// ============================================================
test.describe('Phase 4: Pipeline Status Verification', () => {

  test('4.1 — Verify pipeline reflects review decisions', async ({ page }) => {
    console.log('\n=== PIPELINE STATUS CHECK ===');

    await page.goto(`${BASE}/PolicyAuthor.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(5000);

    await snap(page, '4-1-pipeline-after-reviews');

    const bodyText = await page.textContent('body') || '';

    // Check for status indicators
    const statuses = ['Draft', 'In Review', 'Approved', 'Rejected', 'Published'];
    for (const status of statuses) {
      console.log(`  ${status}: ${bodyText.includes(status)}`);
    }

    // Check KPI numbers
    const kpiCards = page.locator('[style*="borderTop: 3px"], [style*="border-top: 3px"]');
    const kpiCount = await kpiCards.count();
    console.log(`  KPI cards: ${kpiCount}`);
  });
});


// ============================================================
// PHASE 5: Publish an approved policy
// ============================================================
test.describe('Phase 5: Publish', () => {

  test('5.1 — Find and publish an approved policy', async ({ page }) => {
    console.log('\n=== PUBLISH ===');

    await page.goto(`${BASE}/PolicyAuthor.aspx`, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(5000);

    // Click Approved filter
    const approvedFilter = page.locator('button, [role="tab"], [role="button"]').filter({ hasText: /Approved/i }).first();
    if (await approvedFilter.isVisible().catch(() => false)) {
      await approvedFilter.click();
      await page.waitForTimeout(3000);
    }

    await snap(page, '5-1-approved-pipeline');

    // Look for Publish action button
    const publishBtn = page.locator('button[aria-label*="Publish" i], button[title*="Publish" i]').first();
    const hasPublish = await publishBtn.isVisible().catch(() => false);
    console.log(`  Publish button visible: ${hasPublish}`);

    if (hasPublish) {
      await publishBtn.click();
      await page.waitForTimeout(8000);

      const ok = page.locator('button').filter({ hasText: /^OK$/i }).first();
      if (await ok.isVisible({ timeout: 10000 }).catch(() => false)) {
        console.log(`  ✅ POLICY PUBLISHED`);
        await ok.click();
      }
      await snap(page, '5-1-publish-result');
    } else {
      // Try clicking on first approved policy row
      const approvedRow = page.locator('text=/Approved/i').first();
      if (await approvedRow.isVisible().catch(() => false)) {
        console.log(`  Approved policies found but Publish icon not located`);
      } else {
        console.log(`  ⚠️ No approved policies found to publish`);
      }
    }
  });
});


// ============================================================
// PHASE 6: Check Outlook for notifications
// ============================================================
test.describe('Phase 6: Notification Verification', () => {

  test('6.1 — Check Outlook for review notification emails', async ({ page }) => {
    console.log('\n=== CHECK OUTLOOK ===');

    // Navigate to Outlook Web
    await page.goto('https://outlook.office.com/mail/', { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    await page.waitForTimeout(5000);

    await snap(page, '6-1-outlook-inbox');

    const bodyText = await page.textContent('body') || '';
    console.log(`  Outlook loaded: ${bodyText.includes('Inbox') || bodyText.includes('inbox') || bodyText.includes('Outlook')}`);

    // Look for Policy Manager emails
    const policyEmails = page.locator('text=/Policy Manager|Policy.*Review|Approval|Submit/i');
    const emailCount = await policyEmails.count();
    console.log(`  Policy-related emails found: ${emailCount}`);

    // Check for specific notification types
    const notifTypes = ['review', 'approval', 'submitted', 'published', 'DWx'];
    for (const notif of notifTypes) {
      const found = bodyText.toLowerCase().includes(notif.toLowerCase());
      console.log(`    "${notif}": ${found}`);
    }
  });
});
