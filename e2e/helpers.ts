import { Page, expect } from '@playwright/test';
import * as path from 'path';

// ============================================================
// Policy Manager E2E Test Helpers
// ============================================================

const BASE_URL = 'https://mf7m.sharepoint.com/sites/PolicyManager';
const SITE_PAGES = `${BASE_URL}/SitePages`;

// Page URLs
export const PAGES = {
  HUB: `${SITE_PAGES}/PolicyHub.aspx`,
  MY_POLICIES: `${SITE_PAGES}/MyPolicies.aspx`,
  BUILDER: `${SITE_PAGES}/PolicyBuilder.aspx`,
  AUTHOR: `${SITE_PAGES}/PolicyAuthor.aspx`,
  DETAILS: `${SITE_PAGES}/PolicyDetails.aspx`,
  ADMIN: `${SITE_PAGES}/PolicyAdmin.aspx`,
  SEARCH: `${SITE_PAGES}/PolicySearch.aspx`,
  ANALYTICS: `${SITE_PAGES}/PolicyAnalytics.aspx`,
  DISTRIBUTION: `${SITE_PAGES}/PolicyDistribution.aspx`,
  HELP: `${SITE_PAGES}/PolicyHelp.aspx`,
};

// Screenshot counter to keep budget under control
let screenshotCount = 0;
const MAX_SCREENSHOTS = 30;

/**
 * Take a screenshot and save to disk. Does NOT return image data.
 * Keeps file sizes small by using the 1280x720 viewport.
 */
export async function screenshot(page: Page, name: string): Promise<string> {
  if (screenshotCount >= MAX_SCREENSHOTS) {
    console.log(`[screenshot] SKIPPED "${name}" — budget exhausted (${MAX_SCREENSHOTS} max)`);
    return '';
  }
  const filePath = path.join(process.cwd(), `e2e-${name}.png`);
  await page.screenshot({ path: filePath, fullPage: false });
  screenshotCount++;
  console.log(`[screenshot] Saved: e2e-${name}.png (${screenshotCount}/${MAX_SCREENSHOTS})`);
  return filePath;
}

/**
 * Wait for SharePoint page to fully load (spinner gone, content visible)
 */
export async function waitForSPPageLoad(page: Page, timeoutMs = 30000): Promise<void> {
  // Wait for network to settle
  await page.waitForLoadState('networkidle', { timeout: timeoutMs }).catch(() => {});

  // Wait for SP canvas zone to appear (indicates webpart rendered)
  await page.waitForSelector(
    '[data-automation-id="CanvasZone"], .CanvasZone, [class*="policyHub"], [class*="policyAdmin"], [class*="policyAuthor"], [class*="policyDetails"], [class*="myPolicies"]',
    { timeout: timeoutMs }
  ).catch(() => {});

  // Extra settle time for React rendering
  await page.waitForTimeout(2000);
}

/**
 * Navigate to a Policy Manager page and wait for it to load
 */
export async function navigateTo(page: Page, url: string): Promise<void> {
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await waitForSPPageLoad(page);
}

/**
 * Navigate to Policy Builder (create new policy wizard)
 */
export async function goToBuilder(page: Page): Promise<void> {
  await navigateTo(page, PAGES.BUILDER);
}

/**
 * Navigate to Author Pipeline
 */
export async function goToAuthorPipeline(page: Page, tab?: string): Promise<void> {
  const url = tab ? `${PAGES.AUTHOR}?tab=${tab}` : PAGES.AUTHOR;
  await navigateTo(page, url);
}

/**
 * Navigate to Policy Details
 */
export async function goToDetails(page: Page, policyId: number, mode?: string): Promise<void> {
  let url = `${PAGES.DETAILS}?policyId=${policyId}`;
  if (mode) url += `&mode=${mode}`;
  await navigateTo(page, url);
}

/**
 * Select a wizard mode (Fast Track or Standard Wizard) from the pre-wizard screen.
 * PolicyBuilder.aspx shows a mode selection screen first: "How would you like to create this policy?"
 */
export async function selectWizardMode(page: Page, mode: 'fast-track' | 'standard'): Promise<void> {
  // Wait for the mode selection screen to appear
  await page.waitForSelector('text=/How would you like|Fast Track|Standard Wizard/i', { timeout: 30000 });

  const label = mode === 'fast-track' ? 'Fast Track' : 'Standard Wizard';
  // The mode cards are large clickable areas with the mode title
  const modeCard = page.locator(`div[role="button"], div`).filter({ hasText: label }).first();

  // Try clicking — if exact match fails, try broader selector
  const visible = await modeCard.isVisible({ timeout: 5000 }).catch(() => false);
  if (visible) {
    await modeCard.click();
    await page.waitForTimeout(2000);
  } else {
    // Fallback: click by text
    await page.getByText(label).first().click();
    await page.waitForTimeout(2000);
  }
}

/**
 * Select a creation method type in Step 0 of the Standard Wizard.
 * The type strip is a horizontal row of buttons with icon + label.
 * Uses partial text match since buttons contain icon name + label + description.
 */
export async function selectCreationMethod(page: Page, method: string): Promise<void> {
  // Wait for the wizard to fully render
  await page.waitForTimeout(1000);

  // The type strip buttons are div[role="button"] in a flex container
  // Each contains an Icon + label text. Use a looser text match.
  const methodButton = page.locator('div[role="button"]').filter({ hasText: method });

  // If there are multiple matches (e.g. "Rich Text" could match elsewhere),
  // look specifically in the horizontal type strip (flex container near top)
  const count = await methodButton.count();
  if (count > 0) {
    // Click the first match — the type strip buttons are rendered first
    await methodButton.first().click();
    await page.waitForTimeout(500);
  } else {
    // Fallback: try clicking by exact text content
    const textMatch = page.getByText(method, { exact: true });
    await textMatch.first().click({ timeout: 10000 });
    await page.waitForTimeout(500);
  }
}

/**
 * Select the "Blank" card (first card in template grid) or a named template
 */
export async function selectBlankOrTemplate(page: Page, templateName?: string): Promise<void> {
  if (templateName) {
    const templateCard = page.locator('div[role="button"]').filter({ hasText: templateName });
    await templateCard.first().click();
  } else {
    // Click the Blank card — it's always the first one with "Blank" in the text
    const blankCard = page.locator('div[role="button"]').filter({ hasText: /^Blank/ });
    await blankCard.first().click();
  }
  await page.waitForTimeout(500);
}

/**
 * Click the Next button in the wizard.
 * The button contains "Next" text + a ChevronRight icon.
 */
export async function clickNext(page: Page): Promise<void> {
  // Try multiple selectors — the button may render differently
  const nextBtn = page.locator('button').filter({ hasText: /Next/ }).last();
  const visible = await nextBtn.isVisible({ timeout: 10000 }).catch(() => false);
  if (visible) {
    await nextBtn.click();
  } else {
    // Fallback: find by text content directly
    await page.getByText('Next').last().click();
  }
  await page.waitForTimeout(2000);
}

/**
 * Click the Previous button in the wizard
 */
export async function clickPrevious(page: Page): Promise<void> {
  const prevBtn = page.locator('button').filter({ hasText: /Previous/ });
  await prevBtn.first().click();
  await page.waitForTimeout(1000);
}

/**
 * Fill in Step 1 Basic Info fields
 */
export async function fillBasicInfo(page: Page, name: string, category?: string, summary?: string): Promise<void> {
  // Fill policy name
  const nameField = page.locator('input').filter({ hasText: '' }).locator('xpath=//input[contains(@placeholder,"policy") or contains(@placeholder,"Policy") or contains(@placeholder,"name") or contains(@placeholder,"Name") or contains(@placeholder,"title") or contains(@placeholder,"Title")]').first();

  // Try multiple selectors for the name field
  const nameInput = page.locator('input[placeholder*="name" i], input[placeholder*="title" i], input[placeholder*="policy" i]').first();
  if (await nameInput.isVisible().catch(() => false)) {
    await nameInput.clear();
    await nameInput.fill(name);
  } else {
    // Fallback: find by label
    const labelledInput = page.getByLabel(/policy name|title/i).first();
    if (await labelledInput.isVisible().catch(() => false)) {
      await labelledInput.clear();
      await labelledInput.fill(name);
    }
  }

  // Fill summary if provided
  if (summary) {
    const summaryField = page.locator('textarea').first();
    if (await summaryField.isVisible().catch(() => false)) {
      await summaryField.clear();
      await summaryField.fill(summary);
    }
  }

  await page.waitForTimeout(500);
}

/**
 * Click Save as Draft button
 */
export async function clickSaveDraft(page: Page): Promise<void> {
  const saveBtn = page.locator('button').filter({ hasText: /Save.*Draft|Save$/i });
  await saveBtn.first().click();
  await page.waitForTimeout(3000);
}

/**
 * Click Submit for Review button
 */
export async function clickSubmitForReview(page: Page): Promise<void> {
  const submitBtn = page.locator('button').filter({ hasText: /Submit.*Review/i });
  await submitBtn.first().click();
  await page.waitForTimeout(3000);
}

/**
 * Get the count of visible elements matching a selector
 */
export async function countVisible(page: Page, selector: string): Promise<number> {
  const elements = page.locator(selector);
  const count = await elements.count();
  let visible = 0;
  for (let i = 0; i < count; i++) {
    if (await elements.nth(i).isVisible().catch(() => false)) {
      visible++;
    }
  }
  return visible;
}

/**
 * Check if text is visible on the page
 */
export async function hasText(page: Page, text: string | RegExp): Promise<boolean> {
  const locator = typeof text === 'string'
    ? page.locator(`text="${text}"`)
    : page.locator(`text=${text}`);
  return locator.first().isVisible({ timeout: 5000 }).catch(() => false);
}

/**
 * Wait for text to appear on the page
 */
export async function waitForText(page: Page, text: string | RegExp, timeoutMs = 15000): Promise<void> {
  const locator = typeof text === 'string'
    ? page.locator(`text="${text}"`)
    : page.locator(`text=${text}`);
  await locator.first().waitFor({ state: 'visible', timeout: timeoutMs });
}

/**
 * Log a test result line to console
 */
export function logResult(test: string, result: 'PASS' | 'FAIL' | 'SKIP' | 'INFO', detail: string): void {
  const icon = { PASS: '✅', FAIL: '❌', SKIP: '⏭️', INFO: 'ℹ️' }[result];
  console.log(`${icon} ${result} | ${test} | ${detail}`);
}
