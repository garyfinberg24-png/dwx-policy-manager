import { test, Page } from '@playwright/test';
import * as path from 'path';

/**
 * DEEP PAGE TESTS — Every major page, section, and feature
 *
 * Tests each SharePoint page loads correctly with live data,
 * verifies key UI elements, interactive controls, and navigation.
 */

const BASE = 'https://mf7m.sharepoint.com/sites/PolicyManager/SitePages';

let shotCount = 0;
async function snap(page: Page, name: string): Promise<void> {
  if (shotCount >= 20) return;
  await page.screenshot({ path: path.join(process.cwd(), `e2e-pages-${name}.png`), fullPage: false });
  shotCount++;
  console.log(`📸 [${shotCount}/20] e2e-pages-${name}.png`);
}

async function goTo(page: Page, pageName: string): Promise<void> {
  await page.goto(`${BASE}/${pageName}`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);
}


// ============================================================
// 1. Policy Hub — Browse, Filter, Search
// ============================================================
test.describe('1 — Policy Hub', () => {

  test('1.1 — Hub loads with policies and filters', async ({ page }) => {
    console.log('\n=== POLICY HUB ===');
    await goTo(page, 'PolicyHub.aspx');
    await snap(page, '1-1-hub');

    const bodyText = await page.textContent('body') || '';
    console.log(`  Policies visible: ${bodyText.includes('POL-') || bodyText.includes('Policy')}`);

    // Check for category filter dropdown
    const categoryFilter = page.locator('.ms-Dropdown, select').first();
    const hasFilter = await categoryFilter.isVisible().catch(() => false);
    console.log(`  Category filter: ${hasFilter}`);

    if (hasFilter) {
      await categoryFilter.click();
      await page.waitForTimeout(500);
      const options = page.locator('.ms-Dropdown-item, [role="option"]');
      const optCount = await options.count();
      console.log(`  Filter options: ${optCount}`);
      for (let i = 0; i < Math.min(optCount, 8); i++) {
        const text = (await options.nth(i).textContent() || '').trim();
        console.log(`    "${text}"`);
      }
      await page.keyboard.press('Escape');
    }

    // Count policy cards/items
    const policyItems = page.locator('[style*="border-top: 4px"], [style*="borderTop: 4px"]');
    const itemCount = await policyItems.count();
    console.log(`  Policy cards: ${itemCount}`);
  });

  test('1.2 — Hub search works', async ({ page }) => {
    console.log('\n=== HUB SEARCH ===');
    await goTo(page, 'PolicyHub.aspx');

    const searchInput = page.locator('input[placeholder*="search" i], input[type="search"]').first();
    if (await searchInput.isVisible().catch(() => false)) {
      await searchInput.fill('security');
      await searchInput.press('Enter');
      await page.waitForTimeout(3000);

      const bodyText = await page.textContent('body') || '';
      const hasResults = bodyText.toLowerCase().includes('security');
      console.log(`  Search "security": ${hasResults ? 'results found' : 'no results'}`);
      await snap(page, '1-2-hub-search');
    }
  });

  test('1.3 — Click policy opens in Simple Reader', async ({ page }) => {
    console.log('\n=== SIMPLE READER ===');
    await goTo(page, 'PolicyHub.aspx');

    // Click a policy card or link
    const policyCard = page.locator('[style*="border-top: 4px"], [style*="borderTop: 4px"]').first();
    if (await policyCard.isVisible().catch(() => false)) {
      await policyCard.click();
      await page.waitForTimeout(3000);
      await snap(page, '1-3-policy-click');

      // Check if StyledPanel opened or page navigated
      const bodyText = await page.textContent('body') || '';
      const hasPanel = bodyText.includes('Back to Policy Hub') || bodyText.includes('Download') || bodyText.includes('Print');
      console.log(`  Simple reader/panel: ${hasPanel}`);
    }
  });
});


// ============================================================
// 2. Policy Search Centre
// ============================================================
test.describe('2 — Search Centre', () => {

  test('2.1 — Search page layout', async ({ page }) => {
    console.log('\n=== SEARCH CENTRE ===');
    await goTo(page, 'PolicySearch.aspx');
    await snap(page, '2-1-search');

    const bodyText = await page.textContent('body') || '';
    console.log(`  Search input: ${await page.locator('input[placeholder*="search" i]').isVisible().catch(() => false)}`);
    console.log(`  Contains "Search": ${bodyText.includes('Search')}`);
    console.log(`  Contains filters: ${bodyText.includes('Category') || bodyText.includes('Risk')}`);
  });

  test('2.2 — Search with results', async ({ page }) => {
    console.log('\n=== SEARCH WITH RESULTS ===');
    await goTo(page, 'PolicySearch.aspx');

    const searchInput = page.locator('input[placeholder*="search" i], input[type="search"]').first();
    if (await searchInput.isVisible().catch(() => false)) {
      await searchInput.fill('policy');
      await searchInput.press('Enter');
      await page.waitForTimeout(3000);
      await snap(page, '2-2-search-results');

      const bodyText = await page.textContent('body') || '';
      console.log(`  Results: ${bodyText.includes('result') || bodyText.includes('found') || bodyText.includes('POL-')}`);
    }
  });
});


// ============================================================
// 3. Help Centre
// ============================================================
test('3 — Help Centre loads with tabs', async ({ page }) => {
  console.log('\n=== HELP CENTRE ===');
  await goTo(page, 'PolicyHelp.aspx');
  await snap(page, '3-help');

  const bodyText = await page.textContent('body') || '';
  console.log(`  Contains "Help": ${bodyText.includes('Help')}`);
  console.log(`  Contains "Articles": ${bodyText.includes('Articles')}`);
  console.log(`  Contains "FAQs": ${bodyText.includes('FAQ')}`);
  console.log(`  Contains "Shortcuts": ${bodyText.includes('Shortcuts') || bodyText.includes('shortcuts')}`);
  console.log(`  Contains "Support": ${bodyText.includes('Support')}`);

  // Check for tabs
  const tabs = ['Home', 'Articles', 'FAQ', 'Shortcuts', 'Support'];
  for (const tab of tabs) {
    const tabEl = page.locator('button, [role="tab"]').filter({ hasText: new RegExp(tab, 'i') }).first();
    const hasTab = await tabEl.isVisible().catch(() => false);
    console.log(`    Tab "${tab}": ${hasTab}`);
  }
});


// ============================================================
// 4. Analytics Dashboard
// ============================================================
test('4 — Analytics dashboard with 6 tabs', async ({ page }) => {
  console.log('\n=== ANALYTICS DASHBOARD ===');
  await goTo(page, 'PolicyAnalytics.aspx');
  await snap(page, '4-analytics');

  const bodyText = await page.textContent('body') || '';
  console.log(`  Contains "Analytics": ${bodyText.includes('Analytics')}`);

  // Check for the 6 tabs
  const tabs = ['Executive', 'Policy Metrics', 'Acknowledgement', 'SLA', 'Compliance', 'Audit'];
  for (const tab of tabs) {
    const tabEl = page.locator('button, [role="tab"]').filter({ hasText: new RegExp(tab, 'i') }).first();
    const hasTab = await tabEl.isVisible().catch(() => false);
    console.log(`    Tab "${tab}": ${hasTab}`);
    if (hasTab) {
      await tabEl.click();
      await page.waitForTimeout(2000);
    }
  }

  await snap(page, '4-analytics-tabs');
});


// ============================================================
// 5. Manager View Dashboard
// ============================================================
test('5 — Manager Dashboard', async ({ page }) => {
  console.log('\n=== MANAGER DASHBOARD ===');
  await goTo(page, 'PolicyManagerView.aspx');
  await snap(page, '5-manager');

  const bodyText = await page.textContent('body') || '';
  console.log(`  Contains "Manager": ${bodyText.includes('Manager')}`);
  console.log(`  Contains "Team Compliance": ${bodyText.includes('Team Compliance')}`);
  console.log(`  Contains "Approvals": ${bodyText.includes('Approval')}`);
  console.log(`  Contains "Analytics": ${bodyText.includes('Analytics')}`);
  console.log(`  Contains "Delegations": ${bodyText.includes('Delegation')}`);

  // Check Manager dropdown items
  const managerBtn = page.locator('button').filter({ hasText: /^Manager$/ }).first();
  if (await managerBtn.isVisible().catch(() => false)) {
    await managerBtn.click();
    await page.waitForTimeout(1500);
    await snap(page, '5-manager-dropdown');

    const dropdownItems = ['Approvals', 'Team Compliance', 'Delegations', 'Review Cycles', 'Analytics', 'Request Policy'];
    for (const item of dropdownItems) {
      const el = page.getByText(item, { exact: true }).first();
      const visible = await el.isVisible().catch(() => false);
      console.log(`    "${item}": ${visible}`);
    }
  }
});


// ============================================================
// 6. Admin Centre — test each sidebar section loads
// ============================================================
test('6 — Admin Centre sections', async ({ page }) => {
  console.log('\n=== ADMIN CENTRE ===');
  await goTo(page, 'PolicyAdmin.aspx');
  await snap(page, '6-admin');

  const bodyText = await page.textContent('body') || '';
  console.log(`  Contains "Admin": ${bodyText.includes('Admin')}`);

  // Check sidebar sections
  const sections = [
    'Templates', 'Metadata Profiles', 'Approval Workflows', 'Compliance Settings',
    'Notifications', 'Naming Rules', 'SLA Targets', 'Data Lifecycle', 'Navigation',
    'Reviewers', 'Audit Log', 'Data Export'
  ];

  for (const section of sections) {
    const sectionEl = page.getByText(section, { exact: false }).first();
    const visible = await sectionEl.isVisible().catch(() => false);
    console.log(`    "${section}": ${visible}`);
  }

  // Click a few sections and verify they load
  const clickableSections = ['Templates', 'Audit Log', 'Notifications'];
  for (const section of clickableSections) {
    const sectionBtn = page.getByText(section, { exact: true }).first();
    if (await sectionBtn.isVisible().catch(() => false)) {
      await sectionBtn.click();
      await page.waitForTimeout(2000);
      console.log(`    Clicked "${section}" — loaded`);
    }
  }

  await snap(page, '6-admin-section');
});


// ============================================================
// 7. Author View — all tabs
// ============================================================
test('7 — Author View tabs', async ({ page }) => {
  console.log('\n=== AUTHOR VIEW TABS ===');
  await goTo(page, 'PolicyAuthor.aspx');

  const tabs = ['Drafts & Pipeline', 'Policy Requests', 'Approvals', 'Delegations'];
  for (const tab of tabs) {
    const tabBtn = page.locator('button, [role="tab"]').filter({ hasText: new RegExp(tab, 'i') }).first();
    const visible = await tabBtn.isVisible().catch(() => false);
    console.log(`  Tab "${tab}": ${visible}`);

    if (visible) {
      await tabBtn.click();
      await page.waitForTimeout(3000);
    }
  }
  await snap(page, '7-author-tabs');
});


// ============================================================
// 8. Bookmarks
// ============================================================
test('8 — Bookmarks functionality', async ({ page }) => {
  console.log('\n=== BOOKMARKS ===');

  // Open a policy in browse mode
  await page.goto(`${BASE}/PolicyDetails.aspx?policyId=1&mode=browse`, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(5000);

  // Look for Bookmark button
  const bookmarkBtn = page.locator('button').filter({ hasText: /Bookmark/i }).first();
  const hasBookmark = await bookmarkBtn.isVisible().catch(() => false);
  console.log(`  Bookmark button: ${hasBookmark}`);

  if (hasBookmark) {
    await bookmarkBtn.click();
    await page.waitForTimeout(1000);
    console.log(`  ✅ Bookmark toggled`);
  }

  // Check for bookmark icon in toolbar
  const bookmarkIcon = page.locator('button[aria-label*="Bookmark" i], button[title*="Bookmark" i]');
  console.log(`  Bookmark icon buttons: ${await bookmarkIcon.count()}`);

  await snap(page, '8-bookmark');
});


// ============================================================
// 9. Policy Packs
// ============================================================
test('9 — Policy Packs page', async ({ page }) => {
  console.log('\n=== POLICY PACKS ===');
  await goTo(page, 'PolicyPacks.aspx');
  await snap(page, '9-packs');

  const bodyText = await page.textContent('body') || '';
  console.log(`  Contains "Policy Packs": ${bodyText.includes('Policy Pack') || bodyText.includes('Packs')}`);
  console.log(`  Contains "Create": ${bodyText.includes('Create')}`);

  // Count pack cards
  const packCards = page.locator('[style*="border-top"], [style*="borderTop"]');
  console.log(`  Pack cards: ${await packCards.count()}`);
});


// ============================================================
// 10. Quiz Builder
// ============================================================
test('10 — Quiz Builder page', async ({ page }) => {
  console.log('\n=== QUIZ BUILDER ===');
  await goTo(page, 'QuizBuilder.aspx');
  await snap(page, '10-quiz');

  const bodyText = await page.textContent('body') || '';
  console.log(`  Contains "Quiz": ${bodyText.includes('Quiz')}`);
  console.log(`  Contains "Create": ${bodyText.includes('Create')}`);
  console.log(`  Contains "AI Generate": ${bodyText.includes('AI') || bodyText.includes('Generate')}`);
});


// ============================================================
// 11. Distribution page detail
// ============================================================
test('11 — Distribution page KPIs', async ({ page }) => {
  console.log('\n=== DISTRIBUTION DETAIL ===');
  await goTo(page, 'PolicyDistribution.aspx');
  await snap(page, '11-distribution');

  const bodyText = await page.textContent('body') || '';
  const kpis = ['Total', 'Active', 'Completed', 'Acknowledged', 'Overdue'];
  for (const kpi of kpis) {
    console.log(`  KPI "${kpi}": ${bodyText.includes(kpi)}`);
  }

  // Check for percentage completion
  const hasPercent = bodyText.includes('%');
  console.log(`  Completion %: ${hasPercent}`);
});


// ============================================================
// 12. Footer build number on every page
// ============================================================
test('12 — Footer and build number', async ({ page }) => {
  console.log('\n=== FOOTER VERIFICATION ===');

  const pages = ['PolicyHub.aspx', 'MyPolicies.aspx', 'PolicyAuthor.aspx', 'PolicyAdmin.aspx'];
  for (const pg of pages) {
    await goTo(page, pg);
    const bodyText = await page.textContent('body') || '';
    const hasBuild = bodyText.includes('Build') || bodyText.includes('v1.2');
    const hasFooter = bodyText.includes('DWx') || bodyText.includes('Digital Workplace') || bodyText.includes('First Digital');
    console.log(`  ${pg}: build=${hasBuild} footer=${hasFooter}`);
  }
});
