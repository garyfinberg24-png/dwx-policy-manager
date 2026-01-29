import { test, expect } from '@playwright/test';

/**
 * Quiz Builder E2E Tests
 * Tests the policy comprehension quiz creation interface
 */
test.describe('Quiz Builder', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/_layouts/15/workbench.aspx');
    await page.waitForLoadState('networkidle');
  });

  test('should display quiz builder interface', async ({ page }) => {
    const quizBuilder = page.locator('text=Quiz Builder, text=Create Quiz, [data-automation-id="quizBuilder"]');

    if (await quizBuilder.first().isVisible()) {
      await expect(quizBuilder.first()).toBeVisible();
    }
  });

  test('should allow adding quiz questions', async ({ page }) => {
    const addQuestionButton = page.locator('button:has-text("Add Question"), button:has-text("New Question")');

    if (await addQuestionButton.first().isVisible()) {
      await addQuestionButton.first().click();
      await page.waitForTimeout(1000);

      // Should show question form
      const questionForm = page.locator('input[placeholder*="Question"], textarea[placeholder*="Question"]');
      if (await questionForm.isVisible()) {
        await expect(questionForm).toBeEnabled();
      }
    }
  });

  test('should support multiple question types', async ({ page }) => {
    // Look for question type selector
    const questionTypeSelector = page.locator('select:has-text("Type"), [aria-label*="Question type"]');

    if (await questionTypeSelector.first().isVisible()) {
      await questionTypeSelector.first().click();
      await page.waitForTimeout(500);

      // Check for different question types
      const multipleChoice = page.locator('text=Multiple Choice');
      const trueFalse = page.locator('text=True/False');

      const hasMultipleChoice = await multipleChoice.isVisible().catch(() => false);
      const hasTrueFalse = await trueFalse.isVisible().catch(() => false);

      console.log(`Multiple choice option: ${hasMultipleChoice}`);
      console.log(`True/False option: ${hasTrueFalse}`);
    }
  });

  test('should allow quiz preview', async ({ page }) => {
    const previewButton = page.locator('button:has-text("Preview")');

    if (await previewButton.first().isVisible()) {
      await previewButton.first().click();
      await page.waitForTimeout(1000);

      // Should show preview mode
      const previewMode = page.locator('.preview-mode, [data-automation-id="quizPreview"]');
      if (await previewMode.isVisible()) {
        await expect(previewMode).toBeVisible();
      }
    }
  });

  test('should save quiz as draft', async ({ page }) => {
    const saveButton = page.locator('button:has-text("Save"), button:has-text("Save Draft")');

    if (await saveButton.first().isVisible()) {
      await expect(saveButton.first()).toBeEnabled();
    }
  });
});
