import { test, expect, type Page } from '@playwright/test';
import { EditorPage } from '../helpers/editor-page';

const FIXTURE = 'fixtures/issue-318-textbox-support.docx';

async function loadFixture(page: Page) {
  const editor = new EditorPage(page);
  const overrideUrl = process.env.DOCX_EDITOR_E2E_URL;
  if (overrideUrl) {
    await page.goto(new URL('/?e2e=1', overrideUrl).toString());
  } else {
    await editor.goto();
  }
  await editor.waitForReady();
  await page.locator('input[type="file"][accept=".docx"]').setInputFiles(`e2e/${FIXTURE}`);
  await page.waitForSelector('.paged-editor__pages');
  await page.waitForSelector('[data-page-number]');
  await page.waitForSelector('.layout-textbox');
}

test.describe('Issue #318 text box support', () => {
  test('renders body, header, VML, and table-cell text boxes', async ({ page }) => {
    await loadFixture(page);

    await expect(
      page.locator('.layout-textbox').filter({ hasText: 'Body WPS Text Box' })
    ).toBeVisible();
    await expect(
      page.locator('.layout-textbox').filter({ hasText: 'Legacy VML Text Box' })
    ).toBeVisible();
    await expect(
      page.locator('.layout-page-header .layout-textbox').filter({ hasText: 'Header WPS Text Box' })
    ).toBeVisible();
    await expect(
      page.locator('.layout-table-cell .layout-textbox').filter({ hasText: 'Table Cell Text Box' })
    ).toBeVisible();
  });
});
