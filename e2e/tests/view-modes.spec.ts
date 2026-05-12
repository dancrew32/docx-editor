/**
 * E2E coverage for document page view modes.
 */

import { test, expect, type Page } from '@playwright/test';
import { EditorPage } from '../helpers/editor-page';
import * as path from 'path';

async function loadMultiPageFixture(page: Page): Promise<EditorPage> {
  const editor = new EditorPage(page);
  await editor.goto();
  await editor.waitForReady();
  await editor.loadDocxFile(path.resolve(process.cwd(), 'e2e/fixtures/issue-68-large.docx'));
  await page.waitForFunction(() => (window.__DOCX_EDITOR_E2E__?.getTotalPages() ?? 0) > 1, {
    timeout: 10000,
  });
  return editor;
}

async function getPageRects(page: Page) {
  return page.locator('.layout-page').evaluateAll((pages) =>
    pages.map((el) => {
      const rect = el.getBoundingClientRect();
      return {
        left: rect.left,
        right: rect.right,
        top: rect.top,
        bottom: rect.bottom,
        width: rect.width,
        height: rect.height,
      };
    })
  );
}

async function selectPageViewMode(page: Page, mode: 'onePage' | 'multiplePages' | 'pageWidth') {
  await page.getByTestId('page-view-mode-control').click();
  await page.getByTestId(`page-view-mode-${mode}`).click();
}

test.describe('document page view modes', () => {
  test('defaults to one-page vertical stacking on wide viewports', async ({ page }) => {
    await page.setViewportSize({ width: 1900, height: 1000 });
    await loadMultiPageFixture(page);

    await expect(page.getByTestId('page-view-mode-control')).toBeVisible();
    const rects = await getPageRects(page);
    expect(rects.length).toBeGreaterThan(1);
    expect(rects[1].top).toBeGreaterThan(rects[0].bottom);
  });

  test('multiple-pages mode wraps pages side by side and preserves click-to-type', async ({
    page,
  }) => {
    await page.setViewportSize({ width: 1900, height: 1000 });
    await loadMultiPageFixture(page);

    await selectPageViewMode(page, 'multiplePages');
    await page.waitForFunction(() => {
      const pages = Array.from(document.querySelectorAll<HTMLElement>('.layout-page'));
      if (pages.length < 2) return false;
      const first = pages[0].getBoundingClientRect();
      const second = pages[1].getBoundingClientRect();
      return Math.abs(first.top - second.top) < 8 && second.left > first.right;
    });

    const secondPageText = page
      .locator('.layout-page')
      .nth(1)
      .locator('span[data-pm-start][data-pm-end]')
      .first();
    await secondPageText.click();
    await page.keyboard.type('ViewModeSmoke');
    await expect(page.locator('.ProseMirror')).toContainText('ViewModeSmoke');
  });

  test('page-width mode fits the page to the viewport and responds to resize', async ({ page }) => {
    await page.setViewportSize({ width: 720, height: 900 });
    await loadMultiPageFixture(page);

    await selectPageViewMode(page, 'pageWidth');
    await page.waitForFunction(() => {
      const pageEl = document.querySelector('.layout-page');
      let scroller = pageEl?.parentElement;
      while (scroller && getComputedStyle(scroller).overflowY !== 'auto') {
        scroller = scroller.parentElement;
      }
      if (!pageEl || !scroller) return false;
      return pageEl.getBoundingClientRect().width < scroller.getBoundingClientRect().width;
    });
    const narrowWidth = (await getPageRects(page))[0].width;
    expect(narrowWidth).toBeLessThan(760);

    await page.setViewportSize({ width: 1200, height: 900 });
    await page.waitForFunction((previousWidth) => {
      const pageEl = document.querySelector('.layout-page');
      if (!pageEl) return false;
      return pageEl.getBoundingClientRect().width > previousWidth + 80;
    }, narrowWidth);
  });
});
