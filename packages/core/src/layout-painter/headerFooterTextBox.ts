import type { TextBoxBlock, TextBoxFragment, TextBoxMeasure } from '../layout-engine/types';
import { isFloatingTextBoxBlock } from '../layout-engine/textBoxFlow';
import { emuToPixels } from '../utils/units';
import type { HeaderFooterLayoutInfo, RenderContext } from './renderPage';
import { renderTextBoxFragment } from './renderTextBox';

export function renderHeaderFooterTextBox(
  block: TextBoxBlock,
  measure: TextBoxMeasure,
  cursorY: number,
  context: RenderContext,
  layout: HeaderFooterLayoutInfo,
  doc: Document
): { element: HTMLElement; advanceHeight: number } {
  const isFloating = isFloatingTextBoxBlock(block);
  const fragment: TextBoxFragment = {
    kind: 'textBox',
    blockId: block.id,
    x: 0,
    y: cursorY,
    width: measure.width,
    height: measure.height,
    pmStart: block.pmStart,
    pmEnd: block.pmEnd,
    isFloating,
    zIndex: isFloating ? (block.wrapType === 'behind' ? -1 : 1) : undefined,
  };

  const element = renderTextBoxFragment(fragment, block, measure, context, { document: doc });
  const position = isFloating
    ? resolveHeaderFooterTextBoxPosition(block, measure, cursorY, layout)
    : { left: 0, top: cursorY };

  element.style.left = `${position.left}px`;
  element.style.top = `${position.top}px`;

  return {
    element,
    advanceHeight: isFloating ? 0 : measure.height,
  };
}

function resolveHeaderFooterTextBoxPosition(
  block: TextBoxBlock,
  measure: TextBoxMeasure,
  cursorY: number,
  layout: HeaderFooterLayoutInfo
): { left: number; top: number } {
  return {
    left: resolveHorizontalPosition(block, measure, layout),
    top: resolveVerticalPosition(block, measure, cursorY, layout),
  };
}

function resolveHorizontalPosition(
  block: TextBoxBlock,
  measure: TextBoxMeasure,
  layout: HeaderFooterLayoutInfo
): number {
  const horizontal = block.position?.horizontal;
  if (!horizontal) return 0;

  const align = horizontal.align;
  const offsetPx =
    horizontal.posOffset !== undefined ? emuToPixels(horizontal.posOffset) : undefined;

  if (horizontal.relativeTo === 'page') {
    if (offsetPx !== undefined) return offsetPx - layout.flowLeft;
    if (align === 'right') return layout.pageWidth - measure.width - layout.flowLeft;
    if (align === 'center') return (layout.pageWidth - measure.width) / 2 - layout.flowLeft;
    if (align === 'left') return -layout.flowLeft;
  }

  if (horizontal.relativeTo === 'margin') {
    const marginWidth = layout.pageWidth - layout.margins.left - layout.margins.right;
    if (offsetPx !== undefined) return layout.margins.left + offsetPx - layout.flowLeft;
    if (align === 'right')
      return layout.margins.left + marginWidth - measure.width - layout.flowLeft;
    if (align === 'center')
      return layout.margins.left + (marginWidth - measure.width) / 2 - layout.flowLeft;
    if (align === 'left') return layout.margins.left - layout.flowLeft;
  }

  if (offsetPx !== undefined) return offsetPx;
  if (align === 'right') return layout.contentWidth - measure.width;
  if (align === 'center') return (layout.contentWidth - measure.width) / 2;
  return 0;
}

function resolveVerticalPosition(
  block: TextBoxBlock,
  measure: TextBoxMeasure,
  cursorY: number,
  layout: HeaderFooterLayoutInfo
): number {
  const vertical = block.position?.vertical;
  if (!vertical) return cursorY;

  const align = vertical.align;
  const offsetPx = vertical.posOffset !== undefined ? emuToPixels(vertical.posOffset) : undefined;

  if (vertical.relativeTo === 'page') {
    if (offsetPx !== undefined) return offsetPx - layout.flowTop;
    if (align === 'bottom') return layout.pageHeight - measure.height - layout.flowTop;
    if (align === 'center') return (layout.pageHeight - measure.height) / 2 - layout.flowTop;
    if (align === 'top') return -layout.flowTop;
  }

  if (vertical.relativeTo === 'margin') {
    const marginHeight = layout.pageHeight - layout.margins.top - layout.margins.bottom;
    if (offsetPx !== undefined) return layout.margins.top + offsetPx - layout.flowTop;
    if (align === 'bottom')
      return layout.margins.top + marginHeight - measure.height - layout.flowTop;
    if (align === 'center')
      return layout.margins.top + (marginHeight - measure.height) / 2 - layout.flowTop;
    if (align === 'top') return layout.margins.top - layout.flowTop;
  }

  if (offsetPx !== undefined) return cursorY + offsetPx;
  return cursorY;
}
