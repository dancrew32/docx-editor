/**
 * Legacy VML text box parsing.
 *
 * Modern Word emits text boxes as WPS drawings (`wps:wsp`), but older Word
 * documents and some templates still carry them as `w:pict` / `v:shape`
 * elements with a `v:textbox` child. We map the supported VML subset into the
 * same TextBox model used by WPS so the rest of the pipeline stays shared.
 */

import type {
  ImagePosition,
  ImageSize,
  ImageWrap,
  ShapeFill,
  ShapeOutline,
  TextBox,
} from '../types/document';
import { pixelsToEmu } from '../utils/units';
import { findAllDeep, findDeep, getAttribute, getLocalName, type XmlElement } from './xmlParser';

const DEFAULT_TEXTBOX_SIZE: ImageSize = {
  width: pixelsToEmu(200),
  height: pixelsToEmu(80),
};

type StyleMap = Record<string, string>;

export function parseVMLTextBoxes(root: XmlElement): Array<{
  textBox: TextBox;
  txbxContent: XmlElement;
}> {
  const shapes = collectVMLShapes(root);
  const textBoxes: Array<{ textBox: TextBox; txbxContent: XmlElement }> = [];

  for (const shapeEl of shapes) {
    const textboxEl = findDeep(shapeEl, 'v', 'textbox');
    if (!textboxEl) continue;

    const txbxContent = findDeep(textboxEl, 'w', 'txbxContent');
    if (!txbxContent) continue;

    textBoxes.push({
      textBox: parseVMLTextBoxShape(shapeEl, textboxEl),
      txbxContent,
    });
  }

  return textBoxes;
}

function collectVMLShapes(root: XmlElement): XmlElement[] {
  const localName = getLocalName(root.name ?? '');
  const seen = new Set<XmlElement>();
  const shapes: XmlElement[] = [];

  function addShape(shape: XmlElement): void {
    if (seen.has(shape)) return;
    seen.add(shape);
    shapes.push(shape);
  }

  if (localName === 'shape' || localName === 'rect' || localName === 'roundrect') {
    addShape(root);
  }
  for (const shape of findAllDeep(root, 'v', 'shape')) addShape(shape);
  for (const shape of findAllDeep(root, 'v', 'rect')) addShape(shape);
  for (const shape of findAllDeep(root, 'v', 'roundrect')) addShape(shape);

  return shapes;
}

function parseVMLTextBoxShape(shapeEl: XmlElement, textboxEl: XmlElement): TextBox {
  const style = parseStyle(getAttribute(shapeEl, null, 'style') ?? '');
  const size = parseVMLSize(style);
  const fill = parseVMLFill(shapeEl);
  const outline = parseVMLOutline(shapeEl);
  const margins = parseVMLTextBoxInsets(textboxEl);
  const position = parseVMLPosition(style);
  const wrap = parseVMLWrap(shapeEl, style, !!position);
  const id = getAttribute(shapeEl, null, 'id') ?? undefined;

  const textBox: TextBox = {
    type: 'textBox',
    id,
    size,
    content: [],
  };

  if (position) textBox.position = position;
  if (wrap) textBox.wrap = wrap;
  if (fill) textBox.fill = fill;
  if (outline) textBox.outline = outline;
  if (margins) textBox.margins = margins;

  return textBox;
}

function parseStyle(style: string): StyleMap {
  const result: StyleMap = {};
  for (const declaration of style.split(';')) {
    const [rawKey, ...rawValue] = declaration.split(':');
    if (!rawKey || rawValue.length === 0) continue;
    const key = rawKey.trim().toLowerCase();
    const value = rawValue.join(':').trim();
    if (key && value) result[key] = value;
  }
  return result;
}

function parseVMLSize(style: StyleMap): ImageSize {
  return {
    width: parseCssLengthToEmu(style.width) ?? DEFAULT_TEXTBOX_SIZE.width,
    height: parseCssLengthToEmu(style.height) ?? DEFAULT_TEXTBOX_SIZE.height,
  };
}

function parseCssLengthToEmu(value: string | undefined): number | undefined {
  if (!value) return undefined;
  const match = value.trim().match(/^(-?\d+(?:\.\d+)?)([a-z%]*)$/i);
  if (!match) return undefined;

  const numeric = Number(match[1]);
  if (!Number.isFinite(numeric)) return undefined;

  const unit = match[2].toLowerCase();
  switch (unit) {
    case 'pt':
      return Math.round(numeric * 12700);
    case 'in':
      return Math.round(numeric * 914400);
    case 'cm':
      return Math.round((numeric / 2.54) * 914400);
    case 'mm':
      return Math.round((numeric / 25.4) * 914400);
    case 'pc':
      return Math.round(numeric * 12 * 12700);
    case 'px':
    case '':
      return pixelsToEmu(numeric);
    default:
      return undefined;
  }
}

function parseVMLFill(shapeEl: XmlElement): ShapeFill | undefined {
  const filled = getAttribute(shapeEl, null, 'filled');
  if (filled === 'f' || filled === 'false' || filled === '0') {
    return { type: 'none' };
  }

  const fillColor = normalizeVMLColor(getAttribute(shapeEl, null, 'fillcolor'));
  if (!fillColor) return undefined;

  return {
    type: 'solid',
    color: { rgb: fillColor },
  };
}

function parseVMLOutline(shapeEl: XmlElement): ShapeOutline | undefined {
  const stroked = getAttribute(shapeEl, null, 'stroked');
  if (stroked === 'f' || stroked === 'false' || stroked === '0') {
    return undefined;
  }

  const outline: ShapeOutline = {
    style: 'solid',
  };

  const strokeColor = normalizeVMLColor(getAttribute(shapeEl, null, 'strokecolor'));
  if (strokeColor) outline.color = { rgb: strokeColor };

  const strokeWeight = parseCssLengthToEmu(
    getAttribute(shapeEl, null, 'strokeweight') ?? undefined
  );
  if (strokeWeight !== undefined) outline.width = strokeWeight;

  return Object.keys(outline).length > 1 ? outline : undefined;
}

function normalizeVMLColor(value: string | null): string | undefined {
  if (!value) return undefined;
  const trimmed = value.trim();
  if (!trimmed || trimmed === 'none') return undefined;
  if (/^#[0-9a-f]{6}$/i.test(trimmed)) return trimmed.slice(1).toUpperCase();
  if (/^[0-9a-f]{6}$/i.test(trimmed)) return trimmed.toUpperCase();
  return undefined;
}

function parseVMLTextBoxInsets(textboxEl: XmlElement): TextBox['margins'] | undefined {
  const inset = getAttribute(textboxEl, null, 'inset');
  if (!inset) return undefined;

  const parts = inset.split(',').map((part) => parseCssLengthToEmu(part.trim()));
  if (parts.every((part) => part === undefined)) return undefined;

  return {
    left: parts[0],
    top: parts[1],
    right: parts[2],
    bottom: parts[3],
  };
}

function parseVMLPosition(style: StyleMap): ImagePosition | undefined {
  const horizontalOffset = parseCssLengthToEmu(style['margin-left'] ?? style.left);
  const verticalOffset = parseCssLengthToEmu(style['margin-top'] ?? style.top);
  const isPositioned =
    style.position === 'absolute' ||
    style.position === 'relative' ||
    horizontalOffset !== undefined ||
    verticalOffset !== undefined;

  if (!isPositioned) return undefined;

  return {
    horizontal: {
      relativeTo: mapVMLHorizontalRelative(style['mso-position-horizontal-relative']),
      posOffset: horizontalOffset ?? 0,
    },
    vertical: {
      relativeTo: mapVMLVerticalRelative(style['mso-position-vertical-relative']),
      posOffset: verticalOffset ?? 0,
    },
  };
}

function mapVMLHorizontalRelative(
  value: string | undefined
): ImagePosition['horizontal']['relativeTo'] {
  switch (value) {
    case 'page':
      return 'page';
    case 'margin':
      return 'margin';
    case 'char':
      return 'character';
    case 'text':
    default:
      return 'column';
  }
}

function mapVMLVerticalRelative(
  value: string | undefined
): ImagePosition['vertical']['relativeTo'] {
  switch (value) {
    case 'page':
      return 'page';
    case 'margin':
      return 'margin';
    case 'line':
      return 'line';
    case 'text':
    default:
      return 'paragraph';
  }
}

function parseVMLWrap(
  shapeEl: XmlElement,
  style: StyleMap,
  hasPosition: boolean
): ImageWrap | undefined {
  const wrapEl = findDeep(shapeEl, 'w10', 'wrap');
  const rawType = getAttribute(wrapEl, null, 'type') ?? style['mso-wrap-style'];
  const type = mapVMLWrapType(rawType, hasPosition);
  if (!type) return undefined;

  const wrap: ImageWrap = { type };
  const side = getAttribute(wrapEl, null, 'side');
  const wrapText = mapVMLWrapSide(side);
  if (wrapText) wrap.wrapText = wrapText;

  return wrap;
}

function mapVMLWrapType(
  value: string | null | undefined,
  hasPosition: boolean
): ImageWrap['type'] | undefined {
  switch (value) {
    case 'none':
      return 'inFront';
    case 'topAndBottom':
      return 'topAndBottom';
    case 'tight':
      return 'tight';
    case 'through':
      return 'through';
    case 'square':
      return 'square';
    default:
      return hasPosition ? 'square' : undefined;
  }
}

function mapVMLWrapSide(value: string | null): ImageWrap['wrapText'] | undefined {
  switch (value) {
    case 'left':
      return 'left';
    case 'right':
      return 'right';
    case 'largest':
      return 'largest';
    case 'both':
    case 'bothSides':
      return 'bothSides';
    default:
      return undefined;
  }
}
