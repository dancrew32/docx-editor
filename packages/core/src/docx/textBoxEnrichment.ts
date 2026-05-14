/**
 * Shared text box enrichment for parsed paragraphs.
 *
 * `parseRun` intentionally skips non-image drawings and legacy VML objects
 * because parsing their text body needs paragraph parsing, which would create
 * a circular dependency at run level. This pass runs after a paragraph is
 * parsed and injects textbox-bearing shapes back into paragraph runs so the
 * existing ProseMirror and layout pipeline can treat them as TextBox nodes.
 */

import type {
  MediaFile,
  Paragraph,
  RelationshipMap,
  Run,
  Shape,
  ShapeContent,
  TextBox,
  Theme,
} from '../types/document';
import type { NumberingMap } from './numberingParser';
import { parseParagraph } from './paragraphParser';
import type { StyleMap } from './styleParser';
import { findDeep, getChildElements, getLocalName, type XmlElement } from './xmlParser';
import { getTextBoxContentElement, parseTextBox, parseTextBoxContent } from './textBoxParser';
import { parseVMLTextBoxes } from './vmlTextBoxParser';

export interface TextBoxEnrichmentOptions {
  inHeaderFooter?: boolean;
}

export function enrichParagraphTextBoxes(
  paragraph: Paragraph,
  paraXml: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  options: TextBoxEnrichmentOptions = {}
): void {
  const xmlRuns = getChildElements(paraXml).filter(
    (child) => getLocalName(child.name ?? '') === 'r'
  );
  if (xmlRuns.length === 0) return;

  let parsedRunIndex = 0;
  for (const xmlRun of xmlRuns) {
    const shapeContents = extractTextBoxShapeContents(
      xmlRun,
      styles,
      theme,
      numbering,
      rels,
      media,
      options
    );

    if (shapeContents.length > 0) {
      const parsedRun = getParsedRun(paragraph, parsedRunIndex);
      if (parsedRun) {
        parsedRun.content.push(...shapeContents);
      } else {
        paragraph.content.push({
          type: 'run',
          content: shapeContents,
        });
      }
    }

    parsedRunIndex++;
  }
}

function getParsedRun(paragraph: Paragraph, runIndex: number): Run | null {
  let seenRuns = 0;
  for (const content of paragraph.content) {
    if (content.type !== 'run') continue;
    if (seenRuns === runIndex) return content;
    seenRuns++;
  }
  return null;
}

function extractTextBoxShapeContents(
  xmlRun: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  options: TextBoxEnrichmentOptions
): ShapeContent[] {
  const shapeContents: ShapeContent[] = [];

  for (const runEl of getChildElements(xmlRun)) {
    shapeContents.push(
      ...extractTextBoxShapeContentsFromElement(
        runEl,
        styles,
        theme,
        numbering,
        rels,
        media,
        options
      )
    );
  }

  return shapeContents;
}

function extractTextBoxShapeContentsFromElement(
  element: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  options: TextBoxEnrichmentOptions
): ShapeContent[] {
  const localName = getLocalName(element.name ?? '');

  if (localName === 'drawing') {
    const shapeContent = parseWpsDrawingShapeContent(
      element,
      styles,
      theme,
      numbering,
      rels,
      media,
      options
    );
    return shapeContent ? [shapeContent] : [];
  }

  if (localName === 'pict' || localName === 'object') {
    return parseVmlShapeContents(element, styles, theme, numbering, rels, media, options);
  }

  if (localName === 'AlternateContent') {
    return parseAlternateContentShapeContents(
      element,
      styles,
      theme,
      numbering,
      rels,
      media,
      options
    );
  }

  return [];
}

function parseAlternateContentShapeContents(
  alternateContent: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  options: TextBoxEnrichmentOptions
): ShapeContent[] {
  const choices = getChildElements(alternateContent).filter(
    (el) => getLocalName(el.name ?? '') === 'Choice'
  );
  const fallback = getChildElements(alternateContent).find(
    (el) => getLocalName(el.name ?? '') === 'Fallback'
  );

  for (const choice of choices) {
    const parsed = parseAlternateBranch(choice, styles, theme, numbering, rels, media, options);
    if (parsed.length > 0) return parsed;
  }

  return fallback
    ? parseAlternateBranch(fallback, styles, theme, numbering, rels, media, options)
    : [];
}

function parseAlternateBranch(
  branch: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  options: TextBoxEnrichmentOptions
): ShapeContent[] {
  const shapeContents: ShapeContent[] = [];
  for (const child of getChildElements(branch)) {
    shapeContents.push(
      ...extractTextBoxShapeContentsFromElement(
        child,
        styles,
        theme,
        numbering,
        rels,
        media,
        options
      )
    );
  }
  return shapeContents;
}

function parseWpsDrawingShapeContent(
  drawingEl: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  options: TextBoxEnrichmentOptions
): ShapeContent | null {
  const textBox = parseTextBox(drawingEl);
  if (!textBox) return null;

  const wsp = findDeep(drawingEl, 'wps', 'wsp');
  const txbxContentEl = wsp ? getTextBoxContentElement(wsp) : null;
  if (txbxContentEl) {
    textBox.content = parseTextBoxParagraphs(
      txbxContentEl,
      styles,
      theme,
      numbering,
      rels,
      media,
      options
    );
  }

  return textBoxToShapeContent(textBox);
}

function parseVmlShapeContents(
  element: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  options: TextBoxEnrichmentOptions
): ShapeContent[] {
  return parseVMLTextBoxes(element).map(({ textBox, txbxContent }) => {
    textBox.content = parseTextBoxParagraphs(
      txbxContent,
      styles,
      theme,
      numbering,
      rels,
      media,
      options
    );
    return textBoxToShapeContent(textBox);
  });
}

function parseTextBoxParagraphs(
  txbxContentEl: XmlElement,
  styles: StyleMap | null,
  theme: Theme | null,
  numbering: NumberingMap | null,
  rels: RelationshipMap | null,
  media: Map<string, MediaFile> | null,
  options: TextBoxEnrichmentOptions
): Paragraph[] {
  return parseTextBoxContent(
    txbxContentEl,
    (node, pStyles, pTheme, pNumbering, pRels) => {
      const paragraph = parseParagraph(node, pStyles, pTheme, pNumbering, pRels ?? null, media, {
        inHeaderFooter: options.inHeaderFooter,
      });
      enrichParagraphTextBoxes(
        paragraph,
        node,
        pStyles,
        pTheme,
        pNumbering,
        pRels ?? null,
        media,
        options
      );
      return paragraph;
    },
    null,
    styles,
    theme,
    numbering,
    rels ?? undefined,
    media ?? undefined
  );
}

function textBoxToShapeContent(textBox: TextBox): ShapeContent {
  const shape: Shape = {
    type: 'shape',
    shapeType: 'textBox',
    size: textBox.size,
    position: textBox.position,
    wrap: textBox.wrap,
    fill: textBox.fill,
    outline: textBox.outline,
    textBody: {
      content: textBox.content,
      margins: textBox.margins,
    },
  };
  if (textBox.id) shape.id = textBox.id;

  return { type: 'shape', shape };
}
