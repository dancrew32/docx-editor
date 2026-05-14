import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import type { Node as PMNode } from 'prosemirror-model';
import { parseDocx } from '../../../docx/parser';
import { repackDocx } from '../../../docx/rezip';
import type { BlockContent, Paragraph } from '../../../types';
import { headerFooterToProseDoc, toProseDoc } from '../toProseDoc';
import { fromProseDoc } from '../fromProseDoc';

const FIXTURE = resolve(process.cwd(), 'e2e/fixtures/issue-318-textbox-support.docx');

describe('issue #318 text box support', () => {
  test('imports WPS and VML text boxes from body, headers, and table cells', async () => {
    const doc = await loadFixture();
    const pmDoc = toProseDoc(doc, { styles: doc.package.styles });
    const bodyTextBoxes = collectTextBoxes(pmDoc);

    expect(bodyTextBoxes.map((node) => node.textContent)).toEqual([
      'Body WPS Text Box',
      'Legacy VML Text Box',
      'Table Cell Text Box',
    ]);

    const tableCellTextBoxes = collectTextBoxesInside(pmDoc, 'tableCell');
    expect(tableCellTextBoxes.map((node) => node.textContent)).toContain('Table Cell Text Box');

    const header = doc.package.headers?.values().next().value;
    expect(header).toBeDefined();
    const headerDoc = headerFooterToProseDoc(header!.content, {
      styles: doc.package.styles,
      theme: doc.package.theme,
    });
    expect(collectTextBoxes(headerDoc).map((node) => node.textContent)).toEqual([
      'Header WPS Text Box',
    ]);
  });

  test('exports imported body and table-cell text boxes as shape runs', async () => {
    const doc = await loadFixture();
    const pmDoc = toProseDoc(doc, { styles: doc.package.styles });
    const roundTripped = fromProseDoc(pmDoc, doc);
    const serializedTexts = collectShapeTextBodyText(roundTripped.package.document.content);

    expect(serializedTexts).toContain('Body WPS Text Box');
    expect(serializedTexts).toContain('Legacy VML Text Box');
    expect(serializedTexts).toContain('Table Cell Text Box');
  });

  test('repackages imported text boxes as WPS txbx content', async () => {
    const doc = await loadFixture();
    const pmDoc = toProseDoc(doc, { styles: doc.package.styles });
    const roundTripped = fromProseDoc(pmDoc, doc);
    const repacked = await repackDocx(roundTripped, { updateModifiedDate: false });
    const zip = await JSZip.loadAsync(repacked);
    const documentXml = await zip.file('word/document.xml')!.async('text');

    expect(documentXml).toContain('<wps:txbx>');
    expect(documentXml).toContain('Body WPS Text Box');
    expect(documentXml).toContain('Legacy VML Text Box');
    expect(documentXml).toContain('Table Cell Text Box');
  });
});

async function loadFixture() {
  const buffer = readFileSync(FIXTURE);
  return await parseDocx(
    buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength)
  );
}

function collectTextBoxes(root: PMNode): PMNode[] {
  const nodes: PMNode[] = [];
  root.descendants((node) => {
    if (node.type.name === 'textBox') {
      nodes.push(node);
    }
  });
  return nodes;
}

function collectTextBoxesInside(root: PMNode, containerType: string): PMNode[] {
  const nodes: PMNode[] = [];
  root.descendants((node) => {
    if (node.type.name !== containerType) return;
    node.descendants((child) => {
      if (child.type.name === 'textBox') {
        nodes.push(child);
      }
    });
  });
  return nodes;
}

function collectShapeTextBodyText(blocks: BlockContent[]): string[] {
  const texts: string[] = [];

  for (const block of blocks) {
    if (block.type === 'paragraph') {
      collectParagraphShapeText(block, texts);
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          texts.push(...collectShapeTextBodyText(cell.content));
        }
      }
    }
  }

  return texts;
}

function collectParagraphShapeText(paragraph: Paragraph, texts: string[]): void {
  for (const paragraphContent of paragraph.content) {
    if (paragraphContent.type !== 'run') continue;
    for (const runContent of paragraphContent.content) {
      if (runContent.type !== 'shape') continue;
      const text = (runContent.shape.textBody?.content ?? []).map(readParagraphText).join('\n');
      if (text) texts.push(text);
    }
  }
}

function readParagraphText(paragraph: Paragraph): string {
  const pieces: string[] = [];
  for (const paragraphContent of paragraph.content) {
    if (paragraphContent.type !== 'run') continue;
    for (const runContent of paragraphContent.content) {
      if (runContent.type === 'text' && runContent.text) {
        pieces.push(runContent.text);
      }
    }
  }
  return pieces.join('');
}
