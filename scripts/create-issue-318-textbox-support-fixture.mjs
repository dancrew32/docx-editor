/**
 * Create a synthetic DOCX fixture for issue #318 text box support.
 *
 * The generated document covers WPS text boxes in the body, header, and a
 * table cell, plus a legacy VML text box. All text and package metadata are
 * synthetic.
 *
 * Run: bun scripts/create-issue-318-textbox-support-fixture.mjs
 */

import JSZip from 'jszip';
import * as fs from 'node:fs';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT = path.join(ROOT, 'e2e/fixtures/issue-318-textbox-support.docx');
const FIXTURE_DATE = new Date('2026-01-01T00:00:00Z');

const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  v: 'urn:schemas-microsoft-com:vml',
  w10: 'urn:schemas-microsoft-com:office:word',
};

const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;

const RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;

const DOCUMENT_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>`;

const CORE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:dcmitype="http://purl.org/dc/dcmitype/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Issue 318 Synthetic Text Box Support</dc:title>
  <dc:creator>docx-editor fixture generator</dc:creator>
  <cp:lastModifiedBy>docx-editor fixture generator</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>`;

const APP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>docx-editor fixture generator</Application>
</Properties>`;

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${NS.w}">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="24"/></w:rPr></w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="24"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr><w:spacing w:before="240" w:after="240"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:color w:val="2F75B5"/><w:sz w:val="36"/></w:rPr>
  </w:style>
</w:styles>`;

function escapeXml(value) {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}

function paragraph(text, opts = {}) {
  const style = opts.style ? `<w:pStyle w:val="${opts.style}"/>` : '';
  const spacing =
    opts.spacing ?? '<w:spacing w:before="0" w:after="160" w:line="276" w:lineRule="auto"/>';
  const color = opts.color ? `<w:color w:val="${opts.color}"/>` : '';
  const size = opts.size ? `<w:sz w:val="${opts.size}"/>` : '<w:sz w:val="24"/>';
  return `<w:p>
    <w:pPr>${style}${spacing}<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>${color}${size}</w:rPr></w:pPr>
    <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>${color}${size}</w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>
  </w:p>`;
}

function textBoxParagraph(text) {
  return `<w:p>
    <w:pPr><w:spacing w:before="0" w:after="0"/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:pPr>
    <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr><w:t>${escapeXml(text)}</w:t></w:r>
  </w:p>`;
}

function wpsShape({ id, text, cx = 1645920, cy = 457200, fill = 'EAF3F8' }) {
  const fillXml =
    fill === 'none' ? '<a:noFill/>' : `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>`;
  return `<a:graphic xmlns:a="${NS.a}">
    <a:graphicData uri="${NS.wps}">
      <wps:wsp xmlns:wps="${NS.wps}">
        <wps:cNvSpPr txBox="1"/>
        <wps:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          ${fillXml}
          <a:ln w="12700"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill></a:ln>
        </wps:spPr>
        <wps:txbx><w:txbxContent>${textBoxParagraph(text)}</w:txbxContent></wps:txbx>
        <wps:bodyPr lIns="45720" tIns="45720" rIns="45720" bIns="45720"/>
      </wps:wsp>
    </a:graphicData>
  </a:graphic>`;
}

function inlineWpsTextBox({ id, text, cx, cy, fill }) {
  return `<w:drawing>
    <wp:inline distT="0" distB="0" distL="0" distR="0">
      <wp:extent cx="${cx}" cy="${cy}"/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>
      <wp:docPr id="${id}" name="Synthetic inline text box ${id}"/>
      <wp:cNvGraphicFramePr/>
      ${wpsShape({ id, text, cx, cy, fill })}
    </wp:inline>
  </w:drawing>`;
}

function anchoredWpsTextBox({ id, text, cx, cy, x = 0, y = 0, fill = 'FFF2CC' }) {
  return `<w:drawing>
    <wp:anchor distT="0" distB="0" distL="0" distR="0"
      simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0"
      layoutInCell="1" allowOverlap="1">
      <wp:simplePos x="0" y="0"/>
      <wp:positionH relativeFrom="margin"><wp:posOffset>${x}</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="margin"><wp:posOffset>${y}</wp:posOffset></wp:positionV>
      <wp:extent cx="${cx}" cy="${cy}"/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>
      <wp:wrapSquare wrapText="bothSides"/>
      <wp:docPr id="${id}" name="Synthetic anchored text box ${id}"/>
      <wp:cNvGraphicFramePr/>
      ${wpsShape({ id, text, cx, cy, fill })}
    </wp:anchor>
  </w:drawing>`;
}

function vmlTextBox({ id, text }) {
  return `<w:pict>
    <v:shape id="${id}" type="#_x0000_t202"
      style="position:absolute;margin-left:0pt;margin-top:0pt;width:132pt;height:36pt;z-index:1"
      fillcolor="#FCE4D6" strokecolor="#C00000" strokeweight="1pt">
      <v:textbox inset="3pt,3pt,3pt,3pt">
        <w:txbxContent>${textBoxParagraph(text)}</w:txbxContent>
      </v:textbox>
      <w10:wrap type="square"/>
    </v:shape>
  </w:pict>`;
}

function drawingOnlyParagraph(drawingXml) {
  return `<w:p>
    <w:pPr><w:spacing w:before="0" w:after="160"/></w:pPr>
    <w:r>${drawingXml}</w:r>
  </w:p>`;
}

const BODY_WPS = inlineWpsTextBox({
  id: 31801,
  text: 'Body WPS Text Box',
  cx: 1645920,
  cy: 457200,
  fill: 'EAF3F8',
});

const HEADER_WPS = anchoredWpsTextBox({
  id: 31802,
  text: 'Header WPS Text Box',
  cx: 1828800,
  cy: 365760,
  x: 0,
  y: 0,
  fill: 'FFF2CC',
});

const TABLE_WPS = inlineWpsTextBox({
  id: 31803,
  text: 'Table Cell Text Box',
  cx: 1645920,
  cy: 457200,
  fill: 'E2F0D9',
});

const BODY_VML = vmlTextBox({
  id: 'VML31801',
  text: 'Legacy VML Text Box',
});

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="${NS.w}"
  xmlns:r="${NS.r}"
  xmlns:wp="${NS.wp}"
  xmlns:a="${NS.a}"
  xmlns:wps="${NS.wps}"
  xmlns:v="${NS.v}"
  xmlns:w10="${NS.w10}">
  <w:body>
    ${paragraph('Synthetic Text Box Support', { style: 'Heading1', color: '2F75B5', size: '36' })}
    ${paragraph('The following body paragraph contains a DrawingML text box in a run that has no visible text.')}
    ${drawingOnlyParagraph(BODY_WPS)}
    ${paragraph('The next body paragraph contains a legacy VML text box.')}
    ${drawingOnlyParagraph(BODY_VML)}
    ${paragraph('The table below checks that text boxes inside table cells survive the document pipeline.')}
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="0" w:type="auto"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/>
          <w:left w:val="single" w:sz="8" w:space="0" w:color="000000"/>
          <w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>
          <w:right w:val="single" w:sz="8" w:space="0" w:color="000000"/>
          <w:insideH w:val="single" w:sz="8" w:space="0" w:color="000000"/>
          <w:insideV w:val="single" w:sz="8" w:space="0" w:color="000000"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="4320"/><w:gridCol w:w="4320"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="4320" w:type="dxa"/></w:tcPr>
          ${paragraph('Cell with text box:')}
          ${drawingOnlyParagraph(TABLE_WPS)}
        </w:tc>
        <w:tc>
          <w:tcPr><w:tcW w:w="4320" w:type="dxa"/></w:tcPr>
          ${paragraph('Reference cell text.')}
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId2"/>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

const HEADER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr
  xmlns:w="${NS.w}"
  xmlns:r="${NS.r}"
  xmlns:wp="${NS.wp}"
  xmlns:a="${NS.a}"
  xmlns:wps="${NS.wps}">
  ${drawingOnlyParagraph(HEADER_WPS)}
</w:hdr>`;

const zip = new JSZip();
const zipOptions = { date: FIXTURE_DATE, createFolders: false };
zip.file('[Content_Types].xml', CONTENT_TYPES_XML, zipOptions);
zip.file('_rels/.rels', RELS_XML, zipOptions);
zip.file('docProps/core.xml', CORE_XML, zipOptions);
zip.file('docProps/app.xml', APP_XML, zipOptions);
zip.file('word/_rels/document.xml.rels', DOCUMENT_RELS_XML, zipOptions);
zip.file('word/styles.xml', STYLES_XML, zipOptions);
zip.file('word/document.xml', DOCUMENT_XML, zipOptions);
zip.file('word/header1.xml', HEADER_XML, zipOptions);

const buffer = await zip.generateAsync({ type: 'nodebuffer' });
fs.writeFileSync(OUT, buffer);
console.log(`Created ${OUT}`);
