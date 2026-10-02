/**
 * A text box anchored in a body paragraph or in a table cell's paragraph: it loads (the
 * document and cell rules have room for the textBox block), its words are read and shown on
 * the page, and saving puts Word's own XML for the box back unchanged in the paragraph that
 * held it. A header being edited keeps its text boxes the same way.
 *
 * The .docx is built here from made-up text: a title, a "Notes: " paragraph (or a table of
 * photograph captions) holding an anchored text box (Word's form: the drawing, with the VML
 * picture as the fallback) and a closing paragraph.
 */

import { describe, test, expect } from 'bun:test';
import JSZip from 'jszip';
import { EditorState } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import { parseDocx } from '../../docx/parser';
import { repackDocx } from '../../docx/rezip';
import { toFlowBlocks } from '../../layout-bridge/toFlowBlocks';
import type { ParagraphBlock, TableBlock } from '../../layout-engine/types';
import type { Document } from '../../types/document';
import { headerFooterToProseDoc, toProseDoc } from './toProseDoc';
import { fromProseDoc, proseDocToBlocks } from './fromProseDoc';

const TITLE = 'MADE-UP PHOTOGRAPHIC REPORT';
const NOTES = 'Notes: ';
const TEXT_BOX_WORDS = 'Text box: made-up draft copy';
const CLOSING = 'Made-up closing line.';
const CAPTION = 'Photo 1: ';
const OTHER_CAPTION = 'Photo 2: made-up sounding tape';

const NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
  'xmlns:v="urn:schemas-microsoft-com:vml"',
  'xmlns:o="urn:schemas-microsoft-com:office:office"',
  'mc:Ignorable="wps"',
].join(' ');

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const paragraph = (text: string) => `<w:p>${run(text)}</w:p>`;

const TEXT_BOX_CONTENT = `<w:txbxContent>${paragraph(TEXT_BOX_WORDS)}</w:txbxContent>`;
const TEXT_BOX =
  '<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>' +
  '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" ' +
  'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>3200400</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="1828800" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>' +
  '<wp:docPr id="100" name="Text Box 100"/><wp:cNvGraphicFramePr/><a:graphic>' +
  '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>' +
  '<wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="457200"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>' +
  `</wps:spPr><wps:txbx>${TEXT_BOX_CONTENT}</wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" anchor="t"/>` +
  '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict>' +
  '<v:shape id="Text Box 100" style="position:absolute;margin-left:252pt;width:144pt;height:36pt">' +
  `<v:textbox>${TEXT_BOX_CONTENT}</v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent>`;
const TEXT_BOX_RUN = `<w:r>${TEXT_BOX}</w:r>`;

const cell = (content: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
/** A table of two photograph captions, the first one's paragraph holding the text box. */
const CAPTIONS_WITH_BOX =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>' +
  `<w:tr>${cell(`<w:p>${run(CAPTION)}${TEXT_BOX_RUN}</w:p>`)}${cell(paragraph(OTHER_CAPTION))}</w:tr></w:tbl>`;

/** The text box XML with what is inside each w:txbxContent left out. */
const frameOf = (xml: string) =>
  xml.replace(/<w:txbxContent>.*?<\/w:txbxContent>/g, '<w:txbxContent/>');

const CONTENT_TYPES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '</Types>';

const PACKAGE_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  '</Relationships>';

/** A made-up report whose middle part (a paragraph or a table) is `middle`. */
async function reportDocx(middle = `<w:p>${run(NOTES)}${TEXT_BOX_RUN}</w:p>`) {
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NAMESPACES}><w:body>` +
    `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>${TITLE}</w:t></w:r></w:p>` +
    middle +
    paragraph(CLOSING) +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>' +
    '</w:sectPr></w:body></w:document>';
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CONTENT_TYPES);
  zip.file('_rels/.rels', PACKAGE_RELS);
  zip.file('word/document.xml', documentXml);
  return zip.generateAsync({ type: 'arraybuffer' });
}

async function documentXmlOf(buffer: ArrayBuffer): Promise<string> {
  const file = (await JSZip.loadAsync(buffer)).file('word/document.xml');
  if (!file) throw new Error('No word/document.xml');
  return file.async('text');
}

/** Document position just after the given words. */
function positionAfter(doc: PMNode, words: string): number {
  let found = -1;
  doc.descendants((node, pos) => {
    const at = node.isText ? (node.text ?? '').indexOf(words) : -1;
    if (found < 0 && at >= 0) found = pos + at + words.length;
  });
  if (found < 0) throw new Error(`"${words}" not in the document`);
  return found;
}

/** Type `extra` after `words` in the editor and save the result. */
async function saveAfterTyping(doc: Document, words: string, extra: string): Promise<string> {
  const pmDoc = toProseDoc(doc);
  const state = EditorState.create({ doc: pmDoc });
  const edited = state.apply(state.tr.insertText(extra, positionAfter(pmDoc, words))).doc;
  edited.check();
  return documentXmlOf(await repackDocx(fromProseDoc(edited, doc)));
}

describe('text box anchored in a body paragraph', () => {
  test("the parser reads the text box's words", async () => {
    const doc = await parseDocx(await reportDocx(), { preloadFonts: false });
    const notes = doc.package.document.content[1];
    if (notes.type !== 'paragraph') throw new Error('expected the Notes paragraph');
    const shapes = notes.content.flatMap((item) =>
      item.type === 'run' ? item.content.filter((part) => part.type === 'shape') : []
    );
    expect(shapes).toHaveLength(1);
    const boxParagraphs = shapes[0].type === 'shape' ? shapes[0].shape.textBody?.content : [];
    expect(boxParagraphs?.[0]?.content).toEqual([
      expect.objectContaining({
        type: 'run',
        content: [expect.objectContaining({ type: 'text', text: TEXT_BOX_WORDS })],
      }),
    ]);
  });

  test('the editor document is valid, with the box after its paragraph', async () => {
    const doc = await parseDocx(await reportDocx(), { preloadFonts: false });
    const pmDoc = toProseDoc(doc);

    expect(() => pmDoc.check()).not.toThrow();
    // A replace over the whole document validates it the way the editor's transactions do
    const state = EditorState.create({ doc: pmDoc });
    expect(() => state.tr.replaceWith(0, pmDoc.content.size, pmDoc.content)).not.toThrow();

    const names: string[] = [];
    pmDoc.forEach((node) => names.push(node.type.name));
    expect(names).toEqual(['paragraph', 'paragraph', 'textBox', 'paragraph']);
    expect(pmDoc.child(2).textContent).toBe(TEXT_BOX_WORDS);
  });

  test("the page shows the box's words, framed, at the box's place in the document", async () => {
    const doc = await parseDocx(await reportDocx(), { preloadFonts: false });
    const pmDoc = toProseDoc(doc);
    const boxStart = positionAfter(pmDoc, NOTES) + 1; // the Notes paragraph closes, the box opens

    const blocks = toFlowBlocks(pmDoc, { pageContentWidth: 600 }) as ParagraphBlock[];
    const textOf = (block: ParagraphBlock) =>
      block.runs.map((r) => ('text' in r ? r.text : '')).join('');
    expect(blocks.map(textOf)).toEqual([TITLE, NOTES, TEXT_BOX_WORDS, CLOSING]);

    const box = blocks[2];
    expect(box.pmStart).toBe(boxStart + 1);
    expect(box.attrs?.borders?.left?.color).toBe('#000000');
  });

  test('saving without edits keeps the document exactly as it was', async () => {
    const buffer = await reportDocx();
    const doc = await parseDocx(buffer, { preloadFonts: false });
    const saved = await repackDocx(doc);
    expect(await documentXmlOf(saved)).toBe(await documentXmlOf(buffer));
  });

  test("an edit elsewhere keeps the text box's XML unchanged, in its place in its paragraph", async () => {
    const doc = await parseDocx(await reportDocx(), { preloadFonts: false });
    const xml = await saveAfterTyping(doc, CLOSING, ' Edited.');

    expect(xml).toContain(`${NOTES}</w:t></w:r>${TEXT_BOX_RUN}</w:p>`);
    expect(xml.split('<mc:AlternateContent>')).toHaveLength(2);
    expect(xml).toContain(`${CLOSING} Edited.`);
  });

  test('a box held before or inside the text goes back to the same place', async () => {
    const atStart = await parseDocx(await reportDocx(`<w:p>${TEXT_BOX_RUN}${run(NOTES)}</w:p>`), {
      preloadFonts: false,
    });
    expect(await saveAfterTyping(atStart, CLOSING, ' Edited.')).toContain(
      `<w:p>${TEXT_BOX_RUN}<w:r><w:t xml:space="preserve">${NOTES}</w:t></w:r></w:p>`
    );

    // The text either side has the same formatting, so saving rebuilds it as one run: the
    // box must cut that run where it was.
    const inside = await parseDocx(
      await reportDocx(`<w:p>${run('Before the box, ')}${TEXT_BOX_RUN}${run('after it.')}</w:p>`),
      { preloadFonts: false }
    );
    expect(await saveAfterTyping(inside, CLOSING, ' Edited.')).toContain(
      `Before the box, </w:t></w:r>${TEXT_BOX_RUN}<w:r><w:t>after it.</w:t></w:r></w:p>`
    );
  });

  test("an edit inside the box rewrites only the box's words, in the drawing and its fallback", async () => {
    const doc = await parseDocx(await reportDocx(), { preloadFonts: false });
    const xml = await saveAfterTyping(doc, TEXT_BOX_WORDS, ' (edited)');

    const saved = xml.match(/<mc:AlternateContent>.*<\/mc:AlternateContent>/)?.[0] ?? '';
    expect(frameOf(saved)).toBe(frameOf(TEXT_BOX));
    expect(saved.split(`${TEXT_BOX_WORDS} (edited)`)).toHaveLength(3);
    expect(xml).toContain(`${NOTES}</w:t></w:r><w:r>${TEXT_BOX.slice(0, 40)}`);
  });
});

describe('text box anchored in a table cell', () => {
  test('the editor document is valid, with the box after its paragraph in the cell', async () => {
    const doc = await parseDocx(await reportDocx(CAPTIONS_WITH_BOX), { preloadFonts: false });
    const pmDoc = toProseDoc(doc);

    expect(() => pmDoc.check()).not.toThrow();
    const firstCell = pmDoc.child(1).child(0).child(0);
    const names: string[] = [];
    firstCell.forEach((node) => names.push(node.type.name));
    expect(names).toEqual(['paragraph', 'textBox']);
    expect(firstCell.child(1).textContent).toBe(TEXT_BOX_WORDS);
  });

  test("the page shows the box's words inside its cell, at the box's place in the document", async () => {
    const doc = await parseDocx(await reportDocx(CAPTIONS_WITH_BOX), { preloadFonts: false });
    const pmDoc = toProseDoc(doc);
    const boxStart = positionAfter(pmDoc, CAPTION) + 1; // the caption paragraph closes, the box opens

    const table = toFlowBlocks(pmDoc, { pageContentWidth: 600 })[1] as TableBlock;
    const cellBlocks = table.rows[0].cells[0].blocks as ParagraphBlock[];
    const textOf = (block: ParagraphBlock) =>
      block.runs.map((r) => ('text' in r ? r.text : '')).join('');
    expect(cellBlocks.map(textOf)).toEqual([CAPTION, TEXT_BOX_WORDS]);
    expect(cellBlocks[1].pmStart).toBe(boxStart + 1);
  });

  test("an edit elsewhere keeps the text box's XML unchanged, in its place in its cell", async () => {
    const doc = await parseDocx(await reportDocx(CAPTIONS_WITH_BOX), { preloadFonts: false });
    const xml = await saveAfterTyping(doc, CLOSING, ' Edited.');

    expect(xml).toContain(`${CAPTION}</w:t></w:r>${TEXT_BOX_RUN}</w:p></w:tc>`);
    expect(xml.split('<mc:AlternateContent>')).toHaveLength(2);
    expect(xml).toContain(OTHER_CAPTION);
  });

  test("an edit inside the box rewrites only the box's words, in the drawing and its fallback", async () => {
    const doc = await parseDocx(await reportDocx(CAPTIONS_WITH_BOX), { preloadFonts: false });
    const xml = await saveAfterTyping(doc, TEXT_BOX_WORDS, ' (edited)');

    const saved = xml.match(/<mc:AlternateContent>.*<\/mc:AlternateContent>/)?.[0] ?? '';
    expect(frameOf(saved)).toBe(frameOf(TEXT_BOX));
    expect(saved.split(`${TEXT_BOX_WORDS} (edited)`)).toHaveLength(3);
    expect(xml).toContain(`${CAPTION}</w:t></w:r><w:r>${TEXT_BOX.slice(0, 40)}`);
  });
});

describe('text box in a header being edited', () => {
  test('the header editor shows the box, and saving the header keeps its XML unchanged', async () => {
    const doc = await parseDocx(await reportDocx(), { preloadFonts: false });
    const notes = doc.package.document.content[1]; // the same paragraph, as a header would hold it
    if (notes.type !== 'paragraph') throw new Error('expected the Notes paragraph');

    const pmDoc = headerFooterToProseDoc([notes]);
    expect(() => pmDoc.check()).not.toThrow();
    expect(pmDoc.child(1).type.name).toBe('textBox');
    expect(pmDoc.child(1).textContent).toBe(TEXT_BOX_WORDS);

    const [saved] = proseDocToBlocks(pmDoc);
    if (saved.type !== 'paragraph') throw new Error('expected the paragraph back');
    const boxes = saved.content.flatMap((item) =>
      item.type === 'run' ? item.content.filter((part) => part.type === 'shape') : []
    );
    expect(boxes).toEqual([
      expect.objectContaining({ originalXml: TEXT_BOX, textBodyChanged: false }),
    ]);
  });
});
