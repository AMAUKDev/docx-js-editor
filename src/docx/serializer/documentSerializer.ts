/**
 * Document Serializer - Serialize complete document.xml
 *
 * Converts Document objects back to valid document.xml OOXML format.
 * Combines all content (paragraphs, tables) with section properties
 * and proper namespace declarations.
 *
 * OOXML Reference:
 * - Document root: w:document
 * - Document body: w:body
 * - Section properties: w:sectPr
 */

import type {
  Document,
  DocumentBody,
  BlockContent,
  SectionProperties,
  HeaderReference,
  FooterReference,
  FootnoteProperties,
  EndnoteProperties,
  BorderSpec,
} from '../../types/document';

import { serializeParagraph } from './paragraphSerializer';
import { resetDocPrIdCounter } from './runSerializer';
import { serializeTable } from './tableSerializer';

// ============================================================================
// XML NAMESPACES
// ============================================================================

/**
 * Standard OOXML namespaces for document.xml
 */
const NAMESPACES = {
  wpc: 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
  cx: 'http://schemas.microsoft.com/office/drawing/2014/chartex',
  cx1: 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex',
  cx2: 'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex',
  cx3: 'http://schemas.microsoft.com/office/drawing/2016/5/9/chartex',
  cx4: 'http://schemas.microsoft.com/office/drawing/2016/5/10/chartex',
  cx5: 'http://schemas.microsoft.com/office/drawing/2016/5/11/chartex',
  cx6: 'http://schemas.microsoft.com/office/drawing/2016/5/12/chartex',
  cx7: 'http://schemas.microsoft.com/office/drawing/2016/5/13/chartex',
  cx8: 'http://schemas.microsoft.com/office/drawing/2016/5/14/chartex',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  aink: 'http://schemas.microsoft.com/office/drawing/2016/ink',
  am3d: 'http://schemas.microsoft.com/office/drawing/2017/model3d',
  o: 'urn:schemas-microsoft-com:office:office',
  oel: 'http://schemas.microsoft.com/office/2019/extlst',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  v: 'urn:schemas-microsoft-com:vml',
  wp14: 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  w10: 'urn:schemas-microsoft-com:office:word',
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
  w15: 'http://schemas.microsoft.com/office/word/2012/wordml',
  w16cex: 'http://schemas.microsoft.com/office/word/2018/wordml/cex',
  w16cid: 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
  w16: 'http://schemas.microsoft.com/office/word/2018/wordml',
  w16sdtdh: 'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash',
  w16se: 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
  w16sdtfl: 'http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock',
  w16du: 'http://schemas.microsoft.com/office/word/2023/wordml/word16du',
  wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
  wpi: 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
  wne: 'http://schemas.microsoft.com/office/word/2006/wordml',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
};

/**
 * Build namespace declaration string for document element
 */
function buildNamespaceDeclarations(): string {
  // Include ALL namespaces to match Word's output.
  // Missing namespace declarations can cause Word to reject documents
  // that contain raw XML fragments (e.g. rawXml sectPr) referencing
  // those namespace prefixes.
  return Object.entries(NAMESPACES)
    .map(([prefix, uri]) => `xmlns:${prefix}="${uri}"`)
    .join(' ');
}

// ============================================================================
// XML ESCAPING
// ============================================================================

// ============================================================================
// BORDER SERIALIZATION
// ============================================================================

/**
 * Serialize a border element
 */
function serializeBorder(border: BorderSpec | undefined, elementName: string): string {
  if (!border || border.style === 'none' || border.style === 'nil') {
    return '';
  }

  const attrs: string[] = [`w:val="${border.style}"`];

  if (border.size !== undefined) {
    attrs.push(`w:sz="${border.size}"`);
  }

  if (border.space !== undefined) {
    attrs.push(`w:space="${border.space}"`);
  }

  if (border.color) {
    if (border.color.auto) {
      attrs.push('w:color="auto"');
    } else if (border.color.rgb) {
      attrs.push(`w:color="${border.color.rgb}"`);
    }

    if (border.color.themeColor) {
      attrs.push(`w:themeColor="${border.color.themeColor}"`);
    }

    if (border.color.themeTint) {
      attrs.push(`w:themeTint="${border.color.themeTint}"`);
    }

    if (border.color.themeShade) {
      attrs.push(`w:themeShade="${border.color.themeShade}"`);
    }
  }

  if (border.shadow) {
    attrs.push('w:shadow="true"');
  }

  if (border.frame) {
    attrs.push('w:frame="true"');
  }

  return `<w:${elementName} ${attrs.join(' ')}/>`;
}

// ============================================================================
// SECTION PROPERTIES SERIALIZATION
// ============================================================================

/**
 * Serialize header reference (w:headerReference)
 */
function serializeHeaderReference(ref: HeaderReference): string {
  const attrs: string[] = [`w:type="${ref.type}"`, `r:id="${ref.rId}"`];

  return `<w:headerReference ${attrs.join(' ')}/>`;
}

/**
 * Serialize footer reference (w:footerReference)
 */
function serializeFooterReference(ref: FooterReference): string {
  const attrs: string[] = [`w:type="${ref.type}"`, `r:id="${ref.rId}"`];

  return `<w:footerReference ${attrs.join(' ')}/>`;
}

/**
 * Serialize footnote properties (w:footnotePr)
 */
function serializeFootnoteProperties(props: FootnoteProperties | undefined): string {
  if (!props) return '';

  const parts: string[] = [];

  if (props.position) {
    parts.push(`<w:pos w:val="${props.position}"/>`);
  }

  if (props.numFmt) {
    parts.push(`<w:numFmt w:val="${props.numFmt}"/>`);
  }

  if (props.numStart !== undefined) {
    parts.push(`<w:numStart w:val="${props.numStart}"/>`);
  }

  if (props.numRestart) {
    parts.push(`<w:numRestart w:val="${props.numRestart}"/>`);
  }

  if (parts.length === 0) return '';

  return `<w:footnotePr>${parts.join('')}</w:footnotePr>`;
}

/**
 * Serialize endnote properties (w:endnotePr)
 */
function serializeEndnoteProperties(props: EndnoteProperties | undefined): string {
  if (!props) return '';

  const parts: string[] = [];

  if (props.position) {
    parts.push(`<w:pos w:val="${props.position}"/>`);
  }

  if (props.numFmt) {
    parts.push(`<w:numFmt w:val="${props.numFmt}"/>`);
  }

  if (props.numStart !== undefined) {
    parts.push(`<w:numStart w:val="${props.numStart}"/>`);
  }

  if (props.numRestart) {
    parts.push(`<w:numRestart w:val="${props.numRestart}"/>`);
  }

  if (parts.length === 0) return '';

  return `<w:endnotePr>${parts.join('')}</w:endnotePr>`;
}

/**
 * Serialize page size (w:pgSz)
 */
function serializePageSize(props: SectionProperties): string {
  const attrs: string[] = [];

  if (props.pageWidth !== undefined) {
    attrs.push(`w:w="${props.pageWidth}"`);
  }

  if (props.pageHeight !== undefined) {
    attrs.push(`w:h="${props.pageHeight}"`);
  }

  if (props.orientation === 'landscape') {
    attrs.push('w:orient="landscape"');
  }

  if (attrs.length === 0) return '';

  return `<w:pgSz ${attrs.join(' ')}/>`;
}

/**
 * Serialize page margins (w:pgMar)
 */
function serializePageMargins(props: SectionProperties): string {
  const attrs: string[] = [];

  if (props.marginTop !== undefined) {
    attrs.push(`w:top="${props.marginTop}"`);
  }

  if (props.marginRight !== undefined) {
    attrs.push(`w:right="${props.marginRight}"`);
  }

  if (props.marginBottom !== undefined) {
    attrs.push(`w:bottom="${props.marginBottom}"`);
  }

  if (props.marginLeft !== undefined) {
    attrs.push(`w:left="${props.marginLeft}"`);
  }

  if (props.headerDistance !== undefined) {
    attrs.push(`w:header="${props.headerDistance}"`);
  }

  if (props.footerDistance !== undefined) {
    attrs.push(`w:footer="${props.footerDistance}"`);
  }

  if (props.gutter !== undefined) {
    attrs.push(`w:gutter="${props.gutter}"`);
  }

  if (attrs.length === 0) return '';

  return `<w:pgMar ${attrs.join(' ')}/>`;
}

/**
 * Serialize columns (w:cols)
 */
function serializeColumns(props: SectionProperties): string {
  if (!props.columnCount && !props.columns?.length) return '';

  const attrs: string[] = [];

  if (props.columnCount !== undefined && props.columnCount > 1) {
    attrs.push(`w:num="${props.columnCount}"`);
  }

  if (props.columnSpace !== undefined) {
    attrs.push(`w:space="${props.columnSpace}"`);
  }

  if (props.equalWidth !== undefined) {
    attrs.push(`w:equalWidth="${props.equalWidth ? '1' : '0'}"`);
  }

  if (props.separator) {
    attrs.push('w:sep="1"');
  }

  // Individual column definitions
  let colElements = '';
  if (props.columns && props.columns.length > 0) {
    colElements = props.columns
      .map((col) => {
        const colAttrs: string[] = [];
        if (col.width !== undefined) {
          colAttrs.push(`w:w="${col.width}"`);
        }
        if (col.space !== undefined) {
          colAttrs.push(`w:space="${col.space}"`);
        }
        return `<w:col ${colAttrs.join(' ')}/>`;
      })
      .join('');
  }

  if (attrs.length === 0 && !colElements) return '';

  const attrsStr = attrs.length > 0 ? ' ' + attrs.join(' ') : '';
  return `<w:cols${attrsStr}>${colElements}</w:cols>`;
}

/**
 * Serialize line numbers (w:lnNumType)
 */
function serializeLineNumbers(props: SectionProperties): string {
  if (!props.lineNumbers) return '';

  const ln = props.lineNumbers;
  const attrs: string[] = [];

  if (ln.countBy !== undefined) {
    attrs.push(`w:countBy="${ln.countBy}"`);
  }

  if (ln.start !== undefined) {
    attrs.push(`w:start="${ln.start}"`);
  }

  if (ln.distance !== undefined) {
    attrs.push(`w:distance="${ln.distance}"`);
  }

  if (ln.restart) {
    attrs.push(`w:restart="${ln.restart}"`);
  }

  if (attrs.length === 0) return '';

  return `<w:lnNumType ${attrs.join(' ')}/>`;
}

/**
 * Serialize page borders (w:pgBorders)
 */
function serializePageBorders(props: SectionProperties): string {
  if (!props.pageBorders) return '';

  const pb = props.pageBorders;
  const attrs: string[] = [];
  const borderElements: string[] = [];

  if (pb.display) {
    attrs.push(`w:display="${pb.display}"`);
  }

  if (pb.offsetFrom) {
    attrs.push(`w:offsetFrom="${pb.offsetFrom}"`);
  }

  if (pb.zOrder) {
    attrs.push(`w:zOrder="${pb.zOrder}"`);
  }

  if (pb.top) {
    const topXml = serializeBorder(pb.top, 'top');
    if (topXml) borderElements.push(topXml);
  }

  if (pb.left) {
    const leftXml = serializeBorder(pb.left, 'left');
    if (leftXml) borderElements.push(leftXml);
  }

  if (pb.bottom) {
    const bottomXml = serializeBorder(pb.bottom, 'bottom');
    if (bottomXml) borderElements.push(bottomXml);
  }

  if (pb.right) {
    const rightXml = serializeBorder(pb.right, 'right');
    if (rightXml) borderElements.push(rightXml);
  }

  if (borderElements.length === 0) return '';

  const attrsStr = attrs.length > 0 ? ' ' + attrs.join(' ') : '';
  return `<w:pgBorders${attrsStr}>${borderElements.join('')}</w:pgBorders>`;
}

/**
 * Serialize document grid (w:docGrid)
 */
function serializeDocGrid(props: SectionProperties): string {
  if (!props.docGrid) return '';

  const dg = props.docGrid;
  const attrs: string[] = [];

  if (dg.type) {
    attrs.push(`w:type="${dg.type}"`);
  }

  if (dg.linePitch !== undefined) {
    attrs.push(`w:linePitch="${dg.linePitch}"`);
  }

  if (dg.charSpace !== undefined) {
    attrs.push(`w:charSpace="${dg.charSpace}"`);
  }

  if (attrs.length === 0) return '';

  return `<w:docGrid ${attrs.join(' ')}/>`;
}

/**
 * Serialize section properties (w:sectPr)
 */
export function serializeSectionProperties(props: SectionProperties | undefined): string {
  if (!props) return '';

  // Use raw XML when available for lossless round-trip.
  // This preserves elements not modelled in our type (pgNumType, formProt, etc.)
  // and maintains correct OOXML schema element ordering.
  if (props.rawXml) {
    return props.rawXml;
  }

  const parts: string[] = [];

  // Header references
  if (props.headerReferences) {
    for (const ref of props.headerReferences) {
      parts.push(serializeHeaderReference(ref));
    }
  }

  // Footer references
  if (props.footerReferences) {
    for (const ref of props.footerReferences) {
      parts.push(serializeFooterReference(ref));
    }
  }

  // Footnote properties
  const footnotePrXml = serializeFootnoteProperties(props.footnotePr);
  if (footnotePrXml) {
    parts.push(footnotePrXml);
  }

  // Endnote properties
  const endnotePrXml = serializeEndnoteProperties(props.endnotePr);
  if (endnotePrXml) {
    parts.push(endnotePrXml);
  }

  // Section type
  if (props.sectionStart) {
    parts.push(`<w:type w:val="${props.sectionStart}"/>`);
  }

  // Page size
  const pgSzXml = serializePageSize(props);
  if (pgSzXml) {
    parts.push(pgSzXml);
  }

  // Page margins
  const pgMarXml = serializePageMargins(props);
  if (pgMarXml) {
    parts.push(pgMarXml);
  }

  // Paper source
  if (props.paperSrcFirst !== undefined || props.paperSrcOther !== undefined) {
    const attrs: string[] = [];
    if (props.paperSrcFirst !== undefined) {
      attrs.push(`w:first="${props.paperSrcFirst}"`);
    }
    if (props.paperSrcOther !== undefined) {
      attrs.push(`w:other="${props.paperSrcOther}"`);
    }
    parts.push(`<w:paperSrc ${attrs.join(' ')}/>`);
  }

  // Page borders
  const pgBordersXml = serializePageBorders(props);
  if (pgBordersXml) {
    parts.push(pgBordersXml);
  }

  // Line numbers
  const lnNumXml = serializeLineNumbers(props);
  if (lnNumXml) {
    parts.push(lnNumXml);
  }

  // Columns
  const colsXml = serializeColumns(props);
  if (colsXml) {
    parts.push(colsXml);
  }

  // Document grid
  const docGridXml = serializeDocGrid(props);
  if (docGridXml) {
    parts.push(docGridXml);
  }

  // Vertical alignment
  if (props.verticalAlign) {
    parts.push(`<w:vAlign w:val="${props.verticalAlign}"/>`);
  }

  // Bidirectional
  if (props.bidi) {
    parts.push('<w:bidi/>');
  }

  // Title page (different first page header/footer)
  if (props.titlePg) {
    parts.push('<w:titlePg/>');
  }

  // Even and odd headers
  if (props.evenAndOddHeaders) {
    parts.push('<w:evenAndOddHeaders/>');
  }

  if (parts.length === 0) return '';

  return `<w:sectPr>${parts.join('')}</w:sectPr>`;
}

// ============================================================================
// CONTENT SERIALIZATION
// ============================================================================

/**
 * Serialize a single block content item (paragraph or table)
 */
function serializeBlockContent(block: BlockContent): string {
  if (block.type === 'paragraph') {
    return serializeParagraph(block);
  } else if (block.type === 'table') {
    return serializeTable(block);
  } else if (block.type === 'blockSdt') {
    // Block-level SDT: wrap content in w:sdt
    const contentXml = block.content.map((b) => serializeBlockContent(b)).join('');
    const props = block.properties;
    const prParts: string[] = [];
    if (props.alias) prParts.push(`<w:alias w:val="${props.alias}"/>`);
    if (props.tag) prParts.push(`<w:tag w:val="${props.tag}"/>`);
    return `<w:sdt><w:sdtPr>${prParts.join('')}</w:sdtPr><w:sdtContent>${contentXml}</w:sdtContent></w:sdt>`;
  }
  return '';
}

/**
 * Check if any paragraph in the content tree has locked=true.
 */
function hasLockedParagraphs(content: BlockContent[]): boolean {
  for (const block of content) {
    if (block.type === 'paragraph' && block.formatting?.locked) return true;
    if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          if (hasLockedParagraphs(cell.content)) return true;
        }
      }
    }
    if (block.type === 'blockSdt' && hasLockedParagraphs(block.content)) return true;
  }
  return false;
}

/**
 * Serialize document body content.
 *
 * When locked paragraphs exist, wraps unlocked paragraphs with
 * w:permStart/w:permEnd so Word's document protection (set in settings.xml)
 * allows editing only the unlocked regions.
 */
function serializeBodyContent(content: BlockContent[]): string {
  const hasLocked = hasLockedParagraphs(content);
  if (!hasLocked) {
    return content.map((block) => serializeBlockContent(block)).join('');
  }

  // Insert permission ranges around unlocked paragraph runs
  const parts: string[] = [];
  let permId = 100;
  let inUnlockedRun = false;

  for (const block of content) {
    const isUnlockedParagraph = block.type === 'paragraph' && !block.formatting?.locked;

    if (isUnlockedParagraph && !inUnlockedRun) {
      parts.push(`<w:permStart w:id="${permId}" w:edGrp="everyone"/>`);
      inUnlockedRun = true;
    } else if (!isUnlockedParagraph && inUnlockedRun) {
      parts.push(`<w:permEnd w:id="${permId}"/>`);
      permId++;
      inUnlockedRun = false;
    }

    parts.push(serializeBlockContent(block));
  }

  if (inUnlockedRun) {
    parts.push(`<w:permEnd w:id="${permId}"/>`);
  }

  return parts.join('');
}

/**
 * Check if a Document has any locked paragraphs in its body content.
 * Used by rezip to decide whether to inject w:documentProtection into settings.xml.
 */
export function documentHasLockedParagraphs(doc: Document): boolean {
  return hasLockedParagraphs(doc.package.document.content);
}

// ============================================================================
// MAIN DOCUMENT SERIALIZATION
// ============================================================================

/**
 * Serialize a DocumentBody to document.xml body content
 *
 * @param body - The document body to serialize
 * @returns XML string for the body element (without body tags)
 */
export function serializeDocumentBody(body: DocumentBody): string {
  const parts: string[] = [];

  // Serialize all content blocks
  parts.push(serializeBodyContent(body.content));

  // Final section properties (at the end of body)
  if (body.finalSectionProperties) {
    parts.push(serializeSectionProperties(body.finalSectionProperties));
  }

  return parts.join('');
}

/**
 * Serialize a complete Document to valid document.xml
 *
 * @param doc - The document to serialize
 * @returns Complete XML string for document.xml
 */
export function serializeDocument(doc: Document): string {
  // Reset the docPr ID counter so each serialization pass produces unique IDs
  resetDocPrIdCounter();

  const parts: string[] = [];

  // XML declaration
  parts.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');

  // Document element with namespaces — use original tag if available for lossless round-trip
  if (doc.package.document.rawDocumentTag) {
    parts.push(doc.package.document.rawDocumentTag);
  } else {
    const nsDecl = buildNamespaceDeclarations();
    parts.push(
      `<w:document ${nsDecl} mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14">`
    );
  }

  // Document body
  parts.push('<w:body>');
  let bodyXml = serializeDocumentBody(doc.package.document);
  bodyXml = wrapTocWithFieldCodes(bodyXml);
  bodyXml = ensureFieldBalance(bodyXml);
  bodyXml = ensureBookmarkBalance(bodyXml);
  bodyXml = deduplicateParaIds(bodyXml);
  parts.push(bodyXml);
  parts.push('</w:body>');

  // Close document element
  parts.push('</w:document>');

  return parts.join('');
}

/**
 * Wrap consecutive TOC-styled paragraphs with TOC field codes so Word
 * recognizes the section as an updatable Table of Contents.
 *
 * Matches the structure Word produces natively:
 *   <w:p pStyle="TOC1">                    ← first TOC entry
 *     <w:pPr>...</w:pPr>
 *     <w:r> fldChar begin </w:r>           ← TOC field begin
 *     <w:r> instrText TOC \o "1-N" ... </w:r>
 *     <w:r> fldChar separate </w:r>
 *     <w:hyperlink anchor="_TocNNN">       ← entry content with PAGEREF on page num
 *       ...text...
 *       <w:r> fldChar begin </w:r>         ← PAGEREF field
 *       <w:r> instrText PAGEREF _TocNNN \h </w:r>
 *       <w:r> fldChar separate </w:r>
 *       <w:r> <w:t>3</w:t> </w:r>
 *       <w:r> fldChar end </w:r>
 *     </w:hyperlink>
 *   </w:p>
 *   ... more TOC entries with PAGEREF fields ...
 *   <w:p pStyle="TOC1">                    ← last TOC entry
 *     ...content...
 *     <w:r> fldChar end </w:r>             ← TOC field end
 *   </w:p>
 */
function wrapTocWithFieldCodes(bodyXml: string): string {
  // Find ALL TOC1-9 paragraphs anywhere in the document
  const tocParagraphPattern = /<w:p\b[^>]*>([\s\S]*?)<\/w:p>/g;
  const allTocEntries: { xml: string; start: number; end: number; anchor: string | null }[] = [];
  let m;

  while ((m = tocParagraphPattern.exec(bodyXml)) !== null) {
    // Check if this paragraph has a TOC1-9 style
    if (!/w:val="TOC\d"/.test(m[0])) continue;
    // Skip if it already has fldChar (already wrapped — e.g. from original DOCX)
    if (/fldCharType/.test(m[0])) continue;

    // Extract the _Toc anchor from hyperlinks in this paragraph
    const anchorMatch = m[0].match(/w:anchor="(_Toc\d+)"/);

    allTocEntries.push({
      xml: m[0],
      start: m.index,
      end: m.index + m[0].length,
      anchor: anchorMatch ? anchorMatch[1] : null,
    });
  }

  if (allTocEntries.length === 0) return bodyXml;

  // Find consecutive runs of TOC paragraphs (there should be one block)
  // For now, take the first consecutive run
  const tocBlock: typeof allTocEntries = [allTocEntries[0]];
  for (let i = 1; i < allTocEntries.length; i++) {
    const gap = bodyXml.slice(tocBlock[tocBlock.length - 1].end, allTocEntries[i].start).trim();
    if (gap.length === 0) {
      tocBlock.push(allTocEntries[i]);
    } else {
      break;
    }
  }

  // Detect max TOC level from entries
  const allXml = tocBlock.map((e) => e.xml).join('');
  const levelMatches = allXml.match(/w:val="TOC(\d)"/g) || [];
  const maxLevel =
    levelMatches.length > 0
      ? Math.max(...levelMatches.map((m2) => parseInt(m2.match(/\d/)![0])))
      : 3; // fallback only if no TOC levels detected

  // TOC field begin runs (injected into first entry after </w:pPr>)
  const fieldBeginRuns =
    `<w:r><w:rPr><w:noProof/></w:rPr>` +
    `<w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:rPr><w:noProof/></w:rPr>` +
    `<w:instrText xml:space="preserve"> TOC \\o "1-${maxLevel}" \\h \\z \\u </w:instrText></w:r>` +
    `<w:r><w:rPr><w:noProof/></w:rPr>` +
    `<w:fldChar w:fldCharType="separate"/></w:r>`;

  // TOC field end run (appended to last entry before </w:p>)
  const fieldEndRun = `<w:r><w:rPr><w:noProof/></w:rPr>` + `<w:fldChar w:fldCharType="end"/></w:r>`;

  // Process entries in reverse order so string positions remain valid
  let result = bodyXml;

  for (let i = tocBlock.length - 1; i >= 0; i--) {
    const entry = tocBlock[i];
    let entryXml = entry.xml;

    // For each hyperlink with a _Toc anchor, wrap the last <w:t>NUMBER</w:t> in PAGEREF
    if (entry.anchor) {
      // The entry may have multiple hyperlinks (number, title, page number as separate hyperlinks)
      // or one hyperlink wrapping everything. Find the LAST hyperlink with this anchor —
      // its text content is the page number.
      const anchor = entry.anchor;
      const hyperlinkPattern = new RegExp(
        `<w:hyperlink\\s+w:anchor="${anchor}"[^>]*>[\\s\\S]*?</w:hyperlink>`,
        'g'
      );
      const hyperlinks: { match: string; index: number }[] = [];
      let hm;
      while ((hm = hyperlinkPattern.exec(entryXml)) !== null) {
        hyperlinks.push({ match: hm[0], index: hm.index });
      }

      if (hyperlinks.length > 0) {
        // The last hyperlink contains the page number
        const lastHL = hyperlinks[hyperlinks.length - 1];
        const pageRefWrapped =
          `<w:hyperlink w:anchor="${anchor}" w:history="1">` +
          `<w:r><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>` +
          `<w:r><w:rPr><w:webHidden/></w:rPr>` +
          `<w:instrText xml:space="preserve"> PAGEREF ${anchor} \\h </w:instrText></w:r>` +
          `<w:r><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>` +
          // Extract just the text runs from the original hyperlink
          lastHL.match.replace(/<w:hyperlink[^>]*>/, '').replace(/<\/w:hyperlink>/, '') +
          `<w:r><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>` +
          `</w:hyperlink>`;
        entryXml =
          entryXml.slice(0, lastHL.index) +
          pageRefWrapped +
          entryXml.slice(lastHL.index + lastHL.match.length);
      }
    }

    // Inject TOC field begin after </w:pPr> in the first entry
    if (i === 0) {
      const pprEnd = entryXml.indexOf('</w:pPr>');
      if (pprEnd >= 0) {
        const insertPos = pprEnd + '</w:pPr>'.length;
        entryXml = entryXml.slice(0, insertPos) + fieldBeginRuns + entryXml.slice(insertPos);
      }
    }

    // Inject TOC field end before </w:p> in the last entry
    if (i === tocBlock.length - 1) {
      const closeP = entryXml.lastIndexOf('</w:p>');
      entryXml = entryXml.slice(0, closeP) + fieldEndRun + entryXml.slice(closeP);
    }

    // Replace in the result string
    result = result.slice(0, entry.start) + entryXml + result.slice(entry.end);
  }

  return result;
}

/**
 * Ensure fldChar begin/end elements are balanced in the serialized body XML.
 *
 * Multi-paragraph fields (TOC, INDEX, etc.) can lose their outer wrapper
 * during the parse→PM→serialize round-trip, leaving orphaned fldChar end
 * elements. Word rejects documents with unbalanced fldChar sequences.
 *
 * This function strips orphaned fldChar end elements (those without a matching
 * begin) by removing the containing <w:r>...</w:r> element.
 */
function ensureFieldBalance(bodyXml: string): string {
  // Find all fldChar elements with their positions
  const fldCharRegex = /<w:fldChar\s+w:fldCharType="(begin|separate|end)"[^/]*\/>/g;
  let depth = 0;
  const orphanedEndPositions: number[] = [];

  let match;
  while ((match = fldCharRegex.exec(bodyXml)) !== null) {
    const charType = match[1];
    if (charType === 'begin') {
      depth++;
    } else if (charType === 'end') {
      depth--;
      if (depth < 0) {
        // This end has no matching begin — mark for removal
        orphanedEndPositions.push(match.index);
        depth = 0;
      }
    }
  }

  if (orphanedEndPositions.length === 0) return bodyXml;

  // Remove orphaned fldChar end elements by stripping their containing <w:r>
  let result = bodyXml;
  // Process in reverse order so positions remain valid after each removal
  for (let i = orphanedEndPositions.length - 1; i >= 0; i--) {
    const pos = orphanedEndPositions[i];
    // Find the enclosing <w:r> ... </w:r>
    const runStart = result.lastIndexOf('<w:r>', pos);
    const runStartWithAttrs = result.lastIndexOf('<w:r ', pos);
    const actualRunStart = Math.max(runStart, runStartWithAttrs);
    const runEnd = result.indexOf('</w:r>', pos);
    if (actualRunStart >= 0 && runEnd > actualRunStart) {
      result = result.slice(0, actualRunStart) + result.slice(runEnd + 6);
    }
  }

  return result;
}

/**
 * Ensure bookmarkStart/bookmarkEnd elements are balanced.
 *
 * Content loss during the parse→PM→serialize round-trip can drop bookmarkStart
 * elements (e.g. when they're in paragraphs with mc:AlternateContent that gets
 * stripped) while leaving the corresponding bookmarkEnd. Word validates bookmark
 * balance and flags orphaned markers as corruption.
 *
 * This function strips orphaned bookmarkEnd elements (those without a matching
 * bookmarkStart) and orphaned bookmarkStart elements (those without a matching
 * bookmarkEnd).
 */
function ensureBookmarkBalance(bodyXml: string): string {
  let result = bodyXml;

  // --- Pass 1: Remove duplicate bookmark IDs (keep only the first occurrence) ---
  // Duplicate bookmark IDs cause Word to reject the file as corrupt.
  // This can happen when paragraphs are split and attrs (including bookmarks)
  // are inadvertently copied to the new paragraph.
  const seenStartIds = new Set<string>();
  result = result.replace(
    /<w:bookmarkStart\s+w:id="(\d+)"\s+w:name="[^"]*"[^/]*\/>/g,
    (fullMatch, id) => {
      if (seenStartIds.has(id)) return ''; // Remove duplicate
      seenStartIds.add(id);
      return fullMatch;
    }
  );
  const seenEndIds = new Set<string>();
  result = result.replace(/<w:bookmarkEnd\s+w:id="(\d+)"\s*\/>/g, (fullMatch, id) => {
    if (seenEndIds.has(id)) return ''; // Remove duplicate
    seenEndIds.add(id);
    return fullMatch;
  });

  // --- Pass 2: Remove orphaned bookmarks (start without end or vice versa) ---
  const startIds = new Set<string>();
  const startRegex = /<w:bookmarkStart\s+w:id="(\d+)"/g;
  let match;
  while ((match = startRegex.exec(result)) !== null) {
    startIds.add(match[1]);
  }

  const endIds = new Set<string>();
  const endRegex = /<w:bookmarkEnd\s+w:id="(\d+)"/g;
  while ((match = endRegex.exec(result)) !== null) {
    endIds.add(match[1]);
  }

  const orphanedEndIds = new Set<string>();
  for (const id of endIds) {
    if (!startIds.has(id)) orphanedEndIds.add(id);
  }

  const orphanedStartIds = new Set<string>();
  for (const id of startIds) {
    if (!endIds.has(id)) orphanedStartIds.add(id);
  }

  if (orphanedEndIds.size === 0 && orphanedStartIds.size === 0) return result;

  for (const id of orphanedEndIds) {
    const re = new RegExp(`<w:bookmarkEnd\\s+w:id="${id}"\\s*/>`, 'g');
    result = result.replace(re, '');
  }

  for (const id of orphanedStartIds) {
    const re = new RegExp(`<w:bookmarkStart\\s+w:id="${id}"\\s+w:name="[^"]*"[^/]*/>`, 'g');
    result = result.replace(re, '');
  }

  return result;
}

/**
 * Remove duplicate w14:paraId values from the body XML.
 *
 * Each paragraph must have a unique paraId; duplicates cause Word to flag
 * the file as corrupt. Keeps the first occurrence and strips the attribute
 * from subsequent paragraphs with the same value.
 */
function deduplicateParaIds(bodyXml: string): string {
  const seen = new Set<string>();
  return bodyXml.replace(/w14:paraId="([A-Fa-f0-9]+)"/g, (fullMatch, id) => {
    if (seen.has(id)) return `w14:paraId="${generateParaId()}"`;
    seen.add(id);
    return fullMatch;
  });
}

/** Generate a random 8-character hex paraId (OOXML format). */
function generateParaId(): string {
  return Math.floor(Math.random() * 0xffffffff)
    .toString(16)
    .toUpperCase()
    .padStart(8, '0');
}

/**
 * Serialize just the document body (useful for partial updates)
 *
 * @param body - The document body to serialize
 * @returns XML string for the w:body element
 */
export function serializeDocumentBodyElement(body: DocumentBody): string {
  return `<w:body>${serializeDocumentBody(body)}</w:body>`;
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Check if document has any content
 */
export function hasDocumentContent(doc: Document): boolean {
  return doc.package.document.content.length > 0;
}

/**
 * Check if document has sections
 */
export function hasDocumentSections(doc: Document): boolean {
  return (doc.package.document.sections?.length ?? 0) > 0;
}

/**
 * Check if document has section properties
 */
export function hasSectionProperties(doc: Document): boolean {
  return doc.package.document.finalSectionProperties !== undefined;
}

/**
 * Get document content count (paragraphs + tables)
 */
export function getDocumentContentCount(doc: Document): number {
  return doc.package.document.content.length;
}

/**
 * Get paragraph count in document
 */
export function getDocumentParagraphCount(doc: Document): number {
  return doc.package.document.content.filter((b) => b.type === 'paragraph').length;
}

/**
 * Get table count in document
 */
export function getDocumentTableCount(doc: Document): number {
  return doc.package.document.content.filter((b) => b.type === 'table').length;
}

/**
 * Get plain text from document (for comparison/debugging)
 */
export function getDocumentPlainText(doc: Document): string {
  const texts: string[] = [];

  for (const block of doc.package.document.content) {
    if (block.type === 'paragraph') {
      for (const content of block.content) {
        if (content.type === 'run') {
          for (const item of content.content) {
            if (item.type === 'text') {
              texts.push(item.text);
            } else if (item.type === 'tab') {
              texts.push('\t');
            } else if (item.type === 'break') {
              texts.push('\n');
            }
          }
        }
      }
      texts.push('\n'); // Paragraph break
    }
  }

  return texts.join('');
}

/**
 * Create an empty document
 */
export function createEmptyDocument(): Document {
  return {
    package: {
      document: {
        content: [],
      },
    },
  };
}

/**
 * Create a simple document with text content
 */
export function createSimpleDocument(
  paragraphs: Array<{ text: string; styleId?: string }>
): Document {
  return {
    package: {
      document: {
        content: paragraphs.map((p) => ({
          type: 'paragraph' as const,
          formatting: p.styleId ? { styleId: p.styleId } : undefined,
          content: [
            {
              type: 'run' as const,
              content: [{ type: 'text' as const, text: p.text }],
            },
          ],
        })),
      },
    },
  };
}

export default serializeDocument;
