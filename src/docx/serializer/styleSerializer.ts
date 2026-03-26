/**
 * Style Serializer — converts Style objects to OOXML XML strings.
 *
 * Used selectively: only dirty/new styles are serialized from objects.
 * Unmodified styles use their _originalXml verbatim.
 */

import type { Style, StyleDefinitions } from '../../types/styles';

function escapeXml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Serialize a single Style object to a <w:style> XML element.
 */
export function serializeStyle(style: Style): string {
  const attrs = [`w:type="${escapeXml(style.type)}"`, `w:styleId="${escapeXml(style.styleId)}"`];
  if (style.default) attrs.push('w:default="1"');

  const parts: string[] = [`<w:style ${attrs.join(' ')}>`];

  if (style.name) parts.push(`<w:name w:val="${escapeXml(style.name)}"/>`);
  if (style.basedOn) parts.push(`<w:basedOn w:val="${escapeXml(style.basedOn)}"/>`);
  if (style.next) parts.push(`<w:next w:val="${escapeXml(style.next)}"/>`);
  if (style.link) parts.push(`<w:link w:val="${escapeXml(style.link)}"/>`);
  if (style.uiPriority != null) parts.push(`<w:uiPriority w:val="${style.uiPriority}"/>`);
  if (style.hidden) parts.push('<w:hidden/>');
  if (style.semiHidden) parts.push('<w:semiHidden/>');
  if (style.unhideWhenUsed) parts.push('<w:unhideWhenUsed/>');
  if (style.qFormat) parts.push('<w:qFormat/>');

  // Paragraph properties
  if (style.pPr) {
    parts.push('<w:pPr>');
    const pPr = style.pPr;
    if (pPr.alignment && pPr.alignment !== 'left') {
      parts.push(`<w:jc w:val="${escapeXml(pPr.alignment)}"/>`);
    }
    if (pPr.spaceBefore != null || pPr.spaceAfter != null || pPr.lineSpacing != null) {
      const spacingAttrs: string[] = [];
      if (pPr.spaceBefore != null) spacingAttrs.push(`w:before="${pPr.spaceBefore}"`);
      if (pPr.spaceAfter != null) spacingAttrs.push(`w:after="${pPr.spaceAfter}"`);
      if (pPr.lineSpacing != null) spacingAttrs.push(`w:line="${pPr.lineSpacing}"`);
      if (pPr.lineSpacingRule) spacingAttrs.push(`w:lineRule="${pPr.lineSpacingRule}"`);
      parts.push(`<w:spacing ${spacingAttrs.join(' ')}/>`);
    }
    if (pPr.indentLeft != null || pPr.indentRight != null || pPr.indentFirstLine != null) {
      const indAttrs: string[] = [];
      if (pPr.indentLeft != null) indAttrs.push(`w:left="${pPr.indentLeft}"`);
      if (pPr.indentRight != null) indAttrs.push(`w:right="${pPr.indentRight}"`);
      if (pPr.indentFirstLine != null) indAttrs.push(`w:firstLine="${pPr.indentFirstLine}"`);
      parts.push(`<w:ind ${indAttrs.join(' ')}/>`);
    }
    parts.push('</w:pPr>');
  }

  // Run properties
  if (style.rPr) {
    parts.push('<w:rPr>');
    const rPr = style.rPr;
    if (rPr.fontFamily?.ascii) {
      const fontAttrs = [`w:ascii="${escapeXml(rPr.fontFamily.ascii)}"`];
      if (rPr.fontFamily.hAnsi) fontAttrs.push(`w:hAnsi="${escapeXml(rPr.fontFamily.hAnsi)}"`);
      if (rPr.fontFamily.cs) fontAttrs.push(`w:cs="${escapeXml(rPr.fontFamily.cs)}"`);
      parts.push(`<w:rFonts ${fontAttrs.join(' ')}/>`);
    }
    if (rPr.bold) parts.push('<w:b/>');
    if (rPr.italic) parts.push('<w:i/>');
    if (rPr.underline) {
      const ulStyle =
        typeof rPr.underline === 'object' ? rPr.underline.style || 'single' : 'single';
      parts.push(`<w:u w:val="${ulStyle}"/>`);
    }
    if (rPr.strike) parts.push('<w:strike/>');
    if (rPr.fontSize) parts.push(`<w:sz w:val="${rPr.fontSize}"/>`);
    if (rPr.color?.rgb) parts.push(`<w:color w:val="${escapeXml(rPr.color.rgb)}"/>`);
    parts.push('</w:rPr>');
  }

  parts.push('</w:style>');
  return parts.join('');
}

/**
 * Reconstruct word/styles.xml with selective serialization.
 *
 * - Unmodified styles: emit _originalXml verbatim
 * - Dirty styles: serialize from Style object
 * - New styles (no _originalXml): serialize from scratch
 */
export function serializeStyleDefinitions(styleDefs: StyleDefinitions): string {
  const parts: string[] = [];

  // Use preserved preamble (everything before first <w:style>)
  if (styleDefs._preamble) {
    parts.push(styleDefs._preamble);
  } else {
    // Fallback: minimal styles.xml header
    parts.push(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
        ' mc:Ignorable="w14 w15">'
    );
  }

  // Emit each style
  for (const style of styleDefs.styles) {
    if (style._dirty || !style._originalXml) {
      // Serialize from object
      parts.push(serializeStyle(style));
    } else {
      // Preserve original XML verbatim
      parts.push(style._originalXml);
    }
  }

  // Use preserved postamble (closing tag)
  if (styleDefs._postamble) {
    parts.push(styleDefs._postamble);
  } else {
    parts.push('</w:styles>');
  }

  return parts.join('');
}
