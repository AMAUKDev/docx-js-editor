/**
 * Parse a raw <w:style> XML string to extract rPr and pPr for live rendering.
 * Used by importStyles to provide both _originalXml (for saving) AND
 * parsed formatting properties (for the editor to apply styles).
 *
 * Handles:
 * - Explicit fonts + theme font references (asciiTheme, hAnsiTheme)
 * - Auto color (w:val="auto") vs explicit RGB vs theme colors
 * - Bold/italic with explicit false (w:val="false" or w:val="0")
 * - Paragraph spacing and alignment
 */
export function parseImportedStyleXml(xml: string): Record<string, any> {
  const result: Record<string, any> = {};

  // ── Run properties (rPr) ──────────────────────────────────────
  const rPr: Record<string, any> = {};

  // Font families — explicit names + theme references
  const rFontsMatch = xml.match(/<w:rFonts\b([^/]*?)\/?>/);
  if (rFontsMatch) {
    const attrs = rFontsMatch[1];
    const fm: Record<string, string> = {};
    const ascii = attrs.match(/w:ascii="([^"]*)"/);
    const hAnsi = attrs.match(/w:hAnsi="([^"]*)"/);
    const asciiTheme = attrs.match(/w:asciiTheme="([^"]*)"/);
    const hAnsiTheme = attrs.match(/w:hAnsiTheme="([^"]*)"/);
    if (ascii) fm.ascii = ascii[1];
    if (hAnsi) fm.hAnsi = hAnsi[1];
    if (asciiTheme) fm.asciiTheme = asciiTheme[1];
    if (hAnsiTheme) fm.hAnsiTheme = hAnsiTheme[1];
    if (Object.keys(fm).length > 0) rPr.fontFamily = fm;
  }

  // Font size (half-points)
  const sizeMatch = xml.match(/<w:sz w:val="(\d+)"/);
  if (sizeMatch) rPr.fontSize = parseInt(sizeMatch[1], 10);

  // Bold — handle w:val="false" and w:val="0" as explicit bold-off
  const boldMatch = xml.match(/<w:b(\s[^/]*?)?\/?>/);
  if (boldMatch) {
    const attrs = boldMatch[1] || '';
    const valMatch = attrs.match(/w:val="([^"]*)"/);
    if (valMatch && (valMatch[1] === 'false' || valMatch[1] === '0')) {
      rPr.bold = false;
    } else {
      rPr.bold = true;
    }
  }

  // Italic — same false/0 handling
  const italicMatch = xml.match(/<w:i(\s[^/]*?)?\/?>/);
  if (italicMatch) {
    const attrs = italicMatch[1] || '';
    const valMatch = attrs.match(/w:val="([^"]*)"/);
    if (valMatch && (valMatch[1] === 'false' || valMatch[1] === '0')) {
      rPr.italic = false;
    } else {
      rPr.italic = true;
    }
  }

  // Color — auto, explicit RGB, theme color
  const colorMatch = xml.match(/<w:color\b([^/]*?)\/?>/);
  if (colorMatch) {
    const attrs = colorMatch[1];
    const valMatch = attrs.match(/w:val="([^"]*)"/);
    const themeColorMatch = attrs.match(/w:themeColor="([^"]*)"/);
    const themeShadeMatch = attrs.match(/w:themeShade="([^"]*)"/);
    const themeTintMatch = attrs.match(/w:themeTint="([^"]*)"/);

    const color: Record<string, any> = {};
    if (valMatch) {
      if (valMatch[1] === 'auto') {
        color.auto = true;
      } else {
        color.rgb = valMatch[1];
      }
    }
    if (themeColorMatch) color.themeColor = themeColorMatch[1];
    if (themeShadeMatch) color.themeShade = themeShadeMatch[1];
    if (themeTintMatch) color.themeTint = themeTintMatch[1];
    if (Object.keys(color).length > 0) rPr.color = color;
  }

  if (Object.keys(rPr).length > 0) result.rPr = rPr;

  // ── Paragraph properties (pPr) ────────────────────────────────
  const pPr: Record<string, any> = {};

  const spacingMatch = xml.match(/<w:spacing\b([^/]*?)\/?>/);
  if (spacingMatch) {
    const attrs = spacingMatch[1];
    const spacing: Record<string, number> = {};
    const before = attrs.match(/w:before="(\d+)"/);
    const after = attrs.match(/w:after="(\d+)"/);
    const line = attrs.match(/w:line="(\d+)"/);
    if (before) spacing.before = parseInt(before[1], 10);
    if (after) spacing.after = parseInt(after[1], 10);
    if (line) spacing.line = parseInt(line[1], 10);
    if (Object.keys(spacing).length > 0) pPr.spacing = spacing;
  }

  const jcMatch = xml.match(/<w:jc w:val="([^"]+)"/);
  if (jcMatch) pPr.alignment = jcMatch[1];

  if (Object.keys(pPr).length > 0) result.pPr = pPr;

  return result;
}
