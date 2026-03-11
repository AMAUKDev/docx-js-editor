/**
 * ProseMirror to FlowBlock Converter
 *
 * Converts a ProseMirror document into FlowBlock[] for the layout engine.
 * Tracks pmStart/pmEnd positions for click-to-position mapping.
 */

import type { Node as PMNode, Mark } from 'prosemirror-model';
import type {
  FlowBlock,
  ParagraphBlock,
  TableBlock,
  TableRow,
  TableCell,
  CellBorders,
  CellBorderSpec,
  ImageBlock,
  PageBreakBlock,
  Run,
  TextRun,
  TabRun,
  ImageRun,
  LineBreakRun,
  FieldRun,
  RunFormatting,
  ParagraphAttrs,
} from '../layout-engine/types';
import type { ParagraphAttrs as PMParagraphAttrs } from '../prosemirror/schema/nodes';
import type {
  TextColorAttrs,
  UnderlineAttrs,
  FontSizeAttrs,
  FontFamilyAttrs,
} from '../prosemirror/schema/marks';
import type { Theme } from '../types/document';
import type { NumberingMap } from '../docx/numberingParser';
import { resolveColor } from '../utils/colorResolver';

/**
 * Options for the conversion.
 */
export type ToFlowBlocksOptions = {
  /** Default font family. */
  defaultFont?: string;
  /** Default font size in points. */
  defaultSize?: number;
  /** Theme for resolving theme colors. */
  theme?: Theme | null;
  /** Page content height in pixels (pageHeight - marginTop - marginBottom). Images taller than this are scaled down to fit. */
  pageContentHeight?: number;
  /** Page content width in pixels (pageWidth - marginLeft - marginRight). Used to cap shape widths. */
  pageContentWidth?: number;
  /** Numbering map for resolving list markers with correct numFmt and lvlText. */
  numberingMap?: NumberingMap | null;
  /**
   * Display mode for template elements.
   * - 'rendered' (default): context tags show label or {tagKey}, loop blocks show styled markers
   * - 'raw': context tags show raw {context.key} text, loop blocks show {% for %} / {% endfor %}
   */
  renderMode?: 'rendered' | 'raw';
  /**
   * Resolved loop preview data for rendered mode expansion.
   * Keys are array names (e.g. "photos"), values are arrays of items with resolved fields.
   * Image fields are objects with {url, name}; other fields are primitive values.
   */
  loopPreviewData?: Record<string, Array<Record<string, unknown>>>;
};

const DEFAULT_FONT = 'Calibri';

/**
 * Constrain image dimensions to fit within the page content area.
 * Scales proportionally if height exceeds pageContentHeight.
 */
function constrainImageToPage(
  width: number,
  height: number,
  pageContentHeight: number | undefined
): { width: number; height: number } {
  if (!pageContentHeight || height <= pageContentHeight) {
    return { width, height };
  }
  const scale = pageContentHeight / height;
  return { width: Math.round(width * scale), height: pageContentHeight };
}

const DEFAULT_SIZE = 11; // points (Word 2007+ default)

/**
 * Convert twips to pixels (1 twip = 1/1440 inch, 1 inch = 96 CSS px).
 * No rounding — precision prevents cumulative layout drift across paragraphs.
 */
function twipsToPixels(twips: number): number {
  return (twips / 1440) * 96;
}

/**
 * Generate a unique block ID.
 */
let blockIdCounter = 0;
function nextBlockId(): string {
  return `block-${++blockIdCounter}`;
}

/**
 * Convert number to Roman numerals.
 */
function toRoman(num: number): string {
  const romanNumerals: [number, string][] = [
    [1000, 'M'],
    [900, 'CM'],
    [500, 'D'],
    [400, 'CD'],
    [100, 'C'],
    [90, 'XC'],
    [50, 'L'],
    [40, 'XL'],
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I'],
  ];
  let result = '';
  for (const [value, symbol] of romanNumerals) {
    while (num >= value) {
      result += symbol;
      num -= value;
    }
  }
  return result;
}

/**
 * Format a number according to OOXML number format (numFmt).
 */
export function formatNumber(value: number, numFmt: string): string {
  switch (numFmt) {
    case 'decimal':
    case 'decimalZero':
      return String(value);
    case 'lowerLetter':
      return String.fromCharCode(96 + ((value - 1) % 26) + 1);
    case 'upperLetter':
      return String.fromCharCode(64 + ((value - 1) % 26) + 1);
    case 'lowerRoman':
      return toRoman(value).toLowerCase();
    case 'upperRoman':
      return toRoman(value);
    case 'bullet':
      return '•';
    default:
      return String(value);
  }
}

/**
 * Format a numbered list marker using the numbering definition's lvlText template.
 * Falls back to simple decimal "1.2.3." format if no numbering map is available.
 */
export function formatNumberedMarker(
  counters: number[],
  level: number,
  numberingMap?: NumberingMap | null,
  numId?: number
): string {
  // If we have the numbering map, use the actual lvlText template
  if (numberingMap && numId != null) {
    const levelDef = numberingMap.getLevel(numId, level);
    if (levelDef?.lvlText) {
      let marker = levelDef.lvlText;
      // Replace %1, %2, etc. with formatted counter values
      for (let lvl = 0; lvl <= level; lvl++) {
        const placeholder = `%${lvl + 1}`;
        if (marker.includes(placeholder)) {
          const value = counters[lvl] ?? 0;
          const lvlDef = numberingMap.getLevel(numId, lvl);
          const formatted = formatNumber(value, lvlDef?.numFmt || 'decimal');
          marker = marker.replace(placeholder, formatted);
        }
      }
      return marker;
    }
  }

  // Fallback: simple decimal format
  const parts: number[] = [];
  for (let i = 0; i <= level; i += 1) {
    const value = counters[i] ?? 0;
    if (value <= 0) break;
    parts.push(value);
  }
  if (parts.length === 0) return '1.';
  return `${parts.join('.')}.`;
}

/**
 * Reset the block ID counter (useful for testing).
 */
export function resetBlockIdCounter(): void {
  blockIdCounter = 0;
}

/**
 * Extract run formatting from ProseMirror marks.
 */
function extractRunFormatting(marks: readonly Mark[], theme?: Theme | null): RunFormatting {
  const formatting: RunFormatting = {};

  for (const mark of marks) {
    switch (mark.type.name) {
      case 'bold':
        formatting.bold = true;
        break;

      case 'italic':
        formatting.italic = true;
        break;

      case 'underline': {
        const attrs = mark.attrs as UnderlineAttrs;
        if (attrs.style || attrs.color) {
          const underlineColor = attrs.color ? resolveColor(attrs.color, theme) : undefined;
          formatting.underline = {
            style: attrs.style,
            color: underlineColor,
          };
        } else {
          formatting.underline = true;
        }
        break;
      }

      case 'strike':
        formatting.strike = true;
        break;

      case 'textColor': {
        const attrs = mark.attrs as TextColorAttrs;
        if (attrs.themeColor || attrs.rgb) {
          formatting.color = resolveColor(
            {
              rgb: attrs.rgb,
              themeColor: attrs.themeColor,
              themeTint: attrs.themeTint,
              themeShade: attrs.themeShade,
            },
            theme
          );
        }
        break;
      }

      case 'highlight':
        formatting.highlight = mark.attrs.color as string;
        break;

      case 'fontSize': {
        const attrs = mark.attrs as FontSizeAttrs;
        // Convert half-points to points
        formatting.fontSize = attrs.size / 2;
        break;
      }

      case 'fontFamily': {
        const attrs = mark.attrs as FontFamilyAttrs;
        formatting.fontFamily = attrs.ascii || attrs.hAnsi;
        break;
      }

      case 'superscript':
        formatting.superscript = true;
        break;

      case 'subscript':
        formatting.subscript = true;
        break;

      case 'hyperlink': {
        const attrs = mark.attrs as { href: string; tooltip?: string };
        formatting.hyperlink = {
          href: attrs.href,
          tooltip: attrs.tooltip,
        };
        break;
      }

      case 'footnoteRef': {
        const attrs = mark.attrs as { id: string | number; noteType?: string };
        const id = typeof attrs.id === 'string' ? parseInt(attrs.id, 10) : attrs.id;
        if (attrs.noteType === 'endnote') {
          formatting.endnoteRefId = id;
        } else {
          formatting.footnoteRefId = id;
        }
        break;
      }
    }
  }

  return formatting;
}

/**
 * Convert a paragraph node to runs.
 */
function paragraphToRuns(node: PMNode, startPos: number, _options: ToFlowBlocksOptions): Run[] {
  const runs: Run[] = [];
  const offset = startPos + 1; // +1 for opening tag
  const theme = _options.theme;

  node.forEach((child, childOffset) => {
    const childPos = offset + childOffset;

    if (child.isText && child.text) {
      // Text node - create text run
      const formatting = extractRunFormatting(child.marks, theme);
      const run: TextRun = {
        kind: 'text',
        text: child.text,
        ...formatting,
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      };
      runs.push(run);
    } else if (child.type.name === 'hardBreak') {
      // Line break
      const run: LineBreakRun = {
        kind: 'lineBreak',
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      };
      runs.push(run);
    } else if (child.type.name === 'tab') {
      // Tab character
      const formatting = extractRunFormatting(child.marks, theme);
      const run: TabRun = {
        kind: 'tab',
        ...formatting,
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      };
      runs.push(run);
    } else if (child.type.name === 'image') {
      // Image within paragraph
      const attrs = child.attrs;
      const constrained = constrainImageToPage(
        (attrs.width as number) || 100,
        (attrs.height as number) || 100,
        _options.pageContentHeight
      );
      const run: ImageRun = {
        kind: 'image',
        src: attrs.src as string,
        width: constrained.width,
        height: constrained.height,
        alt: attrs.alt as string | undefined,
        transform: attrs.transform as string | undefined,
        // Preserve wrap attributes for proper rendering
        wrapType: attrs.wrapType as string | undefined,
        displayMode: attrs.displayMode as 'inline' | 'block' | 'float' | undefined,
        cssFloat: attrs.cssFloat as 'left' | 'right' | 'none' | undefined,
        distTop: attrs.distTop as number | undefined,
        distBottom: attrs.distBottom as number | undefined,
        distLeft: attrs.distLeft as number | undefined,
        distRight: attrs.distRight as number | undefined,
        // Preserve position for page-level floating image positioning
        position: attrs.position as ImageRun['position'] | undefined,
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      };
      runs.push(run);
    } else if (child.type.name === 'field') {
      // Field node — convert to FieldRun for render-time substitution
      const ft = child.attrs.fieldType as string;
      const mappedType: FieldRun['fieldType'] =
        ft === 'PAGE'
          ? 'PAGE'
          : ft === 'NUMPAGES'
            ? 'NUMPAGES'
            : ft === 'DATE'
              ? 'DATE'
              : ft === 'TIME'
                ? 'TIME'
                : 'OTHER';
      const formatting = extractRunFormatting(child.marks, theme);
      const run: FieldRun = {
        kind: 'field',
        fieldType: mappedType,
        fallback: (child.attrs.displayText as string) || '',
        ...formatting,
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      };
      runs.push(run);
    } else if (child.type.name === 'math') {
      // Math node — render as plain text fallback in layout
      const text = (child.attrs.plainText as string) || '[equation]';
      const run: TextRun = {
        kind: 'text',
        text,
        italic: true,
        fontFamily: 'Cambria Math',
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      };
      runs.push(run);
    } else if (child.type.name === 'crossRef') {
      // Cross-reference — render as styled text showing the display text
      const displayText = (child.attrs.displayText as string) || '[ref]';
      const run: TextRun = {
        kind: 'text',
        text: displayText,
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      };
      runs.push(run);
    } else if (child.type.name === 'contextTag') {
      // Context tag — render as plain text showing the tag label or key.
      // Context tags carry the same marks as surrounding text.
      const tagKey = child.attrs.tagKey as string;
      const label = child.attrs.label as string;
      const renderMode = _options.renderMode;
      const displayText = renderMode === 'raw' ? `{${tagKey}}` : label || `{${tagKey}}`;
      const formatting = extractRunFormatting(child.marks, theme);
      const run: TextRun = {
        kind: 'text',
        text: displayText,
        ...formatting,
        // Mark as atomic: this run occupies exactly 1 PM unit (nodeSize=1)
        // even though it may display as multiple characters.
        isAtomicNode: true,
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      };
      runs.push(run);
    } else if (child.type.name === 'shape') {
      // Shape node — render as an inline SVG image
      const attrs = child.attrs;
      const maxW = _options.pageContentWidth || 650;
      const w = Math.min((attrs.width as number) || 100, maxW);
      const h = (attrs.height as number) || 4;
      const shapeType = (attrs.shapeType as string) || 'rect';
      const fillType = (attrs.fillType as string) || 'solid';
      const fillColor = fillType === 'none' ? 'none' : (attrs.fillColor as string) || '#ffffff';
      const strokeWidth = (attrs.outlineWidth as number) || 1;
      const strokeColor = (attrs.outlineColor as string) || '#000000';
      const strokeStyle = (attrs.outlineStyle as string) || 'solid';
      const strokeDash =
        strokeStyle === 'dashed'
          ? ' stroke-dasharray="8 4"'
          : strokeStyle === 'dotted'
            ? ' stroke-dasharray="2 2"'
            : '';

      // Build SVG path based on shape type
      let svgPath: string;
      switch (shapeType) {
        case 'line':
        case 'straightConnector1':
          svgPath = `<line x1="0" y1="${h / 2}" x2="${w}" y2="${h / 2}" />`;
          break;
        case 'ellipse':
        case 'oval':
          svgPath = `<ellipse cx="${w / 2}" cy="${h / 2}" rx="${w / 2}" ry="${h / 2}" />`;
          break;
        default:
          svgPath = `<rect x="0" y="0" width="${w}" height="${h}" />`;
          break;
      }

      const svg =
        `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">` +
        `<g fill="${fillColor}" stroke="${strokeColor}" stroke-width="${strokeWidth}"${strokeDash}>` +
        svgPath +
        `</g></svg>`;

      const dataUrl = `data:image/svg+xml,${encodeURIComponent(svg)}`;
      const run: ImageRun = {
        kind: 'image',
        src: dataUrl,
        width: w,
        height: Math.max(h, strokeWidth + 2),
        displayMode: 'inline',
        pmStart: childPos,
        pmEnd: childPos + child.nodeSize,
      };
      runs.push(run);
    } else if (child.type.name === 'sdt') {
      // SDT (Structured Document Tag / content control) — inline wrapper node.
      // Descend into its children to extract the actual text runs.
      const sdtInnerOffset = childPos + 1; // +1 for opening tag
      child.forEach((sdtChild, sdtChildOffset) => {
        const sdtChildPos = sdtInnerOffset + sdtChildOffset;
        if (sdtChild.isText && sdtChild.text) {
          const formatting = extractRunFormatting(sdtChild.marks, theme);
          const run: TextRun = {
            kind: 'text',
            text: sdtChild.text,
            ...formatting,
            pmStart: sdtChildPos,
            pmEnd: sdtChildPos + sdtChild.nodeSize,
          };
          runs.push(run);
        } else if (sdtChild.type.name === 'hardBreak') {
          const run: LineBreakRun = {
            kind: 'lineBreak',
            pmStart: sdtChildPos,
            pmEnd: sdtChildPos + sdtChild.nodeSize,
          };
          runs.push(run);
        } else if (sdtChild.type.name === 'tab') {
          const formatting = extractRunFormatting(sdtChild.marks, theme);
          const run: TabRun = {
            kind: 'tab',
            ...formatting,
            pmStart: sdtChildPos,
            pmEnd: sdtChildPos + sdtChild.nodeSize,
          };
          runs.push(run);
        } else if (sdtChild.type.name === 'image') {
          const attrs = sdtChild.attrs;
          const sdtConstrained = constrainImageToPage(
            (attrs.width as number) || 100,
            (attrs.height as number) || 100,
            _options.pageContentHeight
          );
          const run: ImageRun = {
            kind: 'image',
            src: attrs.src as string,
            width: sdtConstrained.width,
            height: sdtConstrained.height,
            alt: attrs.alt as string | undefined,
            transform: attrs.transform as string | undefined,
            wrapType: attrs.wrapType as string | undefined,
            displayMode: attrs.displayMode as 'inline' | 'block' | 'float' | undefined,
            cssFloat: attrs.cssFloat as 'left' | 'right' | 'none' | undefined,
            distTop: attrs.distTop as number | undefined,
            distBottom: attrs.distBottom as number | undefined,
            distLeft: attrs.distLeft as number | undefined,
            distRight: attrs.distRight as number | undefined,
            position: attrs.position as ImageRun['position'] | undefined,
            pmStart: sdtChildPos,
            pmEnd: sdtChildPos + sdtChild.nodeSize,
          };
          runs.push(run);
        } else if (sdtChild.type.name === 'contextTag') {
          // Context tag inside SDT wrapper — render as plain text
          const ctTagKey = sdtChild.attrs.tagKey as string;
          const ctLabel = sdtChild.attrs.label as string;
          const ctDisplayText = ctLabel || `{${ctTagKey}}`;
          const ctFormatting = extractRunFormatting(sdtChild.marks, theme);
          const run: TextRun = {
            kind: 'text',
            text: ctDisplayText,
            ...ctFormatting,
            pmStart: sdtChildPos,
            pmEnd: sdtChildPos + sdtChild.nodeSize,
          };
          runs.push(run);
        }
      });
    }
  });

  return runs;
}

/**
 * Convert PM paragraph attrs to layout engine paragraph attrs.
 */
function convertParagraphAttrs(pmAttrs: PMParagraphAttrs): ParagraphAttrs {
  const attrs: ParagraphAttrs = {};

  // Alignment - map DOCX values to CSS-compatible values
  // DOCX uses 'both' for justify, 'distribute' for distributed justify
  if (pmAttrs.alignment) {
    const align = pmAttrs.alignment;
    if (align === 'both' || align === 'distribute') {
      attrs.alignment = 'justify';
    } else if (align === 'left') {
      attrs.alignment = 'left';
    } else if (align === 'center') {
      attrs.alignment = 'center';
    } else if (align === 'right') {
      attrs.alignment = 'right';
    }
    // Other DOCX alignments (mediumKashida, highKashida, lowKashida, thaiDistribute, justify)
    // default to no alignment set (inherits from style or defaults to left)
  }

  // Spacing
  if (pmAttrs.spaceBefore != null || pmAttrs.spaceAfter != null || pmAttrs.lineSpacing != null) {
    attrs.spacing = {};
    if (pmAttrs.spaceBefore != null) {
      attrs.spacing.before = twipsToPixels(pmAttrs.spaceBefore);
    }
    if (pmAttrs.spaceAfter != null) {
      attrs.spacing.after = twipsToPixels(pmAttrs.spaceAfter);
    }
    if (pmAttrs.lineSpacing != null) {
      // Line spacing in twips - convert to multiplier or exact
      if (pmAttrs.lineSpacingRule === 'exact' || pmAttrs.lineSpacingRule === 'atLeast') {
        attrs.spacing.line = twipsToPixels(pmAttrs.lineSpacing);
        attrs.spacing.lineUnit = 'px';
        attrs.spacing.lineRule = pmAttrs.lineSpacingRule;
      } else {
        // Auto - line spacing is in 240ths of a line
        attrs.spacing.line = pmAttrs.lineSpacing / 240;
        attrs.spacing.lineUnit = 'multiplier';
        attrs.spacing.lineRule = 'auto';
      }
    }
  }

  // Indentation - handle list item fallback calculation
  // For list items without explicit indentation, calculate based on level
  let indentLeft = pmAttrs.indentLeft;
  let indentFirstLine = pmAttrs.indentFirstLine;
  let hangingIndent = pmAttrs.hangingIndent;
  if (pmAttrs.numPr?.numId && indentLeft == null) {
    // Fallback: calculate indentation based on level
    // Each level indents 0.5 inch (720 twips) more
    const level = pmAttrs.numPr.ilvl ?? 0;
    // Base indentation: 0.5 inch (720 twips) per level
    // Level 0 = 720 twips, Level 1 = 1440 twips, etc.
    indentLeft = (level + 1) * 720;
    // Default hanging indent of 360 twips for the list marker
    if (indentFirstLine == null) {
      indentFirstLine = -360;
      hangingIndent = true;
    }
  }

  if (indentLeft != null || pmAttrs.indentRight != null || indentFirstLine != null) {
    attrs.indent = {};
    if (indentLeft != null) {
      attrs.indent.left = twipsToPixels(indentLeft);
    }
    if (pmAttrs.indentRight != null) {
      attrs.indent.right = twipsToPixels(pmAttrs.indentRight);
    }
    if (indentFirstLine != null) {
      if (hangingIndent) {
        // Hanging indent: indentFirstLine is stored as negative, convert to positive for rendering
        attrs.indent.hanging = Math.abs(twipsToPixels(indentFirstLine));
      } else {
        attrs.indent.firstLine = twipsToPixels(indentFirstLine);
      }
    }
  }

  // Style ID
  if (pmAttrs.styleId) {
    attrs.styleId = pmAttrs.styleId;
  }

  // Borders
  if (pmAttrs.borders) {
    const borders = pmAttrs.borders;
    attrs.borders = {};

    const convertBorder = (border: typeof borders.top) => {
      if (!border || border.style === 'none' || border.style === 'nil') {
        return undefined;
      }
      // Convert size from eighths of a point to pixels
      // 1 point = 1.333px at 96 DPI, size is in eighths of a point
      const widthPx = border.size ? Math.max(1, Math.round((border.size / 8) * 1.333)) : 1;
      // Convert color
      let color = '#000000';
      if (border.color?.rgb) {
        color = `#${border.color.rgb}`;
      }
      return {
        style: border.style || 'single',
        width: widthPx,
        color,
      };
    };

    if (borders.top) attrs.borders.top = convertBorder(borders.top);
    if (borders.bottom) attrs.borders.bottom = convertBorder(borders.bottom);
    if (borders.left) attrs.borders.left = convertBorder(borders.left);
    if (borders.right) attrs.borders.right = convertBorder(borders.right);
    if (borders.between) attrs.borders.between = convertBorder(borders.between);

    // Only include if at least one border is set
    if (
      !attrs.borders.top &&
      !attrs.borders.bottom &&
      !attrs.borders.left &&
      !attrs.borders.right &&
      !attrs.borders.between
    ) {
      delete attrs.borders;
    }
  }

  // Shading (background color)
  if (pmAttrs.shading?.fill?.rgb) {
    attrs.shading = `#${pmAttrs.shading.fill.rgb}`;
  }

  // Tab stops
  if (pmAttrs.tabs && pmAttrs.tabs.length > 0) {
    attrs.tabs = pmAttrs.tabs.map((tab) => ({
      val: mapTabAlignment(tab.alignment),
      pos: tab.position,
      leader: tab.leader as
        | 'none'
        | 'dot'
        | 'hyphen'
        | 'underscore'
        | 'heavy'
        | 'middleDot'
        | undefined,
    }));
  }

  // Page break control
  if (pmAttrs.pageBreakBefore) {
    attrs.pageBreakBefore = true;
  }
  if (pmAttrs.keepNext) {
    attrs.keepNext = true;
  }
  if (pmAttrs.keepLines) {
    attrs.keepLines = true;
  }
  if (pmAttrs.contextualSpacing) {
    attrs.contextualSpacing = true;
  }
  if (pmAttrs.styleId) {
    attrs.styleId = pmAttrs.styleId;
  }

  // List properties
  if (pmAttrs.numPr) {
    attrs.numPr = {
      numId: pmAttrs.numPr.numId,
      ilvl: pmAttrs.numPr.ilvl,
    };
  }
  if (pmAttrs.listMarker) {
    attrs.listMarker = pmAttrs.listMarker;
  }
  if (pmAttrs.listIsBullet != null) {
    attrs.listIsBullet = pmAttrs.listIsBullet;
  }

  // Default font for empty paragraph measurement (from style's rPr / pPr/rPr)
  const dtf = pmAttrs.defaultTextFormatting as
    | { fontSize?: number; fontFamily?: { ascii?: string; hAnsi?: string } }
    | undefined;
  if (dtf) {
    if (dtf.fontSize != null) {
      // fontSize in TextFormatting is in half-points, convert to points
      attrs.defaultFontSize = dtf.fontSize / 2;
    }
    if (dtf.fontFamily) {
      attrs.defaultFontFamily = (dtf.fontFamily.ascii || dtf.fontFamily.hAnsi) as
        | string
        | undefined;
    }
  }

  // Lock state for selective editing
  if (pmAttrs.locked) {
    attrs.locked = true;
  }

  return attrs;
}

/**
 * Map document TabStopAlignment to layout engine TabAlignment
 */
function mapTabAlignment(
  align: 'left' | 'center' | 'right' | 'decimal' | 'bar' | 'clear' | 'num'
): 'start' | 'end' | 'center' | 'decimal' | 'bar' | 'clear' {
  switch (align) {
    case 'left':
      return 'start';
    case 'right':
      return 'end';
    case 'center':
      return 'center';
    case 'decimal':
      return 'decimal';
    case 'bar':
      return 'bar';
    case 'clear':
      return 'clear';
    case 'num':
      return 'start'; // Number tab treated as left-aligned
    default:
      return 'start';
  }
}

/**
 * Convert a paragraph node to a ParagraphBlock.
 */
function convertParagraph(
  node: PMNode,
  startPos: number,
  options: ToFlowBlocksOptions
): ParagraphBlock {
  const pmAttrs = node.attrs as PMParagraphAttrs;
  const runs = paragraphToRuns(node, startPos, options);
  const attrs = convertParagraphAttrs(pmAttrs);

  return {
    kind: 'paragraph',
    id: nextBlockId(),
    runs,
    attrs,
    pmStart: startPos,
    pmEnd: startPos + node.nodeSize,
  };
}

/**
 * Convert border width from eighths of a point to pixels.
 * OOXML stores border widths in eighths of a point.
 */
function borderWidthToPixels(eighthsOfPoint: number): number {
  // 1 point = 1.333 pixels at 96 DPI
  // eighths of a point: divide by 8 first
  return Math.max(1, Math.round((eighthsOfPoint / 8) * 1.333));
}

// OOXML border style → CSS border-style mapping
const OOXML_TO_CSS_BORDER: Record<string, string> = {
  single: 'solid',
  double: 'double',
  dotted: 'dotted',
  dashed: 'dashed',
  thick: 'solid',
  dashSmallGap: 'dashed',
  dotDash: 'dashed',
  dotDotDash: 'dotted',
  triple: 'double',
  wave: 'solid',
  doubleWave: 'double',
  threeDEmboss: 'ridge',
  threeDEngrave: 'groove',
  outset: 'outset',
  inset: 'inset',
};

/**
 * Extract cell borders from ProseMirror attributes.
 * Borders are full BorderSpec objects with style/size/color.
 */
function extractCellBorders(attrs: Record<string, unknown>): CellBorders | undefined {
  const borders = attrs.borders as Record<
    string,
    { style?: string; size?: number; color?: { rgb?: string } }
  > | null;

  if (!borders) {
    return undefined;
  }

  const result: CellBorders = {};
  const sides = ['top', 'bottom', 'left', 'right'] as const;

  for (const side of sides) {
    const border = borders[side];
    if (!border || !border.style || border.style === 'none' || border.style === 'nil') {
      result[side] = { width: 0, style: 'none' };
      continue;
    }

    const spec: CellBorderSpec = {
      style: OOXML_TO_CSS_BORDER[border.style] || 'solid',
    };
    if (border.color?.rgb) {
      spec.color = `#${border.color.rgb}`;
    }
    if (border.size) {
      spec.width = borderWidthToPixels(border.size);
    }
    result[side] = spec;
  }

  return Object.keys(result).length > 0 ? result : undefined;
}

/**
 * Convert a table cell node.
 */
function convertTableCell(node: PMNode, startPos: number, options: ToFlowBlocksOptions): TableCell {
  const blocks: FlowBlock[] = [];
  let offset = startPos + 1; // +1 for opening tag

  node.forEach((child) => {
    if (child.type.name === 'paragraph') {
      blocks.push(convertParagraph(child, offset, options));
    } else if (child.type.name === 'table') {
      blocks.push(convertTable(child, offset, options));
    }
    offset += child.nodeSize;
  });

  const attrs = node.attrs;

  // Convert cell margins (twips) to pixel padding
  // OOXML spec: 0 top/bottom, 108 twips (~7px) left/right
  // We add 1px top/bottom to match Word's visual rendering (internal leading)
  const margins = attrs.margins as
    | { top?: number; bottom?: number; left?: number; right?: number }
    | undefined;
  const padding = {
    top: margins?.top != null ? twipsToPixels(margins.top) : 1,
    right: margins?.right != null ? twipsToPixels(margins.right) : 7,
    bottom: margins?.bottom != null ? twipsToPixels(margins.bottom) : 1,
    left: margins?.left != null ? twipsToPixels(margins.left) : 7,
  };

  // Convert cell width based on widthType
  let cellWidth: number | undefined;
  let cellWidthPct: number | undefined;
  const rawWidth = attrs.width as number | undefined;
  const cellWidthType = attrs.widthType as string | undefined;
  if (rawWidth) {
    if (cellWidthType === 'pct') {
      // Percentage width — store the raw fiftieths-of-percent value for measureTableBlock
      // to resolve against the table's pixel width. This is needed because grid column
      // mapping produces wrong widths when a row's cells don't cover all grid columns.
      cellWidthPct = rawWidth;
    } else {
      // Default: treat as twips (dxa)
      cellWidth = twipsToPixels(rawWidth);
    }
  }

  return {
    id: nextBlockId(),
    blocks,
    colSpan: attrs.colspan as number,
    rowSpan: attrs.rowspan as number,
    width: cellWidth,
    widthPct: cellWidthPct,
    verticalAlign: attrs.verticalAlign as 'top' | 'center' | 'bottom' | undefined,
    background: attrs.backgroundColor ? `#${attrs.backgroundColor}` : undefined,
    borders: extractCellBorders(attrs as Record<string, unknown>),
    padding,
  };
}

/**
 * Convert a table row node.
 */
function convertTableRow(node: PMNode, startPos: number, options: ToFlowBlocksOptions): TableRow {
  const cells: TableCell[] = [];
  let offset = startPos + 1; // +1 for opening tag

  node.forEach((child) => {
    if (child.type.name === 'tableCell' || child.type.name === 'tableHeader') {
      cells.push(convertTableCell(child, offset, options));
    }
    offset += child.nodeSize;
  });

  const attrs = node.attrs;
  return {
    id: nextBlockId(),
    cells,
    height: attrs.height ? twipsToPixels(attrs.height as number) : undefined,
    heightRule: (attrs.heightRule as 'auto' | 'atLeast' | 'exact') ?? undefined,
    isHeader: attrs.isHeader as boolean | undefined,
  };
}

/**
 * Convert a table node to a TableBlock.
 */
function convertTable(node: PMNode, startPos: number, options: ToFlowBlocksOptions): TableBlock {
  const rows: TableRow[] = [];
  let offset = startPos + 1; // +1 for opening tag

  node.forEach((child) => {
    if (child.type.name === 'tableRow') {
      rows.push(convertTableRow(child, offset, options));
    }
    offset += child.nodeSize;
  });

  // Extract columnWidths from node attributes and convert from twips to pixels
  const columnWidthsTwips = node.attrs.columnWidths as number[] | undefined;
  let columnWidths = columnWidthsTwips?.map(twipsToPixels);

  const width = node.attrs.width as number | undefined;
  const widthType = node.attrs.widthType as string | undefined;

  // Fallback: compute column widths from first row cell widths if table attr is missing
  if (!columnWidths && rows.length > 0) {
    const firstRow = rows[0];
    const cellWidths = firstRow.cells.map((cell) => cell.width);
    // Only use if all cells have widths defined
    if (cellWidths.every((w) => w !== undefined && w > 0)) {
      columnWidths = cellWidths as number[];
    }
  }

  // Extract justification
  const justification = node.attrs.justification as 'left' | 'center' | 'right' | undefined;

  const floating = node.attrs.floating as
    | {
        horzAnchor?: 'margin' | 'page' | 'text';
        vertAnchor?: 'margin' | 'page' | 'text';
        tblpX?: number;
        tblpXSpec?: 'left' | 'center' | 'right' | 'inside' | 'outside';
        tblpY?: number;
        tblpYSpec?: 'top' | 'center' | 'bottom' | 'inside' | 'outside' | 'inline';
        topFromText?: number;
        bottomFromText?: number;
        leftFromText?: number;
        rightFromText?: number;
      }
    | undefined;

  const floatingPx = floating
    ? {
        horzAnchor: floating.horzAnchor,
        vertAnchor: floating.vertAnchor,
        tblpX: floating.tblpX !== undefined ? twipsToPixels(floating.tblpX) : undefined,
        tblpXSpec: floating.tblpXSpec,
        tblpY: floating.tblpY !== undefined ? twipsToPixels(floating.tblpY) : undefined,
        tblpYSpec: floating.tblpYSpec,
        topFromText:
          floating.topFromText !== undefined ? twipsToPixels(floating.topFromText) : undefined,
        bottomFromText:
          floating.bottomFromText !== undefined
            ? twipsToPixels(floating.bottomFromText)
            : undefined,
        leftFromText:
          floating.leftFromText !== undefined ? twipsToPixels(floating.leftFromText) : undefined,
        rightFromText:
          floating.rightFromText !== undefined ? twipsToPixels(floating.rightFromText) : undefined,
      }
    : undefined;

  return {
    kind: 'table',
    id: nextBlockId(),
    rows,
    columnWidths,
    width,
    widthType,
    justification,
    floating: floatingPx,
    pmStart: startPos,
    pmEnd: startPos + node.nodeSize,
  };
}

/**
 * Convert an image node to an ImageBlock.
 */
function convertImage(node: PMNode, startPos: number, pageContentHeight?: number): ImageBlock {
  const attrs = node.attrs;
  const wrapType = attrs.wrapType as string | undefined;

  // Only anchor images with 'behind' or 'inFront' wrap types
  // Other wrap types (square, tight, through, topAndBottom) need text wrapping
  // which we don't support yet, so treat them as block-level images
  const shouldAnchor = wrapType === 'behind' || wrapType === 'inFront';

  const constrained = constrainImageToPage(
    (attrs.width as number) || 100,
    (attrs.height as number) || 100,
    pageContentHeight
  );

  return {
    kind: 'image',
    id: nextBlockId(),
    src: attrs.src as string,
    width: constrained.width,
    height: constrained.height,
    alt: attrs.alt as string | undefined,
    transform: attrs.transform as string | undefined,
    anchor: shouldAnchor
      ? {
          isAnchored: true,
          offsetH: attrs.distLeft as number | undefined,
          offsetV: attrs.distTop as number | undefined,
          behindDoc: wrapType === 'behind',
        }
      : undefined,
    pmStart: startPos,
    pmEnd: startPos + node.nodeSize,
  };
}

/**
 * Parse a loop expression like "photo in photos" into {itemVar, arrayName}.
 * Returns null if the expression doesn't match the expected pattern.
 */
function parseLoopExpr(loopExpr: string): { itemVar: string; arrayName: string } | null {
  const match = loopExpr.trim().match(/^(\w+)\s+in\s+(\w+)$/);
  if (!match) return null;
  return { itemVar: match[1], arrayName: match[2] };
}

/**
 * Substitute {{ itemVar.field }} patterns in a text string with values from a data item.
 * Returns the substituted string, or null if the text contains an image-field reference.
 */
function substituteText(text: string, itemVar: string, dataItem: Record<string, unknown>): string {
  return text.replace(/\{\{\s*(\w+)\.(\w+)\s*\}\}/g, (_match, varPart, field) => {
    if (varPart !== itemVar) return _match;
    const val = dataItem[field];
    if (val == null) return '';
    if (typeof val === 'object') return ''; // image field — handled separately
    return String(val);
  });
}

/**
 * Check if a text contains an image-field reference for itemVar.
 * e.g. "{{ photo.image }}" where dataItem.image is {url, name}
 */
function getImageFieldFromText(
  text: string,
  itemVar: string,
  dataItem: Record<string, unknown>
): { url: string; name: string } | null {
  const re = /\{\{\s*(\w+)\.(\w+)\s*\}\}/g;
  let match;
  while ((match = re.exec(text)) !== null) {
    const [, varPart, field] = match;
    if (varPart !== itemVar) continue;
    const val = dataItem[field];
    if (val && typeof val === 'object' && 'url' in (val as object)) {
      return val as { url: string; name: string };
    }
  }
  return null;
}

/**
 * Clone a FlowBlock tree with template substitutions applied for a specific loop data item.
 * Image references are replaced with actual ImageBlock / ImageRun blocks.
 */
function substituteBlocksForItem(
  templateBlocks: FlowBlock[],
  itemVar: string,
  dataItem: Record<string, unknown>,
  pageContentWidth?: number
): FlowBlock[] {
  const result: FlowBlock[] = [];

  for (const block of templateBlocks) {
    if (block.kind === 'paragraph') {
      const para = block as ParagraphBlock;
      // Check if any run contains an image field reference
      const newRuns: Run[] = [];
      let hasImageReplacement = false;

      for (const run of para.runs) {
        if (run.kind === 'text') {
          const imgField = getImageFieldFromText(run.text, itemVar, dataItem);
          if (imgField) {
            // Replace this run with an image run
            const maxW = pageContentWidth ?? 500;
            const imgRun: ImageRun = {
              kind: 'image',
              src: imgField.url,
              width: Math.min(maxW, 400),
              height: 300,
              alt: imgField.name,
              displayMode: 'inline',
              pmStart: run.pmStart,
              pmEnd: run.pmEnd,
            };
            newRuns.push(imgRun);
            hasImageReplacement = true;
          } else {
            // Regular text substitution
            const substituted = substituteText(run.text, itemVar, dataItem);
            if (substituted !== run.text) {
              newRuns.push({ ...run, text: substituted });
            } else {
              newRuns.push(run);
            }
          }
        } else {
          newRuns.push(run);
        }
      }

      void hasImageReplacement; // unused — kept for future styling
      result.push({
        ...para,
        id: nextBlockId(),
        runs: newRuns,
        attrs: {
          ...para.attrs,
        },
      });
    } else if (block.kind === 'table') {
      const tableBlock = block as TableBlock;
      // Recursively substitute within table cells
      const newRows = tableBlock.rows.map((row) => ({
        ...row,
        id: nextBlockId(),
        cells: row.cells.map((cell) => ({
          ...cell,
          id: nextBlockId(),
          blocks: substituteBlocksForItem(cell.blocks, itemVar, dataItem, pageContentWidth),
        })),
      }));
      result.push({ ...tableBlock, id: nextBlockId(), rows: newRows });
    } else {
      result.push({ ...block, id: nextBlockId() });
    }
  }

  return result;
}

/**
 * Create a subtle divider paragraph block between loop iterations.
 */
function makeLoopDivider(_iterationIndex: number): ParagraphBlock {
  return {
    kind: 'paragraph',
    id: nextBlockId(),
    runs: [],
    attrs: {
      spacing: { before: 4, after: 4 },
    },
    pmStart: -1,
    pmEnd: -1,
  };
}

/**
 * Convert a ProseMirror document to FlowBlock array.
 *
 * Walks the document tree, converting each node to the appropriate block type.
 * Tracks pmStart/pmEnd positions for each block for click-to-position mapping.
 */
export function toFlowBlocks(doc: PMNode, options: ToFlowBlocksOptions = {}): FlowBlock[] {
  const opts: ToFlowBlocksOptions = {
    ...options,
    defaultFont: options.defaultFont ?? DEFAULT_FONT,
    defaultSize: options.defaultSize ?? DEFAULT_SIZE,
  };

  const blocks: FlowBlock[] = [];
  const listCounters = new Map<number, number[]>();

  // Build an array of [node, offset] pairs so we can skip ahead for loop expansion
  const children: Array<{ node: PMNode; pos: number }> = [];
  doc.forEach((node, nodeOffset) => {
    children.push({ node, pos: nodeOffset });
  });

  let i = 0;
  while (i < children.length) {
    const { node, pos } = children[i];

    switch (node.type.name) {
      case 'paragraph':
        {
          const block = convertParagraph(node, pos, opts);
          const pmAttrs = node.attrs as PMParagraphAttrs;

          if (pmAttrs.numPr) {
            const numId = pmAttrs.numPr.numId;
            // numId === 0 means "no numbering" per OOXML spec (ECMA-376)
            if (numId != null && numId !== 0) {
              const level = pmAttrs.numPr.ilvl ?? 0;
              const counters = listCounters.get(numId) ?? new Array(9).fill(0);

              counters[level] = (counters[level] ?? 0) + 1;
              for (let lvl = level + 1; lvl < counters.length; lvl += 1) {
                counters[lvl] = 0;
              }

              listCounters.set(numId, counters);

              // Always recompute numbered markers so reordering updates them.
              // Bullet markers are static, so keep any pre-set value.
              if (pmAttrs.listIsBullet) {
                if (!pmAttrs.listMarker) {
                  block.attrs = { ...block.attrs, listMarker: '•' };
                }
              } else {
                const marker = formatNumberedMarker(counters, level, opts.numberingMap, numId);
                block.attrs = { ...block.attrs, listMarker: marker };
              }
            }
          }

          blocks.push(block);
        }
        i++;
        break;

      case 'table':
        blocks.push(convertTable(node, pos, opts));
        i++;
        break;

      case 'image':
        // Standalone image block (if not inline)
        blocks.push(convertImage(node, pos, opts.pageContentHeight));
        i++;
        break;

      case 'horizontalRule':
      case 'pageBreak': {
        const pb: PageBreakBlock = {
          kind: 'pageBreak',
          id: nextBlockId(),
          pmStart: pos,
          pmEnd: pos + node.nodeSize,
        };
        blocks.push(pb);
        i++;
        break;
      }

      case 'loopBlock': {
        const loopKind = (node.attrs.kind as 'for' | 'endfor') || 'for';
        const loopExpr = (node.attrs.loopExpr as string) || '';
        const renderMode = opts.renderMode;

        // Attempt loop expansion in rendered mode when data is available
        if (loopKind === 'for' && renderMode !== 'raw' && opts.loopPreviewData) {
          const parsed = parseLoopExpr(loopExpr);
          const dataArray = parsed ? opts.loopPreviewData[parsed.arrayName] : undefined;

          if (parsed && Array.isArray(dataArray) && dataArray.length > 0) {
            // Collect template body nodes up to matching endfor
            const templateChildren: Array<{ node: PMNode; pos: number }> = [];
            let endforIndex = -1;
            for (let j = i + 1; j < children.length; j++) {
              const child = children[j];
              if (
                child.node.type.name === 'loopBlock' &&
                (child.node.attrs.kind as string) === 'endfor'
              ) {
                endforIndex = j;
                break;
              }
              templateChildren.push(child);
            }

            if (endforIndex !== -1 && templateChildren.length > 0) {
              // Convert template nodes to FlowBlocks (without loop expansion recursion)
              const templateOpts = { ...opts, loopPreviewData: undefined };
              const templateBlocks: FlowBlock[] = [];
              for (const { node: tNode, pos: tPos } of templateChildren) {
                if (tNode.type.name === 'paragraph') {
                  templateBlocks.push(convertParagraph(tNode, tPos, templateOpts));
                } else if (tNode.type.name === 'table') {
                  templateBlocks.push(convertTable(tNode, tPos, templateOpts));
                } else if (tNode.type.name === 'image') {
                  templateBlocks.push(convertImage(tNode, tPos, templateOpts.pageContentHeight));
                }
              }

              // Expand one copy per data item
              for (let itemIdx = 0; itemIdx < dataArray.length; itemIdx++) {
                if (itemIdx > 0) {
                  blocks.push(makeLoopDivider(itemIdx));
                }
                const dataItem = dataArray[itemIdx] as Record<string, unknown>;
                const expanded = substituteBlocksForItem(
                  templateBlocks,
                  parsed.itemVar,
                  dataItem,
                  opts.pageContentWidth
                );
                for (const b of expanded) {
                  blocks.push(b);
                }
              }

              // Skip past: for-block + template nodes + endfor-block
              i = endforIndex + 1;
              break;
            }
          }
        }

        // Fall-through: render as loop block pill (raw mode or no data)
        const displayText =
          renderMode === 'raw'
            ? loopKind === 'for'
              ? `{% for ${loopExpr} %}`
              : '{% endfor %}'
            : loopKind === 'for'
              ? `for ${loopExpr}`
              : 'end for';

        const loopRun: TextRun = {
          kind: 'text',
          text: displayText,
          isAtomicNode: true,
          pmStart: pos,
          pmEnd: pos + node.nodeSize,
        };

        const loopPillBlock: ParagraphBlock = {
          kind: 'paragraph',
          id: nextBlockId(),
          runs: [loopRun],
          attrs: {
            isLoopBlock: true,
            loopKind,
            loopExpr,
          },
          pmStart: pos,
          pmEnd: pos + node.nodeSize,
        };
        blocks.push(loopPillBlock);
        i++;
        break;
      }

      default:
        i++;
        break;
    }
  }

  return blocks;
}
