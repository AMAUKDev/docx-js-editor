/**
 * Cross-Reference Updater Plugin
 *
 * An appendTransaction plugin that keeps cross-ref displayText and
 * caption figure/table numbers in sync whenever the document changes.
 *
 * Scans the document for:
 * - Headings: paragraphs with styleId matching Heading\d, tracking numbering counters
 * - Captions: paragraphs with styleId === 'Caption', tracking figure/table counts
 * - CrossRef nodes: inline crossRef atoms, checking if displayText is stale
 * - Field nodes (SEQ): inside captions, update displayText to correct number
 * - Table of Contents: TOCHeading + TOC1-9 paragraphs, refreshed on demand
 *
 * When a mismatch is found, dispatches a corrective transaction.
 */

import { Plugin, PluginKey, type Transaction } from 'prosemirror-state';
import type { EditorState } from 'prosemirror-state';
import { Fragment as PMFragment } from 'prosemirror-model';
import type { Mark, Node as PMNode, Schema } from 'prosemirror-model';
import type { NumberingMap } from '../../docx/numberingParser';
import { formatNumberedMarker } from '../../layout-bridge/toFlowBlocks';
import type { Layout } from '../../layout-engine/types';
import { textFormattingToMarks } from '../extensions/marks/markUtils';

export const crossRefUpdaterKey = new PluginKey('crossRefUpdater');

export interface CrossRefUpdaterConfig {
  /** Return the current numbering map (may change over time). */
  getNumberingMap?: () => NumberingMap | null;
  /** Return a style resolver for looking up numPr and spacing from style definitions. */
  getStyleResolver?: () => {
    resolveParagraphStyle(styleId: string): {
      paragraphFormatting?: {
        numPr?: { numId?: number; ilvl?: number };
        spaceBefore?: number;
        spaceAfter?: number;
        lineSpacing?: number;
        lineSpacingRule?: string;
      };
      runFormatting?: import('../../types/document').TextFormatting;
    };
  } | null;
  /** Return the current page layout (for TOC page numbers). */
  getLayout?: () => Layout | null;
  /** When set, TOC entries inherit font/size/spacing from this style instead of per-level TOC1-9 styles. */
  tocStyleOverride?: string;
}

/** Recognised caption prefixes and their sequence counters */
const CAPTION_PREFIXES = ['Figure', 'Table'] as const;
type CaptionPrefix = (typeof CAPTION_PREFIXES)[number];

/**
 * Detect which caption prefix a paragraph starts with.
 * Works for both plain-text captions ("Figure 1: ...") and
 * field-based captions ("Figure " + SEQ_FIELD + ": ...").
 */
function detectCaptionPrefix(node: PMNode): CaptionPrefix | null {
  const text = node.textContent;
  for (const prefix of CAPTION_PREFIXES) {
    if (text.startsWith(prefix + ' ')) return prefix;
  }
  return null;
}

interface CaptionInfo {
  /** The paragraph position in the document */
  pos: number;
  /** The correct sequential number for this caption */
  correctNumber: number;
  /** "Figure" or "Table" */
  prefix: CaptionPrefix;
}

/** Heading entry with all info needed for TOC refresh. */
interface HeadingEntry {
  text: string;
  number: string;
  level: number;
  pmPos: number;
  /** Existing _Toc bookmark name on the heading, if any. */
  bookmarkName: string | null;
}

// =============================================================================
// TOC helpers
// =============================================================================

/**
 * Check whether a paragraph contains a hyperlink mark pointing to a _Toc bookmark.
 * This identifies TOC entry paragraphs even when their styleId has been
 * degraded to "Normal" (e.g. by the StyleEnforcerPlugin).
 */
function hasTocHyperlink(node: PMNode): boolean {
  let found = false;
  node.descendants((child) => {
    if (found) return false;
    if (child.isText) {
      for (const mark of child.marks) {
        if (mark.type.name === 'hyperlink') {
          const href = mark.attrs.href as string | undefined;
          if (href && href.startsWith('#_Toc')) {
            found = true;
            return false;
          }
        }
      }
    }
    return true;
  });
  return found;
}

/**
 * Find the TOC block boundaries in the document.
 * A TOC block is a heading paragraph followed by consecutive `TOC1`-`TOC9` paragraphs.
 * The heading can be styled as `TOCHeading` or be a regular heading (e.g. Heading2)
 * immediately before the first TOC entry.
 *
 * Fallback: if TOC entries have lost their `TOC\d` styleId (e.g. degraded to "Normal"
 * by the StyleEnforcerPlugin), they are detected by having hyperlink marks
 * pointing to `#_Toc...` bookmarks.
 */
function findTocBlock(
  doc: PMNode
): { startPos: number; endPos: number; maxTocLevel: number; existingLeader: string } | null {
  let startPos: number | null = null;
  let endPos: number | null = null;
  let maxTocLevel = 0;
  let inTocBlock = false;
  // Track previous paragraph so we can include a non-TOCHeading heading
  let prevPos: number | null = null;
  // Track how many TOC entries we found via hyperlink fallback (no explicit TOC style)
  let fallbackEntryCount = 0;
  // Capture the leader style from the first TOC entry's right-aligned tab
  let existingLeader = 'none';

  doc.forEach((node, offset) => {
    if (node.type.name === 'paragraph') {
      const styleId = node.attrs.styleId as string | null;

      if (styleId === 'TOCHeading' && startPos === null) {
        startPos = offset;
        endPos = offset + node.nodeSize;
        inTocBlock = true;
        prevPos = offset;
        return;
      }

      const tocMatch = styleId?.match(/^TOC(\d)$/);
      if (tocMatch) {
        const level = parseInt(tocMatch[1]);
        maxTocLevel = Math.max(maxTocLevel, level);
        if (!inTocBlock) {
          // Start of TOC block — include the previous paragraph as the heading
          if (startPos === null) {
            startPos = prevPos ?? offset;
          }
          inTocBlock = true;
          // Capture leader from the first TOC entry's tabs
          const tabs = node.attrs.tabs as Array<{ alignment?: string; leader?: string }> | null;
          if (tabs) {
            const rightTab = tabs.find((t) => t.alignment === 'right');
            if (rightTab?.leader) {
              existingLeader = rightTab.leader;
            }
          }
        }
        endPos = offset + node.nodeSize;
        prevPos = offset;
        return;
      }

      // Fallback: detect TOC entries by _Toc hyperlinks (for degraded styles)
      if (!inTocBlock && hasTocHyperlink(node)) {
        // First TOC entry found via fallback — include previous paragraph as heading
        if (startPos === null) {
          startPos = prevPos ?? offset;
        }
        inTocBlock = true;
        endPos = offset + node.nodeSize;
        fallbackEntryCount++;
        prevPos = offset;
        return;
      }

      if (inTocBlock && hasTocHyperlink(node)) {
        endPos = offset + node.nodeSize;
        fallbackEntryCount++;
        prevPos = offset;
        return;
      }

      if (inTocBlock) {
        inTocBlock = false;
      }
      prevPos = offset;
    } else if (inTocBlock) {
      inTocBlock = false;
    }
  });

  // If we only found entries via fallback (no TOC\d styles), assume TOC1 depth
  if (maxTocLevel === 0 && fallbackEntryCount > 0) {
    maxTocLevel = 1;
  }

  if (startPos !== null && endPos !== null && endPos > startPos && maxTocLevel > 0) {
    return { startPos, endPos, maxTocLevel, existingLeader };
  }
  return null;
}

/**
 * Get the page number for a PM position using the layout.
 */
function getPageForPmPos(layout: Layout, pmPos: number): number {
  for (const page of layout.pages) {
    for (const fragment of page.fragments) {
      if (fragment.pmStart != null && fragment.pmEnd != null) {
        if (pmPos >= fragment.pmStart && pmPos <= fragment.pmEnd) {
          return page.number;
        }
      }
    }
  }
  return 1; // fallback to page 1
}

/**
 * Build a map of refTarget → currentNumber by scanning the document.
 * Also returns caption number mappings and heading entries for TOC refresh.
 */
/** Paragraph target info for bookmark-based cross-references */
interface ParagraphTarget {
  pmPos: number;
  bookmarkName: string | null;
}

function buildReferenceMap(
  state: EditorState,
  config: CrossRefUpdaterConfig
): {
  headingMap: Map<string, string>;
  captionMap: Map<string, string>;
  captions: CaptionInfo[];
  headingEntries: HeadingEntry[];
  /** Map from paragraph text → position + existing _FP_Ref_ bookmark (if any) */
  targetMap: Map<string, ParagraphTarget>;
} {
  const headingMap = new Map<string, string>();
  const captionMap = new Map<string, string>();
  const captions: CaptionInfo[] = [];
  const headingEntries: HeadingEntry[] = [];
  const targetMap = new Map<string, ParagraphTarget>();
  const headingCounters = [0, 0, 0, 0, 0, 0, 0, 0, 0];
  const captionCounters: Record<string, number> = {};

  const numMap = config.getNumberingMap?.() ?? null;
  const styleResolver = config.getStyleResolver?.() ?? null;

  state.doc.descendants((node, pos) => {
    if (node.type.name !== 'paragraph') return true;

    const styleId = node.attrs.styleId as string | null;
    if (!styleId) return true;

    // Heading numbering
    const headingMatch = styleId.match(/^Heading(\d)$/);
    if (headingMatch) {
      const level = parseInt(headingMatch[1]) - 1;

      // Collect text content (excluding numbering marker)
      // Apply allCaps transformation so TOC shows displayed text (matching Word behavior)
      let text = '';
      node.forEach((child) => {
        if (child.isText) {
          const raw = child.text || '';
          const hasAllCaps = child.marks.some((m) => m.type.name === 'allCaps');
          text += hasAllCaps ? raw.toUpperCase() : raw;
        }
      });
      const trimmedText = text.trim();

      // Skip empty headings (e.g. title-page spacers) — they don't participate in numbering
      if (trimmedText.length === 0) return true;

      headingCounters[level]++;
      for (let i = level + 1; i < headingCounters.length; i++) headingCounters[i] = 0;

      let number: string;
      const numPr =
        node.attrs.numPr ??
        styleResolver?.resolveParagraphStyle(styleId)?.paragraphFormatting?.numPr;
      if (numMap && numPr?.numId) {
        number = formatNumberedMarker(headingCounters, level, numMap, numPr.numId);
      } else {
        const parts = headingCounters.slice(0, level + 1).filter((v: number) => v > 0);
        number = parts.join('.');
      }

      headingMap.set(node.textContent, number);

      // Extract bookmarks from heading paragraph
      const bookmarks = (node.attrs.bookmarks as Array<{ id: number; name: string }>) || [];
      const tocBookmark = bookmarks.find((b) => b.name.startsWith('_Toc'));
      const fpRefBookmark = bookmarks.find((b) => b.name.startsWith('_FP_Ref_'));

      // Track this paragraph as a cross-ref target
      targetMap.set(node.textContent, {
        pmPos: pos,
        bookmarkName: fpRefBookmark?.name ?? null,
      });

      headingEntries.push({
        text: trimmedText,
        number,
        level,
        pmPos: pos,
        bookmarkName: tocBookmark?.name ?? null,
      });
    }

    // Caption numbering — detect prefix and count per-prefix
    if (styleId === 'Caption') {
      const prefix = detectCaptionPrefix(node);
      if (prefix) {
        captionCounters[prefix] = (captionCounters[prefix] ?? 0) + 1;
        const count = captionCounters[prefix];
        captionMap.set(node.textContent, `${prefix} ${count}`);
        captions.push({ pos, correctNumber: count, prefix });

        // Track as cross-ref target
        const capBookmarks = (node.attrs.bookmarks as Array<{ id: number; name: string }>) || [];
        const capFpRef = capBookmarks.find((b) => b.name.startsWith('_FP_Ref_'));
        targetMap.set(node.textContent, {
          pmPos: pos,
          bookmarkName: capFpRef?.name ?? null,
        });
      }
    }

    return true;
  });

  return { headingMap, captionMap, captions, headingEntries, targetMap };
}

/**
 * Analyse a Caption paragraph's children to find where the number is
 * and what kind of node holds it.
 *
 * Returns null if the structure doesn't match "Prefix N:" pattern.
 */
function analyseCaptionStructure(
  node: PMNode,
  prefix: CaptionPrefix
): {
  /** 'text' if the whole "Prefix N: " is a single text range, 'field' if N is a field atom */
  kind: 'text' | 'field';
  /** Current number value */
  currentNumber: number;
  /** For 'text': the full prefix match text and offset info */
  textPrefixLength?: number;
  colonSuffix?: string;
  firstTextMarks?: readonly Mark[];
  /** For 'field': the offset of the field node within the paragraph */
  fieldOffset?: number;
} | null {
  // Strategy: walk children and look for the pattern
  // Case 1 (plain text): text node starting with "Figure 1: "
  // Case 2 (field): text("Figure ") + field(displayText="1") + text(": ...")

  const children: { node: PMNode; offset: number }[] = [];
  node.forEach((child, childOffset) => {
    children.push({ node: child, offset: childOffset });
  });

  if (children.length === 0) return null;

  const first = children[0];

  // Case 1: Plain text caption — entire "Prefix N: " is in the first text node
  if (first.node.isText && first.node.text) {
    const regex = new RegExp(`^${prefix} (\\d+)(:\\s?)`);
    const m = first.node.text.match(regex);
    if (m) {
      return {
        kind: 'text',
        currentNumber: parseInt(m[1]),
        textPrefixLength: m[0].length,
        colonSuffix: m[2],
        firstTextMarks: first.node.marks,
      };
    }
  }

  // Case 2: Field-based caption — text("Prefix ") + field + text(": ...")
  if (
    children.length >= 2 &&
    first.node.isText &&
    first.node.text === prefix + ' ' &&
    children[1].node.type.name === 'field'
  ) {
    const fieldNode = children[1].node;
    const displayText = fieldNode.attrs.displayText as string;
    const num = parseInt(displayText);
    if (!isNaN(num)) {
      return {
        kind: 'field',
        currentNumber: num,
        fieldOffset: children[1].offset,
      };
    }
  }

  return null;
}

/**
 * Scan the document and return a transaction that updates all stale
 * cross-ref displayText, caption "Figure/Table N:" prefixes, and the
 * Table of Contents (if present).
 *
 * Returns null when nothing needs updating.
 */
export function refreshAllReferences(
  state: EditorState,
  config: CrossRefUpdaterConfig
): Transaction | null {
  const { headingMap, captionMap, captions, headingEntries, targetMap } = buildReferenceMap(
    state,
    config
  );

  let tr = state.tr;
  let changed = false;

  // Scan for crossRef nodes — update display text and migrate missing bookmarks
  state.doc.descendants((node, pos) => {
    if (node.type.name === 'crossRef') {
      const { refType, refTarget, displayText, bookmarkName } = node.attrs as {
        refType: string;
        refTarget: string;
        displayText: string;
        bookmarkName: string;
      };

      let currentNumber: string | undefined;
      if (refType === 'heading') {
        currentNumber = headingMap.get(refTarget);
      } else if (refType === 'figure') {
        currentNumber = captionMap.get(refTarget);
      }

      const needsDisplayUpdate = currentNumber != null && currentNumber !== displayText;
      const needsBookmark = !bookmarkName && refTarget;

      if (needsDisplayUpdate || needsBookmark) {
        const newAttrs = { ...node.attrs };
        if (needsDisplayUpdate) {
          newAttrs.displayText = currentNumber;
        }

        // Migrate: assign _FP_Ref_ bookmark to target paragraph if missing
        if (needsBookmark) {
          const target = targetMap.get(refTarget);
          if (target) {
            if (target.bookmarkName) {
              // Target already has a bookmark — use it
              newAttrs.bookmarkName = target.bookmarkName;
            } else {
              // Generate a new bookmark on the target paragraph
              const uuid = Math.random().toString(36).substring(2, 10);
              const bmName = `_FP_Ref_${uuid}`;
              const mappedTargetPos = tr.mapping.map(target.pmPos);
              const targetNode = tr.doc.nodeAt(mappedTargetPos);
              if (targetNode && targetNode.type.name === 'paragraph') {
                const existingBm =
                  (targetNode.attrs.bookmarks as Array<{ id: number; name: string }>) || [];
                tr = tr.setNodeMarkup(mappedTargetPos, undefined, {
                  ...targetNode.attrs,
                  bookmarks: [
                    ...existingBm,
                    { id: Math.floor(Math.random() * 2147483647), name: bmName },
                  ],
                });
                // Update target map so other crossRefs to same target reuse it
                target.bookmarkName = bmName;
              }
              newAttrs.bookmarkName = bmName;
            }
          }
        }

        tr = tr.setNodeMarkup(tr.mapping.map(pos), undefined, newAttrs);
        changed = true;
      }
      return false; // atom, no children
    }

    return true;
  });

  // Update caption numbers
  for (const caption of captions) {
    const mappedPos = tr.mapping.map(caption.pos);
    const paragraphNode = tr.doc.nodeAt(mappedPos);
    if (!paragraphNode || paragraphNode.type.name !== 'paragraph') continue;

    const info = analyseCaptionStructure(paragraphNode, caption.prefix);
    if (!info || info.currentNumber === caption.correctNumber) continue;

    if (info.kind === 'text') {
      // Replace "Prefix N: " text range with corrected number
      const prefixStart = mappedPos + 1; // +1 for paragraph open token
      const prefixEnd = prefixStart + info.textPrefixLength!;
      const newPrefix = `${caption.prefix} ${caption.correctNumber}${info.colonSuffix!}`;
      tr = tr.replaceWith(
        prefixStart,
        prefixEnd,
        state.schema.text(newPrefix, info.firstTextMarks!)
      );
      changed = true;
    } else if (info.kind === 'field') {
      // Update the SEQ field's displayText attribute
      const fieldPos = mappedPos + 1 + info.fieldOffset!; // +1 for paragraph open
      tr = tr.setNodeMarkup(fieldPos, undefined, {
        ...(tr.doc.nodeAt(fieldPos)?.attrs ?? {}),
        displayText: String(caption.correctNumber),
      });
      changed = true;
    }
  }

  // Refresh Table of Contents (only when layout is available — i.e. manual refresh)
  const layout = config.getLayout?.() ?? null;
  if (layout && headingEntries.length > 0) {
    const tocBlock = findTocBlock(tr.doc);
    if (tocBlock) {
      const { schema } = state;

      // Filter headings to match the existing TOC depth and exclude empty/TOC headings
      const tocHeadings = headingEntries.filter(
        (h) =>
          h.level < tocBlock.maxTocLevel &&
          h.text.length > 0 &&
          !/^Table of Contents$/i.test(h.text)
      );

      // Calculate tab position for right-aligned page numbers from page layout
      const page0 = layout.pages[0];
      const contentWidthPx = page0 ? page0.size.w - page0.margins.left - page0.margins.right : 650;
      const tabPositionTwips = Math.round(contentWidthPx * 15); // px → twips

      // Ensure all TOC-included headings have _Toc bookmarks (add missing ones)
      for (const h of tocHeadings) {
        if (h.bookmarkName) continue;
        const mappedPos = tr.mapping.map(h.pmPos);
        const paragraphNode = tr.doc.nodeAt(mappedPos);
        if (!paragraphNode || paragraphNode.type.name !== 'paragraph') continue;

        const bookmarkName = `_Toc${Math.floor(100000000 + Math.random() * 900000000)}`;
        h.bookmarkName = bookmarkName;
        const existingBookmarks =
          (paragraphNode.attrs.bookmarks as Array<{ id: number; name: string }>) || [];
        tr = tr.setNodeMarkup(mappedPos, undefined, {
          ...paragraphNode.attrs,
          bookmarks: [
            ...existingBookmarks,
            { id: Math.floor(Math.random() * 2147483647), name: bookmarkName },
          ],
        });
      }

      // Preserve the existing TOC heading paragraph (first node in TOC block)
      const mappedStart = tr.mapping.map(tocBlock.startPos);
      const tocHeadingNode = tr.doc.nodeAt(mappedStart);
      const tocNodes: PMNode[] = [];

      if (tocHeadingNode && tocHeadingNode.type.name === 'paragraph') {
        // Keep the original heading paragraph as-is
        tocNodes.push(tocHeadingNode);
      } else {
        // Fallback: create a new TOC heading
        tocNodes.push(
          schema.node('paragraph', { styleId: 'TOCHeading', alignment: 'center' }, [
            schema.text('Table of Contents', [schema.marks.bold.create()]),
          ])
        );
      }

      // TOC entries — use proper tab nodes so the layout engine renders
      // tab stops with correct width and dot leaders
      const tabType = schema.nodes.tab;
      const styleResolver = config.getStyleResolver?.() ?? null;

      // Resolve override style once (if set) for font/size/spacing on all TOC entries
      const overrideResolved = config.tocStyleOverride
        ? (styleResolver?.resolveParagraphStyle(config.tocStyleOverride) ?? null)
        : null;
      const overrideMarks: Mark[] = overrideResolved?.runFormatting
        ? textFormattingToMarks(overrideResolved.runFormatting, schema as Schema)
        : [];

      for (const h of tocHeadings) {
        const pageNum = getPageForPmPos(layout, h.pmPos);
        const tocStyleId = `TOC${h.level + 1}`;
        const indent = h.level * 720 || null;
        const linkMark = schema.marks.hyperlink.create({
          href: `#${h.bookmarkName}`,
        });

        // Resolve per-entry TOC style run formatting (font size, caps, etc.) so the
        // regenerated entries match the initial DOCX rendering.  The TOC1-9 styles
        // typically specify a font size (e.g. 14pt for TOC1) and may include allCaps.
        // Without these marks the layout bridge falls back to DEFAULT_SIZE (11pt),
        // making refreshed entries appear smaller than on initial load.
        const tocEntryResolved =
          overrideResolved ?? styleResolver?.resolveParagraphStyle(tocStyleId);
        const tocRunMarks: Mark[] = tocEntryResolved?.runFormatting
          ? textFormattingToMarks(tocEntryResolved.runFormatting, schema as Schema)
          : overrideMarks;

        // Combine link mark with style run formatting marks
        const entryMarks = tocRunMarks.length > 0 ? [linkMark, ...tocRunMarks] : [linkMark];

        const entryContent: PMNode[] = [];

        // Heading number + tab + title (or just title if no number)
        if (h.number) {
          entryContent.push(schema.text(h.number, entryMarks));
          if (tabType) {
            entryContent.push(tabType.create());
          }
          entryContent.push(schema.text(h.text, entryMarks));
        } else {
          entryContent.push(schema.text(h.text, entryMarks));
        }

        // Tab node (right-aligned with dot leader) + page number
        if (tabType) {
          entryContent.push(tabType.create());
        }
        entryContent.push(schema.text(String(pageNum), entryMarks));

        // Use the already-resolved per-entry style for paragraph spacing
        const pFmt = tocEntryResolved?.paragraphFormatting;

        tocNodes.push(
          schema.node(
            'paragraph',
            {
              styleId: tocStyleId,
              indentLeft: indent,
              spaceBefore: pFmt?.spaceBefore ?? null,
              spaceAfter: pFmt?.spaceAfter ?? null,
              lineSpacing: pFmt?.lineSpacing ?? null,
              lineSpacingRule: pFmt?.lineSpacingRule ?? null,
              tabs: [
                // Left tab for heading-number → title spacing (720 twips ≈ 0.5 inch)
                { position: 720, alignment: 'left' as const, leader: 'none' as const },
                // Right tab for page numbers — preserve existing leader style
                {
                  position: tabPositionTwips,
                  alignment: 'right' as const,
                  leader: tocBlock.existingLeader as
                    | 'none'
                    | 'dot'
                    | 'hyphen'
                    | 'underscore'
                    | 'heavy'
                    | 'middleDot',
                },
              ],
            },
            entryContent
          )
        );
      }

      // Replace old TOC block with new
      const mappedEnd = tr.mapping.map(tocBlock.endPos);
      tr = tr.replaceWith(mappedStart, mappedEnd, PMFragment.from(tocNodes));
      changed = true;
    }
  }

  return changed ? tr : null;
}

export function createCrossRefUpdaterPlugin(config: CrossRefUpdaterConfig = {}): Plugin {
  return new Plugin({
    key: crossRefUpdaterKey,

    appendTransaction(
      transactions: readonly Transaction[],
      _oldState: EditorState,
      newState: EditorState
    ) {
      // Only run when the document actually changed
      const docChanged = transactions.some((tr) => tr.docChanged);
      if (!docChanged) return null;

      // Capture storedMarks BEFORE refreshAllReferences (which may clear them)
      const savedStoredMarks = newState.storedMarks;

      const tr = refreshAllReferences(newState, config);

      // Preserve storedMarks from the current state so that formatting
      // set by the triggering transaction (e.g. bold after Enter in a heading)
      // isn't lost when this appendTransaction applies its changes.
      if (tr && savedStoredMarks) {
        tr.setStoredMarks(savedStoredMarks);
      }

      // Cross-reference updates must bypass the SelectiveEditablePlugin's
      // filterTransaction — locked paragraphs still need correct numbering.
      if (tr) {
        tr.setMeta('allowLockedEdit', true);
      }

      return tr;
    },
  });
}
