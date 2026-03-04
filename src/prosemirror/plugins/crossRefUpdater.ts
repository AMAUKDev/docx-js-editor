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
 *
 * When a mismatch is found, dispatches a corrective transaction.
 */

import { Plugin, PluginKey, type Transaction } from 'prosemirror-state';
import type { EditorState } from 'prosemirror-state';
import type { Mark, Node as PMNode } from 'prosemirror-model';
import type { NumberingMap } from '../../docx/numberingParser';
import { formatNumberedMarker } from '../../layout-bridge/toFlowBlocks';

export const crossRefUpdaterKey = new PluginKey('crossRefUpdater');

export interface CrossRefUpdaterConfig {
  /** Return the current numbering map (may change over time). */
  getNumberingMap?: () => NumberingMap | null;
  /** Return a style resolver for looking up numPr from style definitions. */
  getStyleResolver?: () => {
    resolveParagraphStyle(styleId: string): {
      paragraphFormatting?: { numPr?: { numId?: number; ilvl?: number } };
    };
  } | null;
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

/**
 * Build a map of refTarget → currentNumber by scanning the document.
 * Also returns caption number mappings.
 */
function buildReferenceMap(
  state: EditorState,
  config: CrossRefUpdaterConfig
): {
  headingMap: Map<string, string>;
  captionMap: Map<string, string>;
  captions: CaptionInfo[];
} {
  const headingMap = new Map<string, string>();
  const captionMap = new Map<string, string>();
  const captions: CaptionInfo[] = [];
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
    }

    // Caption numbering — detect prefix and count per-prefix
    if (styleId === 'Caption') {
      const prefix = detectCaptionPrefix(node);
      if (prefix) {
        captionCounters[prefix] = (captionCounters[prefix] ?? 0) + 1;
        const count = captionCounters[prefix];
        captionMap.set(node.textContent, `${prefix} ${count}`);
        captions.push({ pos, correctNumber: count, prefix });
      }
    }

    return true;
  });

  return { headingMap, captionMap, captions };
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
 * cross-ref displayText and caption "Figure/Table N:" prefixes.
 *
 * Returns null when nothing needs updating.
 */
export function refreshAllReferences(
  state: EditorState,
  config: CrossRefUpdaterConfig
): Transaction | null {
  const { headingMap, captionMap, captions } = buildReferenceMap(state, config);

  let tr = state.tr;
  let changed = false;

  // Scan for crossRef nodes
  state.doc.descendants((node, pos) => {
    if (node.type.name === 'crossRef') {
      const { refType, refTarget, displayText } = node.attrs as {
        refType: string;
        refTarget: string;
        displayText: string;
      };

      let currentNumber: string | undefined;
      if (refType === 'heading') {
        currentNumber = headingMap.get(refTarget);
      } else if (refType === 'figure') {
        currentNumber = captionMap.get(refTarget);
      }

      if (currentNumber != null && currentNumber !== displayText) {
        tr = tr.setNodeMarkup(tr.mapping.map(pos), undefined, {
          ...node.attrs,
          displayText: currentNumber,
        });
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

      return refreshAllReferences(newState, config);
    },
  });
}
