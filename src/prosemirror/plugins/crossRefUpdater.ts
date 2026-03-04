/**
 * Cross-Reference Updater Plugin
 *
 * An appendTransaction plugin that keeps cross-ref displayText and
 * caption figure numbers in sync whenever the document changes.
 *
 * Scans the document for:
 * - Headings: paragraphs with styleId matching Heading\d, tracking numbering counters
 * - Captions: paragraphs with styleId === 'Caption', tracking figure count
 * - CrossRef nodes: inline crossRef atoms, checking if displayText is stale
 *
 * When a mismatch is found, dispatches a corrective transaction.
 */

import { Plugin, PluginKey, type Transaction } from 'prosemirror-state';
import type { EditorState } from 'prosemirror-state';
import type { Mark } from 'prosemirror-model';
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
  captionNumbers: Map<number, number>; // pmPos → correct figure number
} {
  const headingMap = new Map<string, string>();
  const captionMap = new Map<string, string>();
  const captionNumbers = new Map<number, number>();
  const headingCounters = [0, 0, 0, 0, 0, 0, 0, 0, 0];
  let figureCount = 0;

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

    // Caption numbering
    if (styleId === 'Caption') {
      figureCount++;
      captionMap.set(node.textContent, `Figure ${figureCount}`);
      captionNumbers.set(pos, figureCount);
    }

    return true;
  });

  return { headingMap, captionMap, captionNumbers };
}

/**
 * Scan the document and return a transaction that updates all stale
 * cross-ref displayText and caption "Figure N:" prefixes.
 *
 * Returns null when nothing needs updating.
 */
export function refreshAllReferences(
  state: EditorState,
  config: CrossRefUpdaterConfig
): Transaction | null {
  const { headingMap, captionMap, captionNumbers } = buildReferenceMap(state, config);

  let tr = state.tr;
  let changed = false;

  // Scan for crossRef nodes and caption paragraphs that need updating
  state.doc.descendants((node, pos) => {
    // Update crossRef nodes
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
        tr = tr.setNodeMarkup(pos, undefined, {
          ...node.attrs,
          displayText: currentNumber,
        });
        changed = true;
      }
      return false; // atom, no children
    }

    // Update caption figure numbers
    if (node.type.name === 'paragraph' && node.attrs.styleId === 'Caption') {
      const correctNumber = captionNumbers.get(pos);
      if (correctNumber != null) {
        const text = node.textContent;
        const captionMatch = text.match(/^Figure (\d+)(:\s?)/);
        if (captionMatch) {
          const currentNum = parseInt(captionMatch[1]);
          if (currentNum !== correctNumber) {
            // Extract marks from the first text node in the prefix range
            let existingMarks: readonly Mark[] = [];
            node.nodesBetween(0, captionMatch[0].length, (child) => {
              if (child.isText && existingMarks.length === 0) {
                existingMarks = child.marks;
              }
            });

            // Replace "Figure N:" prefix with correct number, preserving marks
            const prefixStart = pos + 1; // +1 for paragraph open token
            const prefixEnd = prefixStart + captionMatch[0].length;
            const newPrefix = `Figure ${correctNumber}${captionMatch[2]}`;
            tr = tr.replaceWith(
              prefixStart,
              prefixEnd,
              state.schema.text(newPrefix, existingMarks)
            );
            changed = true;
          }
        }
      }
    }

    return true;
  });

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
