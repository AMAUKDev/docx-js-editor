/**
 * Ordinal-suffix auto-superscript plugin.
 *
 * Superscripts the letters of an ordinal suffix ("rd" in "3rd") the moment a real
 * word boundary completes the pattern — mirroring Word's AutoFormat-As-You-Type.
 *
 * Deliberately NOT a document-wide rescan (contrast with the auto-tag transform in
 * `../autoTag/`, which re-derives *node* structure from text and therefore needs a
 * portal-side ignore-set to keep an undone conversion from reappearing). A manually
 * reverted ordinal suffix is just baseline text that still matches the same regex —
 * there is no way to tell "never formatted" apart from "deliberately reverted" once
 * you've forgotten how it got that way. So this plugin never looks at old text: it
 * only inspects the range each incoming transaction actually inserted, mapped
 * through that transaction's own step maps. Untouched text is never revisited,
 * which is what makes a manual revert (via the existing Superscript toolbar toggle)
 * hold permanently, with no extra persisted state.
 *
 * Because `appendTransaction` fires for every dispatched transaction on the view
 * regardless of origin, this one plugin covers typing, paste, and portal-driven
 * insertions (e.g. `WordEditor.applyReview`) uniformly — no separate hook needed for
 * each entry point. Document load/import does not go through incremental
 * transactions at all, so pre-existing report text is never retroactively
 * reformatted — that is the intended scope, not an oversight.
 *
 * Known limitation: only scans within the paragraph containing the start of each
 * inserted range, so an ordinal split across paragraphs by a multi-paragraph paste
 * is not detected. Single-paragraph typing, single-paragraph paste, and in-place
 * AMAi text replacement (the paths that matter here) are all within one paragraph.
 *
 * Mark-only steps (`AddMarkStep`/`RemoveMarkStep`, e.g. the user toggling
 * Superscript off) still make `tr.docChanged` true, but their `getMap()` is
 * `StepMap.empty` — an identity map with zero recorded ranges — so `forEach` below
 * never invokes its callback for them. Only steps that actually change document
 * content (`ReplaceStep`/`ReplaceAroundStep`) ever contribute a range, which is what
 * keeps a manual revert from being reinterpreted as "content newly inserted here".
 */

import { Plugin, PluginKey } from 'prosemirror-state';
import type { EditorState, Transaction } from 'prosemirror-state';
import { findOrdinalSuffixMatches } from './ordinalSuffixMatcher';

export const ordinalSuffixKey = new PluginKey('ordinalSuffix');

/** How far back from an insertion to look for the start of a digit run. */
const LOOKBACK = 40;

export interface InsertedRange {
  from: number;
  to: number;
}

/**
 * Positions (in `newState.doc` coordinates) that this batch of transactions actually
 * inserted new content into. Deletions map to a zero-width range and are skipped —
 * only insertions can newly complete an ordinal pattern.
 *
 * Only accounts for step maps *within* a single transaction; a later transaction in
 * the same batch is not re-mapped against an earlier one's steps. Real editing
 * dispatches exactly one transaction per `appendTransaction` call in every case that
 * matters here (typing, paste, portal-driven insert), so this is not a practical gap.
 */
export function collectInsertedRanges(transactions: readonly Transaction[]): InsertedRange[] {
  const ranges: InsertedRange[] = [];
  for (const tr of transactions) {
    if (!tr.docChanged) continue;
    tr.steps.forEach((step, i) => {
      const stepMap = step.getMap();
      stepMap.forEach((_oldStart, _oldEnd, newStart, newEnd) => {
        if (newEnd <= newStart) return;
        const from = tr.mapping.slice(i + 1).map(newStart, -1);
        const to = tr.mapping.slice(i + 1).map(newEnd, 1);
        if (to > from) ranges.push({ from, to });
      });
    });
  }
  return ranges;
}

/**
 * Build a transaction that superscripts every ordinal suffix whose completing
 * boundary character falls inside one of `ranges`. Returns null if nothing to do.
 */
export function buildOrdinalSuffixTransaction(
  newState: EditorState,
  ranges: InsertedRange[],
): Transaction | null {
  const markType = newState.schema.marks.superscript;
  if (!markType || ranges.length === 0) return null;

  const docSize = newState.doc.content.size;
  let tr: Transaction | null = null;
  const applied = new Set<string>();

  for (const range of ranges) {
    const clampedFrom = Math.max(0, Math.min(range.from, docSize));
    const $from = newState.doc.resolve(clampedFrom);
    if (!$from.parent.isTextblock) continue;
    const blockStart = $from.start();
    const blockEnd = $from.end();
    const scanFrom = Math.max(blockStart, range.from - LOOKBACK);
    if (blockEnd <= scanFrom) continue;

    // '￼' (OBJECT REPLACEMENT CHARACTER) stands in for non-text leaf nodes
    // (e.g. a contextTag atom) so offsets into `text` still line up 1:1 with
    // document positions.
    const text = newState.doc.textBetween(scanFrom, blockEnd, '￼', '￼');

    for (const match of findOrdinalSuffixMatches(text)) {
      const absSuffixStart = scanFrom + match.suffixStart;
      const absSuffixEnd = scanFrom + match.suffixEnd;
      // Only act when THIS insertion is what completed the pattern.
      if (absSuffixEnd < range.from || absSuffixEnd > range.to) continue;

      const dedupeKey = `${absSuffixStart}-${absSuffixEnd}`;
      if (applied.has(dedupeKey)) continue;
      applied.add(dedupeKey);

      if (newState.doc.rangeHasMark(absSuffixStart, absSuffixEnd, markType)) continue;

      tr = (tr ?? newState.tr).addMark(absSuffixStart, absSuffixEnd, markType.create());
    }
  }

  if (tr) {
    // Without this, continuing to type immediately after the marked suffix would
    // inherit superscript — ProseMirror's default is that new input takes on the
    // marks of the character just before the cursor. Only the suffix itself should
    // ever be superscript; the boundary character and everything typed after it
    // must not silently carry the mark forward.
    tr.removeStoredMark(markType);
  }

  return tr;
}

export function createOrdinalSuffixPlugin(): Plugin {
  return new Plugin({
    key: ordinalSuffixKey,
    appendTransaction(transactions, _oldState, newState) {
      if (!newState.schema.marks.superscript) return null;
      if (!transactions.some((tr) => tr.docChanged)) return null;

      const ranges = collectInsertedRanges(transactions);
      const tr = buildOrdinalSuffixTransaction(newState, ranges);
      if (tr) {
        // Bypass SelectiveEditablePlugin's filterTransaction — the same escape
        // hatch crossRefUpdater uses for its own corrective transactions.
        tr.setMeta('allowLockedEdit', true);
      }
      return tr;
    },
  });
}
