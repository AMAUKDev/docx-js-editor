import type { Transaction } from 'prosemirror-state';

/**
 * Meta flag for programmatic context-tag label refreshes (applyLabels in
 * DocxEditor). These transactions change the document (setNodeMarkup on
 * contextTag nodes) but represent derived display data, not user work:
 * consumers must not treat them as "the user edited the document" (e.g.
 * flipping an unsaved-changes indicator), and they are excluded from the
 * undo history.
 */
export const CONTEXT_TAG_LABEL_SYNC_META = 'contextTagLabelSync';

/** Mark a transaction as a programmatic label sync (also skips undo history). */
export function markContextTagLabelSync(tr: Transaction): Transaction {
  return tr.setMeta(CONTEXT_TAG_LABEL_SYNC_META, true).setMeta('addToHistory', false);
}

/** True when the transaction is a programmatic context-tag label sync. */
export function isContextTagLabelSync(tr: Transaction): boolean {
  return tr.getMeta(CONTEXT_TAG_LABEL_SYNC_META) === true;
}
