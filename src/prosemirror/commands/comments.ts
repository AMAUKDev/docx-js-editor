/**
 * Comment and Track Changes Commands
 *
 * PM commands for adding/removing comments and accepting/rejecting tracked changes.
 * Ported from eigenpal/docx-editor upstream/main.
 */

import type { Command, EditorState } from 'prosemirror-state';

/**
 * Add a comment mark to the current selection.
 */
export function addCommentMark(commentId: number): Command {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;
    if (empty) return false;
    const commentType = state.schema.marks.comment;
    if (!commentType) return false;
    if (dispatch) {
      dispatch(state.tr.addMark(from, to, commentType.create({ commentId })));
    }
    return true;
  };
}

/**
 * Remove a comment mark by ID from the entire document.
 */
export function removeCommentMark(commentId: number): Command {
  return (state, dispatch) => {
    const commentType = state.schema.marks.comment;
    if (!commentType) return false;
    if (dispatch) {
      const tr = state.tr;
      state.doc.descendants((node, pos) => {
        if (node.isText) {
          for (const mark of node.marks) {
            if (mark.type === commentType && mark.attrs.commentId === commentId) {
              tr.removeMark(pos, pos + node.nodeSize, mark);
            }
          }
        }
      });
      if (tr.steps.length > 0) dispatch(tr);
    }
    return true;
  };
}

/**
 * Resolve a tracked change: accept or reject.
 * - Accept: keep insertions (remove mark), delete deletions (remove text)
 * - Reject: keep deletions (remove mark), delete insertions (remove text)
 */
function resolveChange(from: number, to: number, mode: 'accept' | 'reject'): Command {
  return (state, dispatch) => {
    const insertionType = state.schema.marks.insertion;
    const deletionType = state.schema.marks.deletion;
    if (!insertionType && !deletionType) return false;

    const keepType = mode === 'accept' ? insertionType : deletionType;
    const removeType = mode === 'accept' ? deletionType : insertionType;

    if (dispatch) {
      const tr = state.tr;
      const deleteRanges: Array<{ from: number; to: number }> = [];

      state.doc.nodesBetween(from, to, (node, pos) => {
        if (!node.isText) return;
        const nodeEnd = pos + node.nodeSize;
        const rangeFrom = Math.max(from, pos);
        const rangeTo = Math.min(to, nodeEnd);

        if (removeType && node.marks.some((m) => m.type === removeType)) {
          deleteRanges.push({ from: rangeFrom, to: rangeTo });
        }
        if (keepType && node.marks.some((m) => m.type === keepType)) {
          tr.removeMark(rangeFrom, rangeTo, keepType);
        }
      });

      for (const range of deleteRanges.reverse()) {
        tr.delete(range.from, range.to);
      }

      if (tr.steps.length > 0) dispatch(tr);
    }
    return true;
  };
}

/** Accept a tracked change in the given range (keep insertions, remove deletions). */
export function acceptChange(from: number, to: number): Command {
  return resolveChange(from, to, 'accept');
}

/** Reject a tracked change in the given range (remove insertions, keep deletions). */
export function rejectChange(from: number, to: number): Command {
  return resolveChange(from, to, 'reject');
}

/** Accept all tracked changes in the document. */
export function acceptAllChanges(): Command {
  return (state, dispatch) => acceptChange(0, state.doc.content.size)(state, dispatch);
}

/** Reject all tracked changes in the document. */
export function rejectAllChanges(): Command {
  return (state, dispatch) => rejectChange(0, state.doc.content.size)(state, dispatch);
}

export interface ChangeRange {
  from: number;
  to: number;
  type: 'insertion' | 'deletion';
}

/**
 * Expand a seed range to cover all contiguous text nodes within the same
 * parent that carry the same tracked-change mark type and revisionId.
 * This ensures Accept/Reject operates on the full run, not one text node.
 */
function expandChangeRange(state: EditorState, seed: ChangeRange): ChangeRange {
  const markType =
    seed.type === 'insertion' ? state.schema.marks.insertion : state.schema.marks.deletion;
  if (!markType) return seed;

  // Find the revisionId at the seed position
  let revisionId: number | null = null;
  state.doc.nodesBetween(seed.from, Math.min(seed.to, state.doc.content.size), (node) => {
    if (!node.isText || revisionId !== null) return;
    for (const m of node.marks) {
      if (m.type === markType) {
        revisionId = m.attrs.revisionId as number;
        return;
      }
    }
  });

  // Walk the parent's children to find the contiguous span
  const $pos = state.doc.resolve(seed.from);
  const parent = $pos.parent;
  const parentStart = $pos.start($pos.depth);

  let from = seed.from;
  let to = seed.to;

  // Scan backwards — find earliest contiguous node with same mark+revisionId
  let prevEnd = seed.from;
  parent.forEach((child, offset) => {
    const childPos = parentStart + offset;
    const childEnd = childPos + child.nodeSize;
    if (childEnd <= seed.from) {
      if (
        child.isText &&
        child.marks.some(
          (m) => m.type === markType && (revisionId == null || m.attrs.revisionId === revisionId)
        )
      ) {
        // Contiguous only if this child ends where the next tracked node starts
        if (childEnd === prevEnd || childEnd >= from) {
          from = childPos;
        }
        prevEnd = childPos;
      } else {
        // Gap — reset
        from = seed.from;
        prevEnd = seed.from;
      }
    }
  });

  // Scan forwards — find latest contiguous node with same mark+revisionId
  parent.forEach((child, offset) => {
    const childPos = parentStart + offset;
    if (childPos < seed.to) return;
    if (
      child.isText &&
      child.marks.some(
        (m) => m.type === markType && (revisionId == null || m.attrs.revisionId === revisionId)
      )
    ) {
      to = childPos + child.nodeSize;
    }
  });

  return { from, to, type: seed.type };
}

/** Find the next tracked change after startPos (wraps around). Returns the full contiguous run. */
export function findNextChange(state: EditorState, startPos: number): ChangeRange | null {
  const insertionType = state.schema.marks.insertion;
  const deletionType = state.schema.marks.deletion;
  if (!insertionType && !deletionType) return null;

  let seed: ChangeRange | null = null;

  state.doc.descendants((node, pos) => {
    if (seed) return false;
    if (!node.isText) return;
    if (pos + node.nodeSize <= startPos) return;
    for (const mark of node.marks) {
      if (mark.type === insertionType || mark.type === deletionType) {
        seed = {
          from: Math.max(pos, startPos),
          to: pos + node.nodeSize,
          type: mark.type === insertionType ? 'insertion' : 'deletion',
        };
        return false;
      }
    }
  });

  if (!seed && startPos > 0) return findNextChange(state, 0);
  return seed ? expandChangeRange(state, seed) : null;
}

/** Find the previous tracked change before startPos (wraps around). Returns the full contiguous run. */
export function findPreviousChange(state: EditorState, startPos: number): ChangeRange | null {
  const insertionType = state.schema.marks.insertion;
  const deletionType = state.schema.marks.deletion;
  if (!insertionType && !deletionType) return null;

  let seed: ChangeRange | null = null;

  state.doc.descendants((node, pos) => {
    if (!node.isText) return;
    if (pos >= startPos) return false;
    for (const mark of node.marks) {
      if (mark.type === insertionType || mark.type === deletionType) {
        seed = {
          from: pos,
          to: pos + node.nodeSize,
          type: mark.type === insertionType ? 'insertion' : 'deletion',
        };
      }
    }
  });

  if (!seed && startPos < state.doc.content.size)
    return findPreviousChange(state, state.doc.content.size);
  return seed ? expandChangeRange(state, seed) : null;
}

/** Count all tracked changes in the document. */
export function countChanges(state: EditorState): number {
  const insertionType = state.schema.marks.insertion;
  const deletionType = state.schema.marks.deletion;
  if (!insertionType && !deletionType) return 0;
  let count = 0;
  state.doc.descendants((node) => {
    if (node.isText) {
      for (const mark of node.marks) {
        if (mark.type === insertionType || mark.type === deletionType) count++;
      }
    }
  });
  return count;
}
