/**
 * Selective Editable Plugin
 *
 * ProseMirror plugin that prevents edits to locked paragraphs when
 * locked-editing mode is active. Allows cursor movement and selection
 * everywhere, but blocks content changes (insert, delete, formatting)
 * on paragraphs with `locked: true`.
 */

import { Plugin, PluginKey } from 'prosemirror-state';
import type { Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';

export const selectiveEditableKey = new PluginKey('selectiveEditable');

/**
 * Check whether a transaction modifies content within any locked paragraph.
 *
 * Strategy: for each step in the transaction, check the range of positions
 * affected. If any position in that range falls inside a locked paragraph,
 * reject the entire transaction.
 */
function transactionTouchesLockedParagraph(tr: Transaction): boolean {
  // Only check transactions with document-changing steps
  if (!tr.docChanged) return false;

  const oldDoc = tr.before;

  for (const step of tr.steps) {
    // Use the step's JSON to get from/to positions in the old doc
    const stepJson = step.toJSON() as { from?: number; to?: number; pos?: number };
    const from = stepJson.from ?? stepJson.pos ?? 0;
    const to = stepJson.to ?? from;

    // Check if any paragraph in the affected range is locked
    try {
      oldDoc.nodesBetween(from, Math.min(to, oldDoc.content.size), (node) => {
        if (node.type.name === 'paragraph' && node.attrs.locked) {
          throw new Error('locked'); // Short-circuit via exception
        }
      });
    } catch (e) {
      if (e instanceof Error && e.message === 'locked') {
        return true;
      }
      throw e;
    }

    // Also check if from position is inside a locked paragraph
    // (for insertions at the boundary)
    if (from <= oldDoc.content.size) {
      try {
        const $from = oldDoc.resolve(from);
        for (let d = $from.depth; d >= 0; d--) {
          const ancestor = $from.node(d);
          if (ancestor.type.name === 'paragraph' && ancestor.attrs.locked) {
            throw new Error('locked');
          }
        }
      } catch (e) {
        if (e instanceof Error && e.message === 'locked') {
          return true;
        }
        throw e;
      }
    }
  }

  return false;
}

/**
 * Create the selective editable plugin.
 *
 * When active, this plugin uses `filterTransaction` to reject any
 * transaction that would modify content inside a locked paragraph.
 * Selection-only transactions (cursor movement, selection changes) pass through.
 */
export function createSelectiveEditablePlugin(): Plugin {
  return new Plugin({
    key: selectiveEditableKey,

    filterTransaction(tr) {
      // Allow selection-only changes and meta-only transactions
      if (!tr.docChanged) return true;

      // Allow transactions that are setting node markup (lock/unlock commands from admin)
      // These are identified by having setNodeMarkup steps that only change the locked attr
      if (tr.getMeta('allowLockedEdit')) return true;

      // Check if the transaction touches any locked paragraph
      return !transactionTouchesLockedParagraph(tr);
    },

    props: {
      // Provide a visual hint: don't show text cursor in locked paragraphs
      handleDOMEvents: {
        beforeinput(view: EditorView, event: Event) {
          const inputEvent = event as InputEvent;
          const { state } = view;
          const { $from } = state.selection;

          // Check if the cursor is in a locked paragraph
          for (let d = $from.depth; d >= 0; d--) {
            const node = $from.node(d);
            if (node.type.name === 'paragraph' && node.attrs.locked) {
              // Block the input event for content-changing input types
              const contentChanging = [
                'insertText',
                'insertCompositionText',
                'insertFromPaste',
                'insertFromDrop',
                'deleteContentBackward',
                'deleteContentForward',
                'deleteByCut',
                'deleteByDrag',
                'deleteWordBackward',
                'deleteWordForward',
                'formatBold',
                'formatItalic',
                'formatUnderline',
              ];
              if (contentChanging.includes(inputEvent.inputType)) {
                event.preventDefault();
                return true;
              }
            }
          }
          return false;
        },
      },
    },
  });
}
