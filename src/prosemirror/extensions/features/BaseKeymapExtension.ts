/**
 * Base Keymap Extension — wraps prosemirror-commands baseKeymap
 *
 * Priority: Low (150) — must be the last keymap so other extensions can override keys
 */

import {
  baseKeymap,
  splitBlock,
  deleteSelection,
  joinBackward,
  joinForward,
  selectAll,
  selectParentNode,
} from 'prosemirror-commands';
import type { Mark } from 'prosemirror-model';
import { createExtension } from '../create';
import { textFormattingToMarks } from '../marks/markUtils';
import { getNextStyleId } from '../../styles/styleStore';
import { Priority } from '../types';
import type { ExtensionRuntime, ExtensionContext } from '../types';
import type { Command, Transaction } from 'prosemirror-state';
import { insertPageBreak } from '../../commands/pageBreak';
import type { TextFormatting } from '../../../types/document';

function chainCommands(...commands: Command[]): Command {
  return (state, dispatch, view) => {
    for (const cmd of commands) {
      if (cmd(state, dispatch, view)) {
        return true;
      }
    }
    return false;
  };
}

/**
 * Backspace at the start of a paragraph clears first-line indent / hanging indent
 * before joining with the previous paragraph (matches Word behavior).
 */
const clearIndentOnBackspace: Command = (state, dispatch) => {
  const { $cursor } = state.selection as {
    $cursor?: {
      parentOffset: number;
      parent: { type: { name: string }; attrs: Record<string, unknown> };
      pos: number;
      before: () => number;
    };
  };
  if (!$cursor) return false;

  // Only at the very start of a paragraph
  if ($cursor.parentOffset !== 0) return false;
  if ($cursor.parent.type.name !== 'paragraph') return false;

  const attrs = $cursor.parent.attrs;
  const hasFirstLine = attrs.indentFirstLine != null && (attrs.indentFirstLine as number) > 0;
  const hasHanging = !!attrs.hangingIndent;
  const hasIndentLeft = attrs.indentLeft != null && (attrs.indentLeft as number) > 0;

  if (!hasFirstLine && !hasHanging && !hasIndentLeft) return false;

  if (dispatch) {
    const pos = $cursor.before();
    const tr = state.tr.setNodeMarkup(pos, undefined, {
      ...attrs,
      indentFirstLine: null,
      hangingIndent: null,
      indentLeft: null,
    });
    dispatch(tr.scrollIntoView());
  }
  return true;
};

/**
 * Custom Enter handler: splits the block, inherits style-related attrs,
 * clears paragraph borders, and preserves font marks on the new paragraph.
 *
 * splitBlock creates a new paragraph with default attrs (all null),
 * so we must manually copy style-related attrs from the source paragraph.
 * Word does NOT propagate paragraph borders (w:pBdr) on Enter.
 */
const INHERITED_PARA_ATTRS = [
  'defaultTextFormatting',
  'styleId',
  'lineSpacing',
  'lineSpacingRule',
  'spaceAfter',
  'spaceBefore',
  'contextualSpacing',
] as const;

/** Mark types that should be carried to the new paragraph on Enter. */
const STYLE_MARK_NAMES = new Set([
  'fontFamily',
  'fontSize',
  'textColor',
  'bold',
  'italic',
  'underline',
  'strikethrough',
]);

const splitBlockClearBorders: Command = (state, dispatch, view) => {
  // Capture source paragraph info BEFORE split (splitBlock resets everything)
  const { $from: preSplitFrom } = state.selection;
  const sourcePara = preSplitFrom.parent.type.name === 'paragraph' ? preSplitFrom.parent : null;

  // Collect style marks from the cursor position before splitting.
  // Use storedMarks if set, otherwise resolve from the position.
  const preMarks = state.storedMarks || preSplitFrom.marks();
  const styleMarks = preMarks.filter((m) => STYLE_MARK_NAMES.has(m.type.name));

  // Intercept splitBlock's transaction so we can modify it before dispatch.
  // This ensures attrs + stored marks are set in a single transaction,
  // avoiding a flash where the empty paragraph has no formatting.
  let splitTr: Transaction | null = null;
  const capturingDispatch = dispatch
    ? (tr: Transaction) => {
        splitTr = tr;
      }
    : undefined;

  if (!splitBlock(state, capturingDispatch, view)) {
    return false;
  }

  if (dispatch && splitTr !== null) {
    // After split, cursor is in the new (second) paragraph.
    // Apply attr inheritance, border clearing, and stored marks to the SAME transaction.
    const tr = splitTr as Transaction;
    const { $from } = tr.selection;
    const newPara = $from.parent;

    if (newPara.type.name === 'paragraph') {
      const newAttrs = { ...newPara.attrs };
      let attrsChanged = false;

      // Copy inherited attrs from source paragraph, respecting style.next
      if (sourcePara) {
        for (const key of INHERITED_PARA_ATTRS) {
          const srcVal = sourcePara.attrs[key];
          if (srcVal != null && newAttrs[key] == null) {
            newAttrs[key] = srcVal;
            attrsChanged = true;
          }
        }

        // Check if the source style defines a "next" style (e.g., Heading1.next = Normal)
        const sourceStyleId = sourcePara.attrs.styleId as string | undefined;
        if (sourceStyleId) {
          const nextStyleId = getNextStyleId(sourceStyleId);
          if (nextStyleId && nextStyleId !== sourceStyleId) {
            newAttrs.styleId = nextStyleId;
            attrsChanged = true;
          }
        }
      }

      // Clear identity attrs that must NOT be duplicated across paragraphs.
      // ProseMirror's splitBlock copies all attrs from the source node; these
      // are per-paragraph identifiers that would cause DOCX corruption
      // (duplicate w14:paraId / bookmarkStart ids) if carried to the new node.
      for (const key of ['paraId', 'textId', 'bookmarks', '_originalFormatting'] as const) {
        if (newAttrs[key] != null) {
          newAttrs[key] = null;
          attrsChanged = true;
        }
      }

      // Clear borders (Word does not propagate paragraph borders on Enter)
      if (newAttrs.borders) {
        newAttrs.borders = null;
        attrsChanged = true;
      }

      if (attrsChanged) {
        tr.setNodeMarkup($from.before(), undefined, newAttrs);
      }

      // When Enter is pressed at the START of a paragraph, splitBlock creates
      // an empty paragraph ABOVE the cursor. That empty paragraph is the "new"
      // node and may have lost the source paragraph's styleId. Fix it, and
      // ensure defaultTextFormatting is set so typing in it later inherits
      // the correct font/size/bold/etc from the style.
      const cursorParaStart = $from.before();
      if (cursorParaStart > 0 && sourcePara) {
        const prevPos = tr.doc.resolve(cursorParaStart - 1);
        const prevPara = prevPos.parent.type.name === 'paragraph' ? prevPos.parent : null;
        // Check if the previous paragraph is empty (split at pos 0) and missing style
        if (prevPara && prevPara.content.size === 0) {
          const prevParaPos = prevPos.before();
          const prevAttrs = { ...prevPara.attrs };
          let prevChanged = false;
          for (const key of INHERITED_PARA_ATTRS) {
            const srcVal = sourcePara.attrs[key];
            if (srcVal != null && prevAttrs[key] == null) {
              prevAttrs[key] = srcVal;
              prevChanged = true;
            }
          }
          // Ensure defaultTextFormatting is carried over so that when the
          // user navigates to this empty paragraph and types, ProseMirror's
          // mark resolution picks up the style's font/size/bold/etc.
          if (sourcePara.attrs.defaultTextFormatting && !prevAttrs.defaultTextFormatting) {
            prevAttrs.defaultTextFormatting = sourcePara.attrs.defaultTextFormatting;
            prevChanged = true;
          }
          // Clear identity attrs on the empty paragraph too
          for (const key of ['paraId', 'textId', 'bookmarks', '_originalFormatting'] as const) {
            if (prevAttrs[key] != null) {
              prevAttrs[key] = null;
              prevChanged = true;
            }
          }
          if (prevAttrs.borders) {
            prevAttrs.borders = null;
            prevChanged = true;
          }
          if (prevChanged) {
            tr.setNodeMarkup(prevParaPos, undefined, prevAttrs);
          }
        }
      }

      // For empty paragraphs (Enter at end of line), set stored marks so typed text
      // inherits formatting from the source paragraph (font, size, color, bold, etc.).
      if (newPara.textContent.length === 0) {
        // Determine effective style marks. When text has explicit marks (e.g. user
        // applied a font override), use those. When text inherits formatting from
        // the paragraph style chain (no explicit marks), derive marks from the
        // source paragraph's defaultTextFormatting.
        let effectiveMarks: Mark[] = styleMarks;

        if (effectiveMarks.length === 0 && sourcePara) {
          const dtf = sourcePara.attrs.defaultTextFormatting as TextFormatting | undefined;
          if (dtf) {
            const allMarks = textFormattingToMarks(dtf, state.schema);
            effectiveMarks = allMarks.filter((m) => STYLE_MARK_NAMES.has(m.type.name));
          }
        }

        if (effectiveMarks.length > 0) {
          // Sync defaultTextFormatting with the actual cursor marks so the empty
          // paragraph measurement (used for caret height) matches the stored marks.
          const dtf = { ...(newAttrs.defaultTextFormatting ?? {}) };
          let dtfChanged = false;
          for (const m of effectiveMarks) {
            if (m.type.name === 'fontSize' && m.attrs.size !== dtf.fontSize) {
              dtf.fontSize = m.attrs.size;
              dtfChanged = true;
            }
            if (m.type.name === 'fontFamily') {
              const ascii = m.attrs.ascii as string | undefined;
              if (ascii && (!dtf.fontFamily || dtf.fontFamily.ascii !== ascii)) {
                dtf.fontFamily = { ...dtf.fontFamily, ascii, hAnsi: m.attrs.hAnsi };
                dtfChanged = true;
              }
            }
          }
          if (dtfChanged) {
            tr.setNodeMarkup($from.before(), undefined, {
              ...newAttrs,
              defaultTextFormatting: dtf,
            });
          }

          // IMPORTANT: setStoredMarks MUST be called AFTER all setNodeMarkup calls.
          // setNodeMarkup adds a ReplaceStep which clears storedMarks on the transaction.
          tr.setStoredMarks(effectiveMarks);
        }
      }
    }

    dispatch(tr.scrollIntoView());
  }

  return true;
};

export const BaseKeymapExtension = createExtension({
  name: 'baseKeymap',
  priority: Priority.Low,
  onSchemaReady(_ctx: ExtensionContext): ExtensionRuntime {
    return {
      keyboardShortcuts: {
        // Base keymap provides default editing commands
        ...baseKeymap,
        // Override some keys with better defaults
        Enter: splitBlockClearBorders,
        Backspace: chainCommands(deleteSelection, clearIndentOnBackspace, joinBackward),
        Delete: chainCommands(deleteSelection, joinForward),
        'Mod-Enter': insertPageBreak,
        'Mod-a': selectAll,
        Escape: selectParentNode,
      },
    };
  },
});
