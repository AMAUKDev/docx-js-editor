/**
 * List Extension — list commands + keymaps
 *
 * No schema contribution — lists use paragraph attrs (numPr).
 * Provides: toggle bullet/number, indent/outdent, enter/backspace handling.
 */

import type { Command, EditorState } from 'prosemirror-state';
import type { Mark } from 'prosemirror-model';
import { createExtension } from '../create';
import { textFormattingToMarks } from '../marks/markUtils';
import { Priority } from '../types';
import type { ExtensionRuntime } from '../types';
import type { TextFormatting } from '../../../types/document';

// ============================================================================
// CHAIN COMMANDS HELPER
// ============================================================================

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

// ============================================================================
// LIST COMMANDS
// ============================================================================

function toggleList(numId: number): Command {
  return (state, dispatch) => {
    const { $from, $to } = state.selection;

    const paragraph = $from.parent;
    if (paragraph.type.name !== 'paragraph') return false;

    const currentNumPr = paragraph.attrs.numPr;
    const isInSameList = currentNumPr?.numId === numId;

    if (!dispatch) return true;

    let tr = state.tr;
    const seen = new Set<number>();

    state.doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
      if (node.type.name === 'paragraph' && !seen.has(pos)) {
        seen.add(pos);

        if (isInSameList) {
          tr = tr.setNodeMarkup(pos, undefined, {
            ...node.attrs,
            numPr: null,
            listIsBullet: null,
            listNumFmt: null,
            listMarker: null,
          });
        } else {
          const isBullet = numId === 1;
          tr = tr.setNodeMarkup(pos, undefined, {
            ...node.attrs,
            numPr: { numId, ilvl: node.attrs.numPr?.ilvl || 0 },
            listIsBullet: isBullet,
            listNumFmt: isBullet ? null : 'decimal',
            listMarker: null,
          });
        }
      }
    });

    dispatch(tr.scrollIntoView());
    return true;
  };
}

const toggleBulletList: Command = (state, dispatch) => {
  return toggleList(1)(state, dispatch);
};

const toggleNumberedList: Command = (state, dispatch) => {
  return toggleList(2)(state, dispatch);
};

const increaseListLevel: Command = (state, dispatch) => {
  const { $from } = state.selection;
  const paragraph = $from.parent;

  if (paragraph.type.name !== 'paragraph') return false;
  if (!paragraph.attrs.numPr) return false;

  const currentLevel = paragraph.attrs.numPr.ilvl || 0;
  if (currentLevel >= 8) return false;

  if (!dispatch) return true;

  const paragraphPos = $from.before($from.depth);

  dispatch(
    state.tr
      .setNodeMarkup(paragraphPos, undefined, {
        ...paragraph.attrs,
        numPr: { ...paragraph.attrs.numPr, ilvl: currentLevel + 1 },
        // Clear explicit indentation so layout engine computes from new level
        indentLeft: null,
        indentFirstLine: null,
        hangingIndent: null,
      })
      .scrollIntoView()
  );

  return true;
};

const decreaseListLevel: Command = (state, dispatch) => {
  const { $from } = state.selection;
  const paragraph = $from.parent;

  if (paragraph.type.name !== 'paragraph') return false;
  if (!paragraph.attrs.numPr) return false;

  const currentLevel = paragraph.attrs.numPr.ilvl || 0;

  if (!dispatch) return true;

  const paragraphPos = $from.before($from.depth);

  if (currentLevel <= 0) {
    dispatch(
      state.tr
        .setNodeMarkup(paragraphPos, undefined, {
          ...paragraph.attrs,
          numPr: null,
          listIsBullet: null,
          listNumFmt: null,
          listMarker: null,
          indentLeft: null,
          indentFirstLine: null,
          hangingIndent: null,
        })
        .scrollIntoView()
    );
  } else {
    dispatch(
      state.tr
        .setNodeMarkup(paragraphPos, undefined, {
          ...paragraph.attrs,
          numPr: { ...paragraph.attrs.numPr, ilvl: currentLevel - 1 },
          indentLeft: null,
          indentFirstLine: null,
          hangingIndent: null,
        })
        .scrollIntoView()
    );
  }

  return true;
};

const removeList: Command = (state, dispatch) => {
  const { $from, $to } = state.selection;

  if (!dispatch) return true;

  let tr = state.tr;
  const seen = new Set<number>();

  state.doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
    if (node.type.name === 'paragraph' && node.attrs.numPr && !seen.has(pos)) {
      seen.add(pos);
      tr = tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        numPr: null,
        listIsBullet: null,
        listNumFmt: null,
        listMarker: null,
      });
    }
  });

  dispatch(tr.scrollIntoView());
  return true;
};

// ============================================================================
// LIST QUERY HELPERS (exported for toolbar)
// ============================================================================

export function isInList(state: EditorState): boolean {
  const { $from } = state.selection;
  const paragraph = $from.parent;

  if (paragraph.type.name !== 'paragraph') return false;
  return !!paragraph.attrs.numPr?.numId;
}

export function getListInfo(state: EditorState): { numId: number; ilvl: number } | null {
  const { $from } = state.selection;
  const paragraph = $from.parent;

  if (paragraph.type.name !== 'paragraph') return null;
  if (!paragraph.attrs.numPr?.numId) return null;

  return {
    numId: paragraph.attrs.numPr.numId,
    ilvl: paragraph.attrs.numPr.ilvl || 0,
  };
}

// ============================================================================
// KEYMAP COMMANDS
// ============================================================================

function exitListOnEmptyEnter(): Command {
  return (state, dispatch) => {
    const { $from, empty } = state.selection;
    if (!empty) return false;

    const paragraph = $from.parent;
    if (paragraph.type.name !== 'paragraph') return false;

    const numPr = paragraph.attrs.numPr;
    if (!numPr) return false;

    // Use content.size instead of textContent.length — atomic inline nodes
    // (e.g. context tags) don't contribute to textContent but are real content.
    // Without this, Enter on a list item containing only a context tag would
    // strip the numbering instead of splitting the paragraph.
    if (paragraph.content.size > 0) return false;

    if (dispatch) {
      const tr = state.tr.setNodeMarkup($from.before(), undefined, {
        ...paragraph.attrs,
        numPr: null,
        listIsBullet: null,
        listNumFmt: null,
        listMarker: null,
      });
      dispatch(tr);
    }
    return true;
  };
}

function splitListItem(): Command {
  return (state, dispatch) => {
    const { $from, empty } = state.selection;
    if (!empty) return false;

    const paragraph = $from.parent;
    if (paragraph.type.name !== 'paragraph') return false;

    const numPr = paragraph.attrs.numPr;
    if (!numPr) return false;

    if (dispatch) {
      const { tr } = state;
      const pos = $from.pos;

      // Capture marks BEFORE split so we can preserve formatting on the new paragraph
      const preMarks: readonly Mark[] = state.storedMarks || $from.marks();

      // Clear identity attrs that must NOT be duplicated across paragraphs
      // (paraId, textId, bookmarks, _originalFormatting cause DOCX corruption if copied)
      const newAttrs = { ...paragraph.attrs };
      for (const key of ['paraId', 'textId', 'bookmarks', '_originalFormatting'] as const) {
        if ((newAttrs as Record<string, unknown>)[key] != null) {
          (newAttrs as Record<string, unknown>)[key] = null;
        }
      }

      tr.split(pos, 1, [{ type: state.schema.nodes.paragraph, attrs: newAttrs }]);

      // If new paragraph is empty (Enter at end of line), set storedMarks so typed
      // text inherits formatting (bold, font, size, etc.) from the source paragraph.
      const { $from: postFrom } = tr.selection;
      const newPara = postFrom.parent;
      if (newPara.textContent.length === 0) {
        let effectiveMarks: readonly Mark[] = preMarks;

        // If no explicit marks on cursor, derive from paragraph's defaultTextFormatting
        if (effectiveMarks.length === 0) {
          const dtf = paragraph.attrs.defaultTextFormatting as TextFormatting | undefined;
          if (dtf) {
            effectiveMarks = textFormattingToMarks(dtf, state.schema);
          }
        }

        if (effectiveMarks.length > 0) {
          tr.setStoredMarks(effectiveMarks);
        }
      }

      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

function backspaceExitList(): Command {
  return (state, dispatch) => {
    const { $from, empty } = state.selection;
    if (!empty) return false;

    if ($from.parentOffset !== 0) return false;

    const paragraph = $from.parent;
    if (paragraph.type.name !== 'paragraph') return false;

    const numPr = paragraph.attrs.numPr;
    if (!numPr) return false;

    if (dispatch) {
      const tr = state.tr.setNodeMarkup($from.before(), undefined, {
        ...paragraph.attrs,
        numPr: null,
        listIsBullet: null,
        listNumFmt: null,
        listMarker: null,
      });
      dispatch(tr);
    }
    return true;
  };
}

/**
 * If the paragraph is a Heading in a numbered list, promote it to the next
 * heading level (Heading1 → Heading2) on Tab, or demote (Heading2 → Heading1)
 * on Shift-Tab. Dispatches a custom DOM event so DocxEditor can apply the
 * style through its full resolution pipeline (same as Alt+digit shortcuts).
 * Returns false if not a heading in a list.
 */
function promoteHeadingLevel(): Command {
  return (state, _dispatch) => {
    const { $from } = state.selection;
    const paragraph = $from.parent;
    if (paragraph.type.name !== 'paragraph') return false;

    const styleId = paragraph.attrs.styleId as string | null;
    const numPr = paragraph.attrs.numPr;
    if (!styleId || !numPr) return false;

    const headingMatch = styleId.match(/^Heading(\d)$/);
    if (!headingMatch) return false;

    const level = parseInt(headingMatch[1]);
    if (level >= 9) return false;

    // Dispatch a DOM event for DocxEditor to pick up and apply through
    // its full style resolution pipeline (same as Alt+digit or StylePicker).
    document.dispatchEvent(
      new CustomEvent('docx-apply-style', { detail: { styleId: `Heading${level + 1}` } })
    );
    return true;
  };
}

function demoteHeadingLevel(): Command {
  return (state, _dispatch) => {
    const { $from } = state.selection;
    const paragraph = $from.parent;
    if (paragraph.type.name !== 'paragraph') return false;

    const styleId = paragraph.attrs.styleId as string | null;
    const numPr = paragraph.attrs.numPr;
    if (!styleId || !numPr) return false;

    const headingMatch = styleId.match(/^Heading(\d)$/);
    if (!headingMatch) return false;

    const level = parseInt(headingMatch[1]);
    if (level <= 1) return true; // Already at Heading1 — consume the keystroke, don't remove list

    document.dispatchEvent(
      new CustomEvent('docx-apply-style', { detail: { styleId: `Heading${level - 1}` } })
    );
    return true;
  };
}

function increaseListIndent(): Command {
  return (state, dispatch) => {
    const { $from } = state.selection;
    const paragraph = $from.parent;

    if (paragraph.type.name !== 'paragraph') return false;

    const numPr = paragraph.attrs.numPr;
    if (!numPr) return false;

    const currentLevel = numPr.ilvl ?? 0;
    if (currentLevel >= 8) return false;

    if (dispatch) {
      const tr = state.tr.setNodeMarkup($from.before(), undefined, {
        ...paragraph.attrs,
        numPr: { ...numPr, ilvl: currentLevel + 1 },
        // Clear explicit indentation so layout engine computes from new level
        indentLeft: null,
        indentFirstLine: null,
        hangingIndent: null,
      });
      dispatch(tr);
    }
    return true;
  };
}

function decreaseListIndent(): Command {
  return (state, dispatch) => {
    const { $from } = state.selection;
    const paragraph = $from.parent;

    if (paragraph.type.name !== 'paragraph') return false;

    const numPr = paragraph.attrs.numPr;
    if (!numPr) return false;

    const currentLevel = numPr.ilvl ?? 0;
    if (currentLevel <= 0) {
      if (dispatch) {
        const tr = state.tr.setNodeMarkup($from.before(), undefined, {
          ...paragraph.attrs,
          numPr: null,
          listIsBullet: null,
          listNumFmt: null,
          listMarker: null,
          // Clear list-specific indentation when removing list
          indentLeft: null,
          indentFirstLine: null,
          hangingIndent: null,
        });
        dispatch(tr);
      }
      return true;
    }

    if (dispatch) {
      const tr = state.tr.setNodeMarkup($from.before(), undefined, {
        ...paragraph.attrs,
        numPr: { ...numPr, ilvl: currentLevel - 1 },
        // Clear explicit indentation so layout engine computes from new level
        indentLeft: null,
        indentFirstLine: null,
        hangingIndent: null,
      });
      dispatch(tr);
    }
    return true;
  };
}

function insertTab(): Command {
  return (state, dispatch) => {
    const { schema } = state;
    const tabType = schema.nodes.tab;

    if (!tabType) {
      return false;
    }

    if (dispatch) {
      const tr = state.tr.replaceSelectionWith(tabType.create());
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

// Import goToNextCell/goToPrevCell from table extension for chaining
import { goToNextCell, goToPrevCell } from '../nodes/TableExtension';

// ============================================================================
// EXTENSION
// ============================================================================

export const ListExtension = createExtension({
  name: 'list',
  priority: Priority.High, // Must be before base keymap
  onSchemaReady(): ExtensionRuntime {
    return {
      commands: {
        toggleBulletList: () => toggleBulletList,
        toggleNumberedList: () => toggleNumberedList,
        increaseListLevel: () => increaseListLevel,
        decreaseListLevel: () => decreaseListLevel,
        removeList: () => removeList,
      },
      keyboardShortcuts: {
        Tab: chainCommands(
          goToNextCell(),
          promoteHeadingLevel(),
          increaseListIndent(),
          insertTab()
        ),
        'Shift-Tab': chainCommands(goToPrevCell(), demoteHeadingLevel(), decreaseListIndent()),
        'Shift-Enter': () => false, // Let base keymap handle this
        Enter: chainCommands(exitListOnEmptyEnter(), splitListItem()),
        Backspace: backspaceExitList(),
      },
    };
  },
});
