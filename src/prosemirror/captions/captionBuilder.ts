/**
 * Caption construction.
 *
 * A caption is a paragraph styled `Caption` containing a real Word SEQ field:
 *
 *     text("Table ") + field(SEQ Table \* ARABIC) + text(": ")
 *
 * The SEQ field is what makes numbering behave like Word's — `crossRefUpdater`
 * re-resolves every caption's number on each document change, so a caption inserted
 * in the middle of a document renumbers the ones after it and any cross-references
 * that point at them.
 *
 * These functions build the caption and locate where it goes; they never dispatch.
 * That lets one caller (the toolbar) dispatch immediately while another (an agent
 * applying a batch of edits) folds several inserts into a single transaction, which
 * keeps undo/reject a single step.
 */
import { TextSelection } from 'prosemirror-state';
import type { EditorState, Transaction } from 'prosemirror-state';
import type { Mark, Node as PMNode, Schema } from 'prosemirror-model';
import { createStyleResolver } from '../styles/styleResolver';
import { textFormattingToMarks } from '../extensions/marks/markUtils';
import type { StyleDefinitions } from '../../types/styles';

/**
 * Caption prefixes that carry their own SEQ counter.
 *
 * Kept in step with CAPTION_PREFIXES in
 * `src/prosemirror/plugins/crossRefUpdater.ts` — a prefix missing there is never
 * renumbered, so it would render a number that silently goes stale.
 */
export const CAPTION_PREFIXES = ['Figure', 'Table'] as const;
export type CaptionPrefix = (typeof CAPTION_PREFIXES)[number];

/** The paragraph style every caption carries. */
export const CAPTION_STYLE_ID = 'Caption';

export function isCaptionPrefix(value: string): value is CaptionPrefix {
  return (CAPTION_PREFIXES as readonly string[]).includes(value);
}

/**
 * The block a caption attaches to: the nearest enclosing paragraph or table.
 *
 * Returns the block's position and node, or null when `pmPos` resolves outside one.
 */
export function findCaptionAnchor(
  state: EditorState,
  pmPos: number
): { pos: number; node: PMNode } | null {
  if (pmPos < 0 || pmPos > state.doc.content.size) return null;
  const $pos = state.doc.resolve(pmPos);
  for (let depth = $pos.depth; depth >= 1; depth--) {
    const pos = $pos.before(depth);
    const node = state.doc.nodeAt(pos);
    if (node?.type.name === 'paragraph' || node?.type.name === 'table') {
      return { pos, node };
    }
  }
  const fallbackPos = $pos.before($pos.depth);
  const fallbackNode = state.doc.nodeAt(fallbackPos);
  return fallbackNode ? { pos: fallbackPos, node: fallbackNode } : null;
}

/**
 * How many captions of the same prefix already exist before `beforePos`.
 *
 * This is the caption's number at insertion time. `crossRefUpdater` corrects it
 * afterwards, so it only has to be right enough to render before the next
 * transaction lands.
 */
export function countCaptionsBefore(
  state: EditorState,
  beforePos: number,
  prefix: string
): number {
  let count = 0;
  state.doc.nodesBetween(0, beforePos, (node) => {
    if (node.type.name === 'paragraph' && node.attrs.styleId === CAPTION_STYLE_ID) {
      if (node.textContent.startsWith(prefix + ' ')) count++;
    }
    return true;
  });
  return count;
}

/** Run marks carried by the `Caption` style, so the caption matches the house style. */
function resolveCaptionMarks(schema: Schema, styles: StyleDefinitions | null | undefined): Mark[] {
  if (!styles) return [];
  const resolved = createStyleResolver(styles).resolveParagraphStyle(CAPTION_STYLE_ID);
  return resolved.runFormatting ? textFormattingToMarks(resolved.runFormatting, schema) : [];
}

export interface BuiltCaption {
  /** The caption paragraph, ready to insert. */
  node: PMNode;
  /** Document position the caption should be inserted at. */
  insertPos: number;
  /** Position of the caret after ": ", where a description would be typed. */
  cursorPos: number;
  /** Marks the caption's text carries, for continued typing. */
  marks: Mark[];
  /** The number rendered at insertion time. */
  number: number;
}

/**
 * Build a caption paragraph for the block containing `pmPos`, without touching the
 * document. Returns null when no anchor block can be found.
 *
 * `text` is appended after the ": " separator; omit it to leave the caption open for
 * the user to type into.
 */
export function buildCaption(
  state: EditorState,
  pmPos: number,
  prefix: CaptionPrefix = 'Figure',
  text?: string,
  styles?: StyleDefinitions | null
): BuiltCaption | null {
  const anchor = findCaptionAnchor(state, pmPos);
  if (!anchor) return null;

  const schema = state.schema;
  if (!schema.nodes.field || !schema.nodes.paragraph) return null;

  const insertPos = anchor.pos + anchor.node.nodeSize;
  const number = countCaptionsBefore(state, insertPos, prefix) + 1;
  const marks = resolveCaptionMarks(schema, styles);

  const withMarks = (value: string) =>
    marks.length > 0 ? schema.text(value, marks) : schema.text(value);

  let seqField = schema.nodes.field.create({
    fieldType: 'SEQ',
    instruction: ` SEQ ${prefix} \\* ARABIC `,
    displayText: String(number),
    fieldKind: 'complex',
    dirty: false,
  });
  if (marks.length > 0) seqField = seqField.mark(marks);

  const content: PMNode[] = [withMarks(prefix + ' '), seqField, withMarks(': ')];
  const trailing = (text ?? '').trim();
  if (trailing) content.push(withMarks(trailing));

  const node = schema.nodes.paragraph.create(
    { styleId: CAPTION_STYLE_ID, alignment: 'center' },
    content
  );

  // paragraph open(1) + prefix text + space(1) + field atom(1) + ": "(2)
  const cursorPos = insertPos + 1 + prefix.length + 1 + 1 + 2;

  return { node, insertPos, cursorPos, marks, number };
}

/**
 * Append a caption insert to `tr` and return it, or null when no caption could be
 * built. The caret and stored marks are only set when `placeCursor` is true — an
 * agent inserting several blocks at once should leave the selection alone.
 */
export function appendCaptionToTransaction(
  tr: Transaction,
  state: EditorState,
  pmPos: number,
  prefix: CaptionPrefix = 'Figure',
  text?: string,
  styles?: StyleDefinitions | null,
  placeCursor = false
): Transaction | null {
  const built = buildCaption(state, pmPos, prefix, text, styles);
  if (!built) return null;

  const next = tr.insert(built.insertPos, built.node);
  if (placeCursor) {
    next.setSelection(TextSelection.create(next.doc, built.cursorPos));
    if (built.marks.length > 0) next.setStoredMarks(built.marks);
  }
  return next;
}
