/**
 * Integration tests for the ordinal-suffix auto-superscript plugin, against a real
 * (minimal) ProseMirror schema/state — not mocks. Mirrors the toy-schema convention
 * used by ParagraphChangeTrackerExtension.test.ts.
 */

import { describe, test, expect } from 'bun:test';
import { Schema } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import { createOrdinalSuffixPlugin, collectInsertedRanges } from './ordinalSuffixPlugin';

const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: { group: 'block', content: 'inline*', toDOM: () => ['p', 0] },
    text: { group: 'inline' },
  },
  marks: {
    superscript: {
      parseDOM: [{ tag: 'sup' }],
      toDOM: () => ['sup', 0],
    },
  },
});

function createState(text: string): EditorState {
  const doc = schema.node('doc', null, [
    schema.node('paragraph', null, text ? [schema.text(text)] : []),
  ]);
  return EditorState.create({ doc, plugins: [createOrdinalSuffixPlugin()] });
}

/**
 * Types `text` one character at a time (each char its own transaction), moving the
 * SELECTION first and then calling `insertText(ch)` with no explicit position — this
 * is `Transaction.insertText`'s selection-based branch (`replaceSelectionWith`,
 * `inheritMarks: true`), the same one real single-keystroke input in a live view goes
 * through. The position-based `insertText(ch, pos)` branch used here previously does
 * NOT exercise the same mark-inheritance path and silently missed a real bug (marks
 * leaking onto text typed immediately after an auto-superscripted suffix) that only
 * showed up live in the browser — see ordinalSuffixPlugin.ts's `removeStoredMark` fix.
 */
function type(state: EditorState, text: string, at?: number): EditorState {
  let s = state;
  const startPos = at ?? s.doc.content.size - 1; // inside the (only) paragraph
  s = s.apply(s.tr.setSelection(TextSelection.create(s.doc, startPos)));
  for (const ch of text) {
    const tr = s.tr.insertText(ch);
    s = s.apply(tr);
  }
  return s;
}

/** Inserts `text` as a single transaction, like a paste or an AMAi content op. */
function insertBulk(state: EditorState, text: string, at?: number): EditorState {
  const pos = at ?? state.doc.content.size - 1;
  return state.apply(state.tr.insertText(text, pos));
}

function superscriptRanges(state: EditorState): string[] {
  const out: string[] = [];
  state.doc.descendants((node) => {
    if (!node.isText) return;
    if (node.marks.some((m) => m.type === schema.marks.superscript)) {
      out.push(node.text ?? '');
    }
  });
  return out;
}

describe('createOrdinalSuffixPlugin', () => {
  test('typing "3rd " char-by-char superscripts "rd" only once the space is typed', () => {
    let state = createState('the ');
    state = type(state, '3rd');
    // Not yet — no boundary character has been typed after "rd".
    expect(superscriptRanges(state)).toEqual([]);
    state = type(state, ' ');
    expect(superscriptRanges(state)).toEqual(['rd']);
  });

  test('does NOT superscript a code-like token typed char-by-char ("45th8gh")', () => {
    let state = createState('');
    state = type(state, '45th8gh ');
    expect(superscriptRanges(state)).toEqual([]);
  });

  test('a single bulk insertion (paste / AMAi) superscripts every valid ordinal inside it', () => {
    const state = insertBulk(createState(''), 'on the 21st of May, the 45th8gh code stays plain.');
    expect(superscriptRanges(state)).toEqual(['st']);
  });

  test('a manual revert holds — later typing elsewhere never re-superscripts it', () => {
    let state = createState('the ');
    state = type(state, '3rd ');
    expect(superscriptRanges(state)).toEqual(['rd']);

    // Manually revert: user selects "rd" and toggles Superscript off (removeMark).
    const supMark = schema.marks.superscript;
    let revertPos = -1;
    state.doc.descendants((node, pos) => {
      if (node.isText && node.text === 'rd') revertPos = pos;
    });
    expect(revertPos).toBeGreaterThan(-1);
    state = state.apply(state.tr.removeMark(revertPos, revertPos + 2, supMark));
    expect(superscriptRanges(state)).toEqual([]);

    // Keep typing elsewhere in the same paragraph, including another ordinal.
    state = type(state, 'and the 4th ', state.doc.content.size - 1);

    // The new "4th" gets superscripted; the manually-reverted "rd" stays plain forever.
    expect(superscriptRanges(state)).toEqual(['th']);
    expect(state.doc.textBetween(0, state.doc.content.size)).toContain('3rd');
  });

  test('does not leak the mark into text typed immediately after the suffix (model-level sanity check only)', () => {
    // KNOWN LIMITATION OF THIS TEST: it passes in this pure-model, DOM-free
    // environment REGARDLESS of whether `removeStoredMark` in ordinalSuffixPlugin.ts
    // is present — confirmed by deliberately disabling that line and re-running.
    // The real bug this guards (superscript leaking onto text typed immediately
    // after an auto-superscripted suffix, e.g. typing "the 6th time..." live
    // produced "<sup>th time...</sup>") only reproduces through a REAL browser's
    // contenteditable input path, not through `Transaction.insertText` called
    // directly against a model with no view/DOM — there is nothing here to exercise
    // whatever cursor/mark-boundary behavior the real DOM exhibits. The
    // `removeStoredMark` fix was verified live (typed "the 7th time we tried." in
    // the actual report editor after the fix and confirmed only "th" was marked,
    // versus a leak across "th time." before it) — that live check, not this test,
    // is the authoritative regression guard for this specific behavior. Recommend a
    // real Playwright/e2e test against the live editor as a follow-up so this has
    // automated coverage; kept here anyway as a (weaker) pure-model sanity check.
    let state = createState('');
    state = type(state, 'the 6th time we tried ');
    expect(superscriptRanges(state)).toEqual(['th']);
  });

  test('does not touch text that was never part of the inserted range', () => {
    // "9th" already exists untouched; typing far away must not retroactively mark it.
    let state = createState('page 9th of the report');
    expect(superscriptRanges(state)).toEqual([]);
    state = type(state, '!', state.doc.content.size - 1);
    expect(superscriptRanges(state)).toEqual([]);
  });

  test('collectInsertedRanges handles a multi-step transaction (delete then insert)', () => {
    // A single transaction with two steps — e.g. replacing a selection — must map
    // the insert step's range through the rest of the transaction correctly, not
    // just through its own step.
    const state = createState('the 3nd of May');
    // "3nd" sits at positions 5-8 inside the paragraph (pos 0 = doc start, 1 = para start).
    const from = 5;
    const to = 8;
    const tr = state.tr.delete(from, to).insertText('3rd', from);
    const ranges = collectInsertedRanges([tr]);
    expect(ranges).toHaveLength(1);
    expect(tr.doc.textBetween(ranges[0].from, ranges[0].to)).toBe('3rd');
  });

  test('a delete-then-insert transaction (replacing a selection) auto-superscripts when a real boundary already follows', () => {
    // Fixing a typo — selecting "nd" in "3nd" and typing "rd" — is itself an
    // insertion event. Since a genuine boundary (the pre-existing space before
    // "of") already sits right after the freshly-inserted "rd", the pattern is
    // completed by this edit and should format immediately, same as it would for
    // a paste or an AMAi replacement landing in the same spot.
    let state = createState('the 3nd of May ');
    const from = 5;
    const to = 8; // "3nd"
    state = state.apply(state.tr.delete(from, to).insertText('3rd', from));
    expect(superscriptRanges(state)).toEqual(['rd']);
  });
});
