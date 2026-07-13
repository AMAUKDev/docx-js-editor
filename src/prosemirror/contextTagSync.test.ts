import { describe, expect, it } from 'bun:test';
import { EditorState } from 'prosemirror-state';
import { Schema } from 'prosemirror-model';
import {
  CONTEXT_TAG_LABEL_SYNC_META,
  isContextTagLabelSync,
  markContextTagLabelSync,
} from './contextTagSync';

const schema = new Schema({
  nodes: {
    doc: { content: 'paragraph+' },
    paragraph: { content: 'text*' },
    text: {},
  },
});

function makeState(): EditorState {
  return EditorState.create({
    doc: schema.node('doc', null, [schema.node('paragraph', null, [schema.text('hello')])]),
  });
}

describe('contextTagSync transaction meta', () => {
  it('markContextTagLabelSync sets the sync meta and excludes from history', () => {
    const tr = markContextTagLabelSync(makeState().tr);
    expect(tr.getMeta(CONTEXT_TAG_LABEL_SYNC_META)).toBe(true);
    expect(tr.getMeta('addToHistory')).toBe(false);
  });

  it('isContextTagLabelSync is true only for marked transactions', () => {
    const state = makeState();
    expect(isContextTagLabelSync(markContextTagLabelSync(state.tr))).toBe(true);
    expect(isContextTagLabelSync(state.tr)).toBe(false);
  });

  it('a marked doc-changing transaction still reports docChanged', () => {
    const state = makeState();
    const tr = markContextTagLabelSync(state.tr.insertText('x', 1));
    expect(tr.docChanged).toBe(true);
    expect(isContextTagLabelSync(tr)).toBe(true);
  });
});
