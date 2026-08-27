import { describe, expect, it } from 'bun:test';
import { EditorState } from 'prosemirror-state';
import { Schema, type Node as PMNode } from 'prosemirror-model';
import { buildAutoTagTransaction, collectAutoTagHits } from './autoTagTransform';

const schema = new Schema({
  nodes: {
    doc: { content: 'paragraph+' },
    paragraph: { content: 'inline*', group: 'block' },
    text: { group: 'inline' },
    contextTag: {
      inline: true,
      group: 'inline',
      atom: true,
      attrs: {
        tagKey: { default: '' },
        label: { default: '' },
        removeIfEmpty: { default: false },
        metaId: { default: '' },
      },
    },
  },
});

function docWith(text: string): EditorState {
  const doc = schema.node('doc', null, [
    schema.node('paragraph', null, text ? [schema.text(text)] : []),
  ]);
  return EditorState.create({ schema, doc });
}

/** Read the (tagKey -> label) of every contextTag node in a doc. */
function tags(doc: PMNode): Array<{ tagKey: string; label: string }> {
  const out: Array<{ tagKey: string; label: string }> = [];
  doc.descendants((n) => {
    if (n.type.name === 'contextTag') out.push({ tagKey: n.attrs.tagKey, label: n.attrs.label });
  });
  return out;
}

const MAP = {
  'case.case_no': 'AMA6835',
  'lead_client.account.name': 'Britannia Hong Kong Limited',
};

describe('buildAutoTagTransaction', () => {
  it('converts a canonical value in loose text into a contextTag node', () => {
    const state = docWith('Our ref is AMA6835 today');
    const { tr, applied } = buildAutoTagTransaction(state, MAP);
    expect(applied).toHaveLength(1);
    expect(tr).not.toBeNull();
    const newDoc = state.apply(tr!).doc;
    const found = tags(newDoc);
    expect(found).toEqual([{ tagKey: 'case.case_no', label: 'AMA6835' }]);
    // literal text no longer present as loose text
    expect(newDoc.textBetween(0, newDoc.content.size, '')).not.toContain('AMA6835');
  });

  it('converts multiple distinct values in one transaction', () => {
    const state = docWith('AMA6835 acts for Britannia Hong Kong Limited');
    const { tr, applied } = buildAutoTagTransaction(state, MAP);
    expect(applied).toHaveLength(2);
    const found = tags(state.apply(tr!).doc).map((t) => t.tagKey).sort();
    expect(found).toEqual(['case.case_no', 'lead_client.account.name']);
  });

  it('defers low-confidence (ambiguous) matches by default', () => {
    const state = docWith('state is Pending');
    const { tr, applied, deferred } = buildAutoTagTransaction(state, {
      'case.status': 'Pending',
      'matter.state': 'Pending',
    });
    expect(applied).toHaveLength(0);
    expect(tr).toBeNull();
    expect(deferred).toHaveLength(1);
    expect(deferred[0].value).toBe('Pending');
  });

  it('applies low-confidence matches when accept() approves them (confirm step)', () => {
    const state = docWith('state is Pending');
    const { tr, applied } = buildAutoTagTransaction(
      state,
      { 'case.status': 'Pending', 'matter.state': 'Pending' },
      { accept: (h) => h.tagKey === 'case.status' },
    );
    expect(applied).toHaveLength(1);
    expect(tags(state.apply(tr!).doc)).toEqual([{ tagKey: 'case.status', label: 'Pending' }]);
  });

  it('respects rangeFrom/rangeTo bounding (typing use-case)', () => {
    const state = docWith('AMA6835 then later AMA6835');
    const all = collectAutoTagHits(state, MAP);
    expect(all).toHaveLength(2);
    // bound to only the second occurrence
    const second = all[1];
    const { applied } = buildAutoTagTransaction(state, MAP, {
      rangeFrom: second.from,
      rangeTo: second.to,
    });
    expect(applied).toHaveLength(1);
    expect(applied[0].from).toBe(second.from);
  });

  it('no-ops on an empty document', () => {
    const { tr, applied } = buildAutoTagTransaction(docWith(''), MAP);
    expect(tr).toBeNull();
    expect(applied).toHaveLength(0);
  });
});
