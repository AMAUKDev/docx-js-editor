import { describe, expect, it } from 'bun:test';
import { EditorState } from 'prosemirror-state';
import { Schema, type Node as PMNode } from 'prosemirror-model';
import {
  buildCaption,
  appendCaptionToTransaction,
  countCaptionsBefore,
  findCaptionAnchor,
  isCaptionPrefix,
  CAPTION_STYLE_ID,
} from './captionBuilder';

// A schema carrying only what the caption builder touches: paragraphs with a styleId,
// an atomic `field` node for the SEQ marker, and a table (a caption can anchor to one).
const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: {
      group: 'block',
      content: 'inline*',
      attrs: { styleId: { default: null }, alignment: { default: null } },
    },
    table: { group: 'block', content: 'paragraph*' },
    field: {
      group: 'inline',
      inline: true,
      atom: true,
      attrs: {
        fieldType: { default: null },
        instruction: { default: null },
        displayText: { default: null },
        fieldKind: { default: null },
        dirty: { default: false },
      },
      // textContent of an atom comes from its leafText
      leafText: (node: PMNode) => String(node.attrs.displayText ?? ''),
    },
    text: { group: 'inline' },
  },
});

const para = (text: string, styleId: string | null = null) =>
  schema.nodes.paragraph.create({ styleId }, text ? [schema.text(text)] : []);

/** A caption paragraph as it exists in a real document: "Prefix N: description". */
const captionPara = (prefix: string, number: number, description: string) =>
  schema.nodes.paragraph.create({ styleId: CAPTION_STYLE_ID }, [
    schema.text(`${prefix} `),
    schema.nodes.field.create({
      fieldType: 'SEQ',
      instruction: ` SEQ ${prefix} \\* ARABIC `,
      displayText: String(number),
      fieldKind: 'complex',
    }),
    schema.text(`: ${description}`),
  ]);

const stateOf = (...nodes: PMNode[]) =>
  EditorState.create({ doc: schema.nodes.doc.create(null, nodes) });

/** Position inside the first paragraph whose text matches. */
const posInside = (state: EditorState, text: string) => {
  let pos = -1;
  state.doc.forEach((node, offset) => {
    if (pos === -1 && node.textContent === text) pos = offset + 1;
  });
  return pos;
};

describe('isCaptionPrefix', () => {
  it('accepts only the prefixes crossRefUpdater renumbers', () => {
    expect(isCaptionPrefix('Figure')).toBe(true);
    expect(isCaptionPrefix('Table')).toBe(true);
    // Anything else would render a number that never updates.
    expect(isCaptionPrefix('Exhibit')).toBe(false);
    expect(isCaptionPrefix('figure')).toBe(false);
  });
});

describe('findCaptionAnchor', () => {
  it('anchors to the enclosing paragraph', () => {
    const state = stateOf(para('First'), para('Second'));
    const anchor = findCaptionAnchor(state, posInside(state, 'Second'));
    expect(anchor?.node.textContent).toBe('Second');
  });

  // Depth order matters: a position inside a table cell resolves as
  // paragraph <- table, and the innermost block wins. A caption anchored on a cell
  // therefore attaches to that cell's paragraph, matching the toolbar's behaviour.
  it('anchors to the innermost block when the position is inside a table cell', () => {
    const table = schema.nodes.table.create(null, [para('cell')]);
    const state = stateOf(para('Intro'), table);
    let tablePos = -1;
    state.doc.forEach((node, offset) => {
      if (node.type.name === 'table') tablePos = offset;
    });
    const anchor = findCaptionAnchor(state, tablePos + 2);
    expect(anchor?.node.type.name).toBe('paragraph');
    expect(anchor?.node.textContent).toBe('cell');
  });

  // A position ON the table node itself (not inside a cell) anchors to the table, which
  // is what a caption for a whole table needs.
  it('anchors to the table itself for a position at its boundary', () => {
    const table = schema.nodes.table.create(null, [para('cell')]);
    const state = stateOf(para('Intro'), table);
    let tablePos = -1;
    state.doc.forEach((node, offset) => {
      if (node.type.name === 'table') tablePos = offset;
    });
    const anchor = findCaptionAnchor(state, tablePos + 1);
    expect(anchor?.node.type.name).toBe('table');
  });

  it('returns null for a position outside the document', () => {
    const state = stateOf(para('Only'));
    expect(findCaptionAnchor(state, 9999)).toBeNull();
  });
});

describe('countCaptionsBefore', () => {
  it('counts only captions sharing the prefix', () => {
    const state = stateOf(
      captionPara('Figure', 1, 'A photo'),
      captionPara('Table', 1, 'Some data'),
      captionPara('Figure', 2, 'Another photo'),
      para('Body text')
    );
    expect(countCaptionsBefore(state, state.doc.content.size, 'Figure')).toBe(2);
    expect(countCaptionsBefore(state, state.doc.content.size, 'Table')).toBe(1);
  });

  it('ignores captions after the insertion point', () => {
    const state = stateOf(captionPara('Figure', 1, 'First'), captionPara('Figure', 2, 'Second'));
    let secondPos = 0;
    let seen = 0;
    state.doc.forEach((_node, offset) => {
      seen++;
      if (seen === 2) secondPos = offset;
    });
    expect(countCaptionsBefore(state, secondPos, 'Figure')).toBe(1);
  });

  it('does not count ordinary paragraphs that merely start with the word', () => {
    const state = stateOf(para('Figure it out later'), para('Body'));
    expect(countCaptionsBefore(state, state.doc.content.size, 'Figure')).toBe(0);
  });
});

describe('buildCaption', () => {
  it('builds "Prefix N: " with a real SEQ field', () => {
    const state = stateOf(para('An image paragraph'));
    const built = buildCaption(state, posInside(state, 'An image paragraph'), 'Figure');
    expect(built).not.toBeNull();
    expect(built!.node.attrs.styleId).toBe(CAPTION_STYLE_ID);

    const field = built!.node.child(1);
    expect(field.type.name).toBe('field');
    expect(field.attrs.fieldType).toBe('SEQ');
    expect(field.attrs.instruction).toBe(' SEQ Figure \\* ARABIC ');
    expect(field.attrs.displayText).toBe('1');
  });

  it('numbers after the captions already present', () => {
    const state = stateOf(
      captionPara('Table', 1, 'Existing'),
      para('A table would sit here')
    );
    const built = buildCaption(state, posInside(state, 'A table would sit here'), 'Table');
    expect(built!.number).toBe(2);
    expect(built!.node.child(1).attrs.displayText).toBe('2');
  });

  it('counts each prefix independently', () => {
    const state = stateOf(captionPara('Figure', 1, 'A photo'), para('A table here'));
    const built = buildCaption(state, posInside(state, 'A table here'), 'Table');
    expect(built!.number).toBe(1);
  });

  it('inserts after the anchor block, not inside it', () => {
    const state = stateOf(para('Anchor'), para('Following'));
    const anchor = findCaptionAnchor(state, posInside(state, 'Anchor'))!;
    const built = buildCaption(state, posInside(state, 'Anchor'), 'Figure');
    expect(built!.insertPos).toBe(anchor.pos + anchor.node.nodeSize);
  });

  it('appends description text after the separator when given', () => {
    const state = stateOf(para('Anchor'));
    const built = buildCaption(state, posInside(state, 'Anchor'), 'Table', 'Fuel test results');
    expect(built!.node.textContent).toBe('Table 1: Fuel test results');
  });

  it('leaves the caption open when no text is given', () => {
    const state = stateOf(para('Anchor'));
    const built = buildCaption(state, posInside(state, 'Anchor'), 'Figure');
    expect(built!.node.textContent).toBe('Figure 1: ');
  });

  it('returns null when the position resolves to no block', () => {
    const state = stateOf(para('Anchor'));
    expect(buildCaption(state, 9999, 'Figure')).toBeNull();
  });
});

describe('appendCaptionToTransaction', () => {
  it('adds the caption to the document', () => {
    const state = stateOf(para('Anchor'));
    const tr = appendCaptionToTransaction(state.tr, state, posInside(state, 'Anchor'), 'Table');
    expect(tr).not.toBeNull();
    const texts: string[] = [];
    tr!.doc.forEach((node) => texts.push(node.textContent));
    expect(texts).toEqual(['Anchor', 'Table 1: ']);
  });

  it('does not move the selection unless asked', () => {
    const state = stateOf(para('Anchor'));
    const tr = appendCaptionToTransaction(state.tr, state, posInside(state, 'Anchor'), 'Figure');
    expect(tr!.selectionSet).toBe(false);
  });

  it('places the caret after the separator when asked', () => {
    const state = stateOf(para('Anchor'));
    const tr = appendCaptionToTransaction(
      state.tr,
      state,
      posInside(state, 'Anchor'),
      'Figure',
      undefined,
      undefined,
      true
    );
    expect(tr!.selectionSet).toBe(true);
  });

  // The point of the non-dispatching design: several captions compose into ONE
  // transaction, so an agent's batch of edits stays a single undo/reject step.
  it('composes multiple captions into a single transaction', () => {
    const state = stateOf(para('First'), para('Second'));
    let tr = appendCaptionToTransaction(state.tr, state, posInside(state, 'Second'), 'Figure')!;
    tr = appendCaptionToTransaction(tr, state, posInside(state, 'First'), 'Table')!;
    const texts: string[] = [];
    tr.doc.forEach((node) => texts.push(node.textContent));
    expect(texts).toEqual(['First', 'Table 1: ', 'Second', 'Figure 1: ']);
  });

  it('returns null when no caption could be built', () => {
    const state = stateOf(para('Anchor'));
    expect(appendCaptionToTransaction(state.tr, state, 9999, 'Figure')).toBeNull();
  });
});
