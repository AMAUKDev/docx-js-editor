/**
 * Loop Block Extension — block node representing a docxtemplater loop delimiter
 *
 * Represents `{% for x in y %}` and `{% endfor %}` paragraphs in the editor.
 * Rendered as a visually distinct coloured block so template authors can see
 * loop boundaries. Not directly editable — treated as atomic structural markers.
 *
 * Round-trip: exports back to the original {% %} text paragraph.
 */

import { createNodeExtension } from '../create';

export const LoopBlockExtension = createNodeExtension({
  name: 'loopBlock',
  schemaNodeName: 'loopBlock',
  nodeSpec: {
    group: 'block',
    atom: true,
    selectable: true,
    attrs: {
      /** The loop expression, e.g. "photo in photos". Empty string for endfor. */
      loopExpr: { default: '' },
      /** 'for' for opening delimiter, 'endfor' for closing. */
      kind: { default: 'for' as 'for' | 'endfor' },
    },
    parseDOM: [
      {
        tag: 'div.docx-loop-block',
        getAttrs(dom) {
          const el = dom as HTMLElement;
          return {
            loopExpr: el.dataset.loopExpr || '',
            kind: el.dataset.kind || 'for',
          };
        },
      },
    ],
    toDOM(node) {
      const { loopExpr, kind } = node.attrs as { loopExpr: string; kind: string };
      const label = kind === 'for' ? `for ${loopExpr}` : 'end for';
      return [
        'div',
        {
          class: 'docx-loop-block',
          'data-loop-expr': loopExpr,
          'data-kind': kind,
          style:
            'padding: 4px 8px; border-radius: 4px; font-size: 12px; font-family: monospace; cursor: default;',
        },
        label,
      ];
    },
  },
});
