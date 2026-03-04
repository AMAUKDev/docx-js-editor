/**
 * Cross-Reference Extension — inline reference to heading/figure numbers
 *
 * Renders as the resolved number text (e.g., "1.1" or "Figure 3").
 * When the referenced heading/figure is renumbered, the display updates
 * automatically because toFlowBlocks resolves the number at render time.
 */

import { createNodeExtension } from '../create';

export const CrossRefExtension = createNodeExtension({
  name: 'crossRef',
  schemaNodeName: 'crossRef',
  nodeSpec: {
    inline: true,
    group: 'inline',
    atom: true,
    selectable: true,
    attrs: {
      /** Reference type: "heading" or "figure" */
      refType: { default: 'heading' },
      /** The text content of the referenced heading or caption (used for matching) */
      refTarget: { default: '' },
      /** Current display text (resolved number, e.g., "1.1" or "Figure 3") */
      displayText: { default: '' },
    },
    parseDOM: [
      {
        tag: 'span.docx-cross-ref',
        getAttrs(dom) {
          const el = dom as HTMLElement;
          return {
            refType: el.dataset.refType || 'heading',
            refTarget: el.dataset.refTarget || '',
            displayText: el.textContent || '',
          };
        },
      },
    ],
    toDOM(node) {
      const { refType, refTarget, displayText } = node.attrs as {
        refType: string;
        refTarget: string;
        displayText: string;
      };
      const text = displayText || `[${refType}: ${refTarget || '?'}]`;
      return [
        'span',
        {
          class: `docx-cross-ref docx-cross-ref-${refType}`,
          'data-ref-type': refType,
          'data-ref-target': refTarget,
          style:
            'color: #1a73e8; text-decoration: underline; cursor: pointer; ' +
            'font-family: inherit;',
        },
        text,
      ];
    },
  },
});
