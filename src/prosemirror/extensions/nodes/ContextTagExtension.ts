/**
 * Context Tag Extension — inline placeholder for context variables
 *
 * Renders as a highlighted badge in the editor (e.g., "{case_no}").
 * When context data is available, shows the resolved value.
 * Stored in the DOCX as {tag_key} text (compatible with docxtemplater).
 */

import { createNodeExtension } from '../create';

export const ContextTagExtension = createNodeExtension({
  name: 'contextTag',
  schemaNodeName: 'contextTag',
  nodeSpec: {
    inline: true,
    group: 'inline',
    atom: true,
    selectable: true,
    attrs: {
      /** The context variable key (e.g., "case_no", "vessel", "client") */
      tagKey: { default: '' },
      /** Display label — resolved value or the tag key itself */
      label: { default: '' },
    },
    parseDOM: [
      {
        tag: 'span.docx-context-tag',
        getAttrs(dom) {
          const el = dom as HTMLElement;
          return {
            tagKey: el.dataset.tagKey || '',
            label: el.textContent || '',
          };
        },
      },
    ],
    toDOM(node) {
      const { tagKey, label } = node.attrs as { tagKey: string; label: string };
      const displayText = label || `{${tagKey}}`;
      return [
        'span',
        {
          class: 'docx-context-tag',
          'data-tag-key': tagKey,
          style:
            'background: #e8f0fe; color: #1a73e8; padding: 1px 6px; border-radius: 3px; ' +
            'font-size: 0.9em; font-family: system-ui, sans-serif; white-space: nowrap; ' +
            'border: 1px solid #c4d9f8; cursor: default;',
        },
        displayText,
      ];
    },
  },
});
