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
    // Allow formatting marks (bold, italic, font, size, etc.) but exclude
    // textColor and hyperlink. Theme-based textColor marks on inline atoms
    // can diverge from surrounding text during editing, causing context tags
    // to render in a different color (typically blue from accent1 fallback).
    // By excluding textColor, context tags inherit color from the parent
    // element — which is the correct default behavior for template variables.
    marks:
      'bold italic underline strike fontSize fontFamily superscript subscript ' +
      'allCaps smallCaps characterSpacing emboss imprint textShadow ' +
      'emphasisMark textOutline footnoteRef',
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
      const displayText = String(label || `{${tagKey}}`);
      return [
        'span',
        {
          class: 'docx-context-tag',
          'data-tag-key': tagKey,
          // No background/border styling here — this DOM is the hidden
          // ProseMirror (off-screen). Any background-color leaks to adjacent
          // text during contenteditable mutations, which HighlightExtension's
          // parseDOM then picks up as a highlight mark, causing the "blue
          // background" bug. Visible rendering is handled by layout-painter.
          style: 'white-space: nowrap; cursor: default;',
        },
        displayText,
      ];
    },
  },
});
