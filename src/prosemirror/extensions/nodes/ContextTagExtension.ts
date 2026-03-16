/**
 * Context Tag Extension — inline placeholder for context variables
 *
 * Renders as a highlighted badge in the editor (e.g., "{case_no}").
 * When context data is available, shows the resolved value.
 * Stored in the DOCX as {tag_key} text (compatible with docxtemplater).
 *
 * Each instance carries a unique `metaId` (UUID) so that per-instance
 * properties (like removeIfEmpty) survive DOCX round-trips even when
 * the same tagKey appears multiple times in the document.
 */

import { createNodeExtension } from '../create';

/** Generate a compact UUID v4 (crypto.randomUUID with fallback). */
function generateMetaId(): string {
  if (typeof crypto !== 'undefined' && crypto.randomUUID) {
    return crypto.randomUUID();
  }
  // Fallback for older environments
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    return (c === 'x' ? r : (r & 0x3) | 0x8).toString(16);
  });
}

export { generateMetaId };

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
      /** If true, the entire parent paragraph is removed when this tag has no value */
      removeIfEmpty: { default: false },
      /** If true, the entire table row containing this tag is removed when it has no value */
      removeTableRow: { default: false },
      /** Unique identifier for this tag instance — links to Custom XML Part metadata */
      metaId: { default: '' },
    },
    parseDOM: [
      {
        tag: 'span.docx-context-tag',
        getAttrs(dom) {
          const el = dom as HTMLElement;
          return {
            tagKey: el.dataset.tagKey || '',
            label: el.textContent || '',
            removeIfEmpty: el.dataset.removeIfEmpty === 'true',
            removeTableRow: el.dataset.removeTableRow === 'true',
            metaId: el.dataset.metaId || generateMetaId(),
          };
        },
      },
    ],
    toDOM(node) {
      const { tagKey, label, removeIfEmpty, removeTableRow, metaId } = node.attrs as {
        tagKey: string;
        label: string;
        removeIfEmpty: boolean;
        removeTableRow: boolean;
        metaId: string;
      };
      const displayText = String(label || `{${tagKey}}`);
      const attrs: Record<string, string> = {
        class: 'docx-context-tag',
        'data-tag-key': tagKey,
        // No background/border styling here — this DOM is the hidden
        // ProseMirror (off-screen). Any background-color leaks to adjacent
        // text during contenteditable mutations, which HighlightExtension's
        // parseDOM then picks up as a highlight mark, causing the "blue
        // background" bug. Visible rendering is handled by layout-painter.
        style: 'white-space: nowrap; cursor: default;',
      };
      if (removeIfEmpty) {
        attrs['data-remove-if-empty'] = 'true';
      }
      if (removeTableRow) {
        attrs['data-remove-table-row'] = 'true';
      }
      if (metaId) {
        attrs['data-meta-id'] = metaId;
      }
      return ['span', attrs, displayText];
    },
  },
});
