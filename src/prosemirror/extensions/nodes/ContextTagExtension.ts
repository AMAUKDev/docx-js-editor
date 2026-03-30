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
/**
 * Generate a short unique ID for context tag metadata.
 * Uses 8 hex chars (32 bits of randomness) — sufficient for document-level
 * uniqueness while keeping bookmark names under Word's 40-character limit.
 * (_FP_ctx_ prefix = 9 chars + 8 hex = 17 chars total, well under 40)
 *
 * IMPORTANT: Do NOT use full UUIDs — Word truncates bookmark names longer
 * than ~40 characters, breaking the round-trip tag restoration system.
 */
function generateMetaId(): string {
  if (typeof crypto !== 'undefined' && crypto.getRandomValues) {
    const buf = new Uint8Array(4);
    crypto.getRandomValues(buf);
    return Array.from(buf, (b) => b.toString(16).padStart(2, '0')).join('');
  }
  // Fallback
  return Math.floor(Math.random() * 0xffffffff)
    .toString(16)
    .padStart(8, '0');
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
      /** Preview URL for image-type context tags (rendered as <img> in the editor) */
      imageUrl: { default: '' },
      /** Display width in points for image-type context tags (0 = auto) */
      imageWidth: { default: 0 },
      /** If true, tag placeholder is preserved even when downloading with "Remove unknown tags" */
      alwaysShow: { default: false },
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
