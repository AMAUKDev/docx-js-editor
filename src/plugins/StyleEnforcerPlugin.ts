/**
 * Style Enforcer ProseMirror Plugin
 *
 * Restricts paragraph formatting to a set of allowed style IDs.
 * - On every transaction, remaps disallowed styleId attrs to "Normal"
 * - On every transaction, strips disallowed inline marks (bold, italic, etc.)
 * - On paste, strips all disallowed marks from incoming content
 */

import { Plugin, PluginKey } from 'prosemirror-state';
import { Slice } from 'prosemirror-model';
import type { Node as ProseMirrorNode, Schema } from 'prosemirror-model';

export const styleEnforcerPluginKey = new PluginKey('styleEnforcer');

export const DEFAULT_ALLOWED_STYLE_IDS = [
  'Normal',
  'Heading1',
  'Heading2',
  'Heading3',
  'Heading4',
  'Heading5',
  'Heading6',
  'Heading7',
  'Heading8',
  'Heading9',
  'ListParagraph',
  'Caption',
  'Title',
  'Subtitle',
  'TOCHeading',
  'TOC1',
  'TOC2',
  'TOC3',
  'TOC4',
  'TOC5',
  'TOC6',
  'TOC7',
  'TOC8',
  'TOC9',
];

/** Mark names that are stripped by the style enforcer (manual formatting marks). */
const DISALLOWED_MARK_NAMES = [
  'fontFamily',
  'fontSize',
  'textColor',
  'highlight',
  'bold',
  'italic',
  'underline',
  'strike',
] as const;

/**
 * Strip all disallowed manual-formatting marks from a slice (used for paste filtering).
 */
function stripFontMarksFromSlice(slice: Slice, schema: Schema): Slice {
  const disallowedMarkTypes = DISALLOWED_MARK_NAMES.map((name) => schema.marks[name]).filter(
    Boolean
  );

  if (disallowedMarkTypes.length === 0) return slice;

  const mapped = mapFragment(slice.content, (node) => {
    if (node.isText) {
      let newNode = node;
      for (const markType of disallowedMarkTypes) {
        if (markType.isInSet(newNode.marks)) {
          newNode = newNode.mark(markType.removeFromSet(newNode.marks));
        }
      }
      return newNode;
    }
    return node;
  });

  return new Slice(mapped, slice.openStart, slice.openEnd);
}

/**
 * Recursively map all nodes in a Fragment.
 */
function mapFragment(
  fragment: import('prosemirror-model').Fragment,
  fn: (node: ProseMirrorNode) => ProseMirrorNode
): import('prosemirror-model').Fragment {
  const nodes: ProseMirrorNode[] = [];
  for (let i = 0; i < fragment.childCount; i++) {
    let child = fragment.child(i);
    child = fn(child);
    if (child.content.childCount > 0) {
      child = child.copy(mapFragment(child.content, fn));
    }
    nodes.push(child);
  }
  return (fragment as any).constructor.fromArray(nodes);
}

export interface StyleEnforcerOptions {
  allowedStyleIds?: string[];
}

/**
 * Creates the style enforcer plugin.
 */
export function createStyleEnforcerPlugin(options: StyleEnforcerOptions = {}): Plugin {
  const allowed = new Set(options.allowedStyleIds ?? DEFAULT_ALLOWED_STYLE_IDS);

  return new Plugin({
    key: styleEnforcerPluginKey,

    // appendTransaction: remap disallowed styleId to Normal + strip disallowed marks
    appendTransaction(transactions, _oldState, newState) {
      // Only process if something actually changed
      if (!transactions.some((tr) => tr.docChanged)) return null;

      const savedStoredMarks = newState.storedMarks;
      const tr = newState.tr;
      let changed = false;

      // Resolve disallowed mark types from the schema
      const disallowedMarkTypes = DISALLOWED_MARK_NAMES.map(
        (name) => newState.schema.marks[name]
      ).filter(Boolean);

      newState.doc.descendants((node, pos) => {
        // Remap disallowed styleId to Normal
        if (node.type.name === 'paragraph' && node.attrs.styleId) {
          if (!allowed.has(node.attrs.styleId)) {
            tr.setNodeMarkup(pos, undefined, {
              ...node.attrs,
              styleId: 'Normal',
            });
            changed = true;
          }
        }

        // Strip disallowed marks from text nodes
        if (node.isText && node.marks.length > 0) {
          let newMarks = node.marks;
          for (const markType of disallowedMarkTypes) {
            if (markType.isInSet(newMarks)) {
              newMarks = markType.removeFromSet(newMarks);
            }
          }
          if (newMarks.length !== node.marks.length) {
            // Remove each disallowed mark from this text range
            const from = pos;
            const to = pos + node.nodeSize;
            for (const markType of disallowedMarkTypes) {
              if (markType.isInSet(node.marks)) {
                tr.removeMark(from, to, markType);
              }
            }
            changed = true;
          }
        }
      });

      if (!changed) return null;

      // Also strip disallowed marks from storedMarks (cursor formatting)
      let marksToRestore = savedStoredMarks;
      if (marksToRestore) {
        for (const markType of disallowedMarkTypes) {
          if (markType.isInSet(marksToRestore)) {
            marksToRestore = markType.removeFromSet(marksToRestore);
          }
        }
        tr.setStoredMarks(marksToRestore);
      }

      return tr;
    },

    props: {
      // On paste, strip fontFamily and fontSize marks
      transformPasted(slice, view) {
        return stripFontMarksFromSlice(slice, view.state.schema);
      },
    },
  });
}
