/**
 * Style Enforcer ProseMirror Plugin
 *
 * Restricts paragraph formatting to a set of allowed style IDs.
 * - On every transaction, remaps disallowed styleId attrs to "Normal"
 * - On paste, strips fontFamily and fontSize marks from incoming content
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
  'ListParagraph',
  'Caption',
  'Title',
  'Subtitle',
];

/**
 * Strip fontFamily and fontSize marks from a slice (used for paste filtering).
 */
function stripFontMarksFromSlice(slice: Slice, schema: Schema): Slice {
  const fontFamily = schema.marks['fontFamily'];
  const fontSize = schema.marks['fontSize'];
  if (!fontFamily && !fontSize) return slice;

  const mapped = mapFragment(slice.content, (node) => {
    if (node.isText) {
      let newNode = node;
      if (fontFamily && fontFamily.isInSet(node.marks)) {
        newNode = newNode.mark(fontFamily.removeFromSet(newNode.marks));
      }
      if (fontSize && fontSize.isInSet(newNode.marks)) {
        newNode = newNode.mark(fontSize.removeFromSet(newNode.marks));
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

    // appendTransaction: remap disallowed styleId to Normal
    appendTransaction(transactions, _oldState, newState) {
      // Only process if something actually changed
      if (!transactions.some((tr) => tr.docChanged)) return null;

      const tr = newState.tr;
      let changed = false;

      newState.doc.descendants((node, pos) => {
        if (node.type.name === 'paragraph' && node.attrs.styleId) {
          if (!allowed.has(node.attrs.styleId)) {
            tr.setNodeMarkup(pos, undefined, {
              ...node.attrs,
              styleId: 'Normal',
            });
            changed = true;
          }
        }
      });

      return changed ? tr : null;
    },

    props: {
      // On paste, strip fontFamily and fontSize marks
      transformPasted(slice, view) {
        return stripFontMarksFromSlice(slice, view.state.schema);
      },
    },
  });
}
