/**
 * Protected Region Notify Plugin
 *
 * Fires a callback once per paragraph position per editor session when
 * the user makes a content-changing input inside a protected (locked) paragraph.
 * Does not block the edit — purely advisory notification.
 */

import { Plugin, PluginKey } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';

export const protectedRegionNotifyKey = new PluginKey<Set<number>>('protectedRegionNotify');

const CONTENT_CHANGING_INPUT_TYPES = new Set([
  'insertText',
  'insertCompositionText',
  'insertFromPaste',
  'insertFromDrop',
  'deleteContentBackward',
  'deleteContentForward',
  'deleteByCut',
  'deleteByDrag',
  'deleteWordBackward',
  'deleteWordForward',
  'formatBold',
  'formatItalic',
  'formatUnderline',
]);

export function createProtectedRegionNotifyPlugin(
  callback: (info: {
    protectedBy: string | null;
    protectedAt: string | null;
    protectedReason: string | null;
  }) => void
): Plugin<Set<number>> {
  return new Plugin<Set<number>>({
    key: protectedRegionNotifyKey,

    state: {
      init: () => new Set<number>(),
      apply: (tr, firedPositions) => {
        if (tr.getMeta('resetProtectedNotify')) return new Set<number>();
        return firedPositions;
      },
    },

    props: {
      handleDOMEvents: {
        beforeinput(view: EditorView, event: Event) {
          const inputEvent = event as InputEvent;
          if (!CONTENT_CHANGING_INPUT_TYPES.has(inputEvent.inputType)) return false;

          const { state } = view;
          const { $from } = state.selection;

          for (let d = $from.depth; d >= 0; d--) {
            const node = $from.node(d);
            if (node.type.name === 'paragraph' && node.attrs.locked) {
              const paraPos = $from.start(d) - 1;
              const firedPositions = protectedRegionNotifyKey.getState(state);
              if (firedPositions && !firedPositions.has(paraPos)) {
                firedPositions.add(paraPos);
                callback({
                  protectedBy: node.attrs.protectedBy ?? null,
                  protectedAt: node.attrs.protectedAt ?? null,
                  protectedReason: node.attrs.protectedReason ?? null,
                });
              }
              break;
            }
          }
          return false;
        },
      },
    },
  });
}
