/**
 * Small Caps Mark Extension (w:smallCaps)
 */

import { toggleMark } from 'prosemirror-commands';
import { createMarkExtension } from '../create';
import type { ExtensionContext, ExtensionRuntime } from '../types';

export const SmallCapsExtension = createMarkExtension({
  name: 'smallCaps',
  schemaMarkName: 'smallCaps',
  markSpec: {
    parseDOM: [
      {
        style: 'font-variant',
        getAttrs: (value) => (value === 'small-caps' ? {} : false),
      },
    ],
    toDOM() {
      return ['span', { style: 'font-variant: small-caps' }, 0];
    },
  },
  onSchemaReady(ctx: ExtensionContext): ExtensionRuntime {
    return {
      commands: {
        toggleSmallCaps: () => toggleMark(ctx.schema.marks.smallCaps),
      },
    };
  },
});
