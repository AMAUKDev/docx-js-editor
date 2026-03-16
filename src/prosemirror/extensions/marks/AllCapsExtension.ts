/**
 * All Caps Mark Extension (w:caps)
 */

import { toggleMark } from 'prosemirror-commands';
import { createMarkExtension } from '../create';
import type { ExtensionContext, ExtensionRuntime } from '../types';

export const AllCapsExtension = createMarkExtension({
  name: 'allCaps',
  schemaMarkName: 'allCaps',
  markSpec: {
    parseDOM: [
      {
        style: 'text-transform',
        getAttrs: (value) => (value === 'uppercase' ? {} : false),
      },
    ],
    toDOM() {
      return ['span', { style: 'text-transform: uppercase' }, 0];
    },
  },
  onSchemaReady(ctx: ExtensionContext): ExtensionRuntime {
    return {
      commands: {
        toggleAllCaps: () => toggleMark(ctx.schema.marks.allCaps),
      },
    };
  },
});
