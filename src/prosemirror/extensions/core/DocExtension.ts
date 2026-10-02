/**
 * Doc Extension — top-level document node
 */

import { createNodeExtension } from '../create';

export const DocExtension = createNodeExtension({
  name: 'doc',
  schemaNodeName: 'doc',
  nodeSpec: {
    // textBox: a text box read from a body paragraph stands as a block after it (toProseDoc)
    content: '(paragraph | horizontalRule | pageBreak | table | textBox | loopBlock)+',
  },
});
