/**
 * ProseMirror Plugins
 *
 * Selection tracker plugin for the DOCX editor.
 * Keymap plugins are now provided by the extension system.
 */

export {
  createSelectionTrackerPlugin,
  extractSelectionContext,
  getSelectionContext,
  selectionTrackerKey,
} from './selectionTracker';

export type { SelectionContext, SelectionChangeCallback } from './selectionTracker';

export {
  createCrossRefUpdaterPlugin,
  crossRefUpdaterKey,
  refreshAllReferences,
} from './crossRefUpdater';
export type { CrossRefUpdaterConfig } from './crossRefUpdater';

export { createSelectiveEditablePlugin, selectiveEditableKey } from './SelectiveEditablePlugin';
