/**
 * DocxEditor Component
 *
 * Main component integrating all editor features:
 * - Toolbar for formatting
 * - ProseMirror-based editor for content editing
 * - Zoom control
 * - Error boundary
 * - Loading states
 */

import {
  useRef,
  useCallback,
  useState,
  useEffect,
  useMemo,
  forwardRef,
  useImperativeHandle,
  lazy,
  Suspense,
} from 'react';
import type { CSSProperties, ReactNode } from 'react';
import type { Document, Theme, HeaderFooter } from '../types/document';
import type { Run } from '../types/content';
import { setDocumentStyles, getStyleDef } from '../prosemirror/styles/styleStore';
import { parseImportedStyleXml } from '../utils/parseImportedStyleXml';

import { Toolbar, type SelectionFormatting, type FormattingAction } from './Toolbar';
import { pointsToHalfPoints } from './ui/FontSizePicker';
import { DocumentOutline } from './DocumentOutline';
import type { HeadingInfo } from '../utils/headingCollector';
import { ErrorBoundary, ErrorProvider } from './ErrorBoundary';
import type { TableAction } from './ui/TableToolbar';
import { mapHexToHighlightName } from './toolbarUtils';
import {
  PageNumberIndicator,
  type PageIndicatorPosition,
  type PageIndicatorVariant,
} from './ui/PageNumberIndicator';
import {
  PageNavigator,
  type PageNavigatorPosition,
  type PageNavigatorVariant,
} from './ui/PageNavigator';
import { HorizontalRuler } from './ui/HorizontalRuler';
import { VerticalRuler } from './ui/VerticalRuler';
import { type PrintOptions } from './ui/PrintPreview';
// Dialog hooks and utilities (static imports — lightweight, no UI)
import {
  useFindReplace,
  findInDocument,
  scrollToMatch,
  type FindMatch,
  type FindOptions,
  type FindResult,
} from './dialogs/FindReplaceDialog';
import { useHyperlinkDialog, type HyperlinkData } from './dialogs/HyperlinkDialog';
import type { ImagePositionData } from './dialogs/ImagePositionDialog';
import type { ImagePropertiesData } from './dialogs/ImagePropertiesDialog';
import {
  InlineHeaderFooterEditor,
  type InlineHeaderFooterEditorRef,
} from './InlineHeaderFooterEditor';

// Dialog components (lazy-loaded — only fetched when first opened)
const FindReplaceDialog = lazy(() => import('./dialogs/FindReplaceDialog'));
const HyperlinkDialog = lazy(() => import('./dialogs/HyperlinkDialog'));
const TablePropertiesDialog = lazy(() =>
  import('./dialogs/TablePropertiesDialog').then((m) => ({ default: m.TablePropertiesDialog }))
);
const ImagePositionDialog = lazy(() =>
  import('./dialogs/ImagePositionDialog').then((m) => ({ default: m.ImagePositionDialog }))
);
const ImagePropertiesDialog = lazy(() =>
  import('./dialogs/ImagePropertiesDialog').then((m) => ({ default: m.ImagePropertiesDialog }))
);
const FootnotePropertiesDialog = lazy(() =>
  import('./dialogs/FootnotePropertiesDialog').then((m) => ({
    default: m.FootnotePropertiesDialog,
  }))
);
const StyleEditorDialog = lazy(() =>
  import('./dialogs/StyleEditorDialog').then((m) => ({
    default: m.StyleEditorDialog,
  }))
);
import { MaterialSymbol } from './ui/Icons';
import { getBuiltinTableStyle, type TableStylePreset } from './ui/TableStyleGallery';
import { DocumentAgent } from '../agent/DocumentAgent';
import { DefaultLoadingIndicator, DefaultPlaceholder, ParseError } from './DocxEditorHelpers';
import { parseDocx } from '../docx/parser';
import { createNumberingMap as createNumberingMapFromDefs } from '../docx/numberingParser';
import { formatNumberedMarker } from '../layout-bridge/toFlowBlocks';
import { type DocxInput } from '../utils/docxInput';
import { onFontsLoaded, loadDocumentFonts } from '../utils/fontLoader';
import { executeCommand } from '../agent/executor';
import { useTableSelection } from '../hooks/useTableSelection';
import { useDocumentHistory } from '../hooks/useHistory';

// Extension system
import { createStarterKit } from '../prosemirror/extensions/StarterKit';
import { ExtensionManager } from '../prosemirror/extensions/ExtensionManager';

// Conversion (for HF inline editor save)
import { proseDocToBlocks, fromProseDoc } from '../prosemirror/conversion/fromProseDoc';
import { collectContextTagMetadata } from '../docx/contextTagMetadata';
import { renderDocumentWithBookmarks } from '../docx/renderWithBookmarks';
import type { ContextTagMeta, FPDocumentMeta } from '../types/document';
import { generateMetaId } from '../prosemirror/extensions/nodes/ContextTagExtension';
// ProseMirror editor
import {
  type SelectionState,
  TextSelection,
  extractSelectionState,
  toggleBold,
  toggleItalic,
  toggleUnderline,
  toggleStrike,
  toggleSuperscript,
  toggleSubscript,
  toggleAllCaps,
  toggleSmallCaps,
  setTextColor,
  setHighlight,
  setFontSize,
  setFontFamily,
  setAlignment,
  setLineSpacing,
  toggleBulletList,
  toggleNumberedList,
  increaseIndent,
  decreaseIndent,
  setIndentLeft,
  setIndentRight,
  setIndentFirstLine,
  removeTabStop,
  increaseListLevel,
  decreaseListLevel,
  clearFormatting,
  applyStyle,
  createStyleResolver,
  // Hyperlink commands
  getHyperlinkAttrs,
  getSelectedText,
  setHyperlink,
  removeHyperlink,
  insertHyperlink,
  // Page break command
  insertPageBreak,
  // Table of Contents command
  generateTOC,
  // Table commands (getTableContext is used directly; other commands come
  // from per-instance extensionManager to avoid schema mismatch)
  getTableContext,
  setCellBorder,
  setCellVerticalAlign,
  setCellMargins,
  setCellTextDirection,
  toggleNoWrap,
  setRowHeight,
  toggleHeaderRow,
  distributeColumns,
  autoFitContents,
  setTableProperties,
  applyTableStyle,
  removeTableBorders,
  setAllTableBorders,
  setOutsideTableBorders,
  setInsideTableBorders,
  setCellFillColor,
  setTableBorderColor,
  setTableBorderWidth,
  type TableContextInfo,
} from '../prosemirror';
import { collectHeadings } from '../utils/headingCollector';
import { textFormattingToMarks } from '../prosemirror/extensions/marks/markUtils';

// Paginated editor
import { PagedEditor, type PagedEditorRef } from '../paged-editor/PagedEditor';

// Plugin API types
import type { RenderedDomContext } from '../plugin-api/types';

// Style enforcer plugin
import {
  createStyleEnforcerPlugin,
  DEFAULT_ALLOWED_STYLE_IDS,
} from '../plugins/StyleEnforcerPlugin';
import {
  createCrossRefUpdaterPlugin,
  refreshAllReferences,
} from '../prosemirror/plugins/crossRefUpdater';
import { createSelectiveEditablePlugin } from '../prosemirror/plugins/SelectiveEditablePlugin';

// ============================================================================
// TYPES
// ============================================================================

/**
 * DocxEditor props
 */
export interface DocxEditorProps {
  /** Document data — ArrayBuffer, Uint8Array, Blob, or File */
  documentBuffer?: DocxInput | null;
  /** Pre-parsed document (alternative to documentBuffer) */
  document?: Document | null;
  /** Callback when document is saved */
  onSave?: (buffer: ArrayBuffer) => void;
  /** Callback when document changes */
  onChange?: (document: Document) => void;
  /** Callback when selection changes */
  onSelectionChange?: (state: SelectionState | null) => void;
  /** Callback on error */
  onError?: (error: Error) => void;
  /** Callback when fonts are loaded */
  onFontsLoaded?: () => void;
  /** External ProseMirror plugins (from PluginHost) */
  externalPlugins?: import('prosemirror-state').Plugin[];
  /** Callback when editor view is ready (for PluginHost) */
  onEditorViewReady?: (view: import('prosemirror-view').EditorView) => void;
  /** Theme for styling */
  theme?: Theme | null;
  /** Whether to show toolbar (default: true) */
  showToolbar?: boolean;
  /** Whether to show zoom control (default: true) */
  showZoomControl?: boolean;
  /** Whether to show page number indicator (default: true) */
  showPageNumbers?: boolean;
  /** Whether to enable interactive page navigation (default: true) */
  enablePageNavigation?: boolean;
  /** Position of page number indicator (default: 'bottom-center') */
  pageNumberPosition?: PageIndicatorPosition | PageNavigatorPosition;
  /** Variant of page number indicator (default: 'default') */
  pageNumberVariant?: PageIndicatorVariant | PageNavigatorVariant;
  /** Whether to show page margin guides/boundaries (default: false) */
  showMarginGuides?: boolean;
  /** Color for margin guides (default: '#c0c0c0') */
  marginGuideColor?: string;
  /** Whether to show horizontal ruler (default: false) */
  showRuler?: boolean;
  /** Unit for ruler display (default: 'inch') */
  rulerUnit?: 'inch' | 'cm';
  /** Initial zoom level (default: 1.0) */
  initialZoom?: number;
  /** Whether the editor is read-only. When true, hides toolbar and rulers */
  readOnly?: boolean;
  /** Custom toolbar actions */
  toolbarExtra?: ReactNode;
  /** Show inline comment margin panel alongside pages */
  showCommentPanel?: boolean;
  /** Callback when comment action occurs in the margin panel */
  onCommentAction?: (action: 'reply' | 'resolve' | 'delete' | 'edit', commentId: number) => void;
  /** Bump to force comment margin panel refresh */
  commentPanelKey?: number;
  /** Additional CSS class name */
  className?: string;
  /** Additional inline styles */
  style?: CSSProperties;
  /** Placeholder when no document */
  placeholder?: ReactNode;
  /** Loading indicator */
  loadingIndicator?: ReactNode;
  /** Whether to show the document outline sidebar (default: false) */
  showOutline?: boolean;
  /** Whether to show print button in toolbar (default: true) */
  showPrintButton?: boolean;
  /** Whether to show line spacing picker in toolbar (default: true) */
  showLineSpacingPicker?: boolean;
  /** Whether to show clear formatting button in toolbar (default: true) */
  showClearFormatting?: boolean;
  /** Whether to show insert TOC button in toolbar (default: true) */
  showInsertTOC?: boolean;
  /** Print options for print preview */
  printOptions?: PrintOptions;
  /** Callback when print is triggered */
  onPrint?: () => void;
  /** Callback when content is copied */
  onCopy?: () => void;
  /** Callback when content is cut */
  onCut?: () => void;
  /** Callback when content is pasted */
  onPaste?: () => void;
  /** Editor mode: 'editing' (direct edits), 'suggesting' (track changes), or 'viewing' (read-only). Default: 'editing' */
  mode?: EditorMode;
  /** Callback when the editing mode changes */
  onModeChange?: (mode: EditorMode) => void;
  /**
   * Fired when the cursor moves onto a comment or tracked-change mark.
   * Null is passed when the cursor leaves all such marks.
   * The id follows the pattern `comment-<id>` for comments and
   * `tc-<revisionId>-<type>` for tracked changes.
   */
  onCursorMarkChange?: (markId: string | null) => void;
  /**
   * Callback when rendered DOM context is ready (for plugin overlays).
   * Used by PluginHost to get access to the rendered page DOM for positioning.
   */
  onRenderedDomContextReady?: (context: RenderedDomContext) => void;
  /**
   * Plugin overlays to render inside the editor viewport.
   * Passed from PluginHost to render plugin-specific overlays.
   */
  pluginOverlays?: ReactNode;
  /** When true, restrict formatting to approved styles only (hides font/size/color/bold/italic pickers) */
  restrictedMode?: boolean;
  /** When true, shows the style picker in gallery mode with formatted previews (independent of restrictedMode) */
  styleGalleryMode?: boolean;
  /** Style IDs shown in the gallery when restrictedMode or styleGalleryMode is true */
  allowedStyleIds?: string[];
  /** When false, disables context tag parsing entirely — document text shown as-is */
  enableContextTags?: boolean;
  /** Context tags available for insertion (key → resolved value) */
  contextTags?: Record<string, string>;
  /** Called when document is loaded/parsed with the set of tagKeys found in the doc */
  onContextTagsDiscovered?: (tagKeys: string[]) => void;
  /** Called when re-uploaded document has loop diffs (expanded loops with edits) */
  onLoopDiffsDetected?: (diffs: import('../docx/renderWithBookmarks').LoopDiffReport[]) => void;
  /** When true, the SelectiveEditablePlugin blocks edits to locked paragraphs */
  lockedEditing?: boolean;
  /** Called when user right-clicks a context tag in the editor */
  onContextTagRightClick?: (info: {
    tagKey: string;
    label: string;
    removeIfEmpty: boolean;
    removeTableRow: boolean;
    alwaysShow: boolean;
    imageWidth: number;
    pmPos: number;
    clientX: number;
    clientY: number;
  }) => void;
  /**
   * Resolved loop preview data for rendered mode expansion.
   * Keys are loop array names (e.g. "photos"), values are arrays of items
   * where image fields have been resolved from case_file_id to {url, name}.
   */
  loopPreviewData?: Record<string, Array<Record<string, unknown>>> | null;
  /** Called when the page count changes after layout */
  onPageCountChange?: (pageCount: number) => void;
  /** When true, show gear icons on styles + "Create New Style" in dropdown */
  canModifyStyles?: boolean;
}

/**
 * DocxEditor ref interface
 */
export interface DocxEditorRef {
  /** Get the DocumentAgent for programmatic access */
  getAgent: () => DocumentAgent | null;
  /** Get the current document */
  getDocument: () => Document | null;
  /** Get the editor ref */
  getEditorRef: () => PagedEditorRef | null;
  /** Save the document to buffer */
  save: () => Promise<ArrayBuffer | null>;
  /** Set zoom level */
  setZoom: (zoom: number) => void;
  /** Get current zoom level */
  getZoom: () => number;
  /** Set render mode for template elements ('rendered' or 'raw') */
  setRenderMode: (mode: 'rendered' | 'raw') => void;
  /** Get current render mode */
  getRenderMode: () => 'rendered' | 'raw';
  /** Focus the editor */
  focus: () => void;
  /** Get current page number */
  getCurrentPage: () => number;
  /** Get total page count */
  getTotalPages: () => number;
  /** Scroll to a specific page */
  scrollToPage: (pageNumber: number) => void;
  /** Open print preview */
  openPrintPreview: () => void;
  /** Print the document directly */
  print: () => void;
  /** Get the currently selected text (empty string if no selection) */
  getSelectedText: () => string;
  /** Get the current PM selection range { from, to } */
  getSelectionRange: () => { from: number; to: number } | null;
  /** Add an inline comment wrapping the current selection */
  addComment: (author: string, text: string, from?: number, to?: number) => number | null;
  /** Add a reply comment to an existing comment (no range marker needed) */
  addReplyComment: (parentCommentId: number, author: string, text: string) => number | null;
  /** Update an existing comment's text */
  editComment: (commentId: number, newText: string) => void;
  /** Remove an inline comment by ID (removes highlight + Comment object) */
  removeComment: (commentId: number) => void;
  /** Get all document comments */
  getDocumentComments: () => Array<{
    id: number;
    author: string;
    date?: string;
    text: string;
    anchorText: string;
    parentId?: number;
    done?: boolean;
  }>;
  /** Get Y-positions of comment marks relative to the pages container */
  getCommentPositions: () => Array<{ commentId: number; top: number; height: number }>;
  /** Insert a context tag at the current cursor position */
  insertContextTag: (tagKey: string, label?: string, removeIfEmpty?: boolean) => void;
  /** Insert a cross-reference at the current cursor position */
  insertCrossRef: (
    refType: 'heading' | 'figure',
    refTarget: string,
    displayText: string,
    bookmarkName?: string
  ) => void;
  /** Get all headings and captions in the document for cross-reference picking */
  getReferenceable: () => Array<{
    type: 'heading' | 'figure';
    text: string;
    number: string;
    bookmarkName: string | null;
    pmPos: number;
  }>;
  /** Force-refresh all cross-ref displayText and caption figure numbers */
  refreshNumbering: (tocStyleOverride?: string) => void;
  /** Get available paragraph styles for TOC style picker */
  getAvailableStyles: () => Array<{ styleId: string; name: string }>;
  /** Render document with context tags resolved to plain text, return as DOCX buffer.
   *  unknownTagMode: 'omit' = remove unknowns, 'keep' = leave unknowns, 'raw' = skip all rendering */
  renderToBuffer: (options?: {
    contextTags?: Record<string, string>;
    unknownTagMode?: 'omit' | 'keep' | 'raw';
  }) => Promise<ArrayBuffer | null>;
  /** Lock paragraphs in a position range (admin) */
  lockParagraphs: (from: number, to: number) => void;
  /** Unlock paragraphs in a position range (admin) */
  unlockParagraphs: (from: number, to: number) => void;
  /** Lock all paragraphs in the document (admin) */
  lockAll: () => void;
  /** Unlock all paragraphs in the document (admin) */
  unlockAll: () => void;
  /** Lock all paragraphs in headers and footers */
  lockHeadersFooters: () => void;
  /** Unlock all paragraphs in headers and footers */
  unlockHeadersFooters: () => void;
  /** Update attributes of a context tag node at a given PM position */
  updateContextTagAttrs: (pmPos: number, attrs: Record<string, unknown>) => void;
  /** Import styles into the document (merges/overwrites existing by styleId) */
  importStyles: (
    styles: Array<{
      styleId: string;
      name?: string;
      type?: string;
      basedOn?: string;
      next?: string;
      fontFamily?: string;
      fontSize?: number;
      bold?: boolean;
      italic?: boolean;
      color?: string;
    }>
  ) => void;
  /** Get document-level metadata from the Custom XML Part (template provenance, tocStyle, etc.) */
  getDocumentMeta: () => FPDocumentMeta | undefined;
  /** Set/update document-level metadata (written to Custom XML Part on next save) */
  setDocumentMeta: (meta: FPDocumentMeta) => void;
}

/** Convert a context tag value (which may be a string, object, or null) to a display string.
 *  Image objects `{ case_file_id, name }` are shown as their filename. */
function contextTagDisplayValue(val: unknown): string {
  if (val == null) return '';
  if (typeof val === 'object') {
    const obj = val as Record<string, unknown>;
    if (obj.case_file_id) return (obj.name as string) || '(image)';
    return '';
  }
  return String(val);
}

/** Regex for matching context tag patterns in header/footer text.
 *  Matches both {{ tag.path }} (double-brace) and {tag.path} (single-brace, 0+ dots).
 *  Used for both discovery and substitution in H/F content. */
const HF_CONTEXT_TAG_RE = /\{\{\s*([\w]+(?:\.[\w]+)*)!?\s*\}\}|\{([\w]+(?:\.[\w]+)*)\}/g;

// ─── HEADER/FOOTER CONTEXT TAG SUBSTITUTION ────────────────────────────────
//
// ARCHITECTURE NOTE:
// H/F tag substitution happens in TWO complementary places:
//
// 1. DocxEditor (here): replaceContextTagsInHf() creates NEW HeaderFooter
//    objects with tag patterns replaced by resolved values. These objects are
//    passed as props to PagedEditor for visual display ONLY.
//    The original Document model is NOT mutated — saving uses the original.
//
// 2. PagedEditor: substituteHfContextTags() in convertHeaderFooterToContent()
//    provides a second pass during layout. This handles tags in the FlowBlock
//    text that might not have been caught by Path 1 (e.g., single-brace tags
//    that were already in the saved DOCX as rendered text).
//
// CRITICAL: replaceContextTagsInHf() MUST always return a new object reference
// (via spread), even when no substitutions occurred. This ensures React's
// dependency tracking detects the change when contextTags changes, triggering
// a re-render in PagedEditor. Returning the same reference is the root cause
// of the recurring "footer doesn't update" bug.
// ────────────────────────────────────────────────────────────────────────────

/**
 * Replace context tags across a group of runs, handling tags split across
 * multiple runs (e.g., {primary_asset.display_name} split by proofErr elements).
 *
 * Algorithm: join all text segments into one string, apply regex, then
 * redistribute replaced text back to the original run segments.
 * Returns null if no changes were made.
 */
function _replaceTagsInRunGroup(
  runs: readonly Run[],
  tags: Record<string, string>,
  mode: 'omit' | 'keep'
): Run[] | null {
  // 1. Collect text segments with character offsets in joined string
  interface Seg {
    runIdx: number;
    rcIdx: number;
    start: number;
    end: number;
  }
  const segs: Seg[] = [];
  let pos = 0;
  for (let ri = 0; ri < runs.length; ri++) {
    for (let ci = 0; ci < runs[ri].content.length; ci++) {
      const rc = runs[ri].content[ci];
      if (rc.type === 'text' && rc.text) {
        segs.push({ runIdx: ri, rcIdx: ci, start: pos, end: pos + rc.text.length });
        pos += rc.text.length;
      }
    }
  }
  if (segs.length === 0) return null;

  // 2. Join all text and find tag matches
  const joined = segs
    .map((s) => (runs[s.runIdx].content[s.rcIdx] as { text: string }).text)
    .join('');

  interface TagMatch {
    start: number;
    end: number;
    replacement: string;
  }
  const matches: TagMatch[] = [];
  const re = new RegExp(HF_CONTEXT_TAG_RE.source, HF_CONTEXT_TAG_RE.flags);
  let m: RegExpExecArray | null;
  while ((m = re.exec(joined)) !== null) {
    const rawKey = m[1] || m[2];
    const tagKey = rawKey.startsWith('context.') ? rawKey.slice(8) : rawKey;
    const resolved = contextTagDisplayValue(tags[tagKey]);
    let replacement: string;
    if (resolved !== '') {
      replacement = resolved;
    } else if (mode === 'omit') {
      replacement = '';
    } else {
      replacement = m[0]; // keep original tag text
    }
    if (replacement !== m[0]) {
      matches.push({ start: m.index, end: m.index + m[0].length, replacement });
    }
  }
  if (matches.length === 0) return null;

  // 3. Build new text for each segment by applying matches right-to-left
  const segTexts = segs.map((s) => (runs[s.runIdx].content[s.rcIdx] as { text: string }).text);
  for (let mi = matches.length - 1; mi >= 0; mi--) {
    const tm = matches[mi];
    let firstSeg = -1;
    for (let si = 0; si < segs.length; si++) {
      const seg = segs[si];
      if (seg.end <= tm.start || seg.start >= tm.end) continue; // no overlap
      const localStart = Math.max(0, tm.start - seg.start);
      const localEnd = Math.min(seg.end - seg.start, tm.end - seg.start);
      if (firstSeg === -1) {
        firstSeg = si;
        // First overlapping segment gets the replacement text
        segTexts[si] =
          segTexts[si].slice(0, localStart) + tm.replacement + segTexts[si].slice(localEnd);
      } else {
        // Subsequent segments: remove the matched portion
        segTexts[si] = segTexts[si].slice(0, localStart) + segTexts[si].slice(localEnd);
      }
    }
  }

  // 4. Rebuild runs with updated text
  const newRuns = runs.map((run, ri) => {
    let runChanged = false;
    const newContent = run.content.map((rc, ci) => {
      if (rc.type !== 'text') return rc;
      const si = segs.findIndex((s) => s.runIdx === ri && s.rcIdx === ci);
      if (si === -1) return rc;
      if (segTexts[si] !== rc.text) {
        runChanged = true;
        return { ...rc, text: segTexts[si] };
      }
      return rc;
    });
    return runChanged ? { ...run, content: newContent } : run;
  });

  return newRuns;
}

/**
 * Replace context tag text patterns in a HeaderFooter for visual display.
 * ALWAYS returns a new object to ensure React prop change detection.
 * Does NOT mutate the original — the Document model keeps tag patterns for saving.
 *
 * Handles tags split across multiple runs (e.g., by proofErr elements) and
 * tags inside InlineSdt content controls.
 */
function replaceContextTagsInHf(
  hf: HeaderFooter,
  tags: Record<string, string>,
  mode: 'omit' | 'keep'
): HeaderFooter {
  const newContent = hf.content.map((block) => {
    if (block.type !== 'paragraph') return block;
    let paraChanged = false;

    // First pass: process InlineSdt content controls
    const newParaContent = block.content.map((item) => {
      if (item.type === 'inlineSdt') {
        const sdtRuns = item.content.filter((c): c is Run => c.type === 'run');
        if (sdtRuns.length === 0) return item;
        const replaced = _replaceTagsInRunGroup(sdtRuns, tags, mode);
        if (!replaced) return item;
        paraChanged = true;
        let runIdx = 0;
        const newSdtContent = item.content.map((c) => (c.type === 'run' ? replaced[runIdx++] : c));
        return { ...item, content: newSdtContent };
      }
      return item;
    });

    // Second pass: process direct runs in the paragraph as a group
    const directRuns: Run[] = [];
    const directRunIndices: number[] = [];
    for (let i = 0; i < newParaContent.length; i++) {
      if (newParaContent[i].type === 'run') {
        directRuns.push(newParaContent[i] as Run);
        directRunIndices.push(i);
      }
    }
    if (directRuns.length > 0) {
      const replaced = _replaceTagsInRunGroup(directRuns, tags, mode);
      if (replaced) {
        paraChanged = true;
        for (let i = 0; i < directRunIndices.length; i++) {
          newParaContent[directRunIndices[i]] = replaced[i];
        }
      }
    }

    return paraChanged ? { ...block, content: newParaContent } : block;
  });
  // ALWAYS return a new object — never return `hf` directly.
  // This ensures React detects the prop change even when no tags matched.
  return { ...hf, content: newContent };
}

/** Set locked state on all paragraphs in every header and footer. */
function setHfLocked(
  history: { state: Document | null; push: (d: Document) => void },
  locked: boolean
) {
  const doc = history.state;
  if (!doc?.package) return;
  const pkg = doc.package;

  const updateMap = (map: Map<string, HeaderFooter> | undefined) => {
    if (!map) return map;
    const newMap = new Map<string, HeaderFooter>();
    for (const [rId, hf] of map) {
      const newContent = hf.content.map((block) => {
        if (block.type === 'paragraph') {
          return {
            ...block,
            formatting: { ...block.formatting, locked: locked || undefined },
          };
        }
        return block;
      });
      newMap.set(rId, { ...hf, content: newContent });
    }
    return newMap;
  };

  history.push({
    ...doc,
    package: {
      ...pkg,
      headers: updateMap(pkg.headers),
      footers: updateMap(pkg.footers),
    },
  });
}

/**
 * Editor internal state
 */
interface EditorState {
  isLoading: boolean;
  parseError: string | null;
  zoom: number;
  /** Current selection formatting for toolbar */
  selectionFormatting: SelectionFormatting;
  /** Current page number (1-indexed) */
  currentPage: number;
  /** Total page count */
  totalPages: number;
  /** Paragraph indent data for ruler */
  paragraphIndentLeft: number;
  paragraphIndentRight: number;
  paragraphFirstLineIndent: number;
  paragraphHangingIndent: boolean;
  paragraphTabs: import('../types/document').TabStop[] | null;
  /** ProseMirror table context (for showing table toolbar) */
  pmTableContext: TableContextInfo | null;
  /** Image context when cursor is on an image node */
  pmImageContext: {
    pos: number;
    wrapType: string;
    displayMode: string;
    cssFloat: string | null;
    transform: string | null;
    alt: string | null;
    borderWidth: number | null;
    borderColor: string | null;
    borderStyle: string | null;
  } | null;
  /** Whether to display template tags/loops in rendered or raw mode */
  renderMode: 'rendered' | 'raw';
}

// ============================================================================
// EDITING MODE DROPDOWN (Google Docs-style)
// ============================================================================

export type EditorMode = 'editing' | 'suggesting' | 'viewing';

const EDITING_MODES: readonly { value: EditorMode; label: string; icon: string; desc: string }[] = [
  {
    value: 'editing',
    label: 'Editing',
    icon: 'edit_note',
    desc: 'Edit document directly',
  },
  {
    value: 'suggesting',
    label: 'Suggesting',
    icon: 'rate_review',
    desc: 'Edits become suggestions',
  },
  {
    value: 'viewing',
    label: 'Viewing',
    icon: 'visibility',
    desc: 'Read-only, no edits',
  },
];

function EditingModeDropdown({
  mode,
  onModeChange,
}: {
  mode: EditorMode;
  onModeChange: (mode: EditorMode) => void;
}) {
  const [isOpen, setIsOpen] = useState(false);
  const [compact, setCompact] = useState(false);
  const triggerRef = useRef<HTMLButtonElement>(null);
  const dropdownRef = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState({ top: 0, left: 0 });

  const current = EDITING_MODES.find((m) => m.value === mode)!;

  // Responsive: icon-only below 1400px
  useEffect(() => {
    const mql = window.matchMedia('(max-width: 1400px)');
    setCompact(mql.matches);
    const handler = (e: MediaQueryListEvent) => setCompact(e.matches);
    mql.addEventListener('change', handler);
    return () => mql.removeEventListener('change', handler);
  }, []);

  useEffect(() => {
    if (!isOpen || !triggerRef.current) return;
    const rect = triggerRef.current.getBoundingClientRect();
    // Align dropdown to right edge of trigger so it doesn't overflow the screen
    setPos({ top: rect.bottom + 2, left: rect.right - 220 });
  }, [isOpen]);

  useEffect(() => {
    if (!isOpen) return;
    const close = (e: MouseEvent) => {
      if (
        !triggerRef.current?.contains(e.target as Node) &&
        !dropdownRef.current?.contains(e.target as Node)
      ) {
        setIsOpen(false);
      }
    };
    const esc = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setIsOpen(false);
    };
    document.addEventListener('mousedown', close);
    document.addEventListener('keydown', esc);
    return () => {
      document.removeEventListener('mousedown', close);
      document.removeEventListener('keydown', esc);
    };
  }, [isOpen]);

  return (
    <div style={{ position: 'relative' }}>
      <button
        ref={triggerRef}
        type="button"
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => setIsOpen(!isOpen)}
        title={`${current.label} (Ctrl+Shift+E)`}
        style={{
          display: 'flex',
          alignItems: 'center',
          gap: compact ? 0 : 4,
          padding: compact ? '2px 4px' : '2px 6px 2px 4px',
          border: 'none',
          background: isOpen ? 'var(--doc-hover, #f3f4f6)' : 'transparent',
          borderRadius: 4,
          cursor: 'pointer',
          fontSize: 13,
          fontWeight: 400,
          color: 'var(--doc-text, #374151)',
          whiteSpace: 'nowrap',
          height: 28,
        }}
      >
        <MaterialSymbol name={current.icon} size={18} />
        {!compact && <span>{current.label}</span>}
        <MaterialSymbol name="arrow_drop_down" size={16} />
      </button>

      {isOpen && (
        <div
          ref={dropdownRef}
          onMouseDown={(e) => e.preventDefault()}
          style={{
            position: 'fixed',
            top: pos.top,
            left: pos.left,
            backgroundColor: 'white',
            border: '1px solid var(--doc-border, #d1d5db)',
            borderRadius: 8,
            boxShadow: '0 4px 12px rgba(0, 0, 0, 0.12)',
            padding: '4px 0',
            zIndex: 10000,
            minWidth: 220,
          }}
        >
          {EDITING_MODES.map((m) => (
            <button
              key={m.value}
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => {
                onModeChange(m.value);
                setIsOpen(false);
              }}
              onMouseOver={(e) => {
                (e.currentTarget as HTMLButtonElement).style.backgroundColor =
                  'var(--doc-hover, #f3f4f6)';
              }}
              onMouseOut={(e) => {
                (e.currentTarget as HTMLButtonElement).style.backgroundColor = 'transparent';
              }}
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: 10,
                padding: '8px 12px',
                border: 'none',
                background: 'transparent',
                cursor: 'pointer',
                fontSize: 13,
                color: 'var(--doc-text, #374151)',
                width: '100%',
                textAlign: 'left',
              }}
            >
              <MaterialSymbol name={m.icon} size={20} />
              <span style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-start' }}>
                <span style={{ fontWeight: 500 }}>{m.label}</span>
                <span style={{ fontSize: 11, color: 'var(--doc-text-muted, #9ca3af)' }}>
                  {m.desc}
                </span>
              </span>
              {m.value === mode && (
                <MaterialSymbol
                  name="check"
                  size={18}
                  style={{ marginLeft: 'auto', color: '#1a73e8' }}
                />
              )}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

// ============================================================================
// MAIN COMPONENT
// ============================================================================

/**
 * DocxEditor - Complete DOCX editor component
 */
export const DocxEditor = forwardRef<DocxEditorRef, DocxEditorProps>(function DocxEditor(
  {
    documentBuffer,
    document: initialDocument,
    onSave,
    onChange,
    onSelectionChange,
    onError,
    onFontsLoaded: onFontsLoadedCallback,
    theme,
    showToolbar = true,
    showZoomControl = true,
    showPageNumbers = true,
    enablePageNavigation = true,
    pageNumberPosition = 'bottom-center',
    pageNumberVariant = 'default',
    showMarginGuides: _showMarginGuides = false,
    marginGuideColor: _marginGuideColor,
    showRuler = false,
    rulerUnit = 'inch',
    initialZoom = 1.0,
    readOnly: readOnlyProp = false,
    toolbarExtra,
    className = '',
    style,
    placeholder,
    loadingIndicator,
    showOutline: showOutlineProp = false,
    showPrintButton = true,
    showLineSpacingPicker: showLineSpacingPickerProp = true,
    showClearFormatting: showClearFormattingProp = true,
    showInsertTOC: showInsertTOCProp = true,
    printOptions: _printOptions,
    onPrint,
    onCopy: _onCopy,
    onCut: _onCut,
    onPaste: _onPaste,
    mode: modeProp,
    onModeChange,
    onCursorMarkChange,
    externalPlugins,
    onEditorViewReady,
    onRenderedDomContextReady,
    pluginOverlays,
    restrictedMode = false,
    styleGalleryMode,
    allowedStyleIds: allowedStyleIdsProp,
    enableContextTags = true,
    contextTags,
    onContextTagsDiscovered,
    onLoopDiffsDetected,
    lockedEditing = false,
    onContextTagRightClick,
    loopPreviewData,
    onPageCountChange: onPageCountChangeProp,
    showCommentPanel,
    onCommentAction,
    commentPanelKey,
    canModifyStyles = false,
  },
  ref
) {
  // State
  const [state, setState] = useState<EditorState>({
    isLoading: !!documentBuffer,
    parseError: null,
    zoom: initialZoom,
    selectionFormatting: {},
    currentPage: 1,
    totalPages: 1,
    paragraphIndentLeft: 0,
    paragraphIndentRight: 0,
    paragraphFirstLineIndent: 0,
    paragraphHangingIndent: false,
    paragraphTabs: null,
    pmTableContext: null,
    pmImageContext: null,
    renderMode: 'rendered',
  });

  // Table properties dialog state
  const [tablePropsOpen, setTablePropsOpen] = useState(false);
  // Image position dialog state
  const [imagePositionOpen, setImagePositionOpen] = useState(false);
  // Image properties dialog state
  const [imagePropsOpen, setImagePropsOpen] = useState(false);
  // Footnote properties dialog state
  const [footnotePropsOpen, setFootnotePropsOpen] = useState(false);
  // Style editor dialog state
  const [styleEditorState, setStyleEditorState] = useState<{
    open: boolean;
    mode: 'modify' | 'create';
    styleId?: string;
  }>({ open: false, mode: 'modify' });
  // Header/footer editing state
  const [hfEditPosition, setHfEditPosition] = useState<'header' | 'footer' | null>(null);
  // Editing mode (editing / suggesting / viewing) — controlled or uncontrolled
  const [editingModeInternal, setEditingModeInternal] = useState<EditorMode>(modeProp ?? 'editing');
  const editingMode = modeProp ?? editingModeInternal;
  const setEditingMode = (mode: EditorMode) => {
    if (!modeProp) setEditingModeInternal(mode);
    onModeChange?.(mode);
  };
  // 'viewing' mode acts as read-only
  const readOnly = readOnlyProp || editingMode === 'viewing';
  // Document outline sidebar state
  const [showOutline, setShowOutline] = useState(showOutlineProp);
  const showOutlineRef = useRef(false);
  showOutlineRef.current = showOutline;
  const [outlineHeadings, setHeadingInfos] = useState<HeadingInfo[]>([]);

  // Sync outline visibility when prop changes
  useEffect(() => {
    setShowOutline(showOutlineProp);
    if (showOutlineProp) {
      const view = pagedEditorRef.current?.getView();
      if (view) {
        setHeadingInfos(collectHeadings(view.state.doc));
      }
    }
  }, [showOutlineProp]);

  // History hook for undo/redo - start with null document
  const history = useDocumentHistory<Document | null>(initialDocument || null, {
    maxEntries: 100,
    groupingInterval: 500,
    enableKeyboardShortcuts: true,
  });

  // Extension manager — built once, provides schema + plugins + commands
  const extensionManager = useMemo(() => {
    const disable = enableContextTags ? [] : ['contextTag'];
    const mgr = new ExtensionManager(createStarterKit({ disable }));
    mgr.buildSchema();
    mgr.initializeRuntime();
    return mgr;
  }, [enableContextTags]);

  // Resolve allowed style IDs for restricted mode
  const allowedStyleIds = allowedStyleIdsProp ?? DEFAULT_ALLOWED_STYLE_IDS;

  // Cross-ref auto-updater plugin — uses refs so it always reads the latest data
  // without needing plugin recreation
  const crossRefNumMapRef = useRef<import('../docx/numberingParser').NumberingMap | null>(null);
  const crossRefStyleResolverRef = useRef<ReturnType<typeof createStyleResolver> | null>(null);
  const crossRefUpdaterPlugin = useMemo(
    () =>
      createCrossRefUpdaterPlugin({
        getNumberingMap: () => crossRefNumMapRef.current,
        getStyleResolver: () => crossRefStyleResolverRef.current,
      }),
    []
  );

  // Merge external plugins with style enforcer, cross-ref updater, and selective editable
  const mergedPlugins = useMemo(() => {
    const plugins = externalPlugins ? [...externalPlugins] : [];
    plugins.push(crossRefUpdaterPlugin);
    if (restrictedMode) {
      plugins.push(createStyleEnforcerPlugin({ allowedStyleIds }));
    }
    if (lockedEditing) {
      plugins.push(createSelectiveEditablePlugin());
    }
    return plugins;
  }, [restrictedMode, allowedStyleIds, externalPlugins, crossRefUpdaterPlugin, lockedEditing]);

  // Refs
  const pagedEditorRef = useRef<PagedEditorRef>(null);
  const hfEditorRef = useRef<InlineHeaderFooterEditorRef>(null);
  const agentRef = useRef<DocumentAgent | null>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  // Save the last known selection for restoring after toolbar interactions
  const lastSelectionRef = useRef<{ from: number; to: number } | null>(null);
  const imageInputRef = useRef<HTMLInputElement>(null);
  const editorContentRef = useRef<HTMLDivElement>(null);
  const toolbarWrapperRef = useRef<HTMLDivElement>(null);
  const toolbarRoRef = useRef<ResizeObserver | null>(null);
  const [toolbarHeight, setToolbarHeight] = useState(0);
  // Keep history.state accessible in stable callbacks without stale closures
  const historyStateRef = useRef(history.state);
  historyStateRef.current = history.state;
  // Separate refs for comment mutations — survive history state changes (which
  // create new Document objects and lose in-place mutations).
  const addedCommentsRef = useRef<import('../types/content').Comment[]>([]);
  const deletedCommentIdsRef = useRef<Set<number>>(new Set());
  const editedCommentsRef = useRef<Map<number, string>>(new Map()); // id → newText

  /**
   * Apply pending comment mutations (adds, edits, deletes) to a Document object.
   * Called before save/export to ensure the agent document has the latest state.
   * Uses refs that survive ProseMirror history state changes.
   */
  const applyCommentMutations = useCallback((doc: any) => {
    if (!doc?.package?.document) return;
    const hasMutations =
      addedCommentsRef.current.length > 0 ||
      deletedCommentIdsRef.current.size > 0 ||
      editedCommentsRef.current.size > 0;
    if (!hasMutations) return;

    let comments: import('../types/content').Comment[] = doc.package.document.comments || [];

    // 1. Remove deleted comments
    if (deletedCommentIdsRef.current.size > 0) {
      comments = comments.filter((c) => !deletedCommentIdsRef.current.has(c.id));
    }

    // 2. Add new comments (avoid duplicates)
    if (addedCommentsRef.current.length > 0) {
      const existingIds = new Set(comments.map((c) => c.id));
      for (const c of addedCommentsRef.current) {
        if (!existingIds.has(c.id)) {
          comments.push(c);
        }
      }
    }

    // 3. Apply edits
    if (editedCommentsRef.current.size > 0) {
      comments = comments.map((c) => {
        const newText = editedCommentsRef.current.get(c.id);
        if (newText == null) return c;
        return {
          ...c,
          content: [
            {
              type: 'paragraph' as const,
              content: [
                { type: 'run' as const, content: [{ type: 'text' as const, text: newText }] },
              ],
              formatting: {},
            },
          ],
        };
      });
    }

    doc.package.document.comments = comments;
    doc.commentsModified = true;
  }, []);

  // Remove comments whose marks no longer exist in the document.
  // Called (debounced) from handleDocumentChange so orphaned comments left
  // by Ctrl+A→Delete or similar bulk deletions are cleaned up automatically.
  const cleanOrphanedCommentsTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const cleanOrphanedComments = useCallback(() => {
    const view = pagedEditorRef.current?.getView();
    if (!view) return;
    const { doc: pmDoc, schema } = view.state;
    const commentMarkType = schema.marks.comment;
    if (!commentMarkType) return;

    // Collect all comment IDs still present as PM marks
    const liveIds = new Set<number>();
    pmDoc.descendants((node) => {
      for (const mark of node.marks) {
        if (mark.type === commentMarkType) {
          liveIds.add(mark.attrs.commentId as number);
        }
      }
    });

    // Check the current document's comment list for orphans
    const currentDoc = historyStateRef.current;
    const allComments = currentDoc?.package.document?.comments ?? [];
    for (const c of allComments) {
      if (c.parentId == null && !liveIds.has(c.id) && !deletedCommentIdsRef.current.has(c.id)) {
        // Parent comment with no PM mark — mark as deleted (replies pruned during save)
        deletedCommentIdsRef.current.add(c.id);
      }
    }
    // Also clean addedCommentsRef so newly-added then immediately deleted comments are removed
    addedCommentsRef.current = addedCommentsRef.current.filter(
      (c) => c.parentId != null || liveIds.has(c.id)
    );
  }, []);

  // Keep cross-ref updater plugin refs in sync with latest data
  crossRefNumMapRef.current = history.state?.package.numbering
    ? createNumberingMapFromDefs(history.state.package.numbering)
    : null;
  crossRefStyleResolverRef.current = history.state?.package.styles
    ? createStyleResolver(history.state.package.styles)
    : null;
  // Populate the style store so ProseMirror commands (Enter key, applyStyle)
  // can access style.next and font fallback without threading styles through.
  setDocumentStyles(history.state?.package.styles?.styles ?? []);
  // Track current border color/width for border presets (like Google Docs)
  const borderSpecRef = useRef({ style: 'single', size: 4, color: { rgb: '000000' } });

  // Measure toolbar height for positioning the outline panel below it
  const toolbarRefCallback = useCallback((el: HTMLDivElement | null) => {
    toolbarWrapperRef.current = el;
    // Clean up previous observer
    if (toolbarRoRef.current) {
      toolbarRoRef.current.disconnect();
      toolbarRoRef.current = null;
    }
    if (!el) {
      setToolbarHeight(0);
      return;
    }
    setToolbarHeight(el.offsetHeight);
    const ro = new ResizeObserver(() => {
      setToolbarHeight(el.offsetHeight);
    });
    ro.observe(el);
    toolbarRoRef.current = ro;
  }, []);

  // Cleanup ResizeObserver on unmount
  useEffect(() => {
    return () => {
      toolbarRoRef.current?.disconnect();
    };
  }, []);

  // Helper to get the active editor's view — returns HF editor view when in HF editing mode
  const getActiveEditorView = useCallback(() => {
    if (hfEditPosition && hfEditorRef.current) {
      return hfEditorRef.current.getView();
    }
    return pagedEditorRef.current?.getView();
  }, [hfEditPosition]);

  // Helper to focus the active editor
  const focusActiveEditor = useCallback(() => {
    if (hfEditPosition && hfEditorRef.current) {
      hfEditorRef.current.focus();
    } else {
      pagedEditorRef.current?.focus();
    }
  }, [hfEditPosition]);

  // Helper to undo in the active editor
  const undoActiveEditor = useCallback(() => {
    if (hfEditPosition && hfEditorRef.current) {
      hfEditorRef.current.undo();
    } else {
      pagedEditorRef.current?.undo();
    }
  }, [hfEditPosition]);

  // Helper to redo in the active editor
  const redoActiveEditor = useCallback(() => {
    if (hfEditPosition && hfEditorRef.current) {
      hfEditorRef.current.redo();
    } else {
      pagedEditorRef.current?.redo();
    }
  }, [hfEditPosition]);

  // Find/Replace hook
  const findReplace = useFindReplace();

  // Hyperlink dialog hook
  const hyperlinkDialog = useHyperlinkDialog();

  // Parse document buffer
  useEffect(() => {
    if (!documentBuffer) {
      if (initialDocument) {
        history.reset(initialDocument);
        setState((prev) => ({ ...prev, isLoading: false }));
        // Load fonts for initial document
        loadDocumentFonts(initialDocument).catch((err) => {
          console.warn('Failed to load document fonts:', err);
        });
      }
      return;
    }

    setState((prev) => ({ ...prev, isLoading: true, parseError: null }));

    // Clear comment mutation refs when loading a new document —
    // prevents stale adds/edits/deletes from a previous session being re-applied.
    addedCommentsRef.current = [];
    deletedCommentIdsRef.current = new Set();
    editedCommentsRef.current = new Map();

    const parseDocument = async () => {
      try {
        const doc = await parseDocx(documentBuffer, { enableContextTags });

        // Report loop diffs if present (from re-uploaded expanded loops)
        if (doc.loopDiffReports && doc.loopDiffReports.length > 0) {
          onLoopDiffsDetected?.(doc.loopDiffReports);
        }

        // Reset history with parsed document (clears undo/redo stacks)
        history.reset(doc);
        setState((prev) => ({
          ...prev,
          isLoading: false,
          parseError: null,
        }));

        // Load fonts used in the document from Google Fonts
        loadDocumentFonts(doc).catch((err) => {
          console.warn('Failed to load document fonts:', err);
        });
      } catch (error) {
        const message = error instanceof Error ? error.message : 'Failed to parse document';
        setState((prev) => ({
          ...prev,
          isLoading: false,
          parseError: message,
        }));
        onError?.(error instanceof Error ? error : new Error(message));
      }
    };

    parseDocument();
  }, [documentBuffer, initialDocument, onError]); // eslint-disable-line react-hooks/exhaustive-deps

  // Update document when initialDocument changes
  useEffect(() => {
    if (initialDocument && !documentBuffer) {
      history.reset(initialDocument);
    }
  }, [initialDocument, documentBuffer]); // eslint-disable-line react-hooks/exhaustive-deps

  // Create/update agent when document changes
  useEffect(() => {
    if (history.state) {
      agentRef.current = new DocumentAgent(history.state);
    } else {
      agentRef.current = null;
    }
  }, [history.state]);

  // Listen for font loading
  useEffect(() => {
    const cleanup = onFontsLoaded(() => {
      onFontsLoadedCallback?.();
    });
    return cleanup;
  }, [onFontsLoadedCallback]);

  // Clean up orphaned-comment debounce timer on unmount
  useEffect(() => {
    return () => {
      if (cleanOrphanedCommentsTimerRef.current) {
        clearTimeout(cleanOrphanedCommentsTimerRef.current);
      }
    };
  }, []);

  // Handle document change
  const handleDocumentChange = useCallback(
    (newDocument: Document) => {
      // Mark as dirty so repackDocx re-serializes document.xml instead of using original
      const dirtyDoc = newDocument.contentDirty
        ? newDocument
        : { ...newDocument, contentDirty: true };
      history.push(dirtyDoc);
      onChange?.(dirtyDoc);
      // Update outline headings if sidebar is open
      if (showOutlineRef.current) {
        const view = pagedEditorRef.current?.getView();
        if (view) {
          setHeadingInfos(collectHeadings(view.state.doc));
        }
      }
      // Clean up orphaned comments (debounced — avoid full doc walk on every keystroke)
      if (cleanOrphanedCommentsTimerRef.current) {
        clearTimeout(cleanOrphanedCommentsTimerRef.current);
      }
      cleanOrphanedCommentsTimerRef.current = setTimeout(cleanOrphanedComments, 300);
    },
    [onChange, history, cleanOrphanedComments]
  );

  // Handle selection changes from ProseMirror
  const handleSelectionChange = useCallback(
    (selectionState: SelectionState | null) => {
      // Save selection for restoring after toolbar interactions
      const view = getActiveEditorView();
      if (view) {
        const { from, to } = view.state.selection;
        lastSelectionRef.current = { from, to };
      }

      // Also check table context from ProseMirror
      let pmTableCtx: TableContextInfo | null = null;
      if (view) {
        pmTableCtx = getTableContext(view.state);
        if (!pmTableCtx.isInTable) {
          pmTableCtx = null;
        }
      }

      // Check if cursor is on an image (NodeSelection)
      let pmImageCtx: typeof state.pmImageContext = null;
      if (view) {
        const sel = view.state.selection;
        // NodeSelection has a `node` property
        const selectedNode = (
          sel as { node?: { type: { name: string }; attrs: Record<string, unknown> } }
        ).node;
        if (selectedNode?.type.name === 'image') {
          pmImageCtx = {
            pos: sel.from,
            wrapType: (selectedNode.attrs.wrapType as string) ?? 'inline',
            displayMode: (selectedNode.attrs.displayMode as string) ?? 'inline',
            cssFloat: (selectedNode.attrs.cssFloat as string) ?? null,
            transform: (selectedNode.attrs.transform as string) ?? null,
            alt: (selectedNode.attrs.alt as string) ?? null,
            borderWidth: (selectedNode.attrs.borderWidth as number) ?? null,
            borderColor: (selectedNode.attrs.borderColor as string) ?? null,
            borderStyle: (selectedNode.attrs.borderStyle as string) ?? null,
          };
        }
      }

      if (!selectionState) {
        setState((prev) => ({
          ...prev,
          selectionFormatting: {},
          pmTableContext: pmTableCtx,
          pmImageContext: pmImageCtx,
        }));
        return;
      }

      // Update toolbar formatting from ProseMirror selection
      const { textFormatting, paragraphFormatting } = selectionState;

      // Extract font family (prefer ascii, fall back to hAnsi)
      const fontFamily = textFormatting.fontFamily?.ascii || textFormatting.fontFamily?.hAnsi;

      // Extract text color as hex string
      const textColor = textFormatting.color?.rgb ? `#${textFormatting.color.rgb}` : undefined;

      // Build list state from numPr
      const numPr = paragraphFormatting.numPr;
      const listState = numPr
        ? {
            type: (numPr.numId === 1 ? 'bullet' : 'numbered') as 'bullet' | 'numbered',
            level: numPr.ilvl ?? 0,
            isInList: true,
            numId: numPr.numId,
          }
        : undefined;

      const formatting: SelectionFormatting = {
        bold: textFormatting.bold,
        italic: textFormatting.italic,
        underline: !!textFormatting.underline,
        strike: textFormatting.strike,
        superscript: textFormatting.vertAlign === 'superscript',
        subscript: textFormatting.vertAlign === 'subscript',
        allCaps: textFormatting.allCaps,
        smallCaps: textFormatting.smallCaps,
        fontFamily,
        fontSize: textFormatting.fontSize,
        color: textColor,
        highlight: textFormatting.highlight,
        alignment: paragraphFormatting.alignment,
        lineSpacing: paragraphFormatting.lineSpacing,
        listState,
        styleId: selectionState.styleId ?? undefined,
        indentLeft: paragraphFormatting.indentLeft,
      };
      setState((prev) => ({
        ...prev,
        selectionFormatting: formatting,
        paragraphIndentLeft: paragraphFormatting.indentLeft ?? 0,
        paragraphIndentRight: paragraphFormatting.indentRight ?? 0,
        paragraphFirstLineIndent: paragraphFormatting.indentFirstLine ?? 0,
        paragraphHangingIndent: paragraphFormatting.hangingIndent ?? false,
        paragraphTabs: paragraphFormatting.tabs ?? null,
        pmTableContext: pmTableCtx,
        pmImageContext: pmImageCtx,
      }));

      // Notify parent
      onSelectionChange?.(selectionState);
    },
    [onSelectionChange]
  );

  // Table selection hook
  const tableSelection = useTableSelection({
    document: history.state,
    onChange: handleDocumentChange,
    onSelectionChange: (_context) => {
      // Could notify parent of table selection changes
    },
  });

  // Keyboard shortcuts for Find/Replace (Ctrl+F, Ctrl+H) and delete table selection
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      // Check for Ctrl+F (Find) or Ctrl+H (Replace)
      const isMac = navigator.platform.toUpperCase().indexOf('MAC') >= 0;
      const cmdOrCtrl = isMac ? e.metaKey : e.ctrlKey;

      // Delete selected table from layout selection (non-ProseMirror selection)
      if (!cmdOrCtrl && !e.shiftKey && !e.altKey) {
        if (e.key === 'Delete' || e.key === 'Backspace') {
          // If full table is selected via ProseMirror CellSelection, delete it.
          const view = pagedEditorRef.current?.getView();
          if (view) {
            const sel = view.state.selection as { $anchorCell?: unknown; forEachCell?: unknown };
            const isCellSel = '$anchorCell' in sel && typeof sel.forEachCell === 'function';
            if (isCellSel) {
              const context = getTableContext(view.state);
              if (context.isInTable && context.table) {
                let totalCells = 0;
                context.table.descendants((node) => {
                  if (node.type.name === 'tableCell' || node.type.name === 'tableHeader') {
                    totalCells += 1;
                  }
                });
                let selectedCells = 0;
                (sel as { forEachCell: (fn: () => void) => void }).forEachCell(() => {
                  selectedCells += 1;
                });
                if (totalCells > 0 && selectedCells >= totalCells) {
                  e.preventDefault();
                  extensionManager.getCommands().deleteTable()(view.state, view.dispatch);
                  return;
                }
              }
            }
          }

          if (tableSelection.state.tableIndex !== null) {
            e.preventDefault();
            tableSelection.handleAction('deleteTable');
            return;
          }
        }
      }

      if (cmdOrCtrl && !e.shiftKey && !e.altKey) {
        if (e.key.toLowerCase() === 'f') {
          e.preventDefault();
          // Get selected text if any
          const selection = window.getSelection();
          const selectedText = selection && !selection.isCollapsed ? selection.toString() : '';
          findReplace.openFind(selectedText);
        } else if (e.key.toLowerCase() === 'h') {
          e.preventDefault();
          // Get selected text if any
          const selection = window.getSelection();
          const selectedText = selection && !selection.isCollapsed ? selection.toString() : '';
          findReplace.openReplace(selectedText);
        } else if (e.key.toLowerCase() === 'k') {
          e.preventDefault();
          // Open hyperlink dialog
          const view = pagedEditorRef.current?.getView();
          if (view) {
            const selectedText = getSelectedText(view.state);
            const existingLink = getHyperlinkAttrs(view.state);
            if (existingLink) {
              hyperlinkDialog.openEdit({
                url: existingLink.href,
                displayText: selectedText,
                tooltip: existingLink.tooltip,
              });
            } else {
              hyperlinkDialog.openInsert(selectedText);
            }
          }
        }
      }
    };

    document.addEventListener('keydown', handleKeyDown);
    return () => {
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [findReplace, hyperlinkDialog, tableSelection]);

  // Handle table insert from toolbar
  // Builds the table directly from the view's schema (bypassing the command
  // system) so that allowLockedEdit works reliably in locked-editing mode.
  const handleInsertTable = useCallback(
    (rows: number, columns: number) => {
      const view = getActiveEditorView();
      if (!view) return;

      const { state } = view;
      const { schema } = state;
      const { $from } = state.selection;

      // Find insertion point: after the current block-level node
      let insertPos = $from.pos;
      for (let d = $from.depth; d > 0; d--) {
        const node = $from.node(d);
        if (node.type.name === 'paragraph' || node.type.name === 'table') {
          insertPos = $from.after(d);
          break;
        }
      }

      // Determine content width (default full page; narrower inside a cell)
      let contentWidthTwips = 9360;
      for (let d = $from.depth; d > 0; d--) {
        const node = $from.node(d);
        if (node.type.name === 'tableCell' || node.type.name === 'tableHeader') {
          const cellWidth = node.attrs.width as number | undefined;
          if (cellWidth && cellWidth > 0) {
            contentWidthTwips = Math.max(cellWidth - 216, 360);
          }
          break;
        }
      }

      // Build table node
      const colWidthTwips = Math.floor(contentWidthTwips / columns);
      const defaultBorder = { style: 'single', size: 4, color: { rgb: '000000' } };
      const defaultBorders = {
        top: defaultBorder,
        bottom: defaultBorder,
        left: defaultBorder,
        right: defaultBorder,
      };

      const tableRows = [];
      for (let r = 0; r < rows; r++) {
        const cells = [];
        for (let c = 0; c < columns; c++) {
          const para = schema.nodes.paragraph.create();
          cells.push(
            schema.nodes.tableCell.create(
              {
                colspan: 1,
                rowspan: 1,
                borders: defaultBorders,
                width: colWidthTwips,
                widthType: 'dxa',
              },
              para
            )
          );
        }
        tableRows.push(schema.nodes.tableRow.create({ height: 360, heightRule: 'atLeast' }, cells));
      }

      const tableNode = schema.nodes.table.create(
        {
          columnWidths: Array(columns).fill(colWidthTwips),
          width: contentWidthTwips,
          widthType: 'dxa',
          justification: 'center',
        },
        tableRows
      );
      const emptyParagraph = schema.nodes.paragraph.create();

      // Build & dispatch transaction
      const $insert = state.doc.resolve(insertPos);
      const needsLeadingParagraph = $insert.nodeBefore?.type.name === 'table';
      const insertContent = needsLeadingParagraph
        ? [emptyParagraph, tableNode, schema.nodes.paragraph.create()]
        : [tableNode, emptyParagraph];

      const tr = state.tr;
      tr.setMeta('allowLockedEdit', true);
      tr.insert(insertPos, insertContent);

      // Place cursor inside the first cell: table(+1) > row(+1) > cell(+1) > paragraph(+1) = 4 levels
      let tableNodePos = insertPos;
      if (needsLeadingParagraph) {
        tableNodePos += emptyParagraph.nodeSize;
      }
      try {
        tr.setSelection(TextSelection.create(tr.doc, tableNodePos + 4));
      } catch {
        // If cursor placement fails, the insert still happened — just don't crash
      }

      view.dispatch(tr.scrollIntoView());
      focusActiveEditor();
    },
    [getActiveEditorView, focusActiveEditor]
  );

  // Insert a page break at cursor
  const handleInsertPageBreak = useCallback(() => {
    const view = getActiveEditorView();
    if (!view) return;
    insertPageBreak(view.state, view.dispatch);
    focusActiveEditor();
  }, [getActiveEditorView, focusActiveEditor]);

  // Insert a table of contents at cursor
  const handleInsertTOC = useCallback(() => {
    const view = getActiveEditorView();
    if (!view) return;
    generateTOC(view.state, (tr) => {
      tr.setMeta('allowLockedEdit', true);
      view.dispatch(tr);
    });
    focusActiveEditor();
  }, [getActiveEditorView, focusActiveEditor]);

  // Toggle document outline sidebar
  const handleToggleOutline = useCallback(() => {
    setShowOutline((prev) => {
      if (!prev) {
        // Opening: collect headings immediately
        const view = pagedEditorRef.current?.getView();
        if (view) {
          setHeadingInfos(collectHeadings(view.state.doc));
        }
      }
      return !prev;
    });
  }, []);

  // Navigate to a heading from the outline
  const handleHeadingInfoClick = useCallback((pmPos: number) => {
    pagedEditorRef.current?.scrollToPosition(pmPos);
    // Also set selection to the heading
    pagedEditorRef.current?.setSelection(pmPos + 1);
    pagedEditorRef.current?.focus();
  }, []);

  // Trigger file picker for image insert
  const handleInsertImageClick = useCallback(() => {
    imageInputRef.current?.click();
  }, []);

  // Handle file selection for image insert
  const handleImageFileChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (!file) return;

      const view = getActiveEditorView();
      if (!view) return;

      const reader = new FileReader();
      reader.onload = () => {
        const dataUrl = reader.result as string;

        // Create an Image element to get natural dimensions
        const img = new Image();
        img.onload = () => {
          let width = img.naturalWidth;
          let height = img.naturalHeight;

          // Constrain to reasonable max width (content area of US Letter page at 96dpi)
          const maxWidth = 612; // ~6.375 inches
          if (width > maxWidth) {
            const scale = maxWidth / width;
            width = maxWidth;
            height = Math.round(height * scale);
          }

          const rId = `rId_img_${Date.now()}`;
          const imageNode = view.state.schema.nodes.image.create({
            src: dataUrl,
            alt: file.name,
            width,
            height,
            rId,
            wrapType: 'inline',
            displayMode: 'inline',
          });

          const { from } = view.state.selection;
          const tr = view.state.tr.insert(from, imageNode);
          view.dispatch(tr.scrollIntoView());
          focusActiveEditor();
        };
        img.src = dataUrl;
      };
      reader.readAsDataURL(file);

      // Reset the input so the same file can be selected again
      e.target.value = '';
    },
    [getActiveEditorView, focusActiveEditor]
  );

  // Handle shape insertion
  // Handle image wrap type change
  const handleImageWrapType = useCallback(
    (wrapType: string) => {
      const view = getActiveEditorView();
      if (!view || !state.pmImageContext) return;

      const pos = state.pmImageContext.pos;
      const node = view.state.doc.nodeAt(pos);
      if (!node || node.type.name !== 'image') return;

      // Map wrap type to display mode + cssFloat
      let displayMode = 'inline';
      let cssFloat: string | null = null;

      switch (wrapType) {
        case 'inline':
          displayMode = 'inline';
          cssFloat = null;
          break;
        case 'square':
        case 'tight':
        case 'through':
          displayMode = 'float';
          cssFloat = 'left';
          break;
        case 'topAndBottom':
          displayMode = 'block';
          cssFloat = null;
          break;
        case 'behind':
        case 'inFront':
          displayMode = 'float';
          cssFloat = 'none';
          break;
        case 'wrapLeft':
          displayMode = 'float';
          cssFloat = 'right';
          wrapType = 'square';
          break;
        case 'wrapRight':
          displayMode = 'float';
          cssFloat = 'left';
          wrapType = 'square';
          break;
      }

      const tr = view.state.tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        wrapType,
        displayMode,
        cssFloat,
      });
      view.dispatch(tr.scrollIntoView());
      focusActiveEditor();
    },
    [getActiveEditorView, focusActiveEditor, state.pmImageContext]
  );

  // Handle image transform (rotate/flip)
  const handleImageTransform = useCallback(
    (action: 'rotateCW' | 'rotateCCW' | 'flipH' | 'flipV') => {
      const view = getActiveEditorView();
      if (!view || !state.pmImageContext) return;

      const pos = state.pmImageContext.pos;
      const node = view.state.doc.nodeAt(pos);
      if (!node || node.type.name !== 'image') return;

      const currentTransform = (node.attrs.transform as string) || '';

      // Parse current rotation and flip state
      const rotateMatch = currentTransform.match(/rotate\((-?\d+(?:\.\d+)?)deg\)/);
      let rotation = rotateMatch ? parseFloat(rotateMatch[1]) : 0;
      let hasFlipH = /scaleX\(-1\)/.test(currentTransform);
      let hasFlipV = /scaleY\(-1\)/.test(currentTransform);

      switch (action) {
        case 'rotateCW':
          rotation = (rotation + 90) % 360;
          break;
        case 'rotateCCW':
          rotation = (rotation - 90 + 360) % 360;
          break;
        case 'flipH':
          hasFlipH = !hasFlipH;
          break;
        case 'flipV':
          hasFlipV = !hasFlipV;
          break;
      }

      // Build new transform string
      const parts: string[] = [];
      if (rotation !== 0) parts.push(`rotate(${rotation}deg)`);
      if (hasFlipH) parts.push('scaleX(-1)');
      if (hasFlipV) parts.push('scaleY(-1)');
      const newTransform = parts.length > 0 ? parts.join(' ') : null;

      const tr = view.state.tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        transform: newTransform,
      });
      view.dispatch(tr.scrollIntoView());
      focusActiveEditor();
    },
    [getActiveEditorView, focusActiveEditor, state.pmImageContext]
  );

  // Apply image position changes
  const handleApplyImagePosition = useCallback(
    (data: ImagePositionData) => {
      const view = getActiveEditorView();
      if (!view || !state.pmImageContext) return;

      const pos = state.pmImageContext.pos;
      const node = view.state.doc.nodeAt(pos);
      if (!node || node.type.name !== 'image') return;

      const tr = view.state.tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        position: {
          horizontal: data.horizontal,
          vertical: data.vertical,
        },
        distTop: data.distTop ?? node.attrs.distTop,
        distBottom: data.distBottom ?? node.attrs.distBottom,
        distLeft: data.distLeft ?? node.attrs.distLeft,
        distRight: data.distRight ?? node.attrs.distRight,
      });
      view.dispatch(tr.scrollIntoView());
      focusActiveEditor();
    },
    [getActiveEditorView, focusActiveEditor, state.pmImageContext]
  );

  // Open image properties dialog
  const handleOpenImageProperties = useCallback(() => {
    setImagePropsOpen(true);
  }, []);

  // Apply image properties (alt text + border)
  const handleApplyImageProperties = useCallback(
    (data: ImagePropertiesData) => {
      const view = getActiveEditorView();
      if (!view || !state.pmImageContext) return;

      const pos = state.pmImageContext.pos;
      const node = view.state.doc.nodeAt(pos);
      if (!node || node.type.name !== 'image') return;

      const tr = view.state.tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        alt: data.alt ?? null,
        borderWidth: data.borderWidth ?? null,
        borderColor: data.borderColor ?? null,
        borderStyle: data.borderStyle ?? null,
      });
      view.dispatch(tr.scrollIntoView());
      focusActiveEditor();
    },
    [getActiveEditorView, focusActiveEditor, state.pmImageContext]
  );

  // Handle footnote/endnote properties update
  const handleApplyFootnoteProperties = useCallback(
    (
      footnotePr: import('../types/document').FootnoteProperties,
      endnotePr: import('../types/document').EndnoteProperties
    ) => {
      if (!history.state?.package) return;
      const newDoc = {
        ...history.state.package.document,
        finalSectionProperties: {
          ...history.state.package.document.finalSectionProperties,
          footnotePr,
          endnotePr,
          rawXml: undefined, // Invalidate raw XML since properties changed
        },
      };
      history.push({
        ...history.state,
        package: {
          ...history.state.package,
          document: newDoc,
        },
      });
    },
    [history]
  );

  // Handle table action from Toolbar - use commands from the per-instance
  // extensionManager (NOT the module-level singleton) so that the schema
  // used to create new nodes matches the editor's own schema.
  const handleTableAction = useCallback(
    (action: TableAction) => {
      const view = getActiveEditorView();
      if (!view) return;
      // Get commands from the editor's own extension manager
      const cmds = extensionManager.getCommands();

      switch (action) {
        case 'addRowAbove':
          cmds.addRowAbove()(view.state, view.dispatch);
          break;
        case 'addRowBelow':
          cmds.addRowBelow()(view.state, view.dispatch);
          break;
        case 'addColumnLeft':
          cmds.addColumnLeft()(view.state, view.dispatch);
          break;
        case 'addColumnRight':
          cmds.addColumnRight()(view.state, view.dispatch);
          break;
        case 'deleteRow':
          cmds.deleteRow()(view.state, view.dispatch);
          break;
        case 'deleteColumn':
          cmds.deleteColumn()(view.state, view.dispatch);
          break;
        case 'deleteTable':
          cmds.deleteTable()(view.state, view.dispatch);
          break;
        case 'selectTable':
          cmds.selectTable()(view.state, view.dispatch);
          break;
        case 'selectRow':
          cmds.selectRow()(view.state, view.dispatch);
          break;
        case 'selectColumn':
          cmds.selectColumn()(view.state, view.dispatch);
          break;
        case 'mergeCells':
          cmds.mergeCells()(view.state, view.dispatch);
          break;
        case 'splitCell':
          cmds.splitCell()(view.state, view.dispatch);
          break;
        // Border actions — use current border spec from toolbar
        case 'borderAll':
          setAllTableBorders(view.state, view.dispatch, borderSpecRef.current);
          break;
        case 'borderOutside':
          setOutsideTableBorders(view.state, view.dispatch, borderSpecRef.current);
          break;
        case 'borderInside':
          setInsideTableBorders(view.state, view.dispatch, borderSpecRef.current);
          break;
        case 'borderNone':
          removeTableBorders(view.state, view.dispatch);
          break;
        // Per-side border actions (use current border spec)
        case 'borderTop':
          setCellBorder('top', borderSpecRef.current)(view.state, view.dispatch);
          break;
        case 'borderBottom':
          setCellBorder('bottom', borderSpecRef.current)(view.state, view.dispatch);
          break;
        case 'borderLeft':
          setCellBorder('left', borderSpecRef.current)(view.state, view.dispatch);
          break;
        case 'borderRight':
          setCellBorder('right', borderSpecRef.current)(view.state, view.dispatch);
          break;
        case 'addCaption': {
          // Insert a "Table " + SEQ_FIELD + ": " caption paragraph after the table
          const { state } = view;
          const { $from } = state.selection;
          const schema = state.schema;

          // Walk up from cursor to find the table node
          let tablePos: number | null = null;
          let tableNode: import('prosemirror-model').Node | null = null;
          for (let d = $from.depth; d >= 1; d--) {
            const node = $from.node(d);
            if (node.type.name === 'table') {
              tablePos = $from.before(d);
              tableNode = node;
              break;
            }
          }
          if (tablePos != null && tableNode) {
            const insertPos = tablePos + tableNode.nodeSize;

            // Count existing "Table" captions before insertion position
            let tableCaptionCount = 0;
            state.doc.nodesBetween(0, insertPos, (node) => {
              if (
                node.type.name === 'paragraph' &&
                node.attrs.styleId === 'Caption' &&
                node.textContent.startsWith('Table ')
              ) {
                tableCaptionCount++;
              }
              return true;
            });
            const tableNumber = tableCaptionCount + 1;

            // Resolve Caption style's run formatting and build marks
            let captionMarks: import('prosemirror-model').Mark[] = [];
            const currentDoc = historyStateRef.current;
            if (currentDoc?.package.styles) {
              const styleResolver = createStyleResolver(currentDoc.package.styles);
              const resolved = styleResolver.resolveParagraphStyle('Caption');
              if (resolved.runFormatting) {
                captionMarks = textFormattingToMarks(resolved.runFormatting, schema);
              }
            }

            // Build: text("Table ") + field(SEQ Table) + text(": ")
            const prefixNode =
              captionMarks.length > 0 ? schema.text('Table ', captionMarks) : schema.text('Table ');
            let seqField = schema.nodes.field.create({
              fieldType: 'SEQ',
              instruction: ' SEQ Table \\* ARABIC ',
              displayText: String(tableNumber),
              fieldKind: 'complex',
              dirty: false,
            });
            if (captionMarks.length > 0) {
              seqField = seqField.mark(captionMarks);
            }
            const suffixNode =
              captionMarks.length > 0 ? schema.text(': ', captionMarks) : schema.text(': ');

            const captionParagraph = schema.nodes.paragraph.create(
              { styleId: 'Caption', alignment: 'center' },
              [prefixNode, seqField, suffixNode]
            );
            const tr = state.tr.insert(insertPos, captionParagraph);

            // Place cursor at end of ": " so user can type description
            // paragraph open(1) + "Table "(6) + field atom(1) + ": "(2) = 10
            const cursorPos = insertPos + 1 + 6 + 1 + 2;
            tr.setSelection(TextSelection.create(tr.doc, cursorPos));

            // Set stored marks so continued typing inherits caption formatting
            if (captionMarks.length > 0) {
              tr.setStoredMarks(captionMarks);
            }

            view.dispatch(tr.scrollIntoView());
          }
          break;
        }
        default:
          // Handle complex actions (with parameters)
          if (typeof action === 'object') {
            if (action.type === 'cellFillColor') {
              setCellFillColor(action.color)(view.state, view.dispatch);
            } else if (action.type === 'borderColor') {
              const rgb = action.color.replace(/^#/, '');
              borderSpecRef.current = { ...borderSpecRef.current, color: { rgb } };
              setTableBorderColor(action.color)(view.state, view.dispatch);
            } else if (action.type === 'borderWidth') {
              borderSpecRef.current = { ...borderSpecRef.current, size: action.size };
              setTableBorderWidth(action.size)(view.state, view.dispatch);
            } else if (action.type === 'cellBorder') {
              setCellBorder(action.side, {
                style: action.style,
                size: action.size,
                color: { rgb: action.color.replace(/^#/, '') },
              })(view.state, view.dispatch);
            } else if (action.type === 'cellVerticalAlign') {
              setCellVerticalAlign(action.align)(view.state, view.dispatch);
            } else if (action.type === 'cellMargins') {
              setCellMargins(action.margins)(view.state, view.dispatch);
            } else if (action.type === 'cellTextDirection') {
              setCellTextDirection(action.direction)(view.state, view.dispatch);
            } else if (action.type === 'toggleNoWrap') {
              toggleNoWrap()(view.state, view.dispatch);
            } else if (action.type === 'rowHeight') {
              setRowHeight(action.height, action.rule)(view.state, view.dispatch);
            } else if (action.type === 'toggleHeaderRow') {
              toggleHeaderRow()(view.state, view.dispatch);
            } else if (action.type === 'distributeColumns') {
              distributeColumns()(view.state, view.dispatch);
            } else if (action.type === 'autoFitContents') {
              autoFitContents()(view.state, view.dispatch);
            } else if (action.type === 'openTableProperties') {
              setTablePropsOpen(true);
            } else if (action.type === 'tableProperties') {
              setTableProperties(action.props)(view.state, view.dispatch);
            } else if (action.type === 'applyTableStyle') {
              // Resolve style data from built-in presets or document styles
              let preset: TableStylePreset | undefined = getBuiltinTableStyle(action.styleId);
              const currentDocForTable = historyStateRef.current;
              if (!preset && currentDocForTable?.package.styles) {
                const styleResolver = createStyleResolver(currentDocForTable.package.styles);
                const docStyle = styleResolver.getStyle(action.styleId);
                if (docStyle) {
                  // Convert to preset inline (same as documentStyleToPreset)
                  preset = { id: docStyle.styleId, name: docStyle.name ?? docStyle.styleId };
                  if (docStyle.tblPr?.borders) {
                    const b = docStyle.tblPr.borders;
                    preset.tableBorders = {};
                    for (const side of [
                      'top',
                      'bottom',
                      'left',
                      'right',
                      'insideH',
                      'insideV',
                    ] as const) {
                      const bs = b[side];
                      if (bs) {
                        preset.tableBorders[side] = {
                          style: bs.style,
                          size: bs.size,
                          color: bs.color?.rgb ? { rgb: bs.color.rgb } : undefined,
                        };
                      }
                    }
                  }
                  if (docStyle.tblStylePr) {
                    preset.conditionals = {};
                    for (const cond of docStyle.tblStylePr) {
                      const entry: Record<string, unknown> = {};
                      if (cond.tcPr?.shading?.fill)
                        entry.backgroundColor = `#${cond.tcPr.shading.fill}`;
                      if (cond.tcPr?.borders) {
                        const borders: Record<string, unknown> = {};
                        for (const s of ['top', 'bottom', 'left', 'right'] as const) {
                          const bs2 = cond.tcPr.borders[s];
                          if (bs2)
                            borders[s] = {
                              style: bs2.style,
                              size: bs2.size,
                              color: bs2.color?.rgb ? { rgb: bs2.color.rgb } : undefined,
                            };
                        }
                        entry.borders = borders;
                      }
                      if (cond.rPr?.bold) entry.bold = true;
                      if (cond.rPr?.color?.rgb) entry.color = `#${cond.rPr.color.rgb}`;
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      (preset.conditionals as any)[cond.type] = entry;
                    }
                  }
                  preset.look = { firstRow: true, lastRow: false, noHBand: false, noVBand: true };
                }
              }
              if (preset) {
                applyTableStyle({
                  styleId: preset.id,
                  tableBorders: preset.tableBorders,
                  conditionals: preset.conditionals,
                  look: preset.look,
                })(view.state, view.dispatch);
              }
            }
          } else {
            // Fallback to legacy table selection handler for other actions
            tableSelection.handleAction(action);
          }
      }

      focusActiveEditor();
    },
    [tableSelection, getActiveEditorView, focusActiveEditor]
  );

  // Handle formatting action from toolbar
  const handleFormat = useCallback(
    (action: FormattingAction) => {
      const view = getActiveEditorView();
      if (!view) return;

      // Focus editor first to ensure we can dispatch commands
      view.focus();

      // Restore selection if it was lost during toolbar interaction
      // This happens when user clicks on dropdown menus (font picker, style picker, etc.)
      // Only restore for the body editor — HF editor manages its own selection
      const isBodyEditor = view === pagedEditorRef.current?.getView();
      const { from, to } = view.state.selection;
      const savedSelection = lastSelectionRef.current;

      if (
        isBodyEditor &&
        savedSelection &&
        (from !== savedSelection.from || to !== savedSelection.to)
      ) {
        // Selection was lost (focus moved to dropdown portal) - restore it
        try {
          const tr = view.state.tr.setSelection(
            TextSelection.create(view.state.doc, savedSelection.from, savedSelection.to)
          );
          view.dispatch(tr);
        } catch (e) {
          // If restoration fails (e.g., positions are invalid after doc change), continue with current selection
          console.warn('Could not restore selection:', e);
        }
      }

      // Handle simple toggle actions
      if (action === 'bold') {
        toggleBold(view.state, view.dispatch);
        return;
      }
      if (action === 'italic') {
        toggleItalic(view.state, view.dispatch);
        return;
      }
      if (action === 'underline') {
        toggleUnderline(view.state, view.dispatch);
        return;
      }
      if (action === 'strikethrough') {
        toggleStrike(view.state, view.dispatch);
        return;
      }
      if (action === 'superscript') {
        toggleSuperscript(view.state, view.dispatch);
        return;
      }
      if (action === 'subscript') {
        toggleSubscript(view.state, view.dispatch);
        return;
      }
      if (action === 'allCaps') {
        toggleAllCaps(view.state, view.dispatch);
        return;
      }
      if (action === 'smallCaps') {
        toggleSmallCaps(view.state, view.dispatch);
        return;
      }
      if (action === 'bulletList') {
        toggleBulletList(view.state, view.dispatch);
        return;
      }
      if (action === 'numberedList') {
        toggleNumberedList(view.state, view.dispatch);
        return;
      }
      if (action === 'indent') {
        // Try list indent first, then paragraph indent
        if (!increaseListLevel(view.state, view.dispatch)) {
          increaseIndent()(view.state, view.dispatch);
        }
        return;
      }
      if (action === 'outdent') {
        // Try list outdent first, then paragraph outdent
        if (!decreaseListLevel(view.state, view.dispatch)) {
          decreaseIndent()(view.state, view.dispatch);
        }
        return;
      }
      if (action === 'clearFormatting') {
        clearFormatting(view.state, view.dispatch);
        return;
      }
      if (action === 'insertLink') {
        // Get the selected text for the hyperlink dialog
        const selectedText = getSelectedText(view.state);
        // Check if we're editing an existing link
        const existingLink = getHyperlinkAttrs(view.state);
        if (existingLink) {
          hyperlinkDialog.openEdit({
            url: existingLink.href,
            displayText: selectedText,
            tooltip: existingLink.tooltip,
          });
        } else {
          hyperlinkDialog.openInsert(selectedText);
        }
        return;
      }

      // Handle object-based actions
      if (typeof action === 'object') {
        switch (action.type) {
          case 'alignment':
            setAlignment(action.value)(view.state, view.dispatch);
            break;
          case 'textColor':
            // action.value can be a string like "#FF0000" or a color name
            setTextColor({ rgb: action.value.replace('#', '') })(view.state, view.dispatch);
            break;
          case 'highlightColor': {
            // Convert hex to OOXML named highlight value (e.g., 'FFFF00' → 'yellow')
            const highlightName = action.value ? mapHexToHighlightName(action.value) : '';
            setHighlight(highlightName || action.value)(view.state, view.dispatch);
            break;
          }
          case 'fontSize':
            // Convert points to half-points (OOXML uses half-points for font sizes)
            setFontSize(pointsToHalfPoints(action.value))(view.state, view.dispatch);
            break;
          case 'fontFamily':
            setFontFamily(action.value)(view.state, view.dispatch);
            break;
          case 'lineSpacing':
            setLineSpacing(action.value)(view.state, view.dispatch);
            break;
          case 'applyStyle': {
            // Resolve style to get its formatting properties
            // Use ref to avoid stale closure (handleFormat has [] deps)
            const currentDoc = historyStateRef.current;
            const styleResolver = currentDoc?.package.styles
              ? createStyleResolver(currentDoc.package.styles)
              : null;

            if (styleResolver) {
              const resolved = styleResolver.resolveParagraphStyle(action.value);

              // Resolve list rendering from numbering definitions
              let listNumFmt: import('../types/document').NumberFormat | undefined;
              let listIsBullet: boolean | undefined;
              const numPr = resolved.paragraphFormatting?.numPr;
              if (numPr?.numId && numPr.numId !== 0 && currentDoc?.package.numbering) {
                const numMap = createNumberingMapFromDefs(currentDoc.package.numbering);
                const level = numMap.getLevel(numPr.numId, numPr.ilvl ?? 0);
                if (level) {
                  listNumFmt = level.numFmt;
                  listIsBullet = level.numFmt === 'bullet';
                  // Merge numbering-level indent into resolved paragraph formatting
                  // AMA Heading2/3/4 styles have NO w:ind in their own pPr — indent
                  // comes entirely from the numbering level definitions.
                  if (level.pPr) {
                    if (!resolved.paragraphFormatting) resolved.paragraphFormatting = {};
                    if (
                      level.pPr.indentLeft != null &&
                      resolved.paragraphFormatting.indentLeft == null
                    ) {
                      resolved.paragraphFormatting.indentLeft = level.pPr.indentLeft;
                    }
                    if (
                      level.pPr.hangingIndent != null &&
                      resolved.paragraphFormatting.hangingIndent == null
                    ) {
                      resolved.paragraphFormatting.hangingIndent = level.pPr.hangingIndent;
                    }
                    if (
                      level.pPr.indentFirstLine != null &&
                      resolved.paragraphFormatting.indentFirstLine == null
                    ) {
                      resolved.paragraphFormatting.indentFirstLine = level.pPr.indentFirstLine;
                    }
                  }
                }
              }

              applyStyle(action.value, {
                paragraphFormatting: resolved.paragraphFormatting,
                runFormatting: resolved.runFormatting,
                listNumFmt,
                listIsBullet,
              })(view.state, view.dispatch);
            } else {
              // No styles available, just set the styleId
              applyStyle(action.value)(view.state, view.dispatch);
            }
            break;
          }
        }
      }
    },
    [getActiveEditorView]
  );

  // Handle zoom change
  const handleZoomChange = useCallback((zoom: number) => {
    setState((prev) => ({ ...prev, zoom }));
  }, []);

  // Handle hyperlink dialog submit
  const handleHyperlinkSubmit = useCallback(
    (data: HyperlinkData) => {
      const view = getActiveEditorView();
      if (!view) return;

      const url = data.url || '';
      const tooltip = data.tooltip;

      // Check if we have a selection
      const { empty } = view.state.selection;

      if (empty && data.displayText) {
        // No selection but display text provided - insert new linked text
        insertHyperlink(data.displayText, url, tooltip)(view.state, view.dispatch);
      } else if (!empty) {
        // Have selection - apply hyperlink to it
        setHyperlink(url, tooltip)(view.state, view.dispatch);
      } else if (data.displayText) {
        // Empty selection but display text provided
        insertHyperlink(data.displayText, url, tooltip)(view.state, view.dispatch);
      }

      hyperlinkDialog.close();
      focusActiveEditor();
    },
    [hyperlinkDialog, getActiveEditorView, focusActiveEditor]
  );

  // Handle hyperlink removal
  const handleHyperlinkRemove = useCallback(() => {
    const view = getActiveEditorView();
    if (!view) return;

    removeHyperlink(view.state, view.dispatch);
    hyperlinkDialog.close();
    focusActiveEditor();
  }, [hyperlinkDialog, getActiveEditorView, focusActiveEditor]);

  // Handle margin changes from rulers
  const createMarginHandler = useCallback(
    (property: 'marginLeft' | 'marginRight' | 'marginTop' | 'marginBottom') =>
      (marginTwips: number) => {
        if (!history.state || readOnly) return;
        const newDoc = {
          ...history.state,
          package: {
            ...history.state.package,
            document: {
              ...history.state.package.document,
              finalSectionProperties: {
                ...history.state.package.document.finalSectionProperties,
                [property]: marginTwips,
                rawXml: undefined, // Invalidate raw XML since properties changed
              },
            },
          },
        };
        handleDocumentChange(newDoc);
      },
    [history.state, readOnly, handleDocumentChange]
  );

  const handleLeftMarginChange = useMemo(
    () => createMarginHandler('marginLeft'),
    [createMarginHandler]
  );
  const handleRightMarginChange = useMemo(
    () => createMarginHandler('marginRight'),
    [createMarginHandler]
  );
  const handleTopMarginChange = useMemo(
    () => createMarginHandler('marginTop'),
    [createMarginHandler]
  );
  const handleBottomMarginChange = useMemo(
    () => createMarginHandler('marginBottom'),
    [createMarginHandler]
  );

  // Paragraph indent handlers (for ruler)
  const handleIndentLeftChange = useCallback(
    (twips: number) => {
      const view = getActiveEditorView();
      if (!view) return;
      setIndentLeft(twips)(view.state, view.dispatch);
    },
    [getActiveEditorView]
  );

  const handleIndentRightChange = useCallback(
    (twips: number) => {
      const view = getActiveEditorView();
      if (!view) return;
      setIndentRight(twips)(view.state, view.dispatch);
    },
    [getActiveEditorView]
  );

  const handleFirstLineIndentChange = useCallback(
    (twips: number) => {
      const view = getActiveEditorView();
      if (!view) return;
      // If twips is negative, it's a hanging indent
      if (twips < 0) {
        setIndentFirstLine(-twips, true)(view.state, view.dispatch);
      } else {
        setIndentFirstLine(twips, false)(view.state, view.dispatch);
      }
    },
    [getActiveEditorView]
  );

  const handleTabStopRemove = useCallback(
    (positionTwips: number) => {
      const view = getActiveEditorView();
      if (!view) return;
      removeTabStop(positionTwips)(view.state, view.dispatch);
    },
    [getActiveEditorView]
  );

  // Handle page navigation (from PageNavigator)
  // TODO: Implement page navigation in ProseMirror
  const handlePageNavigate = useCallback((_pageNumber: number) => {
    // Page navigation not yet implemented for ProseMirror
  }, []);

  // Handle save
  const handleSave = useCallback(async (): Promise<ArrayBuffer | null> => {
    if (!agentRef.current) return null;

    try {
      // Collect context tag metadata from the current PM state before saving
      const view = pagedEditorRef.current?.getView();
      if (view?.state.schema.nodes.contextTag) {
        const ctMeta = collectContextTagMetadata(view.state.doc);
        agentRef.current.getDocument().contextTagMetadata =
          Object.keys(ctMeta).length > 0 ? ctMeta : undefined;
      }

      // Apply pending comment mutations (adds, edits, deletes) to the agent document.
      // These refs survive history state changes that would otherwise lose mutations.
      applyCommentMutations(agentRef.current.getDocument());

      const buffer = await agentRef.current.toBuffer();
      onSave?.(buffer);
      return buffer;
    } catch (error) {
      onError?.(error instanceof Error ? error : new Error('Failed to save document'));
      return null;
    }
  }, [onSave, onError]);

  // Handle error from editor
  const handleEditorError = useCallback(
    (error: Error) => {
      onError?.(error);
    },
    [onError]
  );

  const handleDirectPrint = useCallback(() => {
    // Find the pages container and clone its content into a clean print window
    const pagesEl = containerRef.current?.querySelector('.paged-editor__pages');
    if (!pagesEl) {
      window.print();
      onPrint?.();
      return;
    }

    const printWindow = window.open('', '_blank');
    if (!printWindow) {
      // Popup blocked — fall back to window.print()
      window.print();
      onPrint?.();
      return;
    }

    // Collect all @font-face rules from the current page
    const fontFaceRules: string[] = [];
    for (const sheet of Array.from(document.styleSheets)) {
      try {
        for (const rule of Array.from(sheet.cssRules)) {
          if (rule instanceof CSSFontFaceRule) {
            fontFaceRules.push(rule.cssText);
          }
        }
      } catch {
        // Cross-origin stylesheets can't be read — skip
      }
    }

    // Clone pages and remove transforms/shadows
    const pagesClone = pagesEl.cloneNode(true) as HTMLElement;
    pagesClone.style.cssText = 'display: block; margin: 0; padding: 0;';
    for (const page of Array.from(pagesClone.querySelectorAll('.layout-page'))) {
      const el = page as HTMLElement;
      el.style.boxShadow = 'none';
      el.style.margin = '0';
    }

    printWindow.document.write(`<!DOCTYPE html>
<html><head><title>Print</title>
<style>
${fontFaceRules.join('\n')}
* { margin: 0; padding: 0; }
body { background: white; }
.layout-page { break-after: page; }
.layout-page:last-child { break-after: auto; }
@page { margin: 0; size: auto; }
</style>
</head><body>${pagesClone.outerHTML}</body></html>`);
    printWindow.document.close();

    // Wait for fonts/images then print
    printWindow.onload = () => {
      printWindow.print();
      printWindow.close();
    };

    // Fallback if onload doesn't fire (some browsers)
    setTimeout(() => {
      if (!printWindow.closed) {
        printWindow.print();
        printWindow.close();
      }
    }, 1000);

    onPrint?.();
  }, [onPrint]);

  // ============================================================================
  // FIND/REPLACE HANDLERS
  // ============================================================================

  // Store the current find result for navigation
  const findResultRef = useRef<FindResult | null>(null);

  // Handle find operation
  const handleFind = useCallback(
    (searchText: string, options: FindOptions): FindResult | null => {
      if (!history.state || !searchText.trim()) {
        findResultRef.current = null;
        return null;
      }

      const matches = findInDocument(history.state, searchText, options);
      const result: FindResult = {
        matches,
        totalCount: matches.length,
        currentIndex: 0,
      };

      findResultRef.current = result;
      findReplace.setMatches(matches, 0);

      // Scroll to first match
      if (matches.length > 0 && containerRef.current) {
        scrollToMatch(containerRef.current, matches[0]);
      }

      return result;
    },
    [history.state, findReplace]
  );

  // Handle find next
  const handleFindNext = useCallback((): FindMatch | null => {
    if (!findResultRef.current || findResultRef.current.matches.length === 0) {
      return null;
    }

    const newIndex = findReplace.goToNextMatch();
    const match = findResultRef.current.matches[newIndex];

    // Scroll to the match
    if (match && containerRef.current) {
      scrollToMatch(containerRef.current, match);
    }

    return match || null;
  }, [findReplace]);

  // Handle find previous
  const handleFindPrevious = useCallback((): FindMatch | null => {
    if (!findResultRef.current || findResultRef.current.matches.length === 0) {
      return null;
    }

    const newIndex = findReplace.goToPreviousMatch();
    const match = findResultRef.current.matches[newIndex];

    // Scroll to the match
    if (match && containerRef.current) {
      scrollToMatch(containerRef.current, match);
    }

    return match || null;
  }, [findReplace]);

  // Handle replace current match
  const handleReplace = useCallback(
    (replaceText: string): boolean => {
      if (!history.state || !findResultRef.current || findResultRef.current.matches.length === 0) {
        return false;
      }

      const currentMatch = findResultRef.current.matches[findResultRef.current.currentIndex];
      if (!currentMatch) return false;

      // Execute replace command
      try {
        const newDoc = executeCommand(history.state, {
          type: 'replaceText',
          range: {
            start: {
              paragraphIndex: currentMatch.paragraphIndex,
              offset: currentMatch.startOffset,
            },
            end: {
              paragraphIndex: currentMatch.paragraphIndex,
              offset: currentMatch.endOffset,
            },
          },
          text: replaceText,
        });

        handleDocumentChange(newDoc);
        return true;
      } catch (error) {
        console.error('Replace failed:', error);
        return false;
      }
    },
    [history.state, handleDocumentChange]
  );

  // Handle replace all matches
  const handleReplaceAll = useCallback(
    (searchText: string, replaceText: string, options: FindOptions): number => {
      if (!history.state || !searchText.trim()) {
        return 0;
      }

      // Find all matches first
      const matches = findInDocument(history.state, searchText, options);
      if (matches.length === 0) return 0;

      // Replace from end to start to maintain correct indices
      let doc = history.state;
      const sortedMatches = [...matches].sort((a, b) => {
        if (a.paragraphIndex !== b.paragraphIndex) {
          return b.paragraphIndex - a.paragraphIndex;
        }
        return b.startOffset - a.startOffset;
      });

      for (const match of sortedMatches) {
        try {
          doc = executeCommand(doc, {
            type: 'replaceText',
            range: {
              start: {
                paragraphIndex: match.paragraphIndex,
                offset: match.startOffset,
              },
              end: {
                paragraphIndex: match.paragraphIndex,
                offset: match.endOffset,
              },
            },
            text: replaceText,
          });
        } catch (error) {
          console.error('Replace failed for match:', match, error);
        }
      }

      handleDocumentChange(doc);
      findResultRef.current = null;
      findReplace.setMatches([], 0);

      return matches.length;
    },
    [history.state, handleDocumentChange, findReplace]
  );

  // Expose ref methods
  useImperativeHandle(
    ref,
    () => ({
      getAgent: () => agentRef.current,
      getDocument: () => history.state,
      getEditorRef: () => pagedEditorRef.current,
      save: handleSave,
      setZoom: (zoom: number) => setState((prev) => ({ ...prev, zoom })),
      getZoom: () => state.zoom,
      setRenderMode: (mode: 'rendered' | 'raw') =>
        setState((prev) => ({ ...prev, renderMode: mode })),
      getRenderMode: () => state.renderMode,
      focus: () => {
        pagedEditorRef.current?.focus();
      },
      getCurrentPage: () => state.currentPage,
      getTotalPages: () => state.totalPages,
      scrollToPage: (_pageNumber: number) => {
        // TODO: Implement page navigation in ProseMirror
      },
      openPrintPreview: handleDirectPrint,
      print: handleDirectPrint,
      getSelectedText: () => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return '';
        const { from, to } = view.state.selection;
        if (from === to) return '';
        return view.state.doc.textBetween(from, to, ' ');
      },
      getSelectionRange: () => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return null;
        const { from, to } = view.state.selection;
        if (from === to) return null;
        return { from, to };
      },
      addComment: (author: string, text: string, fromPos?: number, toPos?: number) => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return null;
        // Use provided positions or fall back to current selection
        const from = fromPos ?? view.state.selection.from;
        const to = toPos ?? view.state.selection.to;
        if (from === to) return null; // Need a selection

        const schema = view.state.schema;
        const commentMark = schema.marks.comment;
        if (!commentMark) return null;

        // Generate a new comment ID (avoid collision with existing)
        const doc = historyStateRef.current;
        const existingIds = new Set<number>();
        if (doc?.package.document?.comments) {
          for (const c of doc.package.document.comments) {
            existingIds.add(c.id);
          }
        }
        let newId = 1;
        while (existingIds.has(newId)) newId++;

        // Apply comment mark to selection
        const mark = commentMark.create({ commentId: newId });
        const tr = view.state.tr.addMark(from, to, mark);
        view.dispatch(tr);

        // Add Comment object to document model AND to the addedComments ref
        const newComment: import('../types/content').Comment = {
          id: newId,
          author,
          date: new Date().toISOString(),
          content: [
            {
              type: 'paragraph' as const,
              content: [{ type: 'run' as const, content: [{ type: 'text' as const, text }] }],
              formatting: {},
            },
          ],
        };
        if (doc?.package.document) {
          if (!doc.package.document.comments) {
            doc.package.document.comments = [];
          }
          doc.package.document.comments.push(newComment);
          doc.commentsModified = true;
        }
        addedCommentsRef.current = [...addedCommentsRef.current, newComment];

        return newId;
      },
      addReplyComment: (parentCommentId: number, author: string, text: string) => {
        const doc = historyStateRef.current;
        if (!doc?.package.document) return null;

        // Generate a new comment ID (avoid collision with existing)
        const existingIds = new Set<number>();
        const commentsById = new Map<number, import('../types/content').Comment>();
        if (doc.package.document.comments) {
          for (const c of doc.package.document.comments) {
            existingIds.add(c.id);
            commentsById.set(c.id, c);
          }
        }
        for (const c of addedCommentsRef.current) {
          existingIds.add(c.id);
          commentsById.set(c.id, c);
        }
        let newId = 1;
        while (existingIds.has(newId)) newId++;

        // Word only supports 2 levels (parent + replies). If replying to a reply,
        // walk up to the root ancestor so the DOCX always has flat threading.
        let rootParentId = parentCommentId;
        let visited = 0;
        while (visited < 100) {
          const parent = commentsById.get(rootParentId);
          if (!parent || parent.parentId == null) break;
          rootParentId = parent.parentId;
          visited++;
        }

        const replyComment: import('../types/content').Comment = {
          id: newId,
          author,
          date: new Date().toISOString(),
          content: [
            {
              type: 'paragraph' as const,
              content: [{ type: 'run' as const, content: [{ type: 'text' as const, text }] }],
              formatting: {},
            },
          ],
          parentId: rootParentId,
        };

        if (!doc.package.document.comments) {
          doc.package.document.comments = [];
        }
        doc.package.document.comments.push(replyComment);
        doc.commentsModified = true;
        addedCommentsRef.current = [...addedCommentsRef.current, replyComment];

        return newId;
      },
      editComment: (commentId: number, newText: string) => {
        // Track edit in a ref that survives history state changes.
        editedCommentsRef.current.set(commentId, newText);

        // Also apply to current historyState for immediate display in the panel
        // (no PM transaction here, so historyStateRef is stable)
        const doc = historyStateRef.current;
        if (doc?.package.document?.comments) {
          const comments = doc.package.document.comments;
          const idx = comments.findIndex((c) => c.id === commentId);
          if (idx >= 0) {
            comments[idx] = {
              ...comments[idx],
              content: [
                {
                  type: 'paragraph' as const,
                  content: [
                    { type: 'run' as const, content: [{ type: 'text' as const, text: newText }] },
                  ],
                  formatting: {},
                },
              ],
            };
            doc.package.document.comments = [...comments];
            doc.commentsModified = true;
          }
        }

        // Also update addedCommentsRef if this comment was added in this session
        const newContent = [
          {
            type: 'paragraph' as const,
            content: [
              { type: 'run' as const, content: [{ type: 'text' as const, text: newText }] },
            ],
            formatting: {},
          },
        ];
        addedCommentsRef.current = addedCommentsRef.current.map((c) =>
          c.id === commentId ? { ...c, content: newContent } : c
        );
      },
      removeComment: (commentId: number) => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return;
        const schema = view.state.schema;
        const commentMark = schema.marks.comment;
        if (!commentMark) return;

        // Remove comment mark from entire document
        const tr = view.state.tr;
        const mark = commentMark.create({ commentId });
        const docSize = view.state.doc.content.size;
        view.dispatch(tr.removeMark(0, docSize, mark));

        // Track deletion in a ref that survives history state changes.
        // The actual filtering happens at save time via applyCommentMutations.
        deletedCommentIdsRef.current.add(commentId);
        addedCommentsRef.current = addedCommentsRef.current.filter((c) => c.id !== commentId);
        editedCommentsRef.current.delete(commentId);
      },
      getDocumentComments: () => {
        const doc = historyStateRef.current;
        let docComments = doc?.package.document?.comments || [];
        // Filter out deleted comments
        if (deletedCommentIdsRef.current.size > 0) {
          docComments = docComments.filter((c) => !deletedCommentIdsRef.current.has(c.id));
        }
        // Merge with addedCommentsRef to catch comments that haven't propagated to history yet
        const seenIds = new Set(docComments.map((c) => c.id));
        const extra = addedCommentsRef.current.filter((c) => !seenIds.has(c.id));
        let comments = [...docComments, ...extra];
        // Apply edits
        if (editedCommentsRef.current.size > 0) {
          comments = comments.map((c) => {
            const newText = editedCommentsRef.current.get(c.id);
            if (newText == null) return c;
            return {
              ...c,
              content: [
                {
                  type: 'paragraph' as const,
                  content: [
                    { type: 'run' as const, content: [{ type: 'text' as const, text: newText }] },
                  ],
                  formatting: {} as any,
                },
              ],
            };
          });
        }
        const view = pagedEditorRef.current?.getView();

        // Build anchor text map from PM doc comment marks
        const anchorMap = new Map<number, string>();
        if (view) {
          view.state.doc.descendants((node) => {
            if (node.isText) {
              for (const mark of node.marks) {
                if (mark.type.name === 'comment') {
                  const cid = mark.attrs.commentId as number;
                  const existing = anchorMap.get(cid) || '';
                  anchorMap.set(cid, existing + (node.text || ''));
                }
              }
            }
          });
        }

        return comments.map((c) => {
          // Extract text from comment content paragraphs
          let commentText = '';
          if (c.content) {
            for (const para of c.content) {
              for (const item of para.content || []) {
                if ('content' in item) {
                  for (const rc of (item as { content: Array<{ type: string; text?: string }> })
                    .content) {
                    if (rc.type === 'text' && rc.text) commentText += rc.text;
                  }
                }
              }
            }
          }
          return {
            id: c.id,
            author: c.author || 'Unknown',
            date: c.date,
            text: commentText,
            anchorText: anchorMap.get(c.id) || '',
            parentId: c.parentId,
            done: c.done,
          };
        });
      },
      getCommentPositions: () => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return [];

        // Step 1: Find PM position ranges for each comment mark
        const commentRanges = new Map<number, { from: number; to: number }>();
        view.state.doc.descendants((node, pos) => {
          for (const mark of node.marks) {
            if (mark.type.name === 'comment') {
              const cid = mark.attrs.commentId as number;
              const existing = commentRanges.get(cid);
              if (!existing) {
                commentRanges.set(cid, { from: pos, to: pos + node.nodeSize });
              } else {
                existing.to = Math.max(existing.to, pos + node.nodeSize);
              }
            }
          }
        });

        if (commentRanges.size === 0) return [];

        // Step 2: Find visible DOM elements matching those PM positions
        // Query all rendered spans with PM position data attributes
        const pagesContainer = pagedEditorRef.current?.getPagesContainer?.();
        if (!pagesContainer) return [];
        const containerRect = pagesContainer.getBoundingClientRect();
        const scrollTop = pagesContainer.scrollTop || 0;

        const results: Array<{ commentId: number; top: number; height: number }> = [];

        for (const [commentId, range] of commentRanges) {
          // Find the first span whose PM range overlaps the comment range
          const spans = pagesContainer.querySelectorAll('span[data-pm-start][data-pm-end]');
          let bestEl: HTMLElement | null = null;
          for (const span of Array.from(spans)) {
            const pmStart = Number((span as HTMLElement).dataset.pmStart);
            const pmEnd = Number((span as HTMLElement).dataset.pmEnd);
            if (pmStart < range.to && pmEnd > range.from) {
              bestEl = span as HTMLElement;
              break; // Use first overlapping span
            }
          }
          if (bestEl) {
            const rect = bestEl.getBoundingClientRect();
            results.push({
              commentId,
              top: rect.top - containerRect.top + scrollTop,
              height: rect.height,
            });
          }
        }

        return results;
      },
      insertContextTag: (tagKey: string, label?: string, removeIfEmpty?: boolean) => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return;
        const schema = view.state.schema;
        const nodeType = schema.nodes.contextTag;
        if (!nodeType) return;
        // Inherit formatting marks from the cursor position so the tag
        // matches the surrounding text style (font, size, bold, etc.)
        const { $from, empty } = view.state.selection;
        const cursorMarks = view.state.storedMarks || (empty ? $from.marks() : []);
        // Filter to only marks the contextTag node accepts
        const allowedMarks = cursorMarks.filter((m: import('prosemirror-model').Mark) =>
          nodeType.allowsMarkType(m.type)
        );
        const node = nodeType.create(
          {
            tagKey,
            label: label || contextTagDisplayValue(contextTags?.[tagKey]) || '',
            removeIfEmpty: removeIfEmpty ?? false,
            metaId: generateMetaId(),
          },
          null,
          allowedMarks
        );
        const tr = view.state.tr.replaceSelectionWith(node).scrollIntoView();
        tr.setMeta('allowLockedEdit', true);
        view.dispatch(tr);
        pagedEditorRef.current?.focus();
      },
      insertCrossRef: (
        refType: 'heading' | 'figure',
        refTarget: string,
        displayText: string,
        bookmarkName?: string
      ) => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return;
        const schema = view.state.schema;
        const nodeType = schema.nodes.crossRef;
        if (!nodeType) return;

        let bmName = bookmarkName || '';

        // If no bookmark provided, find the target paragraph and ensure it has one
        if (!bmName) {
          // Scan document for matching paragraph
          let targetPos: number | null = null;
          view.state.doc.descendants((node, pos) => {
            if (targetPos !== null) return false;
            if (node.type.name !== 'paragraph') return true;
            if (node.textContent === refTarget || node.textContent.startsWith(refTarget)) {
              targetPos = pos;
              return false;
            }
            return true;
          });

          if (targetPos !== null) {
            const targetNode = view.state.doc.nodeAt(targetPos);
            if (targetNode) {
              const existing = (
                targetNode.attrs.bookmarks as Array<{ id: number; name: string }>
              )?.find((b) => b.name.startsWith('_FP_Ref_'));
              if (existing) {
                bmName = existing.name;
              } else {
                // Generate a new bookmark
                const uuid = Math.random().toString(36).substring(2, 10);
                bmName = `_FP_Ref_${uuid}`;
                const existingBookmarks =
                  (targetNode.attrs.bookmarks as Array<{ id: number; name: string }>) || [];
                const tr = view.state.tr.setNodeMarkup(targetPos, undefined, {
                  ...targetNode.attrs,
                  bookmarks: [
                    ...existingBookmarks,
                    { id: Math.floor(Math.random() * 2147483647), name: bmName },
                  ],
                });
                view.dispatch(tr);
              }
            }
          }
        }

        const node = nodeType.create({ refType, refTarget, displayText, bookmarkName: bmName });
        // Use a fresh transaction since we may have dispatched above
        const currentView = pagedEditorRef.current?.getView();
        if (!currentView) return;
        const tr2 = currentView.state.tr.replaceSelectionWith(node).scrollIntoView();
        currentView.dispatch(tr2);
        pagedEditorRef.current?.focus();
      },
      insertImage: (dataUrl: string, alt: string, width: number, height: number) => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return;
        const { schema } = view.state;
        if (!schema.nodes.image) return;

        // Constrain to content area width
        const maxWidth = 612;
        let w = width;
        let h = height;
        if (w > maxWidth) {
          const scale = maxWidth / w;
          w = maxWidth;
          h = Math.round(h * scale);
        }

        const rId = `rId_img_${Date.now()}`;
        const imageNode = schema.nodes.image.create({
          src: dataUrl,
          alt,
          width: w,
          height: h,
          rId,
          wrapType: 'inline',
          displayMode: 'inline',
        });

        const { from } = view.state.selection;
        const tr = view.state.tr.insert(from, imageNode);
        tr.setMeta('allowLockedEdit', true);

        // Centre the paragraph containing the inserted image
        const $pos = tr.doc.resolve(from);
        for (let d = $pos.depth; d >= 0; d--) {
          if ($pos.node(d).type.name === 'paragraph') {
            const paragraphPos = $pos.before(d);
            const paragraphNode = $pos.node(d);
            tr.setNodeMarkup(paragraphPos, undefined, {
              ...paragraphNode.attrs,
              justification: 'center',
            });
            break;
          }
        }

        view.dispatch(tr.scrollIntoView());
        pagedEditorRef.current?.focus();
      },
      getReferenceable: () => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return [];
        const items: Array<{
          type: 'heading' | 'figure';
          text: string;
          number: string;
          bookmarkName: string | null;
          pmPos: number;
        }> = [];
        const headingCounters = [0, 0, 0, 0, 0, 0, 0, 0, 0];
        let figureCount = 0;

        // Build numbering map and style resolver for lvlText formatting
        const currentDoc = historyStateRef.current;
        const numMap = currentDoc?.package.numbering
          ? createNumberingMapFromDefs(currentDoc.package.numbering)
          : null;
        const styleRes = currentDoc?.package.styles
          ? createStyleResolver(currentDoc.package.styles)
          : null;

        view.state.doc.descendants((node, pos) => {
          if (node.type.name !== 'paragraph') return true;
          const styleId = node.attrs.styleId as string | null;
          if (!styleId) return true;

          // Find existing _FP_Ref_ bookmark on paragraph
          const bookmarks = (node.attrs.bookmarks as Array<{ id: number; name: string }>) || [];
          const fpRefBookmark = bookmarks.find((b) => b.name.startsWith('_FP_Ref_'));

          // Heading numbering
          const headingMatch = styleId.match(/^Heading(\d)$/);
          if (headingMatch) {
            const level = parseInt(headingMatch[1]) - 1;
            headingCounters[level]++;
            for (let i = level + 1; i < headingCounters.length; i++) headingCounters[i] = 0;

            // Use lvlText format from numbering map when available
            let number: string;
            const numPr =
              node.attrs.numPr ??
              styleRes?.resolveParagraphStyle(styleId)?.paragraphFormatting?.numPr;
            if (numMap && numPr?.numId) {
              number = formatNumberedMarker(headingCounters, level, numMap, numPr.numId);
            } else {
              // Fallback: simple dot-join
              const parts = headingCounters.slice(0, level + 1).filter((v) => v > 0);
              number = parts.join('.');
            }
            items.push({
              type: 'heading',
              text: node.textContent,
              number,
              bookmarkName: fpRefBookmark?.name ?? null,
              pmPos: pos,
            });
          }
          // Caption numbering
          if (styleId === 'Caption') {
            figureCount++;
            items.push({
              type: 'figure',
              text: node.textContent,
              number: `Figure ${figureCount}`,
              bookmarkName: fpRefBookmark?.name ?? null,
              pmPos: pos,
            });
          }
          return true;
        });
        return items;
      },
      refreshNumbering: (tocStyleOverride?: string) => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return;
        let tr = refreshAllReferences(view.state, {
          getNumberingMap: () => crossRefNumMapRef.current,
          getStyleResolver: () => crossRefStyleResolverRef.current,
          getLayout: () => pagedEditorRef.current?.getLayout() ?? null,
          tocStyleOverride,
        });

        // Also refresh context tag labels (in case they've gotten stale)
        const tags = contextTags ?? {};
        const ctNodeType = view.state.schema.nodes.contextTag;
        if (ctNodeType) {
          if (!tr) tr = view.state.tr;
          const doc = tr.doc;
          doc.descendants((node, pos) => {
            if (node.type !== ctNodeType) return true;
            const tagKey = node.attrs.tagKey as string;
            const currentLabel = node.attrs.label as string;
            const newLabel = tags[tagKey] ?? '';
            if (currentLabel !== newLabel) {
              // Use pos directly — it's from tr.doc.descendants so it's already
              // in the correct coordinate space. tr.mapping.map(pos) would double-map.
              tr = tr!.setNodeMarkup(pos, undefined, {
                ...node.attrs,
                label: newLabel,
              });
            }
            return false;
          });
        }

        if (tr && tr.docChanged) {
          tr.setMeta('allowLockedEdit', true);
          view.dispatch(tr);
        }
      },
      getAvailableStyles: () => {
        const resolver = crossRefStyleResolverRef.current;
        if (!resolver) return [];
        return resolver.getParagraphStyles().map((s) => ({
          styleId: s.styleId,
          name: s.name ?? s.styleId,
        }));
      },
      renderToBuffer: async (options) => {
        const view = pagedEditorRef.current?.getView();
        const baseDoc = historyStateRef.current;
        if (!view || !baseDoc) return null;

        const tags = options?.contextTags ?? contextTags ?? {};
        const mode = options?.unknownTagMode ?? 'keep';
        const { schema } = view.state;

        // Collect context tag metadata from current PM state (before any replacement)
        let ctMeta: Record<string, ContextTagMeta> | undefined;
        if (schema.nodes.contextTag) {
          const collected = collectContextTagMetadata(view.state.doc);
          ctMeta = Object.keys(collected).length > 0 ? collected : undefined;
        }

        // 'raw' mode: save as-is without any tag replacement (keeps {context.xxx} tags)
        if (mode === 'raw' || !schema.nodes.contextTag) {
          if (agentRef.current && ctMeta) {
            agentRef.current.getDocument().contextTagMetadata = ctMeta;
          }
          // Apply pending comment mutations to the agent document
          if (agentRef.current) {
            applyCommentMutations(agentRef.current.getDocument());
          }
          return agentRef.current?.toBuffer() ?? null;
        }

        // Convert unmodified PM doc → Document model (contextTag atoms become {{ tagKey }} text)
        const renderedDocument = fromProseDoc(view.state.doc, baseDoc);

        // Replace {{ tagKey }} patterns with rendered values + bookmark markers.
        // Bookmarks (_FP_ctx_{metaId}) enable tag restoration on re-upload.
        if (ctMeta && Object.keys(ctMeta).length > 0) {
          renderDocumentWithBookmarks(renderedDocument, { tags, ctMeta, mode });
        }

        // Pass context tag replacements for header/footer XML-level replacement during repack
        // (headers/footers use original XML — replaced at text level, not Document model)
        if (Object.keys(tags).length > 0) {
          renderedDocument.contextTagReplacements = { tags, mode };
        }

        // Propagate commentsModified flag so rezip regenerates comments.xml
        if (baseDoc.commentsModified) {
          renderedDocument.commentsModified = true;
        }

        // Serialize to DOCX buffer
        const tempAgent = new DocumentAgent(renderedDocument);
        return tempAgent.toBuffer();
      },
      lockParagraphs: (from: number, to: number) => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return;
        let tr = view.state.tr;
        tr.setMeta('allowLockedEdit', true);
        const seen = new Set<number>();
        view.state.doc.nodesBetween(from, to, (node, pos) => {
          if (node.type.name === 'paragraph' && !seen.has(pos)) {
            seen.add(pos);
            tr = tr.setNodeMarkup(pos, undefined, { ...node.attrs, locked: true });
          }
        });
        view.dispatch(tr);
      },
      unlockParagraphs: (from: number, to: number) => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return;
        let tr = view.state.tr;
        tr.setMeta('allowLockedEdit', true);
        const seen = new Set<number>();
        view.state.doc.nodesBetween(from, to, (node, pos) => {
          if (node.type.name === 'paragraph' && !seen.has(pos)) {
            seen.add(pos);
            tr = tr.setNodeMarkup(pos, undefined, { ...node.attrs, locked: false });
          }
        });
        view.dispatch(tr);
      },
      lockAll: () => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return;
        let tr = view.state.tr;
        tr.setMeta('allowLockedEdit', true);
        view.state.doc.descendants((node, pos) => {
          if (node.type.name === 'paragraph') {
            tr = tr.setNodeMarkup(pos, undefined, { ...node.attrs, locked: true });
          }
          return true;
        });
        view.dispatch(tr);
      },
      unlockAll: () => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return;
        let tr = view.state.tr;
        tr.setMeta('allowLockedEdit', true);
        view.state.doc.descendants((node, pos) => {
          if (node.type.name === 'paragraph') {
            tr = tr.setNodeMarkup(pos, undefined, { ...node.attrs, locked: false });
          }
          return true;
        });
        view.dispatch(tr);
      },
      lockHeadersFooters: () => {
        setHfLocked(history, true);
      },
      unlockHeadersFooters: () => {
        setHfLocked(history, false);
      },
      updateContextTagAttrs: (pmPos: number, attrs: Record<string, unknown>) => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return;
        const { state } = view;
        const $pos = state.doc.resolve(pmPos);
        const parent = $pos.parent;
        let tagOffset = -1;
        let tagNode: import('prosemirror-model').Node | null = null;
        parent.forEach((node, offset) => {
          const nodeStart = $pos.start() + offset;
          if (
            node.type.name === 'contextTag' &&
            pmPos >= nodeStart &&
            pmPos < nodeStart + node.nodeSize
          ) {
            tagOffset = nodeStart;
            tagNode = node;
          }
        });
        if (tagNode && tagOffset >= 0) {
          const tr = state.tr.setNodeMarkup(tagOffset, undefined, {
            ...(tagNode as any).attrs,
            ...attrs,
          });
          tr.setMeta('allowLockedEdit', true);
          view.dispatch(tr);
        }
      },
      importStyles: (styles) => {
        const doc = history.state;
        if (!doc?.package?.styles) return;
        const existingStyles = doc.package.styles.styles;

        // Build a lookup of all available source styles for dependency resolution
        const sourceById = new Map<string, any>();
        for (const s of styles) {
          sourceById.set(s.styleId, s);
        }

        // Collect the full dependency tree: basedOn, link, next chains
        // TODO: Also import numbering definitions (w:numId) referenced by styles.
        // Numbering lives in word/numbering.xml and needs separate ID remapping logic.
        const toImport = new Set<string>();
        function collectDeps(styleId: string) {
          if (toImport.has(styleId)) return;
          const s = sourceById.get(styleId);
          if (!s) return;
          toImport.add(styleId);
          if (s.basedOn) collectDeps(s.basedOn);
          if (s.link) collectDeps(s.link);
          if (s.next) collectDeps(s.next);
        }
        for (const s of styles) {
          collectDeps(s.styleId);
        }

        // Import each style, preserving _originalXml when available
        for (const styleId of toImport) {
          const s = sourceById.get(styleId);
          if (!s) continue;

          const idx = existingStyles.findIndex((es) => es.styleId === s.styleId);

          let newStyle: any;
          if (s._originalXml) {
            // Parse the _originalXml to extract rPr/pPr for live rendering,
            // while preserving the raw XML for lossless serialization.
            const parsed = parseImportedStyleXml(s._originalXml);
            newStyle = {
              ...parsed,
              styleId: s.styleId,
              type: (s.type || 'paragraph') as 'paragraph',
              name: s.name || s.styleId,
              basedOn: s.basedOn || undefined,
              next: s.next || undefined,
              link: s.link || undefined,
              qFormat: true,
              _originalXml: s._originalXml,
              _dirty: false, // use _originalXml verbatim during serialization
            };
          } else {
            // Fallback: construct from simplified data (old API without _originalXml)
            newStyle = {
              styleId: s.styleId,
              type: (s.type || 'paragraph') as 'paragraph',
              name: s.name || s.styleId,
              basedOn: s.basedOn || undefined,
              next: s.next || undefined,
              qFormat: true,
              rPr: {
                fontFamily: s.fontFamily ? { ascii: s.fontFamily, hAnsi: s.fontFamily } : undefined,
                fontSize: s.fontSize || undefined,
                bold: s.bold || undefined,
                italic: s.italic || undefined,
                color: s.color ? { rgb: s.color } : undefined,
              },
              _dirty: true,
            };
          }

          if (idx >= 0) {
            existingStyles[idx] = newStyle;
          } else {
            existingStyles.push(newStyle);
          }
        }

        // ── Import numbering definitions referenced by imported styles ──
        // Store imported numbering entries on the Document for injection during repack.
        // Remap IDs to avoid collisions with existing numbering in the target.
        const importedNumbering: Array<{
          abstractNumXml: string;
          numXml: string;
          oldNumId: number;
          newNumId: number;
        }> = [];
        const existingNums = doc.package.numbering?.nums || [];
        const existingAbstracts = doc.package.numbering?.abstractNums || [];
        let nextNumId =
          existingNums.length > 0 ? Math.max(...existingNums.map((n: any) => n.numId)) + 1 : 100;
        let nextAbsId =
          existingAbstracts.length > 0
            ? Math.max(...existingAbstracts.map((a: any) => a.abstractNumId)) + 1
            : 100;
        const absIdRemap = new Map<number, number>();
        const numIdRemap = new Map<number, number>();

        for (const styleId of toImport) {
          const s = sourceById.get(styleId);
          if (!s?._abstractNumXml || !s?._numXml) continue;
          const oldAbsId = s._abstractNumId as number;
          const oldNumId = s._numId as number;
          if (numIdRemap.has(oldNumId)) continue;

          let newAbsId: number;
          if (absIdRemap.has(oldAbsId)) {
            newAbsId = absIdRemap.get(oldAbsId)!;
          } else {
            newAbsId = nextAbsId++;
            absIdRemap.set(oldAbsId, newAbsId);
          }

          const newNumId = nextNumId++;
          numIdRemap.set(oldNumId, newNumId);

          const absXml = (s._abstractNumXml as string).replace(
            `w:abstractNumId="${oldAbsId}"`,
            `w:abstractNumId="${newAbsId}"`
          );
          const nXml = (s._numXml as string)
            .replace(`w:numId="${oldNumId}"`, `w:numId="${newNumId}"`)
            .replace(`w:abstractNumId="${oldAbsId}"`, `w:abstractNumId="${newAbsId}"`);

          importedNumbering.push({ abstractNumXml: absXml, numXml: nXml, oldNumId, newNumId });
        }

        // Update numId references in imported styles' _originalXml
        if (numIdRemap.size > 0) {
          for (const styleId of toImport) {
            const s = sourceById.get(styleId);
            if (!s?._numId || !numIdRemap.has(s._numId)) continue;
            const newNumId = numIdRemap.get(s._numId)!;
            const idx = existingStyles.findIndex((es) => es.styleId === styleId);
            if (idx >= 0 && existingStyles[idx]._originalXml) {
              existingStyles[idx]._originalXml = existingStyles[idx]._originalXml.replace(
                `<w:numId w:val="${s._numId}"/>`,
                `<w:numId w:val="${newNumId}"/>`
              );
            }
          }
          // Store on document for rezip to inject into numbering.xml
          (doc as any).importedNumberingEntries = [
            ...((doc as any).importedNumberingEntries || []),
            ...importedNumbering,
          ];
        }

        doc.package.stylesDirty = true;
        doc.package.styles = { ...doc.package.styles, styles: [...existingStyles] };
        setDocumentStyles(doc.package.styles.styles);
        // Force document state update so Toolbar + StylePicker re-read styles
        history.push({ ...doc, package: { ...doc.package } });
      },
      getDocumentMeta: () => {
        return agentRef.current?.getDocument()?.fpDocumentMeta;
      },
      setDocumentMeta: (meta: FPDocumentMeta) => {
        const doc = agentRef.current?.getDocument();
        if (doc) {
          doc.fpDocumentMeta = { ...doc.fpDocumentMeta, ...meta };
        }
      },
    }),
    [
      history.state,
      state.zoom,
      state.renderMode,
      state.currentPage,
      state.totalPages,
      handleSave,
      handleDirectPrint,
      contextTags,
    ]
  );

  // Auto-update context tag labels when contextTags prop changes or document loads.
  // The view may not be ready immediately when isLoading flips to false (the
  // HiddenProseMirror EditorView is created in a child effect), so we poll briefly.
  // Also reports discovered tagKeys back to the parent so it can show them in the panel.
  const onContextTagsDiscoveredRef = useRef(onContextTagsDiscovered);
  onContextTagsDiscoveredRef.current = onContextTagsDiscovered;

  useEffect(() => {
    if (!enableContextTags) return; // Skip all tag processing when disabled
    let cancelled = false;
    let attempts = 0;

    const applyLabels = () => {
      if (cancelled) return;
      const view = pagedEditorRef.current?.getView();
      if (!view) {
        if (attempts++ < 20) setTimeout(applyLabels, 50);
        return;
      }

      const { state: editorState } = view;
      const nodeType = editorState.schema.nodes.contextTag;
      if (!nodeType) return;

      const tags = contextTags ?? {};
      let tr = editorState.tr;
      let changed = false;
      const discoveredKeys = new Set<string>();

      editorState.doc.descendants((node, pos) => {
        if (node.type !== nodeType) return true;
        const tagKey = node.attrs.tagKey as string;
        discoveredKeys.add(tagKey);
        const currentLabel = node.attrs.label as string;
        const currentImageUrl = (node.attrs.imageUrl as string) || '';
        const rawVal = tags[tagKey];
        const newLabel = contextTagDisplayValue(rawVal);
        // Compute imageUrl for image-type values
        const newImageUrl =
          rawVal && typeof rawVal === 'object' && (rawVal as Record<string, unknown>).case_file_id
            ? `/api/dms/case-files/${(rawVal as Record<string, unknown>).case_file_id}/preview/`
            : '';
        if (currentLabel !== newLabel || currentImageUrl !== newImageUrl) {
          tr = tr.setNodeMarkup(pos, undefined, {
            ...node.attrs,
            label: newLabel,
            imageUrl: newImageUrl,
          });
          changed = true;
        }
        return false; // atom node, no children
      });

      if (changed) {
        tr.setMeta('allowLockedEdit', true);
        view.dispatch(tr);
      }

      // Also discover context tags in headers/footers (Document model level)
      // Join text across runs before matching to catch tags split across runs.
      const pkg = historyStateRef.current?.package;
      const _discoverInRuns = (runs: readonly Run[]) => {
        const joined = runs
          .flatMap((r) => r.content)
          .filter((rc) => rc.type === 'text' && (rc as { text: string }).text)
          .map((rc) => (rc as { text: string }).text)
          .join('');
        if (!joined) return;
        for (const m of joined.matchAll(HF_CONTEXT_TAG_RE)) {
          const tagKey = m[1] || m[2];
          if (tagKey) discoveredKeys.add(tagKey);
        }
      };
      const scanHfMap = (map: Map<string, HeaderFooter> | undefined) => {
        if (!map) return;
        for (const [, hf] of map) {
          for (const block of hf.content) {
            if (block.type !== 'paragraph') continue;
            // Discover in direct runs (joined across the paragraph)
            const directRuns = block.content.filter((c): c is Run => c.type === 'run');
            if (directRuns.length > 0) _discoverInRuns(directRuns);
            // Discover in InlineSdt content controls
            for (const item of block.content) {
              if (item.type === 'inlineSdt') {
                const sdtRuns = item.content.filter((c): c is Run => c.type === 'run');
                if (sdtRuns.length > 0) _discoverInRuns(sdtRuns);
              }
            }
          }
        }
      };
      scanHfMap(pkg?.headers);
      scanHfMap(pkg?.footers);

      // Report discovered tagKeys to parent
      if (discoveredKeys.size > 0) {
        onContextTagsDiscoveredRef.current?.([...discoveredKeys]);
      }
    };

    applyLabels();
    return () => {
      cancelled = true;
    };
  }, [contextTags, state.isLoading]);

  // Create numbering map from document numbering definitions
  const numberingMap = useMemo(() => {
    const numbering = history.state?.package.numbering;
    if (!numbering) return null;
    return createNumberingMapFromDefs(numbering);
  }, [history.state?.package.numbering]);

  // Get header and footer content from document.
  // When titlePg is set, page 1 gets 'first' header/footer and pages 2+ get 'default'.
  const { headerContent, footerContent, firstPageHeaderContent, firstPageFooterContent } = useMemo<{
    headerContent: HeaderFooter | null;
    footerContent: HeaderFooter | null;
    firstPageHeaderContent: HeaderFooter | null;
    firstPageFooterContent: HeaderFooter | null;
  }>(() => {
    const empty = {
      headerContent: null,
      footerContent: null,
      firstPageHeaderContent: null,
      firstPageFooterContent: null,
    };
    if (!history.state?.package) return empty;

    const pkg = history.state.package;
    const sectionProps = pkg.document?.finalSectionProperties;
    const headers = pkg.headers;
    const footers = pkg.footers;

    // Resolve header/footer references.
    // In multi-section documents, the final sectPr may not have headerReferences —
    // OOXML sections inherit header/footer refs from earlier sections.
    let headerRefs = sectionProps?.headerReferences;
    let footerRefs = sectionProps?.footerReferences;
    let titlePg = sectionProps?.titlePg;

    if ((!headerRefs || !footerRefs) && pkg.document?.sections) {
      for (let i = pkg.document.sections.length - 1; i >= 0; i--) {
        const sp = pkg.document.sections[i].properties;
        if (!headerRefs && sp?.headerReferences) {
          headerRefs = sp.headerReferences;
          if (titlePg === undefined && sp.titlePg !== undefined) {
            titlePg = sp.titlePg;
          }
        }
        if (!footerRefs && sp?.footerReferences) {
          footerRefs = sp.footerReferences;
        }
      }
    }

    const hasTitlePg = !!titlePg;

    // Resolve headers: when titlePg, page 1 = 'first', pages 2+ = 'default'
    let defaultHeader: HeaderFooter | null = null;
    let firstPageHeader: HeaderFooter | null = null;
    if (headers && headerRefs) {
      const defaultRef = headerRefs.find((r) => r.type === 'default');
      if (defaultRef?.rId) defaultHeader = headers.get(defaultRef.rId) ?? null;

      if (hasTitlePg) {
        const firstRef = headerRefs.find((r) => r.type === 'first');
        if (firstRef?.rId) firstPageHeader = headers.get(firstRef.rId) ?? null;
      }
    }

    // Resolve footers: same logic
    let defaultFooter: HeaderFooter | null = null;
    let firstPageFooter: HeaderFooter | null = null;
    if (footers && footerRefs) {
      const defaultRef = footerRefs.find((r) => r.type === 'default');
      if (defaultRef?.rId) defaultFooter = footers.get(defaultRef.rId) ?? null;

      if (hasTitlePg) {
        const firstRef = footerRefs.find((r) => r.type === 'first');
        if (firstRef?.rId) firstPageFooter = footers.get(firstRef.rId) ?? null;
      }
    }

    // When no titlePg differentiation, headerContent is used for all pages.
    // When titlePg is set, headerContent = default (pages 2+), firstPage* = page 1.
    //
    // Substitute context tags for visual display. The original Document model is
    // NOT mutated — these are new objects for PagedEditor's layout pipeline.
    // replaceContextTagsInHf ALWAYS returns a new reference (even if no tags matched)
    // to ensure React dependency tracking works correctly.
    const tags = contextTags ?? {};
    const hasCtx = Object.keys(tags).length > 0;
    // In raw mode, skip tag substitution — show raw {tag.path} patterns in H/F
    const shouldSubstitute = hasCtx && state.renderMode !== 'raw';
    return {
      headerContent:
        shouldSubstitute && defaultHeader
          ? replaceContextTagsInHf(defaultHeader, tags, 'keep')
          : defaultHeader,
      footerContent:
        shouldSubstitute && defaultFooter
          ? replaceContextTagsInHf(defaultFooter, tags, 'keep')
          : defaultFooter,
      firstPageHeaderContent:
        shouldSubstitute && firstPageHeader
          ? replaceContextTagsInHf(firstPageHeader, tags, 'keep')
          : firstPageHeader,
      firstPageFooterContent:
        shouldSubstitute && firstPageFooter
          ? replaceContextTagsInHf(firstPageFooter, tags, 'keep')
          : firstPageFooter,
    };
  }, [history.state, contextTags, state.renderMode]);

  // Handle header/footer double-click — open editing overlay
  // If no header/footer exists, create an empty one so the user can add content
  const handleHeaderFooterDoubleClick = useCallback(
    (position: 'header' | 'footer') => {
      const hf = position === 'header' ? headerContent : footerContent;
      if (hf) {
        setHfEditPosition(position);
        return;
      }

      // Create empty header/footer for docs that don't have one yet
      if (!history.state?.package) return;
      const pkg = history.state.package;
      const sectionProps = pkg.document?.finalSectionProperties;
      if (!sectionProps) return;

      const rId = `rId_new_${position}`;
      const emptyHf: HeaderFooter = {
        type: position === 'header' ? 'header' : 'footer',
        hdrFtrType: 'default',
        content: [{ type: 'paragraph', content: [] }],
      };

      const mapKey = position === 'header' ? 'headers' : 'footers';
      const newMap = new Map(pkg[mapKey] ?? []);
      newMap.set(rId, emptyHf);

      const refKey = position === 'header' ? 'headerReferences' : 'footerReferences';
      const existingRefs = sectionProps[refKey] ?? [];
      const newRef = { type: 'default' as const, rId };

      const newDoc: Document = {
        ...history.state,
        package: {
          ...pkg,
          [mapKey]: newMap,
          document: pkg.document
            ? {
                ...pkg.document,
                finalSectionProperties: {
                  ...sectionProps,
                  [refKey]: [...existingRefs, newRef],
                  rawXml: undefined, // Invalidate raw XML since properties changed
                },
              }
            : pkg.document,
        },
      };
      history.push(newDoc);
      setHfEditPosition(position);
    },
    [headerContent, footerContent, history]
  );

  // Handle header/footer save — update document package with edited content
  const handleHeaderFooterSave = useCallback(
    (content: (import('../types/document').Paragraph | import('../types/document').Table)[]) => {
      if (!hfEditPosition || !history.state?.package) {
        setHfEditPosition(null);
        return;
      }

      const pkg = history.state.package;
      const sectionProps = pkg.document?.finalSectionProperties;
      const refs =
        hfEditPosition === 'header'
          ? sectionProps?.headerReferences
          : sectionProps?.footerReferences;
      const defaultRef = refs?.find((r) => r.type === 'default');
      const mapKey = hfEditPosition === 'header' ? 'headers' : 'footers';
      const map = pkg[mapKey];

      if (defaultRef?.rId && map) {
        const existing = map.get(defaultRef.rId);
        const updated: HeaderFooter = {
          type: hfEditPosition,
          hdrFtrType: 'default',
          ...existing,
          content,
        };
        const newMap = new Map(map);
        newMap.set(defaultRef.rId, updated);

        const newDoc: Document = {
          ...history.state,
          package: {
            ...pkg,
            [mapKey]: newMap,
          },
        };
        history.push(newDoc);
      }

      setHfEditPosition(null);
    },
    [hfEditPosition, history]
  );

  // Handle body click while in HF editing mode — save + close
  const handleBodyClick = useCallback(() => {
    if (!hfEditPosition) return;
    // Save if dirty, then close
    const view = hfEditorRef.current?.getView();
    if (view) {
      const blocks = proseDocToBlocks(view.state.doc);
      handleHeaderFooterSave(blocks);
    } else {
      setHfEditPosition(null);
    }
  }, [hfEditPosition, handleHeaderFooterSave]);

  // Handle removing the header/footer entirely
  const handleRemoveHeaderFooter = useCallback(() => {
    if (!hfEditPosition || !history.state?.package) {
      setHfEditPosition(null);
      return;
    }

    const pkg = history.state.package;
    const sectionProps = pkg.document?.finalSectionProperties;
    const refKey = hfEditPosition === 'header' ? 'headerReferences' : 'footerReferences';
    const mapKey = hfEditPosition === 'header' ? 'headers' : 'footers';
    const refs = sectionProps?.[refKey];
    const defaultRef = refs?.find((r) => r.type === 'default');

    if (defaultRef?.rId) {
      const newMap = new Map(pkg[mapKey] ?? []);
      newMap.delete(defaultRef.rId);

      const newRefs = (refs ?? []).filter((r) => r.rId !== defaultRef.rId);

      const newDoc: Document = {
        ...history.state,
        package: {
          ...pkg,
          [mapKey]: newMap,
          document: pkg.document
            ? {
                ...pkg.document,
                finalSectionProperties: {
                  ...sectionProps,
                  [refKey]: newRefs,
                  rawXml: undefined, // Invalidate raw XML since properties changed
                },
              }
            : pkg.document,
        },
      };
      history.push(newDoc);
    }

    setHfEditPosition(null);
  }, [hfEditPosition, history]);

  // Get the DOM element for the header/footer area on the first page
  const getHfTargetElement = useCallback((pos: 'header' | 'footer'): HTMLElement | null => {
    const pagesContainer = containerRef.current?.querySelector('.paged-editor__pages');
    if (!pagesContainer) return null;
    const className = pos === 'header' ? '.layout-page-header' : '.layout-page-footer';
    return pagesContainer.querySelector(className);
  }, []);

  // Container styles - using overflow: auto so sticky toolbar works
  const containerStyle: CSSProperties = {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    width: '100%',
    backgroundColor: 'var(--doc-bg-subtle)',
    ...style,
  };

  const mainContentStyle: CSSProperties = {
    display: 'flex',
    flex: 1,
    minHeight: 0, // Allow flex item to shrink below content size
    minWidth: 0, // Allow flex item to shrink below content width on narrow viewports
    flexDirection: 'row',
  };

  const editorContainerStyle: CSSProperties = {
    flex: 1,
    minHeight: 0,
    minWidth: 0, // Allow flex item to shrink below content width on narrow viewports
    overflow: 'auto', // This is the scroll container - sticky toolbar will stick to this
    position: 'relative',
  };

  // Render loading state
  if (state.isLoading) {
    return (
      <div
        className={`ep-root docx-editor docx-editor-loading ${className}`}
        style={containerStyle}
        data-testid="docx-editor"
      >
        {loadingIndicator || <DefaultLoadingIndicator />}
      </div>
    );
  }

  // Render error state
  if (state.parseError) {
    return (
      <div
        className={`ep-root docx-editor docx-editor-error ${className}`}
        style={containerStyle}
        data-testid="docx-editor"
      >
        <ParseError message={state.parseError} />
      </div>
    );
  }

  // Render placeholder when no document
  if (!history.state) {
    return (
      <div
        className={`ep-root docx-editor docx-editor-empty ${className}`}
        style={containerStyle}
        data-testid="docx-editor"
      >
        {placeholder || <DefaultPlaceholder />}
      </div>
    );
  }

  return (
    <ErrorProvider>
      <ErrorBoundary onError={handleEditorError}>
        <div
          ref={containerRef}
          className={`ep-root docx-editor ${className}`}
          style={containerStyle}
          data-testid="docx-editor"
        >
          {/* Main content area */}
          <div style={mainContentStyle}>
            {/* Wrapper for scroll container + outline overlay */}
            <div
              style={{
                position: 'relative',
                flex: 1,
                minHeight: 0,
                minWidth: 0,
                display: 'flex',
                flexDirection: 'column',
              }}
            >
              {/* Editor container - this is the scroll container */}
              <div style={editorContainerStyle}>
                {/* Toolbar - sticky at top of scroll container */}
                {/* Hide toolbar only when readOnly prop is explicitly set (not from viewing mode) */}
                {showToolbar && !readOnlyProp && (
                  <div
                    ref={toolbarRefCallback}
                    className="sticky top-0 z-50 flex flex-col gap-0 bg-white shadow-sm"
                  >
                    <Toolbar
                      currentFormatting={state.selectionFormatting}
                      onFormat={handleFormat}
                      onUndo={undoActiveEditor}
                      onRedo={redoActiveEditor}
                      canUndo={true}
                      canRedo={true}
                      disabled={readOnly}
                      documentStyles={history.state?.package.styles?.styles}
                      theme={history.state?.package.theme || theme}
                      showPrintButton={showPrintButton}
                      showLineSpacingPicker={showLineSpacingPickerProp}
                      showClearFormatting={showClearFormattingProp}
                      onPrint={handleDirectPrint}
                      showZoomControl={showZoomControl}
                      zoom={state.zoom}
                      onZoomChange={handleZoomChange}
                      onRefocusEditor={focusActiveEditor}
                      onInsertTable={handleInsertTable}
                      showTableInsert={true}
                      onInsertImage={handleInsertImageClick}
                      onInsertPageBreak={handleInsertPageBreak}
                      onInsertTOC={showInsertTOCProp ? handleInsertTOC : undefined}
                      imageContext={state.pmImageContext}
                      onImageWrapType={handleImageWrapType}
                      onImageTransform={handleImageTransform}
                      onOpenImageProperties={handleOpenImageProperties}
                      tableContext={state.pmTableContext}
                      onTableAction={handleTableAction}
                      restrictedMode={restrictedMode}
                      styleGalleryMode={styleGalleryMode}
                      allowedStyleIds={
                        restrictedMode || styleGalleryMode ? allowedStyleIds : undefined
                      }
                      numberingMap={numberingMap}
                      canModifyStyles={canModifyStyles}
                      onModifyStyle={(styleId) =>
                        setStyleEditorState({ open: true, mode: 'modify', styleId })
                      }
                      onCreateStyle={() => setStyleEditorState({ open: true, mode: 'create' })}
                      renderMode={state.renderMode}
                      onToggleRenderMode={() =>
                        setState((prev) => ({
                          ...prev,
                          renderMode: prev.renderMode === 'rendered' ? 'raw' : 'rendered',
                        }))
                      }
                    >
                      <EditingModeDropdown
                        mode={editingMode}
                        onModeChange={(mode) => setEditingMode(mode)}
                      />
                      {toolbarExtra}
                    </Toolbar>

                    {/* Horizontal Ruler - sticky with toolbar */}
                    {showRuler && (
                      <div className="flex justify-center px-5 py-1 overflow-x-auto flex-shrink-0 bg-doc-bg">
                        <HorizontalRuler
                          sectionProps={history.state?.package.document?.finalSectionProperties}
                          zoom={state.zoom}
                          unit={rulerUnit}
                          editable={!readOnly}
                          onLeftMarginChange={handleLeftMarginChange}
                          onRightMarginChange={handleRightMarginChange}
                          indentLeft={state.paragraphIndentLeft}
                          indentRight={state.paragraphIndentRight}
                          onIndentLeftChange={handleIndentLeftChange}
                          onIndentRightChange={handleIndentRightChange}
                          showFirstLineIndent={true}
                          firstLineIndent={state.paragraphFirstLineIndent}
                          hangingIndent={state.paragraphHangingIndent}
                          onFirstLineIndentChange={handleFirstLineIndentChange}
                          tabStops={state.paragraphTabs}
                          onTabStopRemove={handleTabStopRemove}
                        />
                      </div>
                    )}
                  </div>
                )}

                {/* Vertical Ruler - fixed on left edge (hidden when readOnly prop is set) */}
                {showRuler && !readOnlyProp && (
                  <div
                    style={{
                      position: 'absolute',
                      left: 0,
                      top: 0,
                      paddingTop: 20,
                      zIndex: 10,
                    }}
                  >
                    <VerticalRuler
                      sectionProps={history.state?.package.document?.finalSectionProperties}
                      zoom={state.zoom}
                      unit={rulerUnit}
                      editable={!readOnly}
                      onTopMarginChange={handleTopMarginChange}
                      onBottomMarginChange={handleBottomMarginChange}
                    />
                  </div>
                )}

                {/* Editor content wrapper */}
                <div style={{ display: 'flex', flex: 1, minHeight: 0, position: 'relative' }}>
                  {/* Editor content area */}
                  <div
                    ref={editorContentRef}
                    style={{ position: 'relative', flex: 1, minWidth: 0 }}
                    onMouseDown={(e) => {
                      // Focus editor when clicking on the background area (not the editor itself)
                      // Using mouseDown for immediate response before focus can be lost
                      if (e.target === e.currentTarget) {
                        e.preventDefault();
                        pagedEditorRef.current?.focus();
                      }
                    }}
                  >
                    <PagedEditor
                      ref={pagedEditorRef}
                      document={history.state}
                      styles={history.state?.package.styles}
                      theme={history.state?.package.theme || theme}
                      sectionProperties={history.state?.package.document?.finalSectionProperties}
                      headerContent={headerContent}
                      footerContent={footerContent}
                      firstPageHeaderContent={firstPageHeaderContent}
                      firstPageFooterContent={firstPageFooterContent}
                      numberingMap={numberingMap}
                      contextTags={contextTags}
                      renderMode={state.renderMode}
                      loopPreviewData={loopPreviewData}
                      onHeaderFooterDoubleClick={handleHeaderFooterDoubleClick}
                      hfEditMode={hfEditPosition}
                      onBodyClick={handleBodyClick}
                      zoom={state.zoom}
                      readOnly={readOnly}
                      extensionManager={extensionManager}
                      onDocumentChange={handleDocumentChange}
                      onSelectionChange={(_from, _to) => {
                        // Extract full selection state from PM and use the standard handler
                        const view = pagedEditorRef.current?.getView();
                        if (view) {
                          const selectionState = extractSelectionState(view.state);
                          handleSelectionChange(selectionState);

                          // Detect comment/tracked-change marks at cursor to fire
                          // onCursorMarkChange callback (for sidebar auto-expansion).
                          // Collect marks from all sources — inclusive:false marks aren't
                          // reported by $from.marks() at boundaries, and empty arrays are
                          // truthy so an OR chain would short-circuit.
                          if (onCursorMarkChange) {
                            const $from = view.state.selection.$from;
                            const marks = [
                              ...(view.state.storedMarks ?? []),
                              ...($from.nodeAfter?.marks ?? []),
                              ...($from.nodeBefore?.marks ?? []),
                              ...$from.marks(),
                            ];
                            let cursorMarkId: string | null = null;
                            for (const mark of marks) {
                              if (mark.type.name === 'comment' && mark.attrs.commentId != null) {
                                cursorMarkId = `comment-${mark.attrs.commentId}`;
                                break;
                              }
                              if (
                                (mark.type.name === 'insertion' || mark.type.name === 'deletion') &&
                                mark.attrs.revisionId != null
                              ) {
                                cursorMarkId = `tc-${mark.attrs.revisionId}-${mark.type.name}`;
                                break;
                              }
                            }
                            onCursorMarkChange(cursorMarkId);
                          }
                        } else {
                          handleSelectionChange(null);
                          onCursorMarkChange?.(null);
                        }
                      }}
                      externalPlugins={mergedPlugins}
                      onReady={(ref) => {
                        onEditorViewReady?.(ref.getView()!);
                      }}
                      onRenderedDomContextReady={onRenderedDomContextReady}
                      pluginOverlays={pluginOverlays}
                      onContextTagRightClick={onContextTagRightClick}
                      showCommentPanel={showCommentPanel}
                      onCommentAction={onCommentAction}
                      commentPanelKey={commentPanelKey}
                      additionalComments={addedCommentsRef.current}
                      deletedCommentIds={deletedCommentIdsRef.current}
                      onPageCountChange={(pageCount) => {
                        setState((prev) => {
                          if (prev.totalPages !== pageCount) {
                            onPageCountChangeProp?.(pageCount);
                            return { ...prev, totalPages: pageCount };
                          }
                          return prev;
                        });
                      }}
                    />

                    {/* Page navigation / indicator */}
                    {showPageNumbers &&
                      state.totalPages > 0 &&
                      (enablePageNavigation ? (
                        <PageNavigator
                          currentPage={state.currentPage}
                          totalPages={state.totalPages}
                          onNavigate={handlePageNavigate}
                          position={pageNumberPosition as PageNavigatorPosition}
                          variant={pageNumberVariant as PageNavigatorVariant}
                          floating
                        />
                      ) : (
                        <PageNumberIndicator
                          currentPage={state.currentPage}
                          totalPages={state.totalPages}
                          position={pageNumberPosition as PageIndicatorPosition}
                          variant={pageNumberVariant as PageIndicatorVariant}
                          floating
                        />
                      ))}

                    {/* Inline Header/Footer Editor — positioned over the target area */}
                    {hfEditPosition &&
                      (hfEditPosition === 'header' ? headerContent : footerContent) &&
                      (() => {
                        const targetEl = getHfTargetElement(hfEditPosition);
                        const parentEl = editorContentRef.current;
                        if (!targetEl || !parentEl) return null;
                        return (
                          <InlineHeaderFooterEditor
                            ref={hfEditorRef}
                            headerFooter={
                              (hfEditPosition === 'header'
                                ? headerContent
                                : footerContent) as HeaderFooter
                            }
                            position={hfEditPosition}
                            styles={history.state?.package.styles}
                            targetElement={targetEl}
                            parentElement={parentEl}
                            onSave={handleHeaderFooterSave}
                            onClose={() => setHfEditPosition(null)}
                            onSelectionChange={handleSelectionChange}
                            onRemove={handleRemoveHeaderFooter}
                          />
                        );
                      })()}
                  </div>
                </div>
                {/* end editor flex wrapper */}
              </div>
              {/* end scroll container */}

              {/* Document outline sidebar — absolutely positioned, doesn't scroll */}
              {showOutline && (
                <DocumentOutline
                  headings={outlineHeadings}
                  onHeadingClick={handleHeadingInfoClick}
                  onClose={() => setShowOutline(false)}
                  topOffset={toolbarHeight}
                />
              )}

              {/* Outline toggle button — absolutely positioned below toolbar */}
              {!showOutline && (
                <button
                  className="docx-outline-nav"
                  onClick={handleToggleOutline}
                  onMouseDown={(e) => e.stopPropagation()}
                  title="Show document outline"
                  style={{
                    position: 'absolute',
                    left: 48,
                    top: toolbarHeight + 12,
                    zIndex: 20,
                    background: 'transparent',
                    border: 'none',
                    borderRadius: '50%',
                    padding: 6,
                    cursor: 'pointer',
                    display: 'flex',
                    alignItems: 'center',
                  }}
                >
                  <MaterialSymbol
                    name="format_list_bulleted"
                    size={20}
                    style={{ color: '#444746' }}
                  />
                </button>
              )}
            </div>
            {/* end wrapper for scroll container + outline */}
          </div>

          {/* Lazy-loaded dialogs — only fetched when first opened */}
          <Suspense fallback={null}>
            {findReplace.state.isOpen && (
              <FindReplaceDialog
                isOpen={findReplace.state.isOpen}
                onClose={findReplace.close}
                onFind={handleFind}
                onFindNext={handleFindNext}
                onFindPrevious={handleFindPrevious}
                onReplace={handleReplace}
                onReplaceAll={handleReplaceAll}
                initialSearchText={findReplace.state.searchText}
                replaceMode={findReplace.state.replaceMode}
                currentResult={findResultRef.current}
              />
            )}
            {hyperlinkDialog.state.isOpen && (
              <HyperlinkDialog
                isOpen={hyperlinkDialog.state.isOpen}
                onClose={hyperlinkDialog.close}
                onSubmit={handleHyperlinkSubmit}
                onRemove={hyperlinkDialog.state.isEditing ? handleHyperlinkRemove : undefined}
                initialData={hyperlinkDialog.state.initialData}
                selectedText={hyperlinkDialog.state.selectedText}
                isEditing={hyperlinkDialog.state.isEditing}
              />
            )}
            {tablePropsOpen && (
              <TablePropertiesDialog
                isOpen={tablePropsOpen}
                onClose={() => setTablePropsOpen(false)}
                onApply={(props) => {
                  const view = getActiveEditorView();
                  if (view) {
                    setTableProperties(props)(view.state, view.dispatch);
                  }
                }}
                currentProps={
                  state.pmTableContext?.table?.attrs as Record<string, unknown> | undefined
                }
              />
            )}
            {imagePositionOpen && (
              <ImagePositionDialog
                isOpen={imagePositionOpen}
                onClose={() => setImagePositionOpen(false)}
                onApply={handleApplyImagePosition}
              />
            )}
            {imagePropsOpen && (
              <ImagePropertiesDialog
                isOpen={imagePropsOpen}
                onClose={() => setImagePropsOpen(false)}
                onApply={handleApplyImageProperties}
                currentData={
                  state.pmImageContext
                    ? {
                        alt: state.pmImageContext.alt ?? undefined,
                        borderWidth: state.pmImageContext.borderWidth ?? undefined,
                        borderColor: state.pmImageContext.borderColor ?? undefined,
                        borderStyle: state.pmImageContext.borderStyle ?? undefined,
                      }
                    : undefined
                }
              />
            )}
            {footnotePropsOpen && (
              <FootnotePropertiesDialog
                isOpen={footnotePropsOpen}
                onClose={() => setFootnotePropsOpen(false)}
                onApply={handleApplyFootnoteProperties}
                footnotePr={history.state?.package.document?.finalSectionProperties?.footnotePr}
                endnotePr={history.state?.package.document?.finalSectionProperties?.endnotePr}
              />
            )}
          </Suspense>
          {/* Style Editor Dialog */}
          {styleEditorState.open && (
            <Suspense fallback={null}>
              <StyleEditorDialog
                isOpen={styleEditorState.open}
                mode={styleEditorState.mode}
                styleId={styleEditorState.styleId}
                initialData={
                  styleEditorState.mode === 'modify' && styleEditorState.styleId
                    ? (() => {
                        const styleDef = getStyleDef(styleEditorState.styleId!);
                        if (!styleDef) return {};
                        return {
                          name: styleDef.name || styleDef.styleId,
                          basedOn: styleDef.basedOn || '',
                          next: styleDef.next || '',
                          fontFamily: styleDef.rPr?.fontFamily?.ascii || '',
                          fontSize: styleDef.rPr?.fontSize || 22,
                          bold: !!styleDef.rPr?.bold,
                          italic: !!styleDef.rPr?.italic,
                          color: styleDef.rPr?.color?.rgb || '000000',
                          alignment:
                            (styleDef.pPr?.alignment as 'left' | 'center' | 'right' | 'justify') ||
                            'left',
                          spaceBefore: styleDef.pPr?.spaceBefore || 0,
                          spaceAfter: styleDef.pPr?.spaceAfter || 0,
                          lineSpacing: styleDef.pPr?.lineSpacing || 240,
                          numPr:
                            styleDef.pPr?.numPr?.numId != null
                              ? {
                                  numId: styleDef.pPr.numPr.numId,
                                  ilvl: styleDef.pPr.numPr.ilvl ?? 0,
                                }
                              : null,
                        };
                      })()
                    : (() => {
                        // Create mode: pre-fill from selection formatting.
                        // selectionFormatting only has EXPLICIT marks — if font comes
                        // from the paragraph style, we need to resolve it from the style.
                        const sf = state.selectionFormatting;
                        const styleId = sf.styleId || 'Normal';

                        // Resolve the paragraph style to get inherited font/size/color
                        const resolved =
                          crossRefStyleResolverRef.current?.resolveParagraphStyle(styleId);
                        const styleFont = resolved?.runFormatting?.fontFamily?.ascii || '';
                        const styleSize = resolved?.runFormatting?.fontSize || 22;
                        const styleBold = !!resolved?.runFormatting?.bold;
                        const styleItalic = !!resolved?.runFormatting?.italic;
                        const styleColor = resolved?.runFormatting?.color?.rgb || '000000';
                        const styleAlignment = resolved?.paragraphFormatting?.alignment || 'left';
                        const styleSpaceBefore = resolved?.paragraphFormatting?.spaceBefore || 0;
                        const styleSpaceAfter = resolved?.paragraphFormatting?.spaceAfter || 0;
                        const styleLineSpacing = resolved?.paragraphFormatting?.lineSpacing || 240;

                        return {
                          // Explicit marks override style-resolved values
                          fontFamily: sf.fontFamily || styleFont,
                          fontSize: sf.fontSize || styleSize,
                          bold: sf.bold != null ? sf.bold : styleBold,
                          italic: sf.italic != null ? sf.italic : styleItalic,
                          underline: !!sf.underline,
                          strikethrough: !!sf.strike,
                          color: sf.color?.replace('#', '') || styleColor,
                          alignment:
                            (sf.alignment as 'left' | 'center' | 'right' | 'justify') ||
                            (styleAlignment as 'left' | 'center' | 'right' | 'justify') ||
                            'left',
                          lineSpacing: sf.lineSpacing || styleLineSpacing,
                          spaceBefore: styleSpaceBefore,
                          spaceAfter: styleSpaceAfter,
                          basedOn: styleId,
                        };
                      })()
                }
                availableStyles={
                  crossRefStyleResolverRef.current
                    ?.getParagraphStyles()
                    .map((s) => ({ styleId: s.styleId, name: s.name || s.styleId })) ?? []
                }
                affectedParagraphCount={0}
                numberingOptions={(() => {
                  const numbering = agentRef.current?.getDocument()?.package.numbering;
                  if (!numbering) return [];
                  const opts: Array<{ value: string; name: string; preview: string }> = [];
                  const seen = new Set<number>();
                  for (const num of numbering.nums || []) {
                    if (seen.has(num.numId)) continue;
                    seen.add(num.numId);
                    const abstract = (numbering.abstractNums || []).find(
                      (a) => a.abstractNumId === num.abstractNumId
                    );
                    if (!abstract) continue;
                    const name = abstract.name || `Numbering #${num.numId}`;
                    const preview = (abstract.levels || [])
                      .slice(0, 4)
                      .map((l) => l.lvlText || '?')
                      .join(' / ');
                    opts.push({ value: String(num.numId), name, preview });
                  }
                  return opts;
                })()}
                onSave={(formData) => {
                  const doc = agentRef.current?.getDocument();
                  if (doc?.package.styles) {
                    const styles = doc.package.styles;
                    if (styleEditorState.mode === 'modify' && styleEditorState.styleId) {
                      const existing = styles.styles.find(
                        (s) => s.styleId === styleEditorState.styleId
                      );
                      if (existing) {
                        existing.rPr = {
                          ...existing.rPr,
                          fontFamily: formData.fontFamily
                            ? { ascii: formData.fontFamily, hAnsi: formData.fontFamily }
                            : existing.rPr?.fontFamily,
                          fontSize: formData.fontSize || existing.rPr?.fontSize,
                          bold: formData.bold || undefined,
                          italic: formData.italic || undefined,
                          strike: formData.strikethrough || undefined,
                          underline: formData.underline ? { style: 'single' } : undefined,
                          color: formData.color ? { rgb: formData.color } : existing.rPr?.color,
                        };
                        existing.pPr = {
                          ...existing.pPr,
                          alignment: formData.alignment === 'justify' ? 'both' : formData.alignment,
                          spaceBefore: formData.spaceBefore,
                          spaceAfter: formData.spaceAfter,
                          lineSpacing: formData.lineSpacing,
                          numPr: formData.numPr || undefined,
                        };
                        if (formData.basedOn) existing.basedOn = formData.basedOn;
                        if (formData.next) existing.next = formData.next;
                        existing._dirty = true;
                      }
                    } else if (styleEditorState.mode === 'create' && formData.name) {
                      styles.styles.push({
                        styleId: formData.name.replace(/\s+/g, '_'),
                        type: 'paragraph' as const,
                        name: formData.name,
                        basedOn: formData.basedOn || undefined,
                        next: formData.next || undefined,
                        qFormat: true,
                        rPr: {
                          fontFamily: formData.fontFamily
                            ? { ascii: formData.fontFamily, hAnsi: formData.fontFamily }
                            : undefined,
                          fontSize: formData.fontSize || undefined,
                          bold: formData.bold || undefined,
                          italic: formData.italic || undefined,
                          strike: formData.strikethrough || undefined,
                          underline: formData.underline ? { style: 'single' as const } : undefined,
                          color: formData.color ? { rgb: formData.color } : undefined,
                        },
                        pPr: {
                          alignment: formData.alignment === 'justify' ? 'both' : formData.alignment,
                          spaceBefore: formData.spaceBefore,
                          spaceAfter: formData.spaceAfter,
                          lineSpacing: formData.lineSpacing,
                          numPr: formData.numPr || undefined,
                        },
                        _dirty: true,
                      });
                    }
                    doc.package.stylesDirty = true;
                    // Create new styles reference to trigger React re-render
                    doc.package.styles = { ...styles, styles: [...styles.styles] };
                    setDocumentStyles(doc.package.styles.styles);
                    // Force document state update so Toolbar re-reads styles
                    history.push({ ...doc, package: { ...doc.package } });

                    // Reapply modified style formatting to all paragraphs using it.
                    // ProseMirror stores formatting as marks on text nodes, not as
                    // live references to style definitions. We must update the marks
                    // to reflect the changed style definition.
                    //
                    // Strategy: for each mark type, check if it spans the ENTIRE
                    // paragraph text (= style-level formatting). If so, replace it.
                    // If a mark only covers part of the paragraph, it's a user-applied
                    // emphasis override — leave it alone.
                    if (styleEditorState.mode === 'modify' && styleEditorState.styleId) {
                      const view = pagedEditorRef.current?.getView();
                      if (view) {
                        const { state, dispatch } = view;
                        const { schema } = state;
                        const tr = state.tr;
                        let changed = false;

                        // Map of mark type name → { enabled, create attrs }
                        const styleMarks: Record<
                          string,
                          {
                            enabled: boolean;
                            create: () => ReturnType<typeof schema.marks.bold.create> | null;
                          }
                        > = {
                          bold: {
                            enabled: !!formData.bold,
                            create: () => schema.marks.bold?.create() ?? null,
                          },
                          italic: {
                            enabled: !!formData.italic,
                            create: () => schema.marks.italic?.create() ?? null,
                          },
                          strike: {
                            enabled: !!formData.strikethrough,
                            create: () => schema.marks.strike?.create() ?? null,
                          },
                          underline: {
                            enabled: !!formData.underline,
                            create: () =>
                              schema.marks.underline?.create({ style: 'single' }) ?? null,
                          },
                          fontSize: {
                            enabled: !!formData.fontSize,
                            create: () =>
                              schema.marks.fontSize?.create({ size: formData.fontSize }) ?? null,
                          },
                          fontFamily: {
                            enabled: !!formData.fontFamily,
                            create: () =>
                              schema.marks.fontFamily?.create({
                                ascii: formData.fontFamily,
                                hAnsi: formData.fontFamily,
                              }) ?? null,
                          },
                          textColor: {
                            enabled: !!formData.color,
                            create: () =>
                              schema.marks.textColor?.create({ rgb: formData.color }) ?? null,
                          },
                        };

                        state.doc.descendants((node, pos) => {
                          if (
                            node.type.name !== 'paragraph' ||
                            node.attrs.styleId !== styleEditorState.styleId
                          )
                            return;

                          const paraStart = pos + 1;
                          const paraEnd = pos + node.nodeSize - 1;
                          const paraTextLen = paraEnd - paraStart;
                          if (paraTextLen <= 0) return;

                          for (const [markName, def] of Object.entries(styleMarks)) {
                            const markType = schema.marks[markName];
                            if (!markType) continue;

                            // Check if this mark spans the entire paragraph
                            let coveredLen = 0;
                            node.forEach((child) => {
                              if (
                                child.marks.some(
                                  (m: { type: { name: string } }) => m.type.name === markName
                                )
                              ) {
                                coveredLen += child.nodeSize;
                              }
                            });
                            const isFullParagraph = coveredLen === paraTextLen;
                            const isPartial = coveredLen > 0 && coveredLen < paraTextLen;

                            // Only touch full-paragraph marks (style-level).
                            // Partial marks are user emphasis — leave them alone.
                            if (isPartial) continue;

                            if (isFullParagraph && !def.enabled) {
                              // Style no longer has this formatting — remove it
                              tr.removeMark(paraStart, paraEnd, markType);
                              changed = true;
                            } else if (isFullParagraph && def.enabled) {
                              // Style still has it — update to new value (e.g. new fontSize)
                              tr.removeMark(paraStart, paraEnd, markType);
                              const mark = def.create();
                              if (mark) tr.addMark(paraStart, paraEnd, mark);
                              changed = true;
                            } else if (!isFullParagraph && def.enabled) {
                              // Style adds new formatting — apply to entire paragraph
                              const mark = def.create();
                              if (mark) tr.addMark(paraStart, paraEnd, mark);
                              changed = true;
                            }
                          }
                        });

                        if (changed) {
                          dispatch(tr);
                        }
                      }
                    }
                  }
                  setStyleEditorState({ open: false, mode: 'modify' });
                }}
                onClose={() => setStyleEditorState({ open: false, mode: 'modify' })}
              />
            </Suspense>
          )}
          {/* InlineHeaderFooterEditor is rendered inside the editor content area (position:relative div) */}
          {/* Hidden file input for image insertion */}
          <input
            ref={imageInputRef}
            type="file"
            accept="image/*"
            style={{ display: 'none' }}
            onChange={handleImageFileChange}
          />
        </div>
      </ErrorBoundary>
    </ErrorProvider>
  );
});

// ============================================================================
// EXPORTS
// ============================================================================

export default DocxEditor;
