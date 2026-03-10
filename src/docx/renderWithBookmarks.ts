/**
 * Render context tags in a Document model with bookmark markers.
 *
 * After fromProseDoc() converts contextTag atoms to {{ tagKey }} text,
 * this module replaces those patterns with rendered values wrapped in
 * OOXML bookmarks (_FP_ctx_{metaId}), enabling round-trip tag restoration.
 */

import type {
  ParagraphContent,
  Run,
  BookmarkStart,
  BookmarkEnd,
  BlockContent,
  Table,
  TableRow,
  TableCell,
} from '../types/content';
import type { Document, ContextTagMeta } from '../types/document';

/** Prefix for context-tag bookmarks */
export const FP_BOOKMARK_PREFIX = '_FP_ctx_';

/** Regex matching {{ context.tag! }} or {context.tag!} patterns in a single text string */
const TAG_PATTERN = /\{\{\s*(context\.[\w.]+)(!?)\s*\}\}|\{(context\.[\w.]+)(!?)\}/g;

/**
 * Ordered list of metaId entries grouped by tagKey, for document-order matching.
 */
function groupMetaByTagKey(
  ctMeta: Record<string, ContextTagMeta>
): Map<string, Array<{ metaId: string; meta: ContextTagMeta }>> {
  const map = new Map<string, Array<{ metaId: string; meta: ContextTagMeta }>>();
  for (const [metaId, meta] of Object.entries(ctMeta)) {
    const tagKey = meta.tagKey;
    if (!tagKey) continue;
    let arr = map.get(tagKey);
    if (!arr) {
      arr = [];
      map.set(tagKey, arr);
    }
    arr.push({ metaId, meta });
  }
  return map;
}

interface RenderOptions {
  /** Context tag values: tagKey → rendered text */
  tags: Record<string, string>;
  /** Context tag metadata: metaId → { tagKey, removeIfEmpty, ... } */
  ctMeta: Record<string, ContextTagMeta>;
  /** Rendering mode */
  mode: 'omit' | 'keep';
  /** Starting bookmark ID (to avoid collisions with existing bookmarks) */
  startBookmarkId?: number;
}

interface RenderResult {
  /** The new paragraph content with bookmarks */
  content: ParagraphContent[];
  /** Whether the paragraph should be removed (removeIfEmpty with no value in omit mode) */
  removeParagraph: boolean;
  /** Next available bookmark ID after this paragraph */
  nextBookmarkId: number;
}

/**
 * Find the highest bookmark ID already used in the document body.
 */
export function findMaxBookmarkId(doc: Document): number {
  let maxId = 0;
  const body = doc.package.document;
  if (!body) return maxId;

  function scanBlocks(blocks: BlockContent[]) {
    for (const block of blocks) {
      if (block.type === 'paragraph') {
        for (const item of block.content) {
          if (item.type === 'bookmarkStart' || item.type === 'bookmarkEnd') {
            if (item.id > maxId) maxId = item.id;
          }
        }
      } else if (block.type === 'table') {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            scanBlocks(cell.content);
          }
        }
      }
    }
  }

  scanBlocks(body.content);
  return maxId;
}

/**
 * Extract all text from a Run.
 */
function getRunText(run: Run): string {
  let text = '';
  for (const c of run.content) {
    if (c.type === 'text') text += c.text;
  }
  return text;
}

/**
 * Create a text Run with the same formatting as the original, but different text.
 */
function createRunWithText(text: string, formatting: Run['formatting']): Run {
  return {
    type: 'run',
    formatting: formatting ? { ...formatting } : undefined,
    content: [{ type: 'text', text }],
  };
}

/**
 * Process a single paragraph's content, replacing tag patterns with rendered
 * values wrapped in bookmarks.
 *
 * Consumes metaIds from `cursors` map in document order.
 */
export function renderParagraphContent(
  content: ParagraphContent[],
  cursors: Map<string, number>,
  metaGroups: Map<string, Array<{ metaId: string; meta: ContextTagMeta }>>,
  options: RenderOptions
): RenderResult {
  let bookmarkId = options.startBookmarkId ?? 100;
  const result: ParagraphContent[] = [];
  let removeParagraph = false;

  for (const item of content) {
    if (item.type !== 'run') {
      // Pass through non-run content (existing bookmarks, fields, etc.)
      result.push(item);
      continue;
    }

    const runText = getRunText(item);
    if (!runText) {
      result.push(item);
      continue;
    }

    // Check if this run contains any tag patterns
    TAG_PATTERN.lastIndex = 0;
    const match = TAG_PATTERN.exec(runText);
    if (!match) {
      result.push(item);
      continue;
    }

    // Run contains tag patterns — split and process
    TAG_PATTERN.lastIndex = 0;
    let lastIndex = 0;
    let localMatch: RegExpExecArray | null;

    while ((localMatch = TAG_PATTERN.exec(runText)) !== null) {
      const fullMatch = localMatch[0];
      const tagKey = localMatch[1] || localMatch[3];
      const rieFlag = localMatch[2] || localMatch[4];
      const matchStart = localMatch.index;

      // Text before the tag pattern
      if (matchStart > lastIndex) {
        result.push(createRunWithText(runText.slice(lastIndex, matchStart), item.formatting));
      }

      // Look up rendered value
      const resolved = options.tags[tagKey];

      // Consume next metaId for this tagKey
      const group = metaGroups.get(tagKey);
      const cursor = cursors.get(tagKey) ?? 0;
      const entry = group?.[cursor];
      const metaId = entry?.metaId;

      // removeIfEmpty: check CustomXML manifest first, fall back to legacy `!` flag
      const removeIfEmpty = entry?.meta?.removeIfEmpty ?? rieFlag === '!';
      if (group && cursor < group.length) {
        cursors.set(tagKey, cursor + 1);
      }

      if (resolved) {
        // Tag resolved — emit bookmark + rendered text
        if (metaId) {
          const bStart: BookmarkStart = {
            type: 'bookmarkStart',
            id: bookmarkId,
            name: `${FP_BOOKMARK_PREFIX}${metaId}`,
          };
          const bEnd: BookmarkEnd = {
            type: 'bookmarkEnd',
            id: bookmarkId,
          };
          result.push(bStart);
          result.push(createRunWithText(resolved, item.formatting));
          result.push(bEnd);
          bookmarkId++;
        } else {
          // No metaId available — just emit rendered text without bookmark
          result.push(createRunWithText(resolved, item.formatting));
        }
      } else if (removeIfEmpty && options.mode === 'omit') {
        // Tag has no value and removeIfEmpty — flag paragraph for removal
        removeParagraph = true;
      } else if (options.mode === 'keep') {
        // Keep mode, unresolved tag — preserve as {tagKey} with bookmark
        const keepText = `{${tagKey}}`;
        if (metaId) {
          const bStart: BookmarkStart = {
            type: 'bookmarkStart',
            id: bookmarkId,
            name: `${FP_BOOKMARK_PREFIX}${metaId}`,
          };
          const bEnd: BookmarkEnd = {
            type: 'bookmarkEnd',
            id: bookmarkId,
          };
          result.push(bStart);
          result.push(createRunWithText(keepText, item.formatting));
          result.push(bEnd);
          bookmarkId++;
        } else {
          result.push(createRunWithText(keepText, item.formatting));
        }
      }
      // 'omit' without removeIfEmpty and no value → skip (omit the text)

      lastIndex = matchStart + fullMatch.length;
    }

    // Text after the last tag pattern
    if (lastIndex < runText.length) {
      result.push(createRunWithText(runText.slice(lastIndex), item.formatting));
    }
  }

  return { content: result, removeParagraph, nextBookmarkId: bookmarkId };
}

/**
 * Render all context tags in a Document with bookmark markers.
 *
 * Mutates the document in-place: replaces {{ tagKey }} text in paragraph
 * runs with rendered values wrapped in bookmarks.
 *
 * Also stores context tag metadata so bookmarks can be resolved on re-upload.
 */
export function renderDocumentWithBookmarks(
  doc: Document,
  options: Omit<RenderOptions, 'startBookmarkId'>
): void {
  const body = doc.package.document;
  if (!body) return;

  const metaGroups = groupMetaByTagKey(options.ctMeta);
  const cursors = new Map<string, number>();
  let bookmarkId = findMaxBookmarkId(doc) + 100; // generous gap

  function processBlocks(blocks: BlockContent[]): BlockContent[] {
    const result: BlockContent[] = [];

    for (const block of blocks) {
      if (block.type === 'paragraph') {
        const rendered = renderParagraphContent(block.content, cursors, metaGroups, {
          ...options,
          startBookmarkId: bookmarkId,
        });
        bookmarkId = rendered.nextBookmarkId;

        if (rendered.removeParagraph) continue; // omit the paragraph
        result.push({ ...block, content: rendered.content });
      } else if (block.type === 'table') {
        const newRows: TableRow[] = block.rows.map((row: TableRow) => ({
          ...row,
          cells: row.cells.map((cell: TableCell) => {
            const processed: (typeof cell.content)[number][] = [];
            for (const cellBlock of cell.content) {
              if (cellBlock.type === 'paragraph') {
                const rendered = renderParagraphContent(cellBlock.content, cursors, metaGroups, {
                  ...options,
                  startBookmarkId: bookmarkId,
                });
                bookmarkId = rendered.nextBookmarkId;
                if (!rendered.removeParagraph) {
                  processed.push({ ...cellBlock, content: rendered.content });
                }
              } else {
                // Nested table — recurse through its rows
                processed.push(cellBlock);
              }
            }
            return { ...cell, content: processed };
          }),
        }));
        result.push({ ...block, rows: newRows } as Table);
      } else {
        result.push(block);
      }
    }

    return result;
  }

  body.content = processBlocks(body.content);

  // Always preserve context tag metadata for bookmark resolution on re-upload
  doc.contextTagMetadata = options.ctMeta;
}

// ============================================================================
// RESTORATION — Convert bookmarked rendered text back to context tag patterns
// ============================================================================

/**
 * Pre-process a paragraph's content to restore context tags from FP bookmarks.
 *
 * Finds `_FP_ctx_{metaId}` bookmark pairs, looks up the tag metadata in the manifest,
 * and replaces the bookmarked text with `{{ tagKey! }}` pattern text so the existing
 * `splitContextTags` logic in toProseDoc.ts can convert them to contextTag atoms.
 *
 * Returns a new content array with bookmarks consumed and tag patterns restored.
 */
export function restoreParagraphContent(
  content: ParagraphContent[],
  manifest: Record<string, ContextTagMeta>
): ParagraphContent[] {
  // Build a map of bookmark ID → metaId for FP bookmarks
  const fpBookmarks = new Map<number, string>(); // bookmarkId → metaId
  for (const item of content) {
    if (item.type === 'bookmarkStart' && item.name.startsWith(FP_BOOKMARK_PREFIX)) {
      const metaId = item.name.slice(FP_BOOKMARK_PREFIX.length);
      fpBookmarks.set(item.id, metaId);
    }
  }

  if (fpBookmarks.size === 0) return content;

  // Process content: replace bookmarked regions with tag patterns
  const result: ParagraphContent[] = [];
  let captureForBookmarkId: number | null = null;
  let capturedMetaId = '';
  let capturedFormatting: Run['formatting'] | undefined;

  for (const item of content) {
    if (item.type === 'bookmarkStart' && fpBookmarks.has(item.id)) {
      // Start capturing — skip the bookmark start itself
      captureForBookmarkId = item.id;
      capturedMetaId = fpBookmarks.get(item.id)!;
      capturedFormatting = undefined;
      continue;
    }

    if (item.type === 'bookmarkEnd' && item.id === captureForBookmarkId) {
      // End capture — emit the tag pattern
      const meta = manifest[capturedMetaId];
      if (meta?.tagKey) {
        // removeIfEmpty is now tracked in CustomXML metadata, not in-band.
        const tagText = `{{ ${meta.tagKey} }}`;
        result.push(createRunWithText(tagText, capturedFormatting));
      }
      // If no manifest entry, the bookmarked text is lost (graceful degradation)
      captureForBookmarkId = null;
      capturedMetaId = '';
      continue;
    }

    if (captureForBookmarkId !== null) {
      // Inside a bookmark capture — grab formatting from first run
      if (item.type === 'run' && !capturedFormatting) {
        capturedFormatting = item.formatting;
      }
      // Skip the captured content (it was the rendered value)
      continue;
    }

    // Pass through everything else
    result.push(item);
  }

  return result;
}

/**
 * Restore context tags from FP bookmarks in the entire Document model.
 *
 * Call this BEFORE toProseDoc() conversion so that `splitContextTags`
 * picks up the restored `{{ tagKey }}` patterns.
 *
 * Mutates the document in-place.
 */
export function restoreContextTagsFromBookmarks(doc: Document): void {
  const manifest = doc.contextTagMetadata;
  if (!manifest || Object.keys(manifest).length === 0) return;

  const body = doc.package.document;
  if (!body) return;

  const m = manifest; // local const for closure narrowing

  function processBlocks(blocks: BlockContent[]): void {
    for (let i = 0; i < blocks.length; i++) {
      const block = blocks[i];
      if (block.type === 'paragraph') {
        block.content = restoreParagraphContent(block.content, m);
      } else if (block.type === 'table') {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            processBlocks(cell.content as BlockContent[]);
          }
        }
      }
    }
  }

  processBlocks(body.content);
}
