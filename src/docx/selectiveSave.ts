/**
 * Selective Save Module
 *
 * Orchestrates selective XML patching for the save flow.
 * Serializes full document.xml, validates patch safety, builds patched XML,
 * and calls updateMultipleFiles() to produce the final DOCX.
 *
 * Returns null on any failure, signaling the caller to fall back to full repack.
 */

import type { Document, BlockContent } from '../types/document';
import { serializeDocument } from './serializer/documentSerializer';
import {
  serializeCommentsWithInfo,
  serializeCommentsExtended,
  serializeCommentsIds,
  serializeCommentsExtensible,
} from './serializer/commentSerializer';
import { buildPatchedDocumentXml } from './selectiveXmlPatch';
import {
  updateMultipleFiles,
  COMMENTS_CONTENT_TYPE,
  COMMENTS_EXTENDED_CONTENT_TYPE,
  COMMENTS_IDS_CONTENT_TYPE,
  COMMENTS_EXTENSIBLE_CONTENT_TYPE,
} from './rezip';
import { RELATIONSHIP_TYPES } from './relsParser';

/**
 * Check if document content has new images (data: URL without rId) or
 * new hyperlinks (href without rId). Combined into a single traversal
 * to avoid walking the block tree twice.
 */
function hasNewImagesOrHyperlinks(blocks: BlockContent[]): boolean {
  for (const block of blocks) {
    if (block.type === 'paragraph') {
      for (const item of block.content) {
        if (item.type === 'run') {
          for (const c of item.content) {
            if (c.type === 'drawing' && c.image?.src?.startsWith('data:') && !c.image?.rId) {
              return true;
            }
          }
        } else if (item.type === 'hyperlink' && item.href && !item.rId && !item.anchor) {
          return true;
        }
      }
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          if (hasNewImagesOrHyperlinks(cell.content)) return true;
        }
      }
    }
  }
  return false;
}

export interface SelectiveSaveOptions {
  /** Changed paragraph IDs to selectively patch */
  changedParaIds: Set<string>;
  /** Whether structural changes occurred (paragraph add/delete) */
  structuralChange: boolean;
  /** Whether any changes affected paragraphs without paraId */
  hasUntrackedChanges: boolean;
}

/**
 * Attempt a selective save — patch only changed paragraphs in document.xml.
 * Also updates comments and comment extension parts so all parts stay in sync.
 *
 * Returns the saved ArrayBuffer, or null if selective save is not possible
 * (caller should fall back to full repack).
 */
export async function attemptSelectiveSave(
  doc: Document,
  originalBuffer: ArrayBuffer,
  options: SelectiveSaveOptions
): Promise<ArrayBuffer | null> {
  const { changedParaIds, structuralChange, hasUntrackedChanges } = options;

  // Bail out conditions — fall back to full repack
  if (structuralChange) return null;
  if (hasUntrackedChanges) return null;
  if (!originalBuffer) return null;

  // Check for new images/hyperlinks that need relationship management
  const content = doc.package.document.content;
  if (hasNewImagesOrHyperlinks(content)) return null;

  // If no changes, just return the original buffer as-is
  if (changedParaIds.size === 0) {
    return originalBuffer;
  }

  try {
    // Get the original document.xml from the ZIP
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(originalBuffer);
    const docXmlFile = zip.file('word/document.xml');
    if (!docXmlFile) return null;
    const originalDocXml = await docXmlFile.async('text');

    // Serialize the full document.xml from the current document model
    const serializedDocXml = serializeDocument(doc);

    // Build the patched document.xml
    const patchedDocXml = buildPatchedDocumentXml(originalDocXml, serializedDocXml, changedParaIds);
    if (!patchedDocXml) return null;

    // Apply the patch using updateMultipleFiles
    const updates = new Map<string, string>();
    updates.set('word/document.xml', patchedDocXml);

    // Always serialize comments + extension parts when the document has comments
    const comments = doc.package.document.comments;
    const hasComments = comments && comments.length > 0;
    if (hasComments) {
      const { xml: commentsXml, paraInfos } = serializeCommentsWithInfo(comments);
      updates.set('word/comments.xml', commentsXml);

      // Write commentsExtended.xml for reply threading (Word/Google Docs interop)
      const extendedXml = serializeCommentsExtended(paraInfos);
      if (extendedXml) updates.set('word/commentsExtended.xml', extendedXml);

      // Write commentsIds.xml for stable IDs (Word Online needs this for replies)
      const idsXml = serializeCommentsIds(paraInfos);
      if (idsXml) updates.set('word/commentsIds.xml', idsXml);

      // Write commentsExtensible.xml for UTC dates (Pages, Word 2016+)
      const extensibleXml = serializeCommentsExtensible(paraInfos, comments);
      if (extensibleXml) updates.set('word/commentsExtensible.xml', extensibleXml);

      // Ensure [Content_Types].xml has Overrides for all comment parts
      const ctFile = zip.file('[Content_Types].xml');
      if (ctFile) {
        let ctXml = await ctFile.async('text');
        let ctChanged = false;
        const ctEntries: [string, string][] = [
          ['/word/comments.xml', COMMENTS_CONTENT_TYPE],
          ['/word/commentsExtended.xml', COMMENTS_EXTENDED_CONTENT_TYPE],
          ['/word/commentsIds.xml', COMMENTS_IDS_CONTENT_TYPE],
          ['/word/commentsExtensible.xml', COMMENTS_EXTENSIBLE_CONTENT_TYPE],
        ];
        for (const [partName, contentType] of ctEntries) {
          if (!ctXml.includes(partName)) {
            ctXml = ctXml.replace(
              '</Types>',
              `<Override PartName="${partName}" ContentType="${contentType}"/></Types>`
            );
            ctChanged = true;
          }
        }
        if (ctChanged) updates.set('[Content_Types].xml', ctXml);
      }

      // Ensure word/_rels/document.xml.rels has Relationships for all comment parts
      const relsPath = 'word/_rels/document.xml.rels';
      const relsFile = zip.file(relsPath);
      if (relsFile) {
        let relsXml = await relsFile.async('text');
        let relsChanged = false;
        const relEntries: [string, string][] = [
          ['comments.xml', RELATIONSHIP_TYPES.comments],
          ['commentsExtended.xml', RELATIONSHIP_TYPES.commentsExtended],
          ['commentsIds.xml', RELATIONSHIP_TYPES.commentsIds],
          ['commentsExtensible.xml', RELATIONSHIP_TYPES.commentsExtensible],
        ];
        for (const [target, type] of relEntries) {
          if (!relsXml.includes(target)) {
            const ids = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map((m) => parseInt(m[1], 10));
            const newRId = `rId${ids.length > 0 ? Math.max(...ids) + 1 : 100}`;
            relsXml = relsXml.replace(
              '</Relationships>',
              `<Relationship Id="${newRId}" Type="${type}" Target="${target}"/></Relationships>`
            );
            relsChanged = true;
          }
        }
        if (relsChanged) updates.set(relsPath, relsXml);
      }
    }

    return await updateMultipleFiles(originalBuffer, updates);
  } catch {
    // Any error — fall back to full repack
    return null;
  }
}
