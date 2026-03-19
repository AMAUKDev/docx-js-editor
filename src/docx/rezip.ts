/**
 * DOCX Repacker - Repack modified document into valid DOCX
 *
 * Takes a Document with modified content and creates a new DOCX file
 * by updating document.xml while preserving all other files from
 * the original ZIP archive.
 *
 * This ensures round-trip fidelity:
 * - styles.xml, theme1.xml, fontTable.xml remain untouched
 * - Media files preserved
 * - Relationships preserved
 * - Only document.xml is updated with new content
 *
 * OOXML Package Structure:
 * - [Content_Types].xml - Content type declarations
 * - _rels/.rels - Package relationships
 * - word/document.xml - Main document (modified)
 * - word/styles.xml - Styles (preserved)
 * - word/theme/theme1.xml - Theme (preserved)
 * - word/numbering.xml - Numbering (preserved)
 * - word/fontTable.xml - Font table (preserved)
 * - word/settings.xml - Settings (preserved)
 * - word/header*.xml - Headers (preserved)
 * - word/footer*.xml - Footers (preserved)
 * - word/footnotes.xml - Footnotes (preserved)
 * - word/endnotes.xml - Endnotes (preserved)
 * - word/media/* - Media files (preserved)
 * - word/_rels/document.xml.rels - Document relationships (preserved)
 * - docProps/* - Document properties (preserved)
 */

import JSZip from 'jszip';
import type { Document } from '../types/document';
import type { BlockContent, Image } from '../types/content';
import { serializeDocument, documentHasLockedParagraphs } from './serializer/documentSerializer';
import { serializeHeaderFooter } from './serializer/headerFooterSerializer';
import { RELATIONSHIP_TYPES } from './relsParser';
import { type RawDocxContent } from './unzip';
import {
  CUSTOM_XML_LEGACY_PATH,
  CUSTOM_XML_PROPS_CONTENT_TYPE,
  FP_DATASTORE_GUID,
  serializeManifest,
} from './contextTagMetadata';
import { FP_BOOKMARK_PREFIX } from './renderWithBookmarks';
import type { Comment } from '../types/content';

// ============================================================================
// COMMENT SERIALIZATION
// ============================================================================

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Serialize Comment objects to OOXML comments.xml format.
 */
function serializeCommentsXml(comments: Comment[]): string {
  const parts: string[] = [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:comments xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ' +
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
      'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
      'xmlns:v="urn:schemas-microsoft-com:vml" ' +
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
      'xmlns:w10="urn:schemas-microsoft-com:office:word" ' +
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
      'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">',
  ];

  for (const comment of comments) {
    const initials =
      comment.initials ||
      comment.author
        ?.split(' ')
        .map((w) => w[0])
        .join('') ||
      '';
    const dateAttr = comment.date ? ` w:date="${escapeXml(comment.date)}"` : '';
    parts.push(
      `<w:comment w:id="${comment.id}" w:author="${escapeXml(comment.author || 'Unknown')}" w:initials="${escapeXml(initials)}"${dateAttr}>`
    );

    // Serialize comment content as paragraphs
    if (comment.content && comment.content.length > 0) {
      for (const para of comment.content) {
        parts.push('<w:p>');
        for (const item of para.content || []) {
          if ('content' in item) {
            const run = item as { content: Array<{ type: string; text?: string }> };
            parts.push('<w:r>');
            for (const rc of run.content) {
              if (rc.type === 'text' && rc.text) {
                parts.push(`<w:t xml:space="preserve">${escapeXml(rc.text)}</w:t>`);
              }
            }
            parts.push('</w:r>');
          }
        }
        parts.push('</w:p>');
      }
    } else {
      parts.push('<w:p><w:r><w:t></w:t></w:r></w:p>');
    }

    parts.push('</w:comment>');
  }

  parts.push('</w:comments>');
  return parts.join('');
}

// ============================================================================
// COMMENT MARKER INJECTION (into original XML without re-serialization)
// ============================================================================

/**
 * Inject commentRangeStart/End markers into original document.xml for new comments.
 * Uses the ProseMirror comment mark positions to find corresponding <w:t> text
 * in the original XML and bracket it with OOXML comment markers.
 *
 * This avoids lossy full re-serialization while still embedding comments.
 */
function injectCommentMarkersIntoXml(xml: string, doc: Document): string {
  const comments = doc.package.document?.comments;
  if (!comments || comments.length === 0) return xml;

  // Extract comment text from each comment to find its location in the XML
  for (const comment of comments) {
    // Check if this comment's markers are already in the XML
    if (xml.includes(`w:id="${comment.id}"`) && xml.includes('commentRangeStart')) {
      continue; // Already has markers
    }

    // Find the first <w:p> in <w:body> and inject comment markers there.
    // This is a simple heuristic — the comment will be attached to the first
    // text run we can find. The PM comment marks handle precise positioning
    // when contentDirty is true (full re-serialization path).
    const bodyStart = xml.indexOf('<w:body');
    if (bodyStart === -1) continue;

    // Find first <w:p> after <w:body>
    const firstPIdx = xml.indexOf('<w:p ', bodyStart);
    if (firstPIdx === -1) continue;

    // Find first <w:r> in that paragraph
    const firstRIdx = xml.indexOf('<w:r>', firstPIdx);
    const firstRIdx2 = xml.indexOf('<w:r ', firstPIdx);
    const rIdx = Math.min(
      firstRIdx === -1 ? Infinity : firstRIdx,
      firstRIdx2 === -1 ? Infinity : firstRIdx2
    );
    if (rIdx === Infinity) continue;

    // Inject commentRangeStart before the first run
    const startMarker = `<w:commentRangeStart w:id="${comment.id}"/>`;
    const endMarker =
      `<w:commentRangeEnd w:id="${comment.id}"/>` +
      `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>` +
      `<w:commentReference w:id="${comment.id}"/></w:r>`;

    // Find the end of this paragraph
    const pEndIdx = xml.indexOf('</w:p>', rIdx);
    if (pEndIdx === -1) continue;

    // Insert: commentRangeStart before first run, commentRangeEnd before </w:p>
    xml =
      xml.slice(0, rIdx) + startMarker + xml.slice(rIdx, pEndIdx) + endMarker + xml.slice(pEndIdx);
  }

  return xml;
}

// ============================================================================
// NEW IMAGE HANDLING
// ============================================================================

/**
 * Collect all images with data-URL src from the document content.
 * These are newly inserted images that need to be added to the ZIP.
 */
function collectNewImages(blocks: BlockContent[], existingRIds: Set<string>): Image[] {
  const images: Image[] = [];

  for (const block of blocks) {
    if (block.type === 'paragraph') {
      for (const item of block.content) {
        if (item.type === 'run') {
          for (const c of item.content) {
            if (c.type === 'drawing' && c.image.src?.startsWith('data:')) {
              // Only treat as "new" if the image doesn't already have a valid rId
              // in the original ZIP's rels. After parse→display round-trip ALL images
              // get data URL srcs, but existing images already have their binary in
              // the ZIP from the file-copy step — only truly new inserts lack an rId.
              if (!c.image.rId || !existingRIds.has(c.image.rId)) {
                images.push(c.image);
              }
            }
          }
        }
      }
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          images.push(...collectNewImages(cell.content, existingRIds));
        }
      }
    }
  }

  return images;
}

/** Map MIME type to file extension (inverse of getContentTypeForExtension) */
const MIME_TO_EXT: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpeg',
  'image/gif': 'gif',
  'image/bmp': 'bmp',
  'image/tiff': 'tiff',
  'image/webp': 'webp',
  'image/svg+xml': 'svg',
};

/**
 * Decode a data URL to binary ArrayBuffer and file extension.
 */
function decodeDataUrl(dataUrl: string): { data: ArrayBuffer; extension: string } {
  const match = dataUrl.match(/^data:([^;]+);base64,(.+)$/);
  if (!match) {
    throw new Error('Invalid data URL');
  }

  const binary = atob(match[2]);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }

  return { data: bytes.buffer, extension: MIME_TO_EXT[match[1]] || 'png' };
}

/**
 * Process newly inserted images: add binary data to ZIP, create relationships,
 * update content types, and rewrite rIds in the document model so the serializer
 * outputs correct references.
 *
 * Mutates the images' rId fields in-place.
 */
async function processNewImages(
  newImages: Image[],
  zip: JSZip,
  compressionLevel: number
): Promise<void> {
  if (newImages.length === 0) return;

  // Read existing relationships
  const relsPath = 'word/_rels/document.xml.rels';
  const relsFile = zip.file(relsPath);
  if (!relsFile) return;
  let relsXml = await relsFile.async('text');

  // Find highest existing rId
  let maxId = 0;
  for (const match of relsXml.matchAll(/Id="rId(\d+)"/g)) {
    const id = parseInt(match[1], 10);
    if (id > maxId) maxId = id;
  }

  // Find highest existing image number in word/media/
  let maxImageNum = 0;
  zip.forEach((relativePath) => {
    const m = relativePath.match(/^word\/media\/image(\d+)\./);
    if (m) {
      const num = parseInt(m[1], 10);
      if (num > maxImageNum) maxImageNum = num;
    }
  });

  const relEntries: string[] = [];
  const extensionsAdded = new Set<string>();

  for (const image of newImages) {
    const { data, extension } = decodeDataUrl(image.src!);

    maxImageNum++;
    maxId++;
    const mediaFilename = `image${maxImageNum}.${extension}`;
    const mediaPath = `word/media/${mediaFilename}`;
    const newRId = `rId${maxId}`;

    // Add binary to ZIP
    zip.file(mediaPath, data, {
      compression: 'DEFLATE',
      compressionOptions: { level: compressionLevel },
    });

    // Build relationship entry
    relEntries.push(
      `<Relationship Id="${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${mediaFilename}"/>`
    );

    extensionsAdded.add(extension);

    // Rewrite the image's rId so the serializer outputs the correct reference
    image.rId = newRId;
  }

  // Update relationships XML
  if (relEntries.length > 0) {
    relsXml = relsXml.replace('</Relationships>', relEntries.join('') + '</Relationships>');
    zip.file(relsPath, relsXml, {
      compression: 'DEFLATE',
      compressionOptions: { level: compressionLevel },
    });
  }

  // Update [Content_Types].xml if new extensions were added
  if (extensionsAdded.size > 0) {
    const ctFile = zip.file('[Content_Types].xml');
    if (ctFile) {
      let ctXml = await ctFile.async('text');
      for (const ext of extensionsAdded) {
        if (!ctXml.includes(`Extension="${ext}"`)) {
          const contentType = getContentTypeForExtension(ext, '');
          ctXml = ctXml.replace(
            '</Types>',
            `<Default Extension="${ext}" ContentType="${contentType}"/></Types>`
          );
        }
      }
      zip.file('[Content_Types].xml', ctXml, {
        compression: 'DEFLATE',
        compressionOptions: { level: compressionLevel },
      });
    }
  }
}

// ============================================================================
// MAIN REPACKER
// ============================================================================

/**
 * Options for repacking DOCX
 */
export interface RepackOptions {
  /** Compression level (0-9, default: 6) */
  compressionLevel?: number;
  /** Whether to update modification date in docProps/core.xml */
  updateModifiedDate?: boolean;
  /** Custom modifier name for lastModifiedBy */
  modifiedBy?: string;
}

/**
 * Repack a Document into a valid DOCX file
 *
 * @param doc - Document with modified content
 * @param options - Optional repack options
 * @returns Promise resolving to DOCX as ArrayBuffer
 * @throws Error if document has no original buffer for round-trip
 */
export async function repackDocx(doc: Document, options: RepackOptions = {}): Promise<ArrayBuffer> {
  // Validate we have an original buffer to base on
  if (!doc.originalBuffer) {
    throw new Error(
      'Cannot repack document: no original buffer for round-trip. ' +
        'Use createDocx() for new documents.'
    );
  }

  const { compressionLevel = 6, updateModifiedDate = true, modifiedBy } = options;

  // Load the original ZIP
  const originalZip = await JSZip.loadAsync(doc.originalBuffer);

  // Create a new ZIP with all original files
  const newZip = new JSZip();

  // Copy all files from original ZIP
  for (const [path, file] of Object.entries(originalZip.files)) {
    // Skip directories — OOXML ZIPs should not have directory entries.
    // JSZip synthesizes directory entries from file paths; re-creating them
    // adds entries that Word may flag as invalid.
    if (file.dir) {
      continue;
    }

    // Get original file content
    const content = await file.async('arraybuffer');

    // Add to new ZIP (we'll update specific files below)
    newZip.file(path, content, {
      compression: 'DEFLATE',
      compressionOptions: { level: compressionLevel },
    });
  }

  // Build a set of existing relationship IDs from the original rels file.
  // This lets us distinguish truly new images from existing ones that just have
  // data URL srcs after the parse→display round-trip.
  const existingRIds = new Set<string>();
  const relsFile = newZip.file('word/_rels/document.xml.rels');
  if (relsFile) {
    const relsXml = await relsFile.async('text');
    for (const match of relsXml.matchAll(/Id="(rId\d+)"/g)) {
      existingRIds.add(match[1]);
    }
  }

  // If the document content hasn't been modified (e.g. user opened and saved without
  // editing), preserve the original document.xml to avoid lossy re-serialization.
  // Our parser strips SDTs, mc:AlternateContent, rsid attrs, etc. — re-serializing
  // unmodified content produces a degraded file that Word may reject as corrupt.
  const hasModifiedComments = (doc as unknown as Record<string, boolean>).commentsModified;
  if (!doc.contentDirty && doc.originalDocumentXml) {
    // No edits made — keep the original document.xml.
    // If comments were added, inject comment markers into the original XML.
    let origXml = doc.originalDocumentXml;
    if (hasModifiedComments && doc.package.document?.comments) {
      origXml = injectCommentMarkersIntoXml(origXml, doc);
    }
    newZip.file('word/document.xml', origXml, {
      compression: 'DEFLATE',
      compressionOptions: { level: compressionLevel },
    });
  } else {
    // Content was modified — process images and re-serialize
    const newImages = collectNewImages(doc.package.document.content, existingRIds);
    await processNewImages(newImages, newZip, compressionLevel);

    const documentXml = serializeDocument(doc);
    newZip.file('word/document.xml', documentXml, {
      compression: 'DEFLATE',
      compressionOptions: { level: compressionLevel },
    });
  }

  // NOTE: We do NOT re-serialize headers/footers (lossy — drops mc:AlternateContent/VML).
  // Instead, we do text-level replacement of context tags directly in the original XML.
  if (doc.contextTagReplacements) {
    const { tags, mode } = doc.contextTagReplacements;
    const ctMeta = doc.contextTagMetadata;

    // Build tagKey → metaId lookup for H/F bookmark generation
    const tagKeyToMetaId = new Map<string, string>();
    if (ctMeta) {
      for (const [metaId, meta] of Object.entries(ctMeta)) {
        if (meta.tagKey && !tagKeyToMetaId.has(meta.tagKey)) {
          tagKeyToMetaId.set(meta.tagKey, metaId);
        }
      }
    }

    // Regex matching a <w:r> containing a context tag in its <w:t>
    // Matches {{ tag.path }} (double brace, 0+ dots) and {tag.path} (single brace, 0+ dots)
    const runWithTagRe =
      /(<w:r\b[^>]*>(?:<w:rPr>[\s\S]*?<\/w:rPr>)?<w:t[^>]*>)([^<]*?)(\{\{\s*([\w]+(?:\.[\w]+)*)!?\s*\}\}|\{([\w]+(?:\.[\w]+)*)\})([^<]*?<\/w:t><\/w:r>)/g;
    const tagRe = /\{\{\s*([\w]+(?:\.[\w]+)*)\s*\}\}|\{([\w]+(?:\.[\w]+)+)\}/g;
    let hfBookmarkId = 10000;

    for (const [path, file] of Object.entries(newZip.files)) {
      if (file.dir) continue;
      if (!/^word\/(header|footer)\d*\.xml$/i.test(path)) continue;
      let xml = await file.async('text');
      let changed = false;

      // First pass: run-level replacement with bookmarks
      xml = xml.replace(
        runWithTagRe,
        (fullMatch, beforeTag, preText, _tagMatch, g4, g5, afterTag) => {
          const rawKey = g4 || g5;
          // Strip legacy "context." prefix so lookup matches the new flat tag map
          const tagKey = rawKey.startsWith('context.') ? rawKey.slice(8) : rawKey;
          const resolved = tags[tagKey];
          const metaId = tagKeyToMetaId.get(tagKey);
          if (resolved) {
            changed = true;
            const renderedRun = `${beforeTag}${preText}${resolved}${afterTag}`;
            if (metaId) {
              const bmStart = `<w:bookmarkStart w:id="${hfBookmarkId}" w:name="${FP_BOOKMARK_PREFIX}${metaId}"/>`;
              const bmEnd = `<w:bookmarkEnd w:id="${hfBookmarkId}"/>`;
              hfBookmarkId++;
              return `${bmStart}${renderedRun}${bmEnd}`;
            }
            return renderedRun;
          }
          if (mode === 'omit') {
            changed = true;
            return `${beforeTag}${preText}${afterTag}`;
          }
          return fullMatch;
        }
      );

      // Second pass: catch remaining tags not in own run (fallback)
      xml = xml.replace(tagRe, (_match, g1, g2) => {
        const rawKey = g1 || g2;
        const tagKey = rawKey.startsWith('context.') ? rawKey.slice(8) : rawKey;
        const resolved = tags[tagKey];
        if (resolved) {
          changed = true;
          return resolved;
        }
        if (mode === 'omit') {
          changed = true;
          return '';
        }
        return _match;
      });

      if (changed) {
        newZip.file(path, xml, {
          compression: 'DEFLATE',
          compressionOptions: { level: compressionLevel },
        });
      }
    }
  }

  // ── Always reconcile comment markers in document.xml with comments.xml ──
  // This prevents "unreadable content" from orphaned commentRangeStart/End
  // markers that reference non-existent comments (e.g., after deleting comments).
  {
    const comments = doc.package.document?.comments || [];
    const validIds = new Set(comments.map((c) => c.id));
    const docXmlFile = newZip.file('word/document.xml');
    if (docXmlFile) {
      let docXml = await docXmlFile.async('text');
      // Check if there are any comment markers at all
      if (docXml.includes('commentRangeStart') || docXml.includes('commentReference')) {
        let changed = false;
        // Remove commentRangeStart/End/Reference for IDs not in our comments array
        docXml = docXml.replace(/<w:commentRangeStart\s+w:id="(\d+)"\/>/g, (m, id) =>
          validIds.has(parseInt(id, 10)) ? m : ((changed = true), '')
        );
        docXml = docXml.replace(/<w:commentRangeEnd\s+w:id="(\d+)"\/>/g, (m, id) =>
          validIds.has(parseInt(id, 10)) ? m : ((changed = true), '')
        );
        docXml = docXml.replace(
          /<w:r><w:rPr><w:rStyle\s+w:val="CommentReference"\/><\/w:rPr><w:commentReference\s+w:id="(\d+)"\/><\/w:r>/g,
          (m, id) => (validIds.has(parseInt(id, 10)) ? m : ((changed = true), ''))
        );
        if (changed) {
          newZip.file('word/document.xml', docXml, {
            compression: 'DEFLATE',
            compressionOptions: { level: compressionLevel },
          });
        }
      }
    }
    // Remove stale comments.xml if no comments exist
    if (comments.length === 0) {
      if (newZip.file('word/comments.xml')) {
        newZip.remove('word/comments.xml');
      }
    }
    if (hasModifiedComments && comments.length > 0) {
      const commentsXml = serializeCommentsXml(comments);
      newZip.file('word/comments.xml', commentsXml, {
        compression: 'DEFLATE',
        compressionOptions: { level: compressionLevel },
      });
      // Ensure [Content_Types].xml has comments entry
      const ctFile = newZip.file('[Content_Types].xml');
      if (ctFile) {
        let ctXml = await ctFile.async('text');
        if (!ctXml.includes('word/comments.xml')) {
          ctXml = ctXml.replace(
            '</Types>',
            '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>'
          );
          newZip.file('[Content_Types].xml', ctXml, {
            compression: 'DEFLATE',
            compressionOptions: { level: compressionLevel },
          });
        }
      }
      // Ensure document.xml.rels has comments relationship
      const docRelsFile = newZip.file('word/_rels/document.xml.rels');
      if (docRelsFile) {
        let docRels = await docRelsFile.async('text');
        if (!docRels.includes('comments.xml')) {
          const existingIds = [...docRels.matchAll(/Id="rId(\d+)"/g)].map((m) =>
            parseInt(m[1], 10)
          );
          const nextId = existingIds.length > 0 ? Math.max(...existingIds) + 1 : 100;
          docRels = docRels.replace(
            '</Relationships>',
            `<Relationship Id="rId${nextId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>`
          );
          newZip.file('word/_rels/document.xml.rels', docRels, {
            compression: 'DEFLATE',
            compressionOptions: { level: compressionLevel },
          });
        }
      }
    }
  }

  // ── Write FP metadata as a proper OOXML Custom XML Data Store item ──
  // Word requires: itemN.xml + itemPropsN.xml + _rels/itemN.xml.rels + relationships
  // Without this structure, Word drops our Custom XML Part on save.
  const hasTagMeta = doc.contextTagMetadata && Object.keys(doc.contextTagMetadata).length > 0;
  const hasDocMeta = doc.fpDocumentMeta && Object.keys(doc.fpDocumentMeta).length > 0;
  const hasLoopMeta = doc.loopMetadata && Object.keys(doc.loopMetadata).length > 0;
  if (hasTagMeta || hasDocMeta || hasLoopMeta) {
    const metaXml = serializeManifest(
      doc.fpDocumentMeta,
      doc.contextTagMetadata,
      hasLoopMeta ? doc.loopMetadata : undefined
    );

    // Find next available item number (avoid colliding with existing customXml items)
    let itemNum = 1;
    while (newZip.file(`customXml/item${itemNum}.xml`)) {
      // Check if this item is ours (has FPMetadata) — if so, reuse the slot
      const existing = await newZip.file(`customXml/item${itemNum}.xml`)!.async('text');
      if (existing.includes('FPMetadata') || existing.includes('financialportal.app')) {
        break; // Reuse this slot
      }
      itemNum++;
    }

    const itemPath = `customXml/item${itemNum}.xml`;
    const propsPath = `customXml/itemProps${itemNum}.xml`;
    const relsPath = `customXml/_rels/item${itemNum}.xml.rels`;

    // 1. Write the data part
    newZip.file(itemPath, metaXml, {
      compression: 'DEFLATE',
      compressionOptions: { level: compressionLevel },
    });

    // 2. Write itemProps (datastore GUID identifies our schema)
    const itemPropsXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="no"?>` +
      `<ds:datastoreItem ds:itemID="${FP_DATASTORE_GUID}" ` +
      `xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">` +
      `<ds:schemaRefs><ds:schemaRef ds:uri="http://financialportal.app/fpMeta"/></ds:schemaRefs>` +
      `</ds:datastoreItem>`;
    newZip.file(propsPath, itemPropsXml, {
      compression: 'DEFLATE',
      compressionOptions: { level: compressionLevel },
    });

    // 3. Write rels (points itemProps to the data item)
    const relsXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" ` +
      `Target="itemProps${itemNum}.xml"/></Relationships>`;
    newZip.file(relsPath, relsXml, {
      compression: 'DEFLATE',
      compressionOptions: { level: compressionLevel },
    });

    // Remove legacy files if present
    if (newZip.file(CUSTOM_XML_LEGACY_PATH)) {
      newZip.remove(CUSTOM_XML_LEGACY_PATH);
    }
    if (newZip.file('customXml/contextTagMeta.xml')) {
      newZip.remove('customXml/contextTagMeta.xml');
    }

    // 4. Update [Content_Types].xml
    const ctFile = newZip.file('[Content_Types].xml');
    if (ctFile) {
      let ctXml = await ctFile.async('text');
      // Remove legacy entries
      ctXml = ctXml.replace(
        /<Override[^>]*PartName="\/customXml\/contextTagMeta\.xml"[^>]*\/>/g,
        ''
      );
      ctXml = ctXml.replace(/<Override[^>]*PartName="\/customXml\/fpMeta\.xml"[^>]*\/>/g, '');
      // Add itemProps content type if not present
      if (!ctXml.includes(propsPath)) {
        ctXml = ctXml.replace(
          '</Types>',
          `<Override PartName="/${propsPath}" ContentType="${CUSTOM_XML_PROPS_CONTENT_TYPE}"/></Types>`
        );
      }
      newZip.file('[Content_Types].xml', ctXml, {
        compression: 'DEFLATE',
        compressionOptions: { level: compressionLevel },
      });
    }

    // 5. Update word/_rels/document.xml.rels — add relationship to our custom XML item
    const docRelsFile = newZip.file('word/_rels/document.xml.rels');
    if (docRelsFile) {
      let docRels = await docRelsFile.async('text');
      const relTarget = `../customXml/item${itemNum}.xml`;
      if (!docRels.includes(relTarget)) {
        // Find next available rId
        const existingIds = [...docRels.matchAll(/Id="rId(\d+)"/g)].map((m) => parseInt(m[1], 10));
        const nextId = existingIds.length > 0 ? Math.max(...existingIds) + 1 : 100;
        docRels = docRels.replace(
          '</Relationships>',
          `<Relationship Id="rId${nextId}" ` +
            `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" ` +
            `Target="${relTarget}"/></Relationships>`
        );
        newZip.file('word/_rels/document.xml.rels', docRels, {
          compression: 'DEFLATE',
          compressionOptions: { level: compressionLevel },
        });
      }
    }
  }

  // If document has locked paragraphs, inject w:documentProtection into settings.xml
  // so Word enforces editing restrictions (only unlocked regions marked with permStart/permEnd are editable)
  if (doc.contentDirty && documentHasLockedParagraphs(doc)) {
    const settingsFile = newZip.file('word/settings.xml');
    if (settingsFile) {
      let settingsXml = await settingsFile.async('text');
      // Remove any existing documentProtection
      settingsXml = settingsXml.replace(/<w:documentProtection[^/]*\/>/g, '');
      // Insert documentProtection before closing </w:settings>
      settingsXml = settingsXml.replace(
        '</w:settings>',
        '<w:documentProtection w:edit="readOnly" w:enforcement="1"/></w:settings>'
      );
      newZip.file('word/settings.xml', settingsXml, {
        compression: 'DEFLATE',
        compressionOptions: { level: compressionLevel },
      });
    }
  } else if (doc.contentDirty) {
    // No locked paragraphs — remove any existing document protection
    const settingsFile = newZip.file('word/settings.xml');
    if (settingsFile) {
      let settingsXml = await settingsFile.async('text');
      if (settingsXml.includes('w:documentProtection')) {
        settingsXml = settingsXml.replace(/<w:documentProtection[^/]*\/>/g, '');
        newZip.file('word/settings.xml', settingsXml, {
          compression: 'DEFLATE',
          compressionOptions: { level: compressionLevel },
        });
      }
    }
  }

  // Optionally update modification date in docProps/core.xml
  if (updateModifiedDate) {
    const corePropsPath = 'docProps/core.xml';
    const corePropsFile = originalZip.file(corePropsPath);

    if (corePropsFile) {
      const originalCoreProps = await corePropsFile.async('text');
      const updatedCoreProps = updateCoreProperties(originalCoreProps, {
        updateModifiedDate,
        modifiedBy,
      });

      newZip.file(corePropsPath, updatedCoreProps, {
        compression: 'DEFLATE',
        compressionOptions: { level: compressionLevel },
      });
    }
  }

  // Generate the new DOCX file
  const arrayBuffer = await newZip.generateAsync({
    type: 'arraybuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: compressionLevel },
  });

  return arrayBuffer;
}

/**
 * Repack a Document using raw content for more control
 *
 * @param doc - Document with modified content
 * @param rawContent - Original raw content from unzipDocx
 * @param options - Optional repack options
 * @returns Promise resolving to DOCX as ArrayBuffer
 */
export async function repackDocxFromRaw(
  doc: Document,
  rawContent: RawDocxContent,
  options: RepackOptions = {}
): Promise<ArrayBuffer> {
  const { compressionLevel = 6, updateModifiedDate = true, modifiedBy } = options;

  // Create a new ZIP with all original files
  const newZip = new JSZip();

  // Copy all files from original ZIP
  for (const [path, file] of Object.entries(rawContent.originalZip.files)) {
    // Skip directories — OOXML ZIPs should not have directory entries.
    // JSZip synthesizes directory entries from file paths; re-creating them
    // adds entries that Word may flag as invalid.
    if (file.dir) {
      continue;
    }

    // Get original file content
    const content = await file.async('arraybuffer');

    // Add to new ZIP
    newZip.file(path, content, {
      compression: 'DEFLATE',
      compressionOptions: { level: compressionLevel },
    });
  }

  // Serialize and update document.xml
  const documentXml = serializeDocument(doc);
  newZip.file('word/document.xml', documentXml, {
    compression: 'DEFLATE',
    compressionOptions: { level: compressionLevel },
  });

  // Apply context tag replacements to header/footer XML (same as repackDocx)
  if (doc.contextTagReplacements) {
    const { tags, mode } = doc.contextTagReplacements;
    const ctMeta = doc.contextTagMetadata;

    // Build tagKey → metaId lookup for H/F bookmark generation
    const tagKeyToMetaId = new Map<string, string>();
    if (ctMeta) {
      for (const [metaId, meta] of Object.entries(ctMeta)) {
        if (meta.tagKey && !tagKeyToMetaId.has(meta.tagKey)) {
          tagKeyToMetaId.set(meta.tagKey, metaId);
        }
      }
    }

    // Regex matching a <w:r> containing a context tag in its <w:t>
    // Matches {{ tag.path }} (double brace, 0+ dots) and {tag.path} (single brace, 0+ dots)
    const runWithTagRe =
      /(<w:r\b[^>]*>(?:<w:rPr>[\s\S]*?<\/w:rPr>)?<w:t[^>]*>)([^<]*?)(\{\{\s*([\w]+(?:\.[\w]+)*)!?\s*\}\}|\{([\w]+(?:\.[\w]+)*)\})([^<]*?<\/w:t><\/w:r>)/g;
    const tagRe = /\{\{\s*([\w]+(?:\.[\w]+)*)\s*\}\}|\{([\w]+(?:\.[\w]+)+)\}/g;
    let hfBookmarkId = 10000;

    for (const [path, file] of Object.entries(newZip.files)) {
      if (file.dir) continue;
      if (!/^word\/(header|footer)\d*\.xml$/i.test(path)) continue;
      let xml = await file.async('text');
      let changed = false;

      // First pass: run-level replacement with bookmarks
      xml = xml.replace(
        runWithTagRe,
        (fullMatch, beforeTag, preText, _tagMatch, g4, g5, afterTag) => {
          const rawKey = g4 || g5;
          // Strip legacy "context." prefix so lookup matches the new flat tag map
          const tagKey = rawKey.startsWith('context.') ? rawKey.slice(8) : rawKey;
          const resolved = tags[tagKey];
          const metaId = tagKeyToMetaId.get(tagKey);
          if (resolved) {
            changed = true;
            const renderedRun = `${beforeTag}${preText}${resolved}${afterTag}`;
            if (metaId) {
              const bmStart = `<w:bookmarkStart w:id="${hfBookmarkId}" w:name="${FP_BOOKMARK_PREFIX}${metaId}"/>`;
              const bmEnd = `<w:bookmarkEnd w:id="${hfBookmarkId}"/>`;
              hfBookmarkId++;
              return `${bmStart}${renderedRun}${bmEnd}`;
            }
            return renderedRun;
          }
          if (mode === 'omit') {
            changed = true;
            return `${beforeTag}${preText}${afterTag}`;
          }
          return fullMatch;
        }
      );

      // Second pass: catch remaining tags not in own run (fallback)
      xml = xml.replace(tagRe, (_match, g1, g2) => {
        const rawKey = g1 || g2;
        const tagKey = rawKey.startsWith('context.') ? rawKey.slice(8) : rawKey;
        const resolved = tags[tagKey];
        if (resolved) {
          changed = true;
          return resolved;
        }
        if (mode === 'omit') {
          changed = true;
          return '';
        }
        return _match;
      });

      if (changed) {
        newZip.file(path, xml, {
          compression: 'DEFLATE',
          compressionOptions: { level: compressionLevel },
        });
      }
    }
  }

  // Write FP metadata as proper OOXML Custom XML (same logic as contentDirty path above)
  const hasTagMetaRaw = doc.contextTagMetadata && Object.keys(doc.contextTagMetadata).length > 0;
  const hasDocMetaRaw = doc.fpDocumentMeta && Object.keys(doc.fpDocumentMeta).length > 0;
  const hasLoopMetaRaw = doc.loopMetadata && Object.keys(doc.loopMetadata).length > 0;
  if (hasTagMetaRaw || hasDocMetaRaw || hasLoopMetaRaw) {
    const metaXml = serializeManifest(
      doc.fpDocumentMeta,
      doc.contextTagMetadata,
      hasLoopMetaRaw ? doc.loopMetadata : undefined
    );

    let itemNum = 1;
    while (newZip.file(`customXml/item${itemNum}.xml`)) {
      const existing = await newZip.file(`customXml/item${itemNum}.xml`)!.async('text');
      if (existing.includes('FPMetadata') || existing.includes('financialportal.app')) break;
      itemNum++;
    }

    const itemPath = `customXml/item${itemNum}.xml`;
    const propsPath = `customXml/itemProps${itemNum}.xml`;
    const relsPath = `customXml/_rels/item${itemNum}.xml.rels`;

    newZip.file(itemPath, metaXml, {
      compression: 'DEFLATE',
      compressionOptions: { level: compressionLevel },
    });

    const itemPropsXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="no"?>` +
      `<ds:datastoreItem ds:itemID="${FP_DATASTORE_GUID}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">` +
      `<ds:schemaRefs><ds:schemaRef ds:uri="http://financialportal.app/fpMeta"/></ds:schemaRefs></ds:datastoreItem>`;
    newZip.file(propsPath, itemPropsXml, {
      compression: 'DEFLATE',
      compressionOptions: { level: compressionLevel },
    });

    const relsXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps${itemNum}.xml"/></Relationships>`;
    newZip.file(relsPath, relsXml, {
      compression: 'DEFLATE',
      compressionOptions: { level: compressionLevel },
    });

    // Remove legacy files
    if (newZip.file(CUSTOM_XML_LEGACY_PATH)) newZip.remove(CUSTOM_XML_LEGACY_PATH);
    if (newZip.file('customXml/contextTagMeta.xml')) newZip.remove('customXml/contextTagMeta.xml');

    const ctFile = newZip.file('[Content_Types].xml');
    if (ctFile) {
      let ctXml = await ctFile.async('text');
      ctXml = ctXml.replace(
        /<Override[^>]*PartName="\/customXml\/contextTagMeta\.xml"[^>]*\/>/g,
        ''
      );
      ctXml = ctXml.replace(/<Override[^>]*PartName="\/customXml\/fpMeta\.xml"[^>]*\/>/g, '');
      if (!ctXml.includes(propsPath)) {
        ctXml = ctXml.replace(
          '</Types>',
          `<Override PartName="/${propsPath}" ContentType="${CUSTOM_XML_PROPS_CONTENT_TYPE}"/></Types>`
        );
      }
      newZip.file('[Content_Types].xml', ctXml, {
        compression: 'DEFLATE',
        compressionOptions: { level: compressionLevel },
      });
    }

    const docRelsFile = newZip.file('word/_rels/document.xml.rels');
    if (docRelsFile) {
      let docRels = await docRelsFile.async('text');
      const relTarget = `../customXml/item${itemNum}.xml`;
      if (!docRels.includes(relTarget)) {
        const existingIds = [...docRels.matchAll(/Id="rId(\d+)"/g)].map((m) => parseInt(m[1], 10));
        const nextId = existingIds.length > 0 ? Math.max(...existingIds) + 1 : 100;
        docRels = docRels.replace(
          '</Relationships>',
          `<Relationship Id="rId${nextId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="${relTarget}"/></Relationships>`
        );
        newZip.file('word/_rels/document.xml.rels', docRels, {
          compression: 'DEFLATE',
          compressionOptions: { level: compressionLevel },
        });
      }
    }
  }

  // Optionally update core properties
  if (updateModifiedDate && rawContent.corePropsXml) {
    const updatedCoreProps = updateCoreProperties(rawContent.corePropsXml, {
      updateModifiedDate,
      modifiedBy,
    });

    newZip.file('docProps/core.xml', updatedCoreProps, {
      compression: 'DEFLATE',
      compressionOptions: { level: compressionLevel },
    });
  }

  // Generate the new DOCX file
  const arrayBuffer = await newZip.generateAsync({
    type: 'arraybuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: compressionLevel },
  });

  return arrayBuffer;
}

// ============================================================================
// SELECTIVE UPDATES
// ============================================================================

/**
 * Update only document.xml in a DOCX buffer (minimal changes)
 *
 * @param originalBuffer - Original DOCX as ArrayBuffer
 * @param newDocumentXml - New document.xml content
 * @param options - Optional repack options
 * @returns Promise resolving to DOCX as ArrayBuffer
 */
export async function updateDocumentXml(
  originalBuffer: ArrayBuffer,
  newDocumentXml: string,
  options: RepackOptions = {}
): Promise<ArrayBuffer> {
  const { compressionLevel = 6 } = options;

  // Load original ZIP
  const zip = await JSZip.loadAsync(originalBuffer);

  // Update document.xml
  zip.file('word/document.xml', newDocumentXml, {
    compression: 'DEFLATE',
    compressionOptions: { level: compressionLevel },
  });

  // Generate new DOCX
  return zip.generateAsync({
    type: 'arraybuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: compressionLevel },
  });
}

/**
 * Update a specific XML file in a DOCX buffer
 *
 * @param originalBuffer - Original DOCX as ArrayBuffer
 * @param path - Path within the ZIP (e.g., "word/styles.xml")
 * @param content - New XML content
 * @param options - Optional repack options
 * @returns Promise resolving to DOCX as ArrayBuffer
 */
export async function updateXmlFile(
  originalBuffer: ArrayBuffer,
  path: string,
  content: string,
  options: RepackOptions = {}
): Promise<ArrayBuffer> {
  const { compressionLevel = 6 } = options;

  const zip = await JSZip.loadAsync(originalBuffer);

  zip.file(path, content, {
    compression: 'DEFLATE',
    compressionOptions: { level: compressionLevel },
  });

  return zip.generateAsync({
    type: 'arraybuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: compressionLevel },
  });
}

/**
 * Update multiple files in a DOCX buffer
 *
 * @param originalBuffer - Original DOCX as ArrayBuffer
 * @param updates - Map of path -> content for files to update
 * @param options - Optional repack options
 * @returns Promise resolving to DOCX as ArrayBuffer
 */
export async function updateMultipleFiles(
  originalBuffer: ArrayBuffer,
  updates: Map<string, string | ArrayBuffer>,
  options: RepackOptions = {}
): Promise<ArrayBuffer> {
  const { compressionLevel = 6 } = options;

  const zip = await JSZip.loadAsync(originalBuffer);

  for (const [path, content] of updates) {
    zip.file(path, content, {
      compression: 'DEFLATE',
      compressionOptions: { level: compressionLevel },
    });
  }

  return zip.generateAsync({
    type: 'arraybuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: compressionLevel },
  });
}

// ============================================================================
// RELATIONSHIP MANAGEMENT
// ============================================================================

/**
 * Add a new relationship to document.xml.rels
 *
 * @param originalBuffer - Original DOCX as ArrayBuffer
 * @param relationship - New relationship to add
 * @returns Promise resolving to { buffer: ArrayBuffer, rId: string }
 */
export async function addRelationship(
  originalBuffer: ArrayBuffer,
  relationship: {
    type: string;
    target: string;
    targetMode?: 'External' | 'Internal';
  }
): Promise<{ buffer: ArrayBuffer; rId: string }> {
  const zip = await JSZip.loadAsync(originalBuffer);

  // Read existing relationships
  const relsPath = 'word/_rels/document.xml.rels';
  const relsFile = zip.file(relsPath);

  if (!relsFile) {
    throw new Error('document.xml.rels not found in DOCX');
  }

  const relsXml = await relsFile.async('text');

  // Find highest existing rId
  const rIdMatches = relsXml.matchAll(/Id="rId(\d+)"/g);
  let maxId = 0;
  for (const match of rIdMatches) {
    const id = parseInt(match[1], 10);
    if (id > maxId) maxId = id;
  }

  // Generate new rId
  const newRId = `rId${maxId + 1}`;

  // Build new relationship element
  const targetModeAttr = relationship.targetMode === 'External' ? ' TargetMode="External"' : '';

  const newRelElement = `<Relationship Id="${newRId}" Type="${relationship.type}" Target="${escapeXmlAttr(relationship.target)}"${targetModeAttr}/>`;

  // Insert before closing tag
  const updatedRelsXml = relsXml.replace('</Relationships>', `${newRelElement}</Relationships>`);

  // Update the ZIP
  zip.file(relsPath, updatedRelsXml);

  const buffer = await zip.generateAsync({
    type: 'arraybuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });

  return { buffer, rId: newRId };
}

/**
 * Add a media file to the DOCX
 *
 * @param originalBuffer - Original DOCX as ArrayBuffer
 * @param filename - Filename for the media (e.g., "image1.png")
 * @param data - Binary data for the media file
 * @param mimeType - MIME type (e.g., "image/png")
 * @returns Promise resolving to { buffer: ArrayBuffer, rId: string, path: string }
 */
export async function addMedia(
  originalBuffer: ArrayBuffer,
  filename: string,
  data: ArrayBuffer,
  mimeType: string
): Promise<{ buffer: ArrayBuffer; rId: string; path: string }> {
  const zip = await JSZip.loadAsync(originalBuffer);

  // Determine media path
  const mediaPath = `word/media/${filename}`;

  // Add media file
  zip.file(mediaPath, data);

  // Add relationship
  const relResult = await addRelationship(await zip.generateAsync({ type: 'arraybuffer' }), {
    type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
    target: `media/${filename}`,
  });

  // Update content types if needed
  const contentTypesFile = zip.file('[Content_Types].xml');
  if (contentTypesFile) {
    const contentTypesXml = await contentTypesFile.async('text');
    const extension = filename.split('.').pop()?.toLowerCase() || '';

    // Check if extension is already registered
    const hasExtension = contentTypesXml.includes(`Extension="${extension}"`);

    if (!hasExtension && extension) {
      // Add content type for this extension
      const contentType = getContentTypeForExtension(extension, mimeType);
      const extensionElement = `<Default Extension="${extension}" ContentType="${contentType}"/>`;

      // Insert after other defaults
      const updatedContentTypes = contentTypesXml.replace(
        '</Types>',
        `${extensionElement}</Types>`
      );

      const finalZip = await JSZip.loadAsync(relResult.buffer);
      finalZip.file('[Content_Types].xml', updatedContentTypes);

      return {
        buffer: await finalZip.generateAsync({
          type: 'arraybuffer',
          compression: 'DEFLATE',
          compressionOptions: { level: 6 },
        }),
        rId: relResult.rId,
        path: mediaPath,
      };
    }
  }

  return {
    buffer: relResult.buffer,
    rId: relResult.rId,
    path: mediaPath,
  };
}

// ============================================================================
// HEADER/FOOTER SERIALIZATION
// ============================================================================

/**
 * Serialize modified headers and footers into the ZIP
 *
 * Maps rId → filename via relationships, then serializes each
 * HeaderFooter object to its corresponding word/header*.xml or word/footer*.xml
 */
// @ts-expect-error Kept for reference but no longer called — our serializer is lossy for headers/footers.
function _serializeHeadersFootersToZip(doc: Document, zip: JSZip, compressionLevel: number): void {
  const rels = doc.package.relationships;
  if (!rels) return;

  const compressionOptions = { level: compressionLevel };

  // Serialize headers
  if (doc.package.headers) {
    for (const [rId, headerFooter] of doc.package.headers.entries()) {
      const rel = rels.get(rId);
      if (rel && rel.type === RELATIONSHIP_TYPES.header && rel.target) {
        const filename = rel.target.startsWith('/') ? rel.target.slice(1) : `word/${rel.target}`;
        const xml = serializeHeaderFooter(headerFooter);
        zip.file(filename, xml, { compression: 'DEFLATE', compressionOptions });
      }
    }
  }

  // Serialize footers
  if (doc.package.footers) {
    for (const [rId, headerFooter] of doc.package.footers.entries()) {
      const rel = rels.get(rId);
      if (rel && rel.type === RELATIONSHIP_TYPES.footer && rel.target) {
        const filename = rel.target.startsWith('/') ? rel.target.slice(1) : `word/${rel.target}`;
        const xml = serializeHeaderFooter(headerFooter);
        zip.file(filename, xml, { compression: 'DEFLATE', compressionOptions });
      }
    }
  }
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Update core properties XML with new modification date
 */
function updateCoreProperties(
  corePropsXml: string,
  options: { updateModifiedDate?: boolean; modifiedBy?: string }
): string {
  let result = corePropsXml;

  if (options.updateModifiedDate) {
    const now = new Date().toISOString();

    // Update dcterms:modified
    if (result.includes('<dcterms:modified')) {
      result = result.replace(
        /<dcterms:modified[^>]*>[^<]*<\/dcterms:modified>/,
        `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>`
      );
    } else {
      // Add modified date if not present
      result = result.replace(
        '</cp:coreProperties>',
        `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`
      );
    }
  }

  if (options.modifiedBy) {
    // Update cp:lastModifiedBy
    if (result.includes('<cp:lastModifiedBy')) {
      result = result.replace(
        /<cp:lastModifiedBy>[^<]*<\/cp:lastModifiedBy>/,
        `<cp:lastModifiedBy>${escapeXmlText(options.modifiedBy)}</cp:lastModifiedBy>`
      );
    } else {
      // Add lastModifiedBy if not present
      result = result.replace(
        '</cp:coreProperties>',
        `<cp:lastModifiedBy>${escapeXmlText(options.modifiedBy)}</cp:lastModifiedBy></cp:coreProperties>`
      );
    }
  }

  return result;
}

/**
 * Escape special XML characters in text content
 */
function escapeXmlText(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
 * Escape special XML characters in attribute values
 */
function escapeXmlAttr(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Get content type for a file extension
 */
function getContentTypeForExtension(extension: string, mimeType: string): string {
  // Use provided mime type or fall back to common types
  if (mimeType) return mimeType;

  const contentTypes: Record<string, string> = {
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    bmp: 'image/bmp',
    tif: 'image/tiff',
    tiff: 'image/tiff',
    svg: 'image/svg+xml',
    webp: 'image/webp',
    wmf: 'image/x-wmf',
    emf: 'image/x-emf',
  };

  return contentTypes[extension] || 'application/octet-stream';
}

// ============================================================================
// VALIDATION
// ============================================================================

/**
 * Validate that a buffer is a valid DOCX file
 *
 * @param buffer - Buffer to validate
 * @returns Promise resolving to validation result
 */
export async function validateDocx(buffer: ArrayBuffer): Promise<{
  valid: boolean;
  errors: string[];
  warnings: string[];
}> {
  const errors: string[] = [];
  const warnings: string[] = [];

  try {
    const zip = await JSZip.loadAsync(buffer);

    // Check for required files
    const requiredFiles = ['[Content_Types].xml', 'word/document.xml'];

    for (const file of requiredFiles) {
      if (!zip.file(file)) {
        errors.push(`Missing required file: ${file}`);
      }
    }

    // Check for recommended files
    const recommendedFiles = ['_rels/.rels', 'word/_rels/document.xml.rels', 'word/styles.xml'];

    for (const file of recommendedFiles) {
      if (!zip.file(file)) {
        warnings.push(`Missing recommended file: ${file}`);
      }
    }

    // Validate document.xml is valid XML
    const docFile = zip.file('word/document.xml');
    if (docFile) {
      const docXml = await docFile.async('text');

      // Basic XML validation
      if (!docXml.includes('<?xml')) {
        warnings.push('document.xml missing XML declaration');
      }

      if (!docXml.includes('<w:document')) {
        errors.push('document.xml missing w:document element');
      }

      if (!docXml.includes('<w:body>')) {
        errors.push('document.xml missing w:body element');
      }
    }

    // Validate Content_Types.xml
    const ctFile = zip.file('[Content_Types].xml');
    if (ctFile) {
      const ctXml = await ctFile.async('text');

      if (
        !ctXml.includes('word/document.xml') &&
        !ctXml.includes(
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
        )
      ) {
        warnings.push('Content_Types.xml may be missing document.xml type declaration');
      }
    }
  } catch (error) {
    errors.push(
      `Failed to read as ZIP: ${error instanceof Error ? error.message : 'Unknown error'}`
    );
  }

  return {
    valid: errors.length === 0,
    errors,
    warnings,
  };
}

/**
 * Check if buffer looks like a DOCX file (quick check)
 *
 * @param buffer - Buffer to check
 * @returns true if buffer starts with ZIP signature
 */
export function isDocxBuffer(buffer: ArrayBuffer): boolean {
  if (buffer.byteLength < 4) return false;

  const view = new Uint8Array(buffer);

  // ZIP file signature: PK (0x50, 0x4B)
  return view[0] === 0x50 && view[1] === 0x4b;
}

// ============================================================================
// CREATE NEW DOCX
// ============================================================================

/**
 * Create a new empty DOCX file
 *
 * @returns Promise resolving to minimal DOCX as ArrayBuffer
 */
export async function createEmptyDocx(): Promise<ArrayBuffer> {
  const zip = new JSZip();

  // Content Types
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`
  );

  // Package relationships
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`
  );

  // Document relationships
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
  );

  // Document
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:t></w:t>
      </w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`
  );

  // Minimal styles
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
        <w:sz w:val="22"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="200" w:line="276" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>`
  );

  // Core properties
  const now = new Date().toISOString();
  zip.file(
    'docProps/core.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>EigenPal DOCX Editor</dc:creator>
  <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`
  );

  // App properties
  zip.file(
    'docProps/app.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>EigenPal DOCX Editor</Application>
  <AppVersion>1.0.0</AppVersion>
</Properties>`
  );

  return zip.generateAsync({
    type: 'arraybuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
}

/**
 * Create a new DOCX from a Document (without requiring original buffer)
 *
 * @param doc - Document to serialize
 * @returns Promise resolving to DOCX as ArrayBuffer
 */
export async function createDocx(doc: Document): Promise<ArrayBuffer> {
  // Start with an empty DOCX
  const emptyBuffer = await createEmptyDocx();

  // Add document as original buffer
  const docWithBuffer: Document = {
    ...doc,
    originalBuffer: emptyBuffer,
  };

  // Repack with the document content
  return repackDocx(docWithBuffer);
}

// ============================================================================
// EXPORTS
// ============================================================================

export default repackDocx;
