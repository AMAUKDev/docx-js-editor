/**
 * Financial Portal Custom XML Part — OOXML ECMA-376
 *
 * Persists document-level metadata and context tag properties inside the DOCX zip
 * as a Custom XML Part at `customXml/fpMeta.xml`.
 *
 * Versioned JSON manifest wrapped in CDATA:
 * ```xml
 * <?xml version="1.0" encoding="UTF-8"?>
 * <FPMetadata xmlns="http://financialportal.app/fpMeta">
 *   <![CDATA[{
 *     "version": 1,
 *     "document": { "templateId": 42, "templateName": "AMA UK Survey", "tocStyle": "Heading1" },
 *     "tags": { "uuid-1": { "tagKey": "context.case_no", "removeIfEmpty": true } }
 *   }]]>
 * </FPMetadata>
 * ```
 *
 * Security:
 * - JSON parsed with try/catch; malformed data → empty defaults
 * - Size limit: rejects payloads > 1 MB
 * - Only primitive values kept per tag entry (no nested objects)
 * - No eval() or Function() — only JSON.parse
 */

import type { ContextTagMeta, FPDocumentMeta } from '../types/document';

/** Maximum allowed size of the JSON payload in bytes */
const MAX_PAYLOAD_SIZE = 1_048_576; // 1 MB

/** Current manifest version */
const MANIFEST_VERSION = 1;

/** Legacy path (kept for backward-compat reading) */
export const CUSTOM_XML_LEGACY_PATH = 'customXml/fpMeta.xml';

/** Content type for registration in [Content_Types].xml */
export const CUSTOM_XML_CONTENT_TYPE = 'application/xml';

/** Content type for itemProps registration */
export const CUSTOM_XML_PROPS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.customXmlProperties+xml';

/** XML namespace for our metadata */
const XMLNS = 'http://financialportal.app/fpMeta';

/**
 * Stable datastore GUID for our Custom XML Part.
 * Word uses this GUID to identify our schema across save/load cycles.
 * Using a fixed GUID ensures Word recognizes and preserves our data.
 */
export const FP_DATASTORE_GUID = '{B7E6D12F-4A8C-4D3E-9F1A-2C5B8E7D9F0A}';

// Re-export types for convenience
export type { ContextTagMeta, FPDocumentMeta };

// ============================================================================
// PARSED RESULT
// ============================================================================

/** Per-item rendered values stored for loop diff detection on re-upload. */
export interface FPLoopItemMeta {
  index: number;
  /** tagKey → rendered string value (e.g. { "photo.caption": "Damage to port bow" }) */
  renderedTags: Record<string, string>;
  /** tagKey → image info for image fields */
  renderedImages: Record<string, { caseFileId: number; width?: number; height?: number }>;
}

/** Metadata for a single loop block, stored in the manifest. */
export interface FPLoopMeta {
  loopExpr: string; // e.g. "photo in photos"
  collectionKey: string; // e.g. "photos"
  itemVar: string; // e.g. "photo"
  templateXml: string; // Original <w:tr> XML with {{ tags }} intact
  items: FPLoopItemMeta[];
}

export interface FPManifest {
  document: FPDocumentMeta;
  tags: Record<string, ContextTagMeta>;
  loops?: Record<string, FPLoopMeta>;
}

// ============================================================================
// PARSING (read from DOCX)
// ============================================================================

/**
 * Parse the FP metadata manifest from the Custom XML Part.
 *
 * @param xml - Raw XML string from customXml/fpMeta.xml (may be null)
 * @returns Parsed manifest with document metadata and tag metadata
 */
export function parseManifest(xml: string | null): FPManifest {
  const empty: FPManifest = { document: {}, tags: {} };
  if (!xml) return empty;

  try {
    // Extract JSON from CDATA section
    const cdataMatch = xml.match(/<!\[CDATA\[([\s\S]*?)\]\]>/);
    let jsonStr: string | null = null;
    if (cdataMatch) {
      jsonStr = cdataMatch[1];
    } else {
      // Fallback: text content between root element tags
      const textMatch = xml.match(
        /<(?:FPMetadata|ContextTagMetadata)[^>]*>([\s\S]*?)<\/(?:FPMetadata|ContextTagMetadata)>/
      );
      if (textMatch) jsonStr = textMatch[1].trim();
    }
    if (!jsonStr) return empty;

    return validateManifest(jsonStr);
  } catch {
    console.warn('[fpMeta] Failed to parse Custom XML Part');
    return empty;
  }
}

/**
 * Backward-compatible: parse tag-only metadata (old format, flat object keyed by metaId).
 */
function migrateLegacyFormat(raw: Record<string, unknown>): FPManifest {
  const tags: Record<string, ContextTagMeta> = {};
  for (const [key, value] of Object.entries(raw)) {
    if (typeof key !== 'string' || key.length === 0 || key.length > 100) continue;
    if (typeof value !== 'object' || value === null || Array.isArray(value)) continue;
    tags[key] = sanitizeTagEntry(value as Record<string, unknown>);
  }
  return { document: {}, tags };
}

function validateManifest(jsonStr: string): FPManifest {
  if (jsonStr.length > MAX_PAYLOAD_SIZE) {
    console.warn('[fpMeta] Payload exceeds size limit, ignoring');
    return { document: {}, tags: {} };
  }

  const raw = JSON.parse(jsonStr);
  if (typeof raw !== 'object' || raw === null || Array.isArray(raw)) {
    return { document: {}, tags: {} };
  }

  // Detect format: versioned manifest vs legacy flat tags
  if (!('version' in raw)) {
    return migrateLegacyFormat(raw as Record<string, unknown>);
  }

  // Versioned manifest
  const result: FPManifest = { document: {}, tags: {} };

  // Parse document metadata
  if (raw.document && typeof raw.document === 'object' && !Array.isArray(raw.document)) {
    const d = raw.document as Record<string, unknown>;
    if (typeof d.templateId === 'number') result.document.templateId = d.templateId;
    if (typeof d.templateName === 'string') result.document.templateName = d.templateName;
    if (typeof d.tocStyle === 'string') result.document.tocStyle = d.tocStyle;
  }

  // Parse tag metadata
  if (raw.tags && typeof raw.tags === 'object' && !Array.isArray(raw.tags)) {
    for (const [key, value] of Object.entries(raw.tags as Record<string, unknown>)) {
      if (typeof key !== 'string' || key.length === 0 || key.length > 100) continue;
      if (typeof value !== 'object' || value === null || Array.isArray(value)) continue;
      result.tags[key] = sanitizeTagEntry(value as Record<string, unknown>);
    }
  }

  // Parse loop metadata (pass through as-is — complex nested structure)
  if (raw.loops && typeof raw.loops === 'object' && !Array.isArray(raw.loops)) {
    result.loops = raw.loops as Record<string, FPLoopMeta>;
  }

  return result;
}

function sanitizeTagEntry(obj: Record<string, unknown>): ContextTagMeta {
  const sanitized: ContextTagMeta = {};
  for (const [prop, val] of Object.entries(obj)) {
    if (typeof val === 'boolean' || typeof val === 'string' || typeof val === 'number') {
      sanitized[prop] = val;
    }
    // Skip nested objects/arrays
  }
  return sanitized;
}

// ============================================================================
// SERIALIZATION (write to DOCX)
// ============================================================================

/**
 * Serialize the full manifest to the Custom XML Part XML string.
 */
export function serializeManifest(
  documentMeta: FPDocumentMeta | undefined,
  tags: Record<string, ContextTagMeta> | undefined,
  loops?: Record<string, FPLoopMeta>
): string {
  const manifest: Record<string, unknown> = {
    version: MANIFEST_VERSION,
    document: documentMeta || {},
    tags: tags || {},
  };
  if (loops && Object.keys(loops).length > 0) {
    manifest.loops = loops;
  }
  const json = JSON.stringify(manifest);
  return (
    `<?xml version="1.0" encoding="UTF-8"?>\n` +
    `<FPMetadata xmlns="${XMLNS}"><![CDATA[${json}]]></FPMetadata>`
  );
}

// Keep old name as alias for existing call sites
export { serializeManifest as serializeContextTagMetadata };

// ============================================================================
// PM DOC COLLECTION (extract tag metadata from ProseMirror doc)
// ============================================================================

/**
 * Walk a ProseMirror document and collect metadata for all contextTag nodes.
 * Each entry is keyed by the node's unique `metaId` (UUID).
 */
export function collectContextTagMetadata(doc: {
  descendants: (
    callback: (node: { type: { name: string }; attrs: Record<string, unknown> }) => boolean | void
  ) => void;
}): Record<string, ContextTagMeta> {
  const result: Record<string, ContextTagMeta> = {};

  doc.descendants((node) => {
    if (node.type.name === 'contextTag') {
      const metaId = node.attrs.metaId as string;
      const tagKey = node.attrs.tagKey as string;
      if (metaId && tagKey) {
        const meta: ContextTagMeta = {
          tagKey,
          removeIfEmpty: !!node.attrs.removeIfEmpty,
          removeTableRow: !!node.attrs.removeTableRow,
        };
        if (node.attrs.imageWidth) meta.imageWidth = node.attrs.imageWidth as number;
        if (node.attrs.alwaysShow) meta.alwaysShow = true;
        result[metaId] = meta;
      }
    }
  });

  return result;
}

// ============================================================================
// LEGACY COMPAT — keep old function name working
// ============================================================================

/** @deprecated Use parseManifest instead */
export function parseContextTagMetadata(xml: string | null): Record<string, ContextTagMeta> {
  return parseManifest(xml).tags;
}
