/**
 * Comprehensive TypeScript types for full DOCX document representation
 *
 * This barrel file re-exports all types from the split modules.
 * Existing imports from './types/document' continue to work unchanged.
 *
 * Module structure:
 * - colors.ts      — Color primitives, borders, shading
 * - formatting.ts  — Text, paragraph, and table formatting properties
 * - lists.ts       — Numbering and list definitions
 * - content.ts     — Content model (runs, images, shapes, tables, paragraphs, sections)
 * - styles.ts      — Styles, theme, fonts, relationships, media
 */

// Color & Styling Primitives
export type { ThemeColorSlot, ColorValue, BorderSpec, ShadingProperties } from './colors';

// Text & Paragraph Formatting
export type {
  UnderlineStyle,
  TextEffect,
  EmphasisMark,
  TextFormatting,
  TabStopAlignment,
  TabLeader,
  TabStop,
  LineSpacingRule,
  ParagraphAlignment,
  ParagraphFormatting,
  TableWidthType,
  TableMeasurement,
  TableBorders,
  CellMargins,
  TableLook,
  FloatingTableProperties,
  TableFormatting,
  TableRowFormatting,
  ConditionalFormatStyle,
  TableCellFormatting,
} from './formatting';

// Lists & Numbering
export type {
  NumberFormat,
  LevelSuffix,
  ListLevel,
  AbstractNumbering,
  NumberingInstance,
  ListRendering,
  NumberingDefinitions,
} from './lists';

// Content Model
export type {
  TextContent,
  TabContent,
  BreakContent,
  SymbolContent,
  NoteReferenceContent,
  FieldCharContent,
  InstrTextContent,
  SoftHyphenContent,
  NoBreakHyphenContent,
  DrawingContent,
  ShapeContent,
  RunContent,
  Run,
  Hyperlink,
  BookmarkStart,
  BookmarkEnd,
  FieldType,
  SimpleField,
  ComplexField,
  Field,
  ImageSize,
  ImageWrap,
  ImagePosition,
  ImageTransform,
  ImagePadding,
  Image,
  ShapeType,
  ShapeFill,
  ShapeOutline,
  ShapeTextBody,
  Shape,
  TextBox,
  TableCell,
  TableRow,
  Table,
  Comment,
  CommentRangeStart,
  CommentRangeEnd,
  MathEquation,
  TrackedChangeInfo,
  Insertion,
  Deletion,
  SdtType,
  SdtProperties,
  InlineSdt,
  BlockSdt,
  ParagraphContent,
  Paragraph,
  HeaderFooterType,
  HeaderReference,
  FooterReference,
  HeaderFooter,
  FootnotePosition,
  EndnotePosition,
  NoteNumberRestart,
  FootnoteProperties,
  EndnoteProperties,
  Footnote,
  Endnote,
  PageOrientation,
  SectionStart,
  VerticalAlign,
  LineNumberRestart,
  Column,
  SectionProperties,
  BlockContent,
  Section,
  DocumentBody,
} from './content';

// Styles, Theme, Fonts, Relationships & Media
export type {
  StyleType,
  Style,
  DocDefaults,
  StyleDefinitions,
  ThemeColorScheme,
  ThemeFont,
  ThemeFontScheme,
  Theme,
  FontInfo,
  FontTable,
  RelationshipType,
  Relationship,
  RelationshipMap,
  MediaFile,
} from './styles';

// ============================================================================
// DOCX PACKAGE & TOP-LEVEL DOCUMENT
// ============================================================================

import type { DocumentBody } from './content';
import type { StyleDefinitions, Theme, FontTable, RelationshipMap, MediaFile } from './styles';
import type { NumberingDefinitions } from './lists';
import type { Footnote, Endnote, HeaderFooter } from './content';

/**
 * Complete DOCX package structure
 */
export interface DocxPackage {
  /** Document body */
  document: DocumentBody;
  /** Style definitions */
  styles?: StyleDefinitions;
  /** Theme */
  theme?: Theme;
  /** Numbering definitions */
  numbering?: NumberingDefinitions;
  /** Font table */
  fontTable?: FontTable;
  /** Footnotes */
  footnotes?: Footnote[];
  /** Endnotes */
  endnotes?: Endnote[];
  /** Headers by relationship ID */
  headers?: Map<string, HeaderFooter>;
  /** Footers by relationship ID */
  footers?: Map<string, HeaderFooter>;
  /** Document relationships */
  relationships?: RelationshipMap;
  /** Media files */
  media?: Map<string, MediaFile>;
  /** Document properties */
  properties?: {
    title?: string;
    subject?: string;
    creator?: string;
    keywords?: string;
    description?: string;
    lastModifiedBy?: string;
    revision?: number;
    created?: Date;
    modified?: Date;
  };
}

/**
 * Complete parsed DOCX document
 */
export interface Document {
  /** DOCX package with all parsed content */
  package: DocxPackage;
  /** Original ArrayBuffer for round-trip */
  originalBuffer?: ArrayBuffer;
  /** Original word/document.xml string for lossless round-trip when unedited */
  originalDocumentXml?: string;
  /** True when document content has been modified via editing (set by fromProseDoc) */
  contentDirty?: boolean;
  /** Detected template variables ({{...}}) */
  templateVariables?: string[];
  /** Parsing warnings/errors */
  warnings?: string[];
  /** Context tag replacements to apply to header/footer XML during repack */
  contextTagReplacements?: {
    tags: Record<string, string>;
    mode: 'omit' | 'keep';
  };
  /**
   * Context tag metadata loaded from the Custom XML Part (customXml/fpMeta.xml).
   * Keyed by metaId (UUID). Each entry stores the tagKey and properties like removeIfEmpty.
   * Written on save, read on load, reconciled against visible tags in the document.
   */
  contextTagMetadata?: Record<string, ContextTagMeta>;
  /**
   * Document-level metadata from the Custom XML Part.
   * Stores template provenance, editor settings, etc.
   */
  fpDocumentMeta?: FPDocumentMeta;
  /**
   * Loop block metadata from the Custom XML Part.
   * Stores template XML and per-item rendered values for round-trip diff detection.
   * Keyed by collection name (e.g. "photos").
   */
  loopMetadata?: Record<string, import('../docx/contextTagMetadata').FPLoopMeta>;
}

/**
 * Metadata for a single context tag instance, persisted in the Custom XML Part.
 * Keyed by the tag's unique metaId (UUID) in the manifest.
 */
export interface ContextTagMeta {
  /** The context variable key (e.g., "context.case_no") — used for reconciliation on load */
  tagKey?: string;
  removeIfEmpty?: boolean;
  removeTableRow?: boolean;
  [key: string]: unknown;
}

/**
 * Document-level metadata persisted in the Custom XML Part.
 * Survives DOCX round-trips (download → edit externally → re-upload).
 */
export interface FPDocumentMeta {
  /** ID of the template this document was created from (DynamicReportTemplate PK) */
  templateId?: number;
  /** Human-readable name of the source template */
  templateName?: string;
  /** Selected TOC/numbering style ID (e.g., "Heading1") */
  tocStyle?: string;
}
