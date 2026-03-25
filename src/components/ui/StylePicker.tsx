/**
 * Style Picker Component (Radix UI)
 *
 * A dropdown selector for applying named paragraph styles using Radix Select.
 * Shows each style with its visual appearance (font size, bold, color).
 * In gallery mode, renders as a "Select Style" dropdown + current-style label,
 * with numbering prefixes for heading styles.
 */

import * as React from 'react';
import { Select, SelectContent, SelectItem, SelectTrigger } from './Select';
import { cn } from '../../lib/utils';
import type { Style, StyleType, Theme } from '../../types/document';
import type { NumberingMap } from '../../docx/numberingParser';
import { formatNumberedMarker } from '../../layout-bridge/toFlowBlocks';

// ============================================================================
// TYPES
// ============================================================================

export interface StyleOption {
  styleId: string;
  name: string;
  type: StyleType;
  isDefault?: boolean;
  qFormat?: boolean;
  priority?: number;
  /** Font size in half-points for visual preview */
  fontSize?: number;
  /** Bold styling */
  bold?: boolean;
  /** Italic styling */
  italic?: boolean;
  /** Text color (RGB hex) */
  color?: string;
  /** Font family for preview rendering */
  fontFamily?: string;
  /** Numbering prefix (e.g. "1.0", "1.1", ".A") */
  numberPrefix?: string;
}

export interface StylePickerProps {
  value?: string;
  onChange?: (styleId: string) => void;
  styles?: Style[];
  theme?: Theme | null;
  disabled?: boolean;
  className?: string;
  width?: number | string;
  /** When true, render as a "Select Style" dropdown + current-style label */
  galleryMode?: boolean;
  /** If provided, only show styles whose styleId is in this list */
  allowedStyleIds?: string[];
  /** Numbering map for computing heading numbering prefixes */
  numberingMap?: NumberingMap | null;
}

// ============================================================================
// DEFAULT STYLES (matching Google Docs order and appearance)
// ============================================================================

const DEFAULT_STYLES: StyleOption[] = [
  {
    styleId: 'Normal',
    name: 'Normal text',
    type: 'paragraph',
    isDefault: true,
    priority: 0,
    qFormat: true,
    fontSize: 22, // 11pt
  },
  {
    styleId: 'Title',
    name: 'Title',
    type: 'paragraph',
    priority: 1,
    qFormat: true,
    fontSize: 52, // 26pt
    bold: true,
  },
  {
    styleId: 'Subtitle',
    name: 'Subtitle',
    type: 'paragraph',
    priority: 2,
    qFormat: true,
    fontSize: 30, // 15pt
    color: '666666', // Gray
  },
  {
    styleId: 'Heading1',
    name: 'Heading 1',
    type: 'paragraph',
    priority: 3,
    qFormat: true,
    fontSize: 40, // 20pt
    bold: true,
  },
  {
    styleId: 'Heading2',
    name: 'Heading 2',
    type: 'paragraph',
    priority: 4,
    qFormat: true,
    fontSize: 32, // 16pt
    bold: true,
  },
  {
    styleId: 'Heading3',
    name: 'Heading 3',
    type: 'paragraph',
    priority: 5,
    qFormat: true,
    fontSize: 28, // 14pt
    bold: true,
  },
];

// ============================================================================
// COMPONENT
// ============================================================================

/** Google Docs heading color */
const HEADING_COLOR = '#4a6c8c';

/**
 * Get inline styles for a style option's visual preview.
 * Derives clamped font size from the style's actual fontSize,
 * and applies fontFamily, bold, italic, and color.
 */
function getStylePreviewCSS(style: StyleOption): React.CSSProperties {
  const css: React.CSSProperties = {};

  // fontSize is in half-points; convert to pt then clamp for dropdown readability
  const pt = style.fontSize ? style.fontSize / 2 : 11;
  const px = Math.min(Math.max(pt, 11), 18);
  css.fontSize = `${px}px`;
  css.lineHeight = '1.3';

  if (style.fontFamily) {
    css.fontFamily = style.fontFamily;
  }

  if (style.bold) {
    css.fontWeight = 'bold';
  }

  if (style.italic) {
    css.fontStyle = 'italic';
  }

  // Use explicit color if provided, otherwise apply heading color for heading styles
  if (style.color) {
    css.color = `#${style.color}`;
  } else if (style.styleId.startsWith('Heading')) {
    css.color = HEADING_COLOR;
  }

  return css;
}

export function StylePicker({
  value,
  onChange,
  styles,
  disabled = false,
  className,
  width = 120,
  galleryMode = false,
  allowedStyleIds,
  numberingMap,
}: StylePickerProps) {
  // Convert document styles to options with visual info
  const styleOptions = React.useMemo(() => {
    let options: StyleOption[];

    if (!styles || styles.length === 0) {
      options = DEFAULT_STYLES;
    } else {
      // Show all paragraph styles that are visible:
      // - Not hidden and not semiHidden, OR
      // - Marked as qFormat (quick format / gallery style)
      const docStyles = styles
        .filter((s) => s.type === 'paragraph')
        .filter((s) => {
          if (s.qFormat) return true;
          // Show explicitly allowed styles even if semi-hidden (e.g. Heading 6+)
          if (allowedStyleIds?.includes(s.styleId)) return true;
          // Always show AMA-prefixed styles (company house styles)
          const id = s.styleId?.toLowerCase() ?? '';
          const name = s.name?.toLowerCase() ?? '';
          if (id.startsWith('ama') || name.startsWith('ama')) return true;
          if (s.hidden || s.semiHidden) return false;
          return true;
        })
        .map((s) => {
          const defaultStyle = DEFAULT_STYLES.find((d) => d.styleId === s.styleId);

          // Compute numbering prefix from style's numPr or reverse pStyle lookup
          let numberPrefix: string | undefined;
          const numPr = s.pPr?.numPr;
          let resolvedNumId: number | undefined;
          let resolvedIlvl: number | undefined;

          if (numPr?.numId != null) {
            // Style has explicit numPr
            resolvedNumId = numPr.numId;
            resolvedIlvl = numPr.ilvl ?? 0;
          } else if (numberingMap) {
            // Reverse lookup: check if a numbering level has pStyle pointing to this style
            const found = numberingMap.findByPStyle(s.styleId);
            if (found) {
              resolvedNumId = found.numId;
              resolvedIlvl = found.ilvl;
            }
          }

          if (resolvedNumId != null && resolvedIlvl != null && numberingMap) {
            // Use dummy counters (all 1s) to get a representative prefix
            const counters: number[] = [];
            for (let i = 0; i <= resolvedIlvl; i++) {
              counters.push(1);
            }
            const marker = formatNumberedMarker(
              counters,
              resolvedIlvl,
              numberingMap,
              resolvedNumId
            );
            if (marker) {
              numberPrefix = marker;
            }
          }

          // Extract font family from rPr
          const fontFamily = s.rPr?.fontFamily?.ascii || undefined;

          return {
            styleId: s.styleId,
            name: s.name || s.styleId,
            type: s.type,
            isDefault: s.default,
            qFormat: s.qFormat,
            priority: s.uiPriority ?? 99,
            // Extract visual properties from rPr, fall back to hardcoded defaults
            fontSize: s.rPr?.fontSize ?? defaultStyle?.fontSize,
            bold: s.rPr?.bold ?? defaultStyle?.bold,
            italic: s.rPr?.italic ?? defaultStyle?.italic,
            color: s.rPr?.color?.rgb ?? defaultStyle?.color,
            fontFamily,
            numberPrefix,
          };
        });

      // Sort by priority
      options = docStyles.sort((a, b) => (a.priority ?? 99) - (b.priority ?? 99));
    }

    // Filter to allowed style IDs if provided.
    // AMA-prefixed styles are always permitted (company house styles).
    if (allowedStyleIds) {
      options = options.filter(
        (s) =>
          allowedStyleIds.includes(s.styleId) ||
          s.styleId.toLowerCase().startsWith('ama') ||
          (s.name ?? '').toLowerCase().startsWith('ama')
      );
    }

    return options;
  }, [styles, allowedStyleIds, numberingMap]);

  // Key to force remount after selection — allows re-applying the same style.
  // After each selection the key increments AND we briefly set a sentinel value
  // so that Radix Select treats the next click on the same style as a real change.
  const [selectKey, setSelectKey] = React.useState(0);
  const [overrideValue, setOverrideValue] = React.useState<string | null>(null);

  const handleValueChange = React.useCallback(
    (newValue: string) => {
      onChange?.(newValue);
      // Set a sentinel value so Radix sees the next identical pick as a change
      setOverrideValue('__reset__');
      requestAnimationFrame(() => {
        setOverrideValue(null);
        setSelectKey((k) => k + 1);
      });
    },
    [onChange]
  );

  const currentValue = value || 'Normal';
  const displayName = styleOptions.find((s) => s.styleId === currentValue)?.name || currentValue;

  // Gallery mode: "Select Style" dropdown + current-style label
  if (galleryMode) {
    return (
      <div className={cn('flex items-center gap-2', className)}>
        <Select
          key={selectKey}
          value={overrideValue ?? currentValue}
          onValueChange={handleValueChange}
          disabled={disabled}
        >
          <SelectTrigger
            className="h-8 text-sm px-3 min-w-[120px]"
            aria-label="Select paragraph style"
          >
            <span className="truncate">Select Style</span>
          </SelectTrigger>
          <SelectContent className="min-w-[280px] max-h-[400px]">
            {styleOptions.map((style) => {
              const isActive = style.styleId === currentValue;
              const previewCSS = getStylePreviewCSS(style);
              return (
                <SelectItem key={style.styleId} value={style.styleId} className="py-2 px-3">
                  <div className="flex items-baseline gap-2 w-full">
                    {/* Numbering prefix column */}
                    <span
                      className="flex-shrink-0 text-right text-slate-400"
                      style={{ width: '44px', fontSize: '12px' }}
                    >
                      {style.numberPrefix || ''}
                    </span>
                    {/* Style name with formatting preview */}
                    <span
                      style={previewCSS}
                      className={cn('truncate', isActive && 'underline decoration-1')}
                    >
                      {style.name}
                    </span>
                  </div>
                </SelectItem>
              );
            })}
          </SelectContent>
        </Select>
        {/* Current style indicator */}
        <span className="text-xs text-slate-500 whitespace-nowrap">{displayName}</span>
      </div>
    );
  }

  // Default dropdown mode
  return (
    <Select
      key={selectKey}
      value={overrideValue ?? currentValue}
      onValueChange={handleValueChange}
      disabled={disabled}
    >
      <SelectTrigger
        className={cn('h-8 text-sm', className)}
        style={{ width: typeof width === 'number' ? `${width}px` : width }}
        aria-label="Select paragraph style"
      >
        <span className="truncate">{displayName}</span>
      </SelectTrigger>
      <SelectContent className="min-w-[260px] max-h-[400px]">
        {styleOptions.map((style) => (
          <SelectItem key={style.styleId} value={style.styleId} className="py-2.5 px-3">
            <span style={getStylePreviewCSS(style)}>{style.name}</span>
          </SelectItem>
        ))}
      </SelectContent>
    </Select>
  );
}
