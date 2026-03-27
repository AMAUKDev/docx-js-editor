/**
 * Style Editor Dialog
 *
 * Two-column modal for modifying existing styles or creating new ones.
 * Left: form fields (name, based-on, next, font, paragraph, indent).
 * Right: live preview + affected paragraph count.
 *
 * Modes:
 * - "modify": Pre-filled from existing style definition. Name is read-only.
 * - "create": Pre-filled from current selection formatting. Name is editable.
 */

import { useState, useEffect, useCallback, useRef } from 'react';
import type { CSSProperties } from 'react';

// ============================================================================
// TYPES
// ============================================================================

export interface StyleEditorData {
  name: string;
  basedOn: string;
  next: string;
  fontFamily: string;
  fontSize: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  color: string;
  alignment: 'left' | 'center' | 'right' | 'justify';
  spaceBefore: number;
  spaceAfter: number;
  lineSpacing: number;
  indentLeft: number;
  indentRight: number;
  indentFirstLine: number;
  /** Numbering reference — null or absent means no numbering */
  numPr?: { numId: number; ilvl: number } | null;
}

/** Numbering definition info for the dropdown */
export interface NumberingOption {
  /** Value for the select option (numId as string, or 'none') */
  value: string;
  /** Display name */
  name: string;
  /** Format preview (e.g., "1. / 1.1. / 1.1.1.") */
  preview: string;
}

export interface StyleEditorProps {
  isOpen: boolean;
  mode: 'modify' | 'create';
  /** Pre-filled data from existing style (modify) or selection (create) */
  initialData?: Partial<StyleEditorData>;
  /** Style ID being modified (modify mode only) */
  styleId?: string;
  /** Available paragraph styles for based-on and next dropdowns */
  availableStyles?: Array<{ styleId: string; name: string }>;
  /** Number of paragraphs that will be affected */
  affectedParagraphCount?: number;
  /** Available numbering definitions for the numbering dropdown */
  numberingOptions?: NumberingOption[];
  onSave: (data: StyleEditorData) => void;
  onClose: () => void;
}

// ============================================================================
// DEFAULT VALUES
// ============================================================================

const DEFAULT_DATA: StyleEditorData = {
  name: '',
  basedOn: 'Normal',
  next: '',
  fontFamily: 'Calibri',
  fontSize: 22, // half-points (11pt)
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  color: '000000',
  alignment: 'left',
  spaceBefore: 0,
  spaceAfter: 160, // 8pt in twentieths of a point
  lineSpacing: 276, // 1.15 lines (240 * 1.15)
  indentLeft: 0,
  indentRight: 0,
  indentFirstLine: 0,
  numPr: null,
};

// ============================================================================
// STYLES
// ============================================================================

const overlayStyle: CSSProperties = {
  position: 'fixed',
  top: 0,
  left: 0,
  right: 0,
  bottom: 0,
  backgroundColor: 'rgba(0,0,0,0.5)',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  zIndex: 10000,
};

const modalStyle: CSSProperties = {
  backgroundColor: '#fff',
  borderRadius: '8px',
  boxShadow: '0 20px 60px rgba(0,0,0,0.3)',
  width: '720px',
  maxHeight: '80vh',
  display: 'flex',
  flexDirection: 'column',
  overflow: 'hidden',
};

const headerStyle: CSSProperties = {
  padding: '16px 20px',
  borderBottom: '1px solid #e5e7eb',
  display: 'flex',
  justifyContent: 'space-between',
  alignItems: 'center',
};

const bodyStyle: CSSProperties = {
  padding: '20px',
  display: 'flex',
  gap: '20px',
  overflow: 'auto',
  flex: 1,
};

const leftColumnStyle: CSSProperties = {
  flex: 1,
  display: 'flex',
  flexDirection: 'column',
  gap: '16px',
};

const rightColumnStyle: CSSProperties = {
  width: '260px',
  display: 'flex',
  flexDirection: 'column',
  gap: '12px',
};

const footerStyle: CSSProperties = {
  padding: '12px 20px',
  borderTop: '1px solid #e5e7eb',
  display: 'flex',
  justifyContent: 'flex-end',
  gap: '8px',
};

const sectionStyle: CSSProperties = {
  borderTop: '1px solid #f3f4f6',
  paddingTop: '12px',
};

const labelStyle: CSSProperties = {
  fontSize: '12px',
  fontWeight: 600,
  color: '#6b7280',
  marginBottom: '4px',
  display: 'block',
};

const inputStyle: CSSProperties = {
  width: '100%',
  padding: '6px 8px',
  border: '1px solid #d1d5db',
  borderRadius: '4px',
  fontSize: '13px',
};

const selectStyle: CSSProperties = {
  ...inputStyle,
  cursor: 'pointer',
};

const rowStyle: CSSProperties = {
  display: 'flex',
  gap: '8px',
  alignItems: 'center',
};

const toggleBtnStyle = (active: boolean): CSSProperties => ({
  padding: '4px 8px',
  border: '1px solid #d1d5db',
  borderRadius: '4px',
  backgroundColor: active ? '#3b82f6' : '#fff',
  color: active ? '#fff' : '#374151',
  cursor: 'pointer',
  fontSize: '13px',
  fontWeight: active ? 700 : 400,
  minWidth: '32px',
  textAlign: 'center',
});

const previewStyle = (data: StyleEditorData): CSSProperties => ({
  padding: '16px',
  border: '1px solid #e5e7eb',
  borderRadius: '6px',
  backgroundColor: '#fafafa',
  minHeight: '80px',
  fontFamily: data.fontFamily || 'Calibri',
  fontSize: `${(data.fontSize || 22) / 2}pt`,
  fontWeight: data.bold ? 700 : 400,
  fontStyle: data.italic ? 'italic' : 'normal',
  textDecoration:
    [data.underline ? 'underline' : '', data.strikethrough ? 'line-through' : '']
      .filter(Boolean)
      .join(' ') || 'none',
  color: `#${data.color || '000000'}`,
  textAlign: data.alignment || 'left',
  lineHeight: data.lineSpacing ? `${(data.lineSpacing / 240).toFixed(2)}` : '1.15',
});

const btnStyle = (variant: 'primary' | 'secondary'): CSSProperties => ({
  padding: '8px 16px',
  border: variant === 'primary' ? 'none' : '1px solid #d1d5db',
  borderRadius: '6px',
  backgroundColor: variant === 'primary' ? '#3b82f6' : '#fff',
  color: variant === 'primary' ? '#fff' : '#374151',
  cursor: 'pointer',
  fontSize: '13px',
  fontWeight: 500,
});

// ============================================================================
// COMPONENT
// ============================================================================

export function StyleEditorDialog({
  isOpen,
  mode,
  initialData,
  styleId,
  availableStyles = [],
  affectedParagraphCount = 0,
  numberingOptions = [],
  onSave,
  onClose,
}: StyleEditorProps) {
  const [data, setData] = useState<StyleEditorData>({ ...DEFAULT_DATA, ...initialData });
  const nameRef = useRef<HTMLInputElement>(null);

  // Reset form when dialog opens
  useEffect(() => {
    if (isOpen) {
      setData({ ...DEFAULT_DATA, ...initialData });
      if (mode === 'create') {
        setTimeout(() => nameRef.current?.focus(), 100);
      }
    }
  }, [isOpen, mode, initialData]);

  const handleChange = useCallback(
    <K extends keyof StyleEditorData>(key: K, value: StyleEditorData[K]) => {
      setData((prev) => ({ ...prev, [key]: value }));
    },
    []
  );

  const handleSave = useCallback(() => {
    if (mode === 'create' && !data.name.trim()) return;
    onSave(data);
  }, [data, mode, onSave]);

  if (!isOpen) return null;

  const title = mode === 'modify' ? `Modify Style: ${styleId || ''}` : 'Create New Style';

  return (
    <div style={overlayStyle} onMouseDown={onClose} data-testid="style-editor-modal">
      <div style={modalStyle} onMouseDown={(e) => e.stopPropagation()}>
        {/* Header */}
        <div style={headerStyle}>
          <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 600 }}>{title}</h3>
          <button
            onClick={onClose}
            style={{
              border: 'none',
              background: 'none',
              fontSize: '18px',
              cursor: 'pointer',
              color: '#9ca3af',
            }}
          >
            ✕
          </button>
        </div>

        {/* Body */}
        <div style={bodyStyle}>
          {/* Left column: form fields */}
          <div style={leftColumnStyle}>
            {/* Name */}
            <div>
              <label style={labelStyle}>Style Name</label>
              <input
                ref={nameRef}
                style={inputStyle}
                value={data.name}
                readOnly={mode === 'modify'}
                onChange={(e) => handleChange('name', e.target.value)}
                placeholder={mode === 'create' ? 'Enter style name...' : undefined}
                data-testid="style-name-input"
              />
            </div>

            {/* Based on & Next */}
            <div style={rowStyle}>
              <div style={{ flex: 1 }}>
                <label style={labelStyle}>Based On</label>
                <select
                  style={selectStyle}
                  value={data.basedOn}
                  onChange={(e) => handleChange('basedOn', e.target.value)}
                >
                  <option value="">(None)</option>
                  {availableStyles.map((s) => (
                    <option key={s.styleId} value={s.styleId}>
                      {s.name}
                    </option>
                  ))}
                </select>
              </div>
              <div style={{ flex: 1 }}>
                <label style={labelStyle}>Next Style</label>
                <select
                  style={selectStyle}
                  value={data.next}
                  onChange={(e) => handleChange('next', e.target.value)}
                >
                  <option value="">(Same)</option>
                  {availableStyles.map((s) => (
                    <option key={s.styleId} value={s.styleId}>
                      {s.name}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            {/* Font section */}
            <div style={sectionStyle}>
              <label style={{ ...labelStyle, marginBottom: '8px' }}>Font</label>
              <div style={rowStyle}>
                <input
                  style={{ ...inputStyle, flex: 2 }}
                  value={data.fontFamily}
                  onChange={(e) => handleChange('fontFamily', e.target.value)}
                  placeholder="Font family"
                  data-testid="style-font-family"
                  aria-label="Font family"
                />
                <input
                  style={{ ...inputStyle, width: '60px', flex: 0 }}
                  type="number"
                  value={Math.round(data.fontSize / 2)}
                  onChange={(e) =>
                    handleChange('fontSize', Math.round(parseFloat(e.target.value) * 2) || 22)
                  }
                  data-testid="style-font-size"
                  aria-label="Font size"
                />
                <span style={{ fontSize: '12px', color: '#9ca3af' }}>pt</span>
              </div>
              <div style={{ ...rowStyle, marginTop: '8px' }}>
                <button
                  type="button"
                  style={toggleBtnStyle(data.bold)}
                  onClick={() => handleChange('bold', !data.bold)}
                >
                  <strong>B</strong>
                </button>
                <button
                  type="button"
                  style={toggleBtnStyle(data.italic)}
                  onClick={() => handleChange('italic', !data.italic)}
                >
                  <em>I</em>
                </button>
                <button
                  type="button"
                  style={toggleBtnStyle(data.underline)}
                  onClick={() => handleChange('underline', !data.underline)}
                >
                  <u>U</u>
                </button>
                <button
                  type="button"
                  style={toggleBtnStyle(data.strikethrough)}
                  onClick={() => handleChange('strikethrough', !data.strikethrough)}
                >
                  <s>S</s>
                </button>
                <div
                  style={{ marginLeft: '8px', display: 'flex', alignItems: 'center', gap: '4px' }}
                >
                  <label style={{ fontSize: '12px', color: '#6b7280' }}>Color:</label>
                  <input
                    type="color"
                    value={`#${data.color}`}
                    onChange={(e) => handleChange('color', e.target.value.replace('#', ''))}
                    style={{ width: '28px', height: '28px', border: 'none', cursor: 'pointer' }}
                  />
                </div>
              </div>
            </div>

            {/* Paragraph section */}
            <div style={sectionStyle}>
              <label style={{ ...labelStyle, marginBottom: '8px' }}>Paragraph</label>
              <div style={rowStyle}>
                <div style={{ flex: 1 }}>
                  <label style={{ ...labelStyle, fontSize: '11px' }}>Alignment</label>
                  <select
                    style={selectStyle}
                    value={data.alignment}
                    onChange={(e) =>
                      handleChange('alignment', e.target.value as StyleEditorData['alignment'])
                    }
                  >
                    <option value="left">Left</option>
                    <option value="center">Center</option>
                    <option value="right">Right</option>
                    <option value="justify">Justify</option>
                  </select>
                </div>
                <div style={{ flex: 1 }}>
                  <label style={{ ...labelStyle, fontSize: '11px' }}>Line Spacing</label>
                  <input
                    style={inputStyle}
                    type="number"
                    step="0.05"
                    value={(data.lineSpacing / 240).toFixed(2)}
                    onChange={(e) =>
                      handleChange(
                        'lineSpacing',
                        Math.round(parseFloat(e.target.value) * 240) || 240
                      )
                    }
                  />
                </div>
              </div>
              <div style={{ ...rowStyle, marginTop: '8px' }}>
                <div style={{ flex: 1 }}>
                  <label style={{ ...labelStyle, fontSize: '11px' }}>Space Before (pt)</label>
                  <input
                    style={inputStyle}
                    type="number"
                    value={Math.round(data.spaceBefore / 20)}
                    onChange={(e) =>
                      handleChange('spaceBefore', Math.round(parseFloat(e.target.value) * 20) || 0)
                    }
                  />
                </div>
                <div style={{ flex: 1 }}>
                  <label style={{ ...labelStyle, fontSize: '11px' }}>Space After (pt)</label>
                  <input
                    style={inputStyle}
                    type="number"
                    value={Math.round(data.spaceAfter / 20)}
                    onChange={(e) =>
                      handleChange('spaceAfter', Math.round(parseFloat(e.target.value) * 20) || 0)
                    }
                  />
                </div>
              </div>
            </div>

            {/* Indent section */}
            <div style={sectionStyle}>
              <label style={{ ...labelStyle, marginBottom: '8px' }}>Indent</label>
              <div style={rowStyle}>
                <div style={{ flex: 1 }}>
                  <label style={{ ...labelStyle, fontSize: '11px' }}>Left (pt)</label>
                  <input
                    style={inputStyle}
                    type="number"
                    value={Math.round(data.indentLeft / 20)}
                    onChange={(e) =>
                      handleChange('indentLeft', Math.round(parseFloat(e.target.value) * 20) || 0)
                    }
                  />
                </div>
                <div style={{ flex: 1 }}>
                  <label style={{ ...labelStyle, fontSize: '11px' }}>Right (pt)</label>
                  <input
                    style={inputStyle}
                    type="number"
                    value={Math.round(data.indentRight / 20)}
                    onChange={(e) =>
                      handleChange('indentRight', Math.round(parseFloat(e.target.value) * 20) || 0)
                    }
                  />
                </div>
                <div style={{ flex: 1 }}>
                  <label style={{ ...labelStyle, fontSize: '11px' }}>First Line (pt)</label>
                  <input
                    style={inputStyle}
                    type="number"
                    value={Math.round(data.indentFirstLine / 20)}
                    onChange={(e) =>
                      handleChange(
                        'indentFirstLine',
                        Math.round(parseFloat(e.target.value) * 20) || 0
                      )
                    }
                  />
                </div>
              </div>
            </div>

            {/* Numbering section */}
            {numberingOptions.length > 0 && (
              <div style={sectionStyle} data-testid="style-numbering-section">
                <label style={{ ...labelStyle, marginBottom: '8px' }}>Numbering</label>
                <div style={rowStyle}>
                  <div style={{ flex: 2 }}>
                    <label style={{ ...labelStyle, fontSize: '11px' }}>Definition</label>
                    <select
                      style={selectStyle}
                      data-testid="style-numbering-dropdown"
                      value={data.numPr ? String(data.numPr.numId) : 'none'}
                      onChange={(e) => {
                        const val = e.target.value;
                        if (val === 'none') {
                          handleChange('numPr', null);
                        } else {
                          const numId = parseInt(val, 10);
                          handleChange('numPr', {
                            numId,
                            ilvl: data.numPr?.ilvl ?? 0,
                          });
                        }
                      }}
                    >
                      <option value="none">None</option>
                      {numberingOptions.map((opt) => (
                        <option key={opt.value} value={opt.value}>
                          {opt.name} — {opt.preview}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div style={{ flex: 1 }}>
                    <label style={{ ...labelStyle, fontSize: '11px' }}>Level</label>
                    <select
                      style={selectStyle}
                      data-testid="style-numbering-level"
                      disabled={!data.numPr}
                      value={data.numPr?.ilvl ?? 0}
                      onChange={(e) => {
                        if (data.numPr) {
                          handleChange('numPr', {
                            ...data.numPr,
                            ilvl: parseInt(e.target.value, 10),
                          });
                        }
                      }}
                    >
                      {[0, 1, 2, 3, 4, 5, 6, 7, 8].map((lvl) => (
                        <option key={lvl} value={lvl}>
                          Level {lvl}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
              </div>
            )}
          </div>

          {/* Right column: preview + paragraph count */}
          <div style={rightColumnStyle}>
            <div>
              <label style={labelStyle}>Preview</label>
              <div style={previewStyle(data)} data-testid="style-preview">
                The quick brown fox jumps over the lazy dog. Sample text for previewing the style
                definition.
              </div>
            </div>
            <div
              data-testid="style-para-count"
              style={{
                padding: '8px 12px',
                backgroundColor: '#f0f9ff',
                borderRadius: '6px',
                fontSize: '12px',
                color: '#0369a1',
              }}
            >
              {mode === 'modify'
                ? `This will update ${affectedParagraphCount} paragraph${affectedParagraphCount !== 1 ? 's' : ''} in this document.`
                : 'New style will be applied to the current selection.'}
            </div>
          </div>
        </div>

        {/* Footer */}
        <div style={footerStyle}>
          <button type="button" style={btnStyle('secondary')} onClick={onClose}>
            Cancel
          </button>
          <button
            type="button"
            style={btnStyle('primary')}
            onClick={handleSave}
            disabled={mode === 'create' && !data.name.trim()}
          >
            {mode === 'modify' ? 'OK' : 'Create Style'}
          </button>
        </div>
      </div>
    </div>
  );
}

export default StyleEditorDialog;
