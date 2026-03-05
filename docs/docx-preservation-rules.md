# DOCX Preservation Rules

What MUST be preserved during the parse → edit → serialize round-trip to avoid Word rejecting the file as corrupt or degrading visual quality.

---

## Critical (Word rejects file without these)

### 1. `wps:bodyPr` on every `wps:wsp`

The OOXML schema requires `<wps:bodyPr>` as a child of every `<wps:wsp>` element, even when the shape has no text content. An empty `<wps:bodyPr/>` is sufficient.

```xml
<!-- CORRECT -->
<wps:wsp>
  <wps:cNvCnPr/>
  <wps:spPr>...</wps:spPr>
  <wps:bodyPr/>              ← required even if empty
</wps:wsp>

<!-- WRONG — Word flags as corrupt -->
<wps:wsp>
  <wps:cNvSpPr/>
  <wps:spPr>...</wps:spPr>
                              ← missing bodyPr
</wps:wsp>
```

### 2. Correct cNv\*Pr element for shape type

Connector shapes MUST use `<wps:cNvCnPr>`. Other shapes use `<wps:cNvSpPr>`.

| Shape types                                                            | Element          |
| ---------------------------------------------------------------------- | ---------------- |
| `straightConnector1`, `bentConnector2-5`, `curvedConnector2-5`, `line` | `<wps:cNvCnPr/>` |
| `rect`, `roundRect`, `ellipse`, `textBox`, all others                  | `<wps:cNvSpPr/>` |

### 3. Unique `w14:paraId` values

Every `<w:p>` must have a unique `w14:paraId`. ProseMirror's `splitBlock` copies all attrs to new paragraphs — the identity attrs (`paraId`, `textId`, `bookmarks`) must be cleared on the new node.

### 4. Unique bookmark IDs

Every `<w:bookmarkStart w:id="N">` must have a unique ID. Duplicate IDs cause corruption. When paragraphs are split, bookmarks must not be duplicated.

### 5. Balanced `fldChar` sequences

Complex fields (`begin` → `separate` → `end`) must be balanced. Every `begin` needs a matching `end`. Nested fields are allowed.

### 6. Balanced bookmarks

Every `<w:bookmarkStart>` must have a matching `<w:bookmarkEnd>` with the same `w:id`.

### 7. Element ordering (xs:sequence)

OOXML uses `xs:sequence` — child elements must appear in schema order. Key orderings:

- **pPr**: `pStyle` → `keepNext` → `numPr` → `pBdr` → `shd` → `tabs` → `spacing` → `ind` → `contextualSpacing` → `jc` → `outlineLvl` → `rPr`
- **rPr**: `rStyle` → `rFonts` → `b` → `i` → `caps` → `strike` → `noProof` → `color` → `spacing` → `sz` → `szCs` → `u` → `lang`
- **tblPr**: `tblStyle` → `tblW` → `jc` → `tblBorders` → `shd` → `tblLayout` → `tblCellMar` → `tblLook`
- **tcPr**: `cnfStyle` → `tcW` → `gridSpan` → `vMerge` → `tcBorders` → `shd` → `vAlign`
- **trPr**: `cnfStyle` → `trHeight` → `tblHeader` → `jc`
- **numPr**: `ilvl` → `numId`

---

## Important for Visual Fidelity (file opens but looks wrong without these)

### 8. `mc:AlternateContent` wrappers for shapes

Original documents wrap `wps:wsp` shapes in `mc:AlternateContent` with a `mc:Choice` (DrawingML) and `mc:Fallback` (VML). Stripping this works (Word can process bare wps content) but **the VML fallback contains rendering details** (exact positioning, shadow effects, gfxdata) that the simplified DrawingML reconstruction loses.

```xml
<!-- BEST — preserve original mc:AlternateContent -->
<mc:AlternateContent>
  <mc:Choice Requires="wps">
    <w:drawing><wp:anchor ...>...<wps:wsp>...</wps:wsp>...</wp:anchor></w:drawing>
  </mc:Choice>
  <mc:Fallback>
    <w:pict><v:shape ...>...</v:shape></w:pict>
  </mc:Fallback>
</mc:AlternateContent>
```

**Strategy**: For unedited shapes, preserve the original `mc:AlternateContent` XML rather than reconstructing from the parsed model.

### 9. `wp:anchor` vs `wp:inline`

Original anchored drawings (`wp:anchor` with position/wrap properties) should not be converted to `wp:inline`. Anchored shapes have absolute positioning that `wp:inline` cannot represent.

### 10. SDT (Structured Document Tag) wrappers

Original SDT wrappers carry metadata: lock state (`contentLocked`), date pickers (`w:date`), aliases, tags. Stripping them loses this functionality. The content is preserved but interactive controls are gone.

### 11. Shape property details

When reconstructing shapes from the parsed model, these original properties are lost:

- `bwMode` attribute on `spPr`
- `a:round` / `a:bevel` / `a:miter` line joins
- `a:headEnd` / `a:tailEnd` line endings
- `a:effectLst` (shadow, glow, reflection effects)
- `a:extLst` with hidden fill/effects
- Custom `o:gfxdata` in VML fallback

### 12. TOC complex fields

The TOC (`TOC \o "1-1" \h \z \u`) with nested `PAGEREF` hyperlinks is a complex field structure. Replacing it with plain text paragraphs works visually but loses the ability for Word to auto-update the TOC.

---

## Preservation Strategy

The safest approach is **preserve original XML for unedited content**:

1. **`contentDirty` flag** — When false, reuse `originalDocumentXml` verbatim (already implemented in `rezip.ts`)
2. **Per-element preservation** — For edited documents, preserve original XML for elements that weren't modified (shapes, headers/footers, images). Headers/footers already use this strategy.
3. **Shape round-trip** — Store original shape XML in the document model and emit it during serialization if the shape wasn't edited, rather than reconstructing from parsed properties.

---

## Quick Reference: What the Current Pipeline Loses

| Feature                  | Original                    | After round-trip               | Impact                    |
| ------------------------ | --------------------------- | ------------------------------ | ------------------------- |
| `wps:bodyPr`             | Present                     | **Now preserved (fixed)**      | Was causing corruption    |
| `cNvCnPr` for connectors | Correct                     | **Now correct (fixed)**        | Was causing corruption    |
| `mc:AlternateContent`    | Present with VML fallback   | Stripped, bare wps inline      | Visual degradation        |
| `wp:anchor` positioning  | Absolute position           | Converted to `wp:inline`       | Layout changes            |
| SDT wrappers             | Present with metadata       | Content-only, wrapper stripped | Lost interactive controls |
| TOC fields               | Complex field with PAGEREF  | Plain text paragraphs          | Can't auto-update TOC     |
| Shape effects            | Full effects (shadow, etc.) | Simplified or missing          | Visual degradation        |
| Body-level bookmarks     | Present                     | Stripped                       | Minor (internal refs)     |
