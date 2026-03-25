/**
 * Module-level style definition store.
 *
 * Populated when a document is loaded, consumed by ProseMirror commands
 * (e.g., Enter key handler) that need access to style.next and fallback resolution.
 *
 * This avoids threading style data through every command/plugin signature.
 */
import type { Style } from '../../types/styles';

let _styles: Style[] = [];

/** Set the current document's style definitions. Call when a document is loaded. */
export function setDocumentStyles(styles: Style[]): void {
  _styles = styles;
}

/** Get a style definition by styleId. */
export function getStyleDef(styleId: string): Style | undefined {
  return _styles.find((s) => s.styleId === styleId);
}

/** Get the 'next' styleId for a given style (what Enter should produce). */
export function getNextStyleId(styleId: string): string | undefined {
  const def = getStyleDef(styleId);
  return def?.next;
}

/**
 * Resolve font fallback: given a style's run formatting, if no fontFamily is defined,
 * walk fallback chain: AMA_normal → AMA_default → Normal → undefined.
 */
export function resolveFontFallback(): { fontFamily?: string; fontSize?: number } | undefined {
  const fallbackIds = ['AMA_normal', 'AMA_default', 'Normal'];
  for (const id of fallbackIds) {
    const def = getStyleDef(id);
    if (!def) continue;
    const rPr = def.rPr;
    if (rPr?.fontFamily?.ascii || rPr?.fontSize) {
      return {
        fontFamily: rPr.fontFamily?.ascii,
        fontSize: rPr.fontSize,
      };
    }
  }
  return undefined;
}

/** Clear the store (e.g., when editor unmounts). */
export function clearDocumentStyles(): void {
  _styles = [];
}
