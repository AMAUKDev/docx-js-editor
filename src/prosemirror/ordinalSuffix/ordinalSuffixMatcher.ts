/**
 * Ordinal-suffix matcher (pure, framework-agnostic).
 *
 * Finds digit-run + ordinal-suffix ("1st", "22nd", "103rd", "9th") sequences that are
 * both preceded and followed by a real word boundary — anything other than a letter,
 * digit, underscore, or hyphen. This keeps a code-like token such as "45th8gh" or
 * "iPhone12th" from ever matching, while still matching at the very start/end of a
 * string (no preceding/following character at all satisfies the boundary check).
 *
 * Lowercase suffix only, matching Word's own AutoFormat behaviour ("3RD" is left
 * alone). No grammatical validation of digit<->suffix pairing (Word doesn't validate
 * "2th" either — it superscripts whatever suffix you actually typed).
 *
 * The trailing boundary must be an ACTUAL character, never merely "nothing typed
 * here yet". Word's AutoFormat only fires reactively when you type a trigger
 * character after the ordinal — it does not format "3rd" sitting at the end of an
 * open paragraph just because nothing follows it yet. Matching on bare end-of-string
 * here would fire the instant "3rd" finishes typing, before any space exists —
 * exactly the premature-formatting bug this whole feature exists to avoid.
 */

export interface OrdinalSuffixMatch {
  /** Start offset of the whole match (start of the digit run), relative to the scanned text. */
  matchStart: number;
  /** Start offset of the suffix letters — the part that gets superscripted. */
  suffixStart: number;
  /** End offset of the suffix letters (exclusive). */
  suffixEnd: number;
}

const ORDINAL_SUFFIX_RE = /(?<![A-Za-z0-9_-])\d+(st|nd|rd|th)(?=[^A-Za-z0-9_-])/g;

/** Find every ordinal-suffix match in `text`. Offsets are relative to `text`. */
export function findOrdinalSuffixMatches(text: string): OrdinalSuffixMatch[] {
  const matches: OrdinalSuffixMatch[] = [];
  ORDINAL_SUFFIX_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = ORDINAL_SUFFIX_RE.exec(text))) {
    const suffix = m[1];
    const matchEnd = m.index + m[0].length;
    matches.push({
      matchStart: m.index,
      suffixStart: matchEnd - suffix.length,
      suffixEnd: matchEnd,
    });
  }
  return matches;
}
