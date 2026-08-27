/**
 * Auto context-tag matcher (pure, framework-agnostic).
 *
 * Given plain text and a complete {tagKey: value} map, returns the ranges of text
 * that match a tag value and should become context-tag nodes. Guards protect against
 * false conversions (too-short, numeric/date-like, common words, ambiguous values).
 *
 * This is the fork-native counterpart of the portal's
 * `assets/dms/shared/utils/autoTagMatcher.js` — keep the two in sync.
 */

export type MatchConfidence = 'high' | 'low';

export interface AutoTagCandidate {
  value: string;
  tagKey: string;
  candidateKeys: string[];
  start: number;
  end: number;
  ambiguous: boolean;
  confidence: MatchConfidence;
}

export interface AutoTagMatchOptions {
  /** Values shorter than this are never candidates. */
  minLength?: number;
  /** Unambiguous values at least this long convert automatically ("high"). */
  minAutoLength?: number;
  /** Skip values that are entirely digits / date-like punctuation. */
  skipNumeric?: boolean;
  /** Lower-cased values to always ignore. */
  stopWords?: Set<string>;
}

export const COMMON_STOP_WORDS = new Set<string>([
  'the', 'and', 'for', 'with', 'from', 'this', 'that', 'case', 'report',
  'client', 'date', 'name', 'yes', 'no', 'true', 'false', 'none', 'n/a',
  'ltd', 'limited', 'company', 'office', 'page', 'total', 'type',
]);

export const DEFAULT_MATCH_OPTIONS: Required<AutoTagMatchOptions> = {
  minLength: 3,
  minAutoLength: 6,
  skipNumeric: true,
  stopWords: COMMON_STOP_WORDS,
};

const NON_MATCHABLE_VALUES = new Set(['None', 'True', 'False', 'N/A', '']);

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function isEligibleValue(value: string, options: Required<AutoTagMatchOptions>): boolean {
  if (value.length < options.minLength) return false;
  if (NON_MATCHABLE_VALUES.has(value)) return false;
  if (value.startsWith('[') || value.startsWith('{')) return false;
  if (value.includes('[Not Found]')) return false;
  if (options.skipNumeric && /^[\d.,\s/:-]+$/.test(value)) return false;
  if (options.stopWords.has(value.toLowerCase())) return false;
  return true;
}

function overlapsClaimed(start: number, end: number, claimed: Array<[number, number]>): boolean {
  for (const [cStart, cEnd] of claimed) {
    if (start < cEnd && end > cStart) return true;
  }
  return false;
}

/**
 * Scan `text` for occurrences of any value in `tagMap` and return guarded,
 * non-overlapping candidates ordered by position (longest value wins a contested span).
 */
export function findAutoTagCandidates(
  text: string,
  tagMap: Record<string, string | null | undefined> | null | undefined,
  options: AutoTagMatchOptions = {},
): AutoTagCandidate[] {
  const opts: Required<AutoTagMatchOptions> = { ...DEFAULT_MATCH_OPTIONS, ...options };
  if (!text || !tagMap) return [];

  const valueToKeys = new Map<string, { value: string; keys: Set<string> }>();
  for (const [tagKey, rawValue] of Object.entries(tagMap)) {
    if (rawValue == null) continue;
    const value = String(rawValue).trim();
    if (!isEligibleValue(value, opts)) continue;
    const entry = valueToKeys.get(value) || { value, keys: new Set<string>() };
    entry.keys.add(tagKey);
    valueToKeys.set(value, entry);
  }

  const entries = [...valueToKeys.values()].sort((a, b) => b.value.length - a.value.length);

  const candidates: AutoTagCandidate[] = [];
  const claimed: Array<[number, number]> = [];
  for (const entry of entries) {
    const pattern = new RegExp(`(?<!\\w)${escapeRegExp(entry.value)}(?!\\w)`, 'g');
    let match: RegExpExecArray | null;
    while ((match = pattern.exec(text)) !== null) {
      const start = match.index;
      const end = start + entry.value.length;
      if (overlapsClaimed(start, end, claimed)) continue;
      claimed.push([start, end]);
      const candidateKeys = [...entry.keys].sort();
      const ambiguous = candidateKeys.length > 1;
      const confidence: MatchConfidence =
        !ambiguous && entry.value.length >= opts.minAutoLength ? 'high' : 'low';
      candidates.push({
        value: entry.value,
        tagKey: candidateKeys[0],
        candidateKeys,
        start,
        end,
        ambiguous,
        confidence,
      });
    }
  }

  candidates.sort((a, b) => a.start - b.start);
  return candidates;
}
