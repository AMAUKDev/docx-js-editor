import { describe, expect, it } from 'bun:test';
import { findOrdinalSuffixMatches } from './ordinalSuffixMatcher';

function slice(text: string, m: { matchStart: number; suffixStart: number; suffixEnd: number }) {
  return {
    whole: text.slice(m.matchStart, m.suffixEnd),
    suffix: text.slice(m.suffixStart, m.suffixEnd),
  };
}

describe('findOrdinalSuffixMatches', () => {
  it('matches an ordinal followed by a space', () => {
    const text = 'the 3rd of May';
    const [m] = findOrdinalSuffixMatches(text);
    expect(slice(text, m)).toEqual({ whole: '3rd', suffix: 'rd' });
  });

  it('does NOT match at the very end of the string with nothing typed after it yet', () => {
    // Word only fires when you type a trigger character AFTER the ordinal — not the
    // instant the ordinal itself finishes. Matching bare end-of-string here would
    // format "3rd" the moment "d" is typed, before any real boundary exists.
    expect(findOrdinalSuffixMatches('the 21st')).toHaveLength(0);
  });

  it('matches an ordinal followed by common sentence punctuation', () => {
    for (const trailing of ['.', ',', ';', ':', '!', '?', ')', ']', '"', "'"]) {
      const text = `the 2nd${trailing} thing`;
      const [m] = findOrdinalSuffixMatches(text);
      expect(slice(text, m).suffix).toBe('nd');
    }
  });

  it('does NOT match a digit+letters code-like token (the request\'s exact example)', () => {
    expect(findOrdinalSuffixMatches('45th8gh')).toHaveLength(0);
  });

  it('does NOT match when followed by another letter or digit', () => {
    expect(findOrdinalSuffixMatches('4thing')).toHaveLength(0);
    expect(findOrdinalSuffixMatches('9th9')).toHaveLength(0);
  });

  it('does NOT match when followed by a hyphen or underscore', () => {
    expect(findOrdinalSuffixMatches('3rd-place')).toHaveLength(0);
    expect(findOrdinalSuffixMatches('3rd_place')).toHaveLength(0);
  });

  it('does NOT match when preceded by a letter, digit, or hyphen (e.g. a version-like token)', () => {
    expect(findOrdinalSuffixMatches('iPhone12th generation')).toHaveLength(0);
    expect(findOrdinalSuffixMatches('x-4th one')).toHaveLength(0);
  });

  it('does NOT match uppercase suffixes', () => {
    expect(findOrdinalSuffixMatches('the 3RD of May')).toHaveLength(0);
  });

  it('does not validate digit<->suffix grammar (matches Word\'s own leniency)', () => {
    const text = '2th ';
    const [m] = findOrdinalSuffixMatches(text);
    expect(m).toBeTruthy();
    expect(slice(text, m).suffix).toBe('th');
  });

  it('finds multiple independent ordinals in one string', () => {
    const text = 'on the 1st and the 22nd of that month, plus the 103rd item';
    const out = findOrdinalSuffixMatches(text);
    expect(out.map((m) => slice(text, m).whole)).toEqual(['1st', '22nd', '103rd']);
  });

  it('returns offsets that slice back to exactly the suffix', () => {
    const text = 'Item 5th.';
    const [m] = findOrdinalSuffixMatches(text);
    expect(text.slice(m.suffixStart, m.suffixEnd)).toBe('th');
    expect(text.slice(m.matchStart, m.suffixStart)).toBe('5');
  });

  it('returns no matches for plain text with no ordinals', () => {
    expect(findOrdinalSuffixMatches('nothing to see here')).toHaveLength(0);
  });
});
