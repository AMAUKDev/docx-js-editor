import { describe, expect, it } from 'bun:test';
import { findAutoTagCandidates, DEFAULT_MATCH_OPTIONS } from './autoTagMatcher';

const TAG_MAP = {
  'case.case_no': 'AMA6835',
  'lead_client.account.name': 'Britannia Hong Kong Limited',
  'case.status': 'Open',
};

describe('findAutoTagCandidates', () => {
  it('matches a canonical case value (AMA6835 -> case.case_no) as high confidence', () => {
    const out = findAutoTagCandidates('Our reference is AMA6835 for this matter.', TAG_MAP);
    const hit = out.find((c) => c.value === 'AMA6835');
    expect(hit).toBeTruthy();
    expect(hit?.tagKey).toBe('case.case_no');
    expect(hit?.confidence).toBe('high');
    expect(hit?.ambiguous).toBe(false);
  });

  it('matches a multi-word alias value (lead_client)', () => {
    const out = findAutoTagCandidates('We act for Britannia Hong Kong Limited here.', TAG_MAP);
    const hit = out.find((c) => c.tagKey === 'lead_client.account.name');
    expect(hit?.value).toBe('Britannia Hong Kong Limited');
    expect(hit?.confidence).toBe('high');
  });

  it('returns offsets that slice back to the value', () => {
    const text = 'Ref AMA6835.';
    const [hit] = findAutoTagCandidates(text, { 'case.case_no': 'AMA6835' });
    expect(text.slice(hit.start, hit.end)).toBe('AMA6835');
  });

  it('respects word boundaries', () => {
    const out = findAutoTagCandidates('XAMA6835X and AMA6835Z', { 'case.case_no': 'AMA6835' });
    expect(out).toHaveLength(0);
  });

  it('skips too-short, numeric, and stop-word values', () => {
    expect(findAutoTagCandidates('id AB', { 'x.y': 'AB' })).toHaveLength(0);
    expect(findAutoTagCandidates('year 2024', { 'x.y': '2024' })).toHaveLength(0);
    expect(findAutoTagCandidates('the report', { 'x.y': 'report' })).toHaveLength(0);
  });

  it('flags ambiguous values (one value, multiple keys) as low confidence', () => {
    const out = findAutoTagCandidates('status is Pending now', {
      'case.status': 'Pending',
      'matter.state': 'Pending',
    });
    const hit = out.find((c) => c.value === 'Pending');
    expect(hit?.ambiguous).toBe(true);
    expect(hit?.candidateKeys).toEqual(['case.status', 'matter.state']);
    expect(hit?.confidence).toBe('low');
  });

  it('marks short unambiguous values low confidence', () => {
    const out = findAutoTagCandidates('port ABC done', { 'case.port': 'ABC' });
    expect(out.find((c) => c.value === 'ABC')?.confidence).toBe('low');
  });

  it('prefers the longest value on a contested span', () => {
    const out = findAutoTagCandidates('by Britannia Hong Kong Limited today', {
      'lead_client.account.name': 'Britannia Hong Kong Limited',
      'port.name': 'Hong Kong',
    });
    expect(out).toHaveLength(1);
    expect(out[0].tagKey).toBe('lead_client.account.name');
  });

  it('finds every occurrence of a repeated value', () => {
    const out = findAutoTagCandidates('AMA6835 ... AMA6835', { 'case.case_no': 'AMA6835' });
    expect(out).toHaveLength(2);
  });

  it('ignores null / sentinel values and empty inputs', () => {
    expect(findAutoTagCandidates('None [Not Found]', { a: null, b: 'None', c: '[Not Found]' })).toHaveLength(0);
    expect(findAutoTagCandidates('', TAG_MAP)).toHaveLength(0);
    expect(findAutoTagCandidates('text', {})).toHaveLength(0);
  });

  it('exposes tunable defaults', () => {
    expect(DEFAULT_MATCH_OPTIONS.minLength).toBe(3);
    expect(DEFAULT_MATCH_OPTIONS.minAutoLength).toBe(6);
  });
});
