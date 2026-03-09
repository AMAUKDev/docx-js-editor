/**
 * Tests for renderWithBookmarks — bookmark generation during context tag rendering.
 */

import { describe, test, expect } from 'bun:test';
import {
  renderParagraphContent,
  renderDocumentWithBookmarks,
  findMaxBookmarkId,
  FP_BOOKMARK_PREFIX,
} from './renderWithBookmarks';
import type {
  Run,
  BookmarkStart,
  BookmarkEnd,
  ParagraphContent,
  Paragraph,
  ContextTagMeta,
  Document,
} from '../types/document';

// ── Helpers ─────────────────────────────────────────────────────

function textRun(text: string, formatting?: Run['formatting']): Run {
  return { type: 'run', formatting, content: [{ type: 'text', text }] };
}

function makeParagraph(content: ParagraphContent[]): Paragraph {
  return { type: 'paragraph', content };
}

function makeDocument(paragraphs: Paragraph[]): Document {
  return {
    package: {
      document: {
        content: paragraphs,
      },
    },
  } as unknown as Document;
}

function getBookmarkStarts(content: ParagraphContent[]): BookmarkStart[] {
  return content.filter((c): c is BookmarkStart => c.type === 'bookmarkStart');
}

function getBookmarkEnds(content: ParagraphContent[]): BookmarkEnd[] {
  return content.filter((c): c is BookmarkEnd => c.type === 'bookmarkEnd');
}

function getRuns(content: ParagraphContent[]): Run[] {
  return content.filter((c): c is Run => c.type === 'run');
}

function getRunTexts(content: ParagraphContent[]): string[] {
  return getRuns(content).map((r) =>
    r.content
      .filter((c) => c.type === 'text')
      .map((c) => (c as { text: string }).text)
      .join('')
  );
}

// ── renderParagraphContent ──────────────────────────────────────

describe('renderParagraphContent', () => {
  const ctMeta: Record<string, ContextTagMeta> = {
    'uuid-1': { tagKey: 'context.case_no' },
    'uuid-2': { tagKey: 'context.vessel' },
    'uuid-3': { tagKey: 'context.case_no' }, // second instance of same tag
    'uuid-4': { tagKey: 'context.empty_tag', removeIfEmpty: true },
  };

  const tags: Record<string, string> = {
    'context.case_no': 'AMA7089',
    'context.vessel': 'ORANGE PHOENIX',
  };

  function render(
    content: ParagraphContent[],
    mode: 'omit' | 'keep' = 'keep',
    cursorsInit?: Map<string, number>
  ) {
    const metaGroups = new Map<string, Array<{ metaId: string; meta: ContextTagMeta }>>();
    for (const [metaId, meta] of Object.entries(ctMeta)) {
      const key = meta.tagKey!;
      if (!metaGroups.has(key)) metaGroups.set(key, []);
      metaGroups.get(key)!.push({ metaId, meta });
    }
    return renderParagraphContent(content, cursorsInit ?? new Map(), metaGroups, {
      tags,
      ctMeta,
      mode,
      startBookmarkId: 100,
    });
  }

  test('resolved tag wrapped in bookmark', () => {
    const content = [textRun('{{ context.case_no }}')];
    const result = render(content);

    expect(result.removeParagraph).toBe(false);
    const starts = getBookmarkStarts(result.content);
    const ends = getBookmarkEnds(result.content);

    expect(starts).toHaveLength(1);
    expect(ends).toHaveLength(1);
    expect(starts[0].name).toBe(`${FP_BOOKMARK_PREFIX}uuid-1`);
    expect(starts[0].id).toBe(ends[0].id);

    const texts = getRunTexts(result.content);
    expect(texts).toEqual(['AMA7089']);
  });

  test('tag with surrounding text preserves surrounding text', () => {
    const content = [textRun('Case: {{ context.case_no }} is active')];
    const result = render(content);

    const texts = getRunTexts(result.content);
    expect(texts).toEqual(['Case: ', 'AMA7089', ' is active']);

    const starts = getBookmarkStarts(result.content);
    expect(starts).toHaveLength(1);
    expect(starts[0].name).toBe(`${FP_BOOKMARK_PREFIX}uuid-1`);
  });

  test('multiple tags in one run', () => {
    const content = [textRun('{{ context.case_no }} on {{ context.vessel }}')];
    const result = render(content);

    const texts = getRunTexts(result.content);
    expect(texts).toEqual(['AMA7089', ' on ', 'ORANGE PHOENIX']);

    const starts = getBookmarkStarts(result.content);
    expect(starts).toHaveLength(2);
    expect(starts[0].name).toBe(`${FP_BOOKMARK_PREFIX}uuid-1`);
    expect(starts[1].name).toBe(`${FP_BOOKMARK_PREFIX}uuid-2`);
  });

  test('same tagKey consumed in document order', () => {
    // Two occurrences of context.case_no — should get uuid-1 then uuid-3
    const content = [textRun('{{ context.case_no }}'), textRun('{{ context.case_no }}')];
    const result = render(content);

    const starts = getBookmarkStarts(result.content);
    expect(starts).toHaveLength(2);
    expect(starts[0].name).toBe(`${FP_BOOKMARK_PREFIX}uuid-1`);
    expect(starts[1].name).toBe(`${FP_BOOKMARK_PREFIX}uuid-3`);
  });

  test('bookmark IDs are unique and sequential', () => {
    const content = [textRun('{{ context.case_no }} {{ context.vessel }}')];
    const result = render(content);

    const starts = getBookmarkStarts(result.content);
    const ends = getBookmarkEnds(result.content);

    expect(starts[0].id).toBe(100);
    expect(starts[1].id).toBe(101);
    expect(ends[0].id).toBe(100);
    expect(ends[1].id).toBe(101);
    expect(result.nextBookmarkId).toBe(102);
  });

  test('unresolved tag in keep mode preserved with bookmark', () => {
    const content = [textRun('{{ context.missing_field }}')];
    const metaWithMissing: Record<string, ContextTagMeta> = {
      'uuid-m': { tagKey: 'context.missing_field' },
    };
    const metaGroups = new Map<string, Array<{ metaId: string; meta: ContextTagMeta }>>();
    metaGroups.set('context.missing_field', [
      { metaId: 'uuid-m', meta: metaWithMissing['uuid-m'] },
    ]);

    const result = renderParagraphContent(content, new Map(), metaGroups, {
      tags: {}, // no resolved value
      ctMeta: metaWithMissing,
      mode: 'keep',
      startBookmarkId: 50,
    });

    const texts = getRunTexts(result.content);
    expect(texts).toEqual(['{context.missing_field}']);

    const starts = getBookmarkStarts(result.content);
    expect(starts).toHaveLength(1);
    expect(starts[0].name).toBe(`${FP_BOOKMARK_PREFIX}uuid-m`);
  });

  test('removeIfEmpty in omit mode flags paragraph for removal', () => {
    const content = [textRun('{{ context.empty_tag! }}')];
    const metaWithRie: Record<string, ContextTagMeta> = {
      'uuid-rie': { tagKey: 'context.empty_tag', removeIfEmpty: true },
    };
    const metaGroups = new Map<string, Array<{ metaId: string; meta: ContextTagMeta }>>();
    metaGroups.set('context.empty_tag', [{ metaId: 'uuid-rie', meta: metaWithRie['uuid-rie'] }]);

    const result = renderParagraphContent(content, new Map(), metaGroups, {
      tags: {}, // no value
      ctMeta: metaWithRie,
      mode: 'omit',
      startBookmarkId: 50,
    });

    expect(result.removeParagraph).toBe(true);
  });

  test('omit mode without removeIfEmpty skips tag silently', () => {
    const content = [textRun('Before {{ context.missing_field }} After')];
    const metaGroups = new Map<string, Array<{ metaId: string; meta: ContextTagMeta }>>();
    metaGroups.set('context.missing_field', [
      { metaId: 'uuid-x', meta: { tagKey: 'context.missing_field' } },
    ]);

    const result = renderParagraphContent(content, new Map(), metaGroups, {
      tags: {}, // no value
      ctMeta: {},
      mode: 'omit',
      startBookmarkId: 50,
    });

    expect(result.removeParagraph).toBe(false);
    const texts = getRunTexts(result.content);
    expect(texts).toEqual(['Before ', ' After']);
  });

  test('non-run content passed through unchanged', () => {
    const existing: BookmarkStart = { type: 'bookmarkStart', id: 1, name: '_GoBack' };
    const existingEnd: BookmarkEnd = { type: 'bookmarkEnd', id: 1 };
    const content: ParagraphContent[] = [existing, textRun('Hello'), existingEnd];

    const result = render(content);
    expect(result.content[0]).toBe(existing);
    expect(result.content[2]).toBe(existingEnd);
  });

  test('run with no tag patterns passed through unchanged', () => {
    const run = textRun('Plain text with no tags');
    const result = render([run]);

    expect(result.content).toHaveLength(1);
    expect(result.content[0]).toBe(run);
  });

  test('preserves run formatting on rendered text', () => {
    const bold = { bold: true, fontSize: 24 };
    const content = [textRun('{{ context.case_no }}', bold)];
    const result = render(content);

    const runs = getRuns(result.content);
    expect(runs).toHaveLength(1);
    expect(runs[0].formatting?.bold).toBe(true);
    expect(runs[0].formatting?.fontSize).toBe(24);
  });

  test('single-brace tag format also works', () => {
    const content = [textRun('{context.case_no}')];
    const result = render(content);

    const texts = getRunTexts(result.content);
    expect(texts).toEqual(['AMA7089']);

    const starts = getBookmarkStarts(result.content);
    expect(starts).toHaveLength(1);
  });
});

// ── renderDocumentWithBookmarks ─────────────────────────────────

describe('renderDocumentWithBookmarks', () => {
  const ctMeta: Record<string, ContextTagMeta> = {
    'uuid-a': { tagKey: 'context.case_no' },
    'uuid-b': { tagKey: 'context.vessel' },
  };

  const tags: Record<string, string> = {
    'context.case_no': 'AMA7089',
    'context.vessel': 'ORANGE PHOENIX',
  };

  test('renders tags and inserts bookmarks in document', () => {
    const doc = makeDocument([
      makeParagraph([textRun('Case: {{ context.case_no }}')]),
      makeParagraph([textRun('Vessel: {{ context.vessel }}')]),
    ]);

    renderDocumentWithBookmarks(doc, { tags, ctMeta, mode: 'keep' });

    const p1 = doc.package.document.content[0] as Paragraph;
    const p2 = doc.package.document.content[1] as Paragraph;

    expect(getBookmarkStarts(p1.content)).toHaveLength(1);
    expect(getBookmarkStarts(p2.content)).toHaveLength(1);
    expect(getRunTexts(p1.content)).toEqual(['Case: ', 'AMA7089']);
    expect(getRunTexts(p2.content)).toEqual(['Vessel: ', 'ORANGE PHOENIX']);
  });

  test('removes paragraph with removeIfEmpty in omit mode', () => {
    const ctMetaRie: Record<string, ContextTagMeta> = {
      'uuid-r': { tagKey: 'context.empty', removeIfEmpty: true },
    };
    const doc = makeDocument([
      makeParagraph([textRun('Keep this')]),
      makeParagraph([textRun('{{ context.empty! }}')]),
      makeParagraph([textRun('And this')]),
    ]);

    renderDocumentWithBookmarks(doc, { tags: {}, ctMeta: ctMetaRie, mode: 'omit' });

    expect(doc.package.document.content).toHaveLength(2);
    const texts1 = getRunTexts((doc.package.document.content[0] as Paragraph).content);
    const texts2 = getRunTexts((doc.package.document.content[1] as Paragraph).content);
    expect(texts1).toEqual(['Keep this']);
    expect(texts2).toEqual(['And this']);
  });

  test('preserves context tag metadata on document', () => {
    const doc = makeDocument([makeParagraph([textRun('{{ context.case_no }}')])]);

    renderDocumentWithBookmarks(doc, { tags, ctMeta, mode: 'keep' });

    expect(doc.contextTagMetadata).toBe(ctMeta);
  });

  test('bookmark IDs do not collide with existing bookmarks', () => {
    const doc = makeDocument([
      makeParagraph([
        { type: 'bookmarkStart', id: 50, name: '_GoBack' } as BookmarkStart,
        textRun('{{ context.case_no }}'),
        { type: 'bookmarkEnd', id: 50 } as BookmarkEnd,
      ]),
    ]);

    renderDocumentWithBookmarks(doc, { tags, ctMeta, mode: 'keep' });

    const starts = getBookmarkStarts((doc.package.document.content[0] as Paragraph).content);
    const fpStarts = starts.filter((s) => s.name.startsWith(FP_BOOKMARK_PREFIX));
    expect(fpStarts).toHaveLength(1);
    expect(fpStarts[0].id).toBeGreaterThan(50);
  });
});

// ── restoreParagraphContent ─────────────────────────────────────

describe('restoreParagraphContent', () => {
  const { restoreParagraphContent } = require('./renderWithBookmarks');

  const manifest: Record<string, ContextTagMeta> = {
    'uuid-a': { tagKey: 'context.case_no' },
    'uuid-b': { tagKey: 'context.vessel', removeIfEmpty: true },
  };

  test('restores bookmarked text to tag pattern', () => {
    const content: ParagraphContent[] = [
      { type: 'bookmarkStart', id: 100, name: '_FP_ctx_uuid-a' } as BookmarkStart,
      textRun('AMA7089'),
      { type: 'bookmarkEnd', id: 100 } as BookmarkEnd,
    ];

    const result = restoreParagraphContent(content, manifest);
    const texts = getRunTexts(result);
    expect(texts).toEqual(['{{ context.case_no }}']);
    expect(getBookmarkStarts(result)).toHaveLength(0);
    expect(getBookmarkEnds(result)).toHaveLength(0);
  });

  test('preserves removeIfEmpty flag in restored tag', () => {
    const content: ParagraphContent[] = [
      { type: 'bookmarkStart', id: 100, name: '_FP_ctx_uuid-b' } as BookmarkStart,
      textRun('ORANGE PHOENIX'),
      { type: 'bookmarkEnd', id: 100 } as BookmarkEnd,
    ];

    const result = restoreParagraphContent(content, manifest);
    const texts = getRunTexts(result);
    expect(texts).toEqual(['{{ context.vessel! }}']);
  });

  test('preserves non-FP bookmarks', () => {
    const content: ParagraphContent[] = [
      { type: 'bookmarkStart', id: 1, name: '_GoBack' } as BookmarkStart,
      textRun('Regular text'),
      { type: 'bookmarkEnd', id: 1 } as BookmarkEnd,
    ];

    const result = restoreParagraphContent(content, manifest);
    expect(result).toHaveLength(3);
    expect((result[0] as BookmarkStart).name).toBe('_GoBack');
  });

  test('graceful degradation: no manifest entry for metaId', () => {
    const content: ParagraphContent[] = [
      textRun('Before '),
      { type: 'bookmarkStart', id: 100, name: '_FP_ctx_unknown-id' } as BookmarkStart,
      textRun('some rendered text'),
      { type: 'bookmarkEnd', id: 100 } as BookmarkEnd,
      textRun(' after'),
    ];

    const result = restoreParagraphContent(content, manifest);
    const texts = getRunTexts(result);
    // Unknown metaId — bookmarked text is dropped (no manifest entry)
    expect(texts).toEqual(['Before ', ' after']);
  });

  test('preserves formatting from bookmarked run', () => {
    const bold = { bold: true, fontSize: 24 };
    const content: ParagraphContent[] = [
      { type: 'bookmarkStart', id: 100, name: '_FP_ctx_uuid-a' } as BookmarkStart,
      textRun('AMA7089', bold),
      { type: 'bookmarkEnd', id: 100 } as BookmarkEnd,
    ];

    const result = restoreParagraphContent(content, manifest);
    const runs = getRuns(result);
    expect(runs).toHaveLength(1);
    expect(runs[0].formatting?.bold).toBe(true);
    expect(runs[0].formatting?.fontSize).toBe(24);
  });

  test('multiple FP bookmarks in one paragraph', () => {
    const content: ParagraphContent[] = [
      textRun('Case: '),
      { type: 'bookmarkStart', id: 100, name: '_FP_ctx_uuid-a' } as BookmarkStart,
      textRun('AMA7089'),
      { type: 'bookmarkEnd', id: 100 } as BookmarkEnd,
      textRun(' Vessel: '),
      { type: 'bookmarkStart', id: 101, name: '_FP_ctx_uuid-b' } as BookmarkStart,
      textRun('ORANGE PHOENIX'),
      { type: 'bookmarkEnd', id: 101 } as BookmarkEnd,
    ];

    const result = restoreParagraphContent(content, manifest);
    const texts = getRunTexts(result);
    expect(texts).toEqual([
      'Case: ',
      '{{ context.case_no }}',
      ' Vessel: ',
      '{{ context.vessel! }}',
    ]);
  });

  test('mixed FP and non-FP bookmarks', () => {
    const content: ParagraphContent[] = [
      { type: 'bookmarkStart', id: 1, name: '_GoBack' } as BookmarkStart,
      { type: 'bookmarkStart', id: 100, name: '_FP_ctx_uuid-a' } as BookmarkStart,
      textRun('AMA7089'),
      { type: 'bookmarkEnd', id: 100 } as BookmarkEnd,
      { type: 'bookmarkEnd', id: 1 } as BookmarkEnd,
    ];

    const result = restoreParagraphContent(content, manifest);
    // _GoBack bookmark preserved, FP bookmark consumed
    const starts = getBookmarkStarts(result);
    expect(starts).toHaveLength(1);
    expect(starts[0].name).toBe('_GoBack');
    const texts = getRunTexts(result);
    expect(texts).toEqual(['{{ context.case_no }}']);
  });

  test('no FP bookmarks returns content unchanged', () => {
    const content: ParagraphContent[] = [textRun('Plain text')];
    const result = restoreParagraphContent(content, manifest);
    expect(result).toBe(content); // same reference — no processing
  });
});

// ── restoreContextTagsFromBookmarks (document-level) ────────────

describe('restoreContextTagsFromBookmarks', () => {
  const { restoreContextTagsFromBookmarks } = require('./renderWithBookmarks');

  test('restores tags across multiple paragraphs', () => {
    const manifest: Record<string, ContextTagMeta> = {
      'uuid-a': { tagKey: 'context.case_no' },
      'uuid-b': { tagKey: 'context.vessel' },
    };
    const doc = makeDocument([
      makeParagraph([
        { type: 'bookmarkStart', id: 100, name: '_FP_ctx_uuid-a' } as BookmarkStart,
        textRun('AMA7089'),
        { type: 'bookmarkEnd', id: 100 } as BookmarkEnd,
      ]),
      makeParagraph([
        { type: 'bookmarkStart', id: 101, name: '_FP_ctx_uuid-b' } as BookmarkStart,
        textRun('ORANGE PHOENIX'),
        { type: 'bookmarkEnd', id: 101 } as BookmarkEnd,
      ]),
    ]);
    doc.contextTagMetadata = manifest;

    restoreContextTagsFromBookmarks(doc);

    const p1 = doc.package.document.content[0] as Paragraph;
    const p2 = doc.package.document.content[1] as Paragraph;
    expect(getRunTexts(p1.content)).toEqual(['{{ context.case_no }}']);
    expect(getRunTexts(p2.content)).toEqual(['{{ context.vessel }}']);
  });

  test('no-op when no manifest present', () => {
    const doc = makeDocument([makeParagraph([textRun('Hello')])]);
    // No contextTagMetadata set
    restoreContextTagsFromBookmarks(doc);
    expect(getRunTexts((doc.package.document.content[0] as Paragraph).content)).toEqual(['Hello']);
  });
});

// ── Round-trip integration tests ────────────────────────────────

describe('round-trip: render → restore', () => {
  test('render then restore produces tag patterns', () => {
    const ctMeta: Record<string, ContextTagMeta> = {
      'uuid-rt1': { tagKey: 'context.case_no' },
      'uuid-rt2': { tagKey: 'context.vessel', removeIfEmpty: true },
    };
    const tags: Record<string, string> = {
      'context.case_no': 'AMA7089',
      'context.vessel': 'ORANGE PHOENIX',
    };

    // Step 1: Create document with tag patterns
    const doc = makeDocument([
      makeParagraph([textRun('Case: {{ context.case_no }}')]),
      makeParagraph([textRun('Vessel: {{ context.vessel! }}')]),
    ]);

    // Step 2: Render with bookmarks
    renderDocumentWithBookmarks(doc, { tags, ctMeta, mode: 'keep' });

    // Verify rendered text is present
    const p1After = doc.package.document.content[0] as Paragraph;
    expect(getRunTexts(p1After.content)).toEqual(['Case: ', 'AMA7089']);

    // Step 3: Restore from bookmarks
    const { restoreContextTagsFromBookmarks } = require('./renderWithBookmarks');
    restoreContextTagsFromBookmarks(doc);

    // Verify tag patterns are restored
    const p1Restored = doc.package.document.content[0] as Paragraph;
    const p2Restored = doc.package.document.content[1] as Paragraph;
    expect(getRunTexts(p1Restored.content)).toEqual(['Case: ', '{{ context.case_no }}']);
    expect(getRunTexts(p2Restored.content)).toEqual(['Vessel: ', '{{ context.vessel! }}']);
  });

  test('render with omit mode then restore', () => {
    const ctMeta: Record<string, ContextTagMeta> = {
      'uuid-o1': { tagKey: 'context.case_no' },
      'uuid-o2': { tagKey: 'context.empty', removeIfEmpty: true },
    };
    const tags: Record<string, string> = {
      'context.case_no': 'AMA7089',
    };

    const doc = makeDocument([
      makeParagraph([textRun('Case: {{ context.case_no }}')]),
      makeParagraph([textRun('{{ context.empty! }}')]),
      makeParagraph([textRun('End')]),
    ]);

    // Render in omit mode — removeIfEmpty paragraph gets removed
    renderDocumentWithBookmarks(doc, { tags, ctMeta, mode: 'omit' });
    expect(doc.package.document.content).toHaveLength(2);

    // Restore — the surviving bookmark should restore the tag
    const { restoreContextTagsFromBookmarks } = require('./renderWithBookmarks');
    restoreContextTagsFromBookmarks(doc);

    const p1 = doc.package.document.content[0] as Paragraph;
    expect(getRunTexts(p1.content)).toEqual(['Case: ', '{{ context.case_no }}']);
  });

  test('same tagKey rendered twice preserves both via metaId', () => {
    const ctMeta: Record<string, ContextTagMeta> = {
      'uuid-d1': { tagKey: 'context.case_no' },
      'uuid-d2': { tagKey: 'context.case_no' },
    };
    const tags: Record<string, string> = {
      'context.case_no': 'AMA7089',
    };

    const doc = makeDocument([
      makeParagraph([textRun('{{ context.case_no }}')]),
      makeParagraph([textRun('{{ context.case_no }}')]),
    ]);

    renderDocumentWithBookmarks(doc, { tags, ctMeta, mode: 'keep' });

    // Both paragraphs should have bookmarks with different metaIds
    const p1 = doc.package.document.content[0] as Paragraph;
    const p2 = doc.package.document.content[1] as Paragraph;
    const s1 = getBookmarkStarts(p1.content);
    const s2 = getBookmarkStarts(p2.content);
    expect(s1[0].name).toBe('_FP_ctx_uuid-d1');
    expect(s2[0].name).toBe('_FP_ctx_uuid-d2');

    // Restore
    const { restoreContextTagsFromBookmarks } = require('./renderWithBookmarks');
    restoreContextTagsFromBookmarks(doc);

    const p1r = doc.package.document.content[0] as Paragraph;
    const p2r = doc.package.document.content[1] as Paragraph;
    expect(getRunTexts(p1r.content)).toEqual(['{{ context.case_no }}']);
    expect(getRunTexts(p2r.content)).toEqual(['{{ context.case_no }}']);
  });
});

// ── findMaxBookmarkId ───────────────────────────────────────────

describe('findMaxBookmarkId', () => {
  test('returns 0 for document with no bookmarks', () => {
    const doc = makeDocument([makeParagraph([textRun('Hello')])]);
    expect(findMaxBookmarkId(doc)).toBe(0);
  });

  test('finds max ID across paragraphs', () => {
    const doc = makeDocument([
      makeParagraph([
        { type: 'bookmarkStart', id: 10, name: 'a' } as BookmarkStart,
        textRun('Hello'),
        { type: 'bookmarkEnd', id: 10 } as BookmarkEnd,
      ]),
      makeParagraph([
        { type: 'bookmarkStart', id: 42, name: 'b' } as BookmarkStart,
        textRun('World'),
        { type: 'bookmarkEnd', id: 42 } as BookmarkEnd,
      ]),
    ]);
    expect(findMaxBookmarkId(doc)).toBe(42);
  });
});
