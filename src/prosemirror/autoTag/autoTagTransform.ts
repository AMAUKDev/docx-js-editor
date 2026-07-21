/**
 * Auto context-tag transform.
 *
 * Wraps loose text that matches a tag value in `contextTag` nodes, for ANY tag key
 * (including canonical `case.*` keys — the node schema places no restriction on
 * `tagKey`). This is the fork-native core of the auto-tagging feature: it runs over
 * a ProseMirror document at import, on typing/paste (via a plugin), or on demand.
 *
 * Positions are computed against the passed EditorState's doc; the returned
 * transaction replaces matched ranges from the end backwards so earlier positions
 * stay valid.
 */

import type { EditorState, Transaction } from 'prosemirror-state';
import type { Mark } from 'prosemirror-model';
import { generateMetaId } from '../extensions/nodes/ContextTagExtension';
import {
  findAutoTagCandidates,
  type AutoTagCandidate,
  type AutoTagMatchOptions,
} from './autoTagMatcher';

export interface AutoTagHit extends AutoTagCandidate {
  /** Absolute document position of the match start. */
  from: number;
  /** Absolute document position of the match end. */
  to: number;
}

export interface BuildAutoTagOptions extends AutoTagMatchOptions {
  /**
   * Convert low-confidence (short / ambiguous) candidates without confirmation.
   * Default false — low-confidence hits are returned in `deferred` for a confirm step.
   */
  applyLowConfidence?: boolean;
  /**
   * Fine-grained accept predicate (used to apply a user's confirm-step choices).
   * When provided it overrides the confidence heuristic for every hit.
   */
  accept?: (hit: AutoTagHit) => boolean;
  /** Restrict scanning to hits fully inside [rangeFrom, rangeTo) (used for typing). */
  rangeFrom?: number;
  rangeTo?: number;
}

export interface AutoTagResult {
  /** Transaction that applies the accepted conversions, or null if nothing applied. */
  tr: Transaction | null;
  /** Candidates that were converted into tag nodes. */
  applied: AutoTagHit[];
  /** Low-confidence candidates left as literal text (awaiting confirmation). */
  deferred: AutoTagHit[];
}

const EMPTY_RESULT: AutoTagResult = { tr: null, applied: [], deferred: [] };

/** Collect every hit in the document (before accept/confidence filtering). */
export function collectAutoTagHits(
  state: EditorState,
  tagMap: Record<string, string | null | undefined> | null | undefined,
  options: BuildAutoTagOptions = {},
): AutoTagHit[] {
  if (!state.schema.nodes.contextTag || !tagMap) return [];
  const { rangeFrom, rangeTo } = options;
  const hits: AutoTagHit[] = [];
  state.doc.descendants((node, pos) => {
    if (!node.isText || !node.text) return;
    const candidates = findAutoTagCandidates(node.text, tagMap, options);
    for (const candidate of candidates) {
      const from = pos + candidate.start;
      const to = pos + candidate.end;
      if (rangeFrom != null && from < rangeFrom) continue;
      if (rangeTo != null && to > rangeTo) continue;
      hits.push({ ...candidate, from, to });
    }
  });
  return hits;
}

/**
 * Build a transaction that converts matching loose text into `contextTag` nodes.
 * By default only high-confidence hits are applied; low-confidence hits are deferred.
 */
export function buildAutoTagTransaction(
  state: EditorState,
  tagMap: Record<string, string | null | undefined> | null | undefined,
  options: BuildAutoTagOptions = {},
): AutoTagResult {
  const nodeType = state.schema.nodes.contextTag;
  if (!nodeType) return EMPTY_RESULT;

  const hits = collectAutoTagHits(state, tagMap, options);
  if (hits.length === 0) return EMPTY_RESULT;

  const applied: AutoTagHit[] = [];
  const deferred: AutoTagHit[] = [];
  for (const hit of hits) {
    let take: boolean;
    if (options.accept) take = options.accept(hit);
    else if (options.applyLowConfidence) take = true;
    else take = hit.confidence === 'high';
    (take ? applied : deferred).push(hit);
  }

  if (applied.length === 0) return { tr: null, applied: [], deferred };

  // Apply from the end backwards so earlier positions remain valid.
  const ordered = [...applied].sort((a, b) => b.from - a.from);
  let tr = state.tr;
  for (const hit of ordered) {
    const marks = state.doc
      .resolve(hit.from)
      .marks()
      .filter((mark: Mark) => nodeType.allowsMarkType(mark.type));
    const tagNode = nodeType.create(
      {
        tagKey: hit.tagKey,
        label: hit.value,
        removeIfEmpty: false,
        metaId: generateMetaId(),
      },
      null,
      marks,
    );
    tr = tr.replaceWith(hit.from, hit.to, tagNode);
  }
  tr.setMeta('allowLockedEdit', true);
  tr.setMeta('autoTag', true);
  return { tr, applied, deferred };
}
