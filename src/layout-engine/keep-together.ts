/**
 * Keep Together Logic - Handle keepNext and keepLines paragraph properties
 *
 * DOCX paragraphs can have keepNext (keep with next paragraph) and keepLines
 * (keep all lines together) properties that affect pagination.
 */

import type { FlowBlock, ParagraphBlock, Measure, ParagraphMeasure } from './types';

/**
 * A chain of consecutive keepNext paragraphs.
 */
export type KeepNextChain = {
  /** Index of the first paragraph in the chain. */
  startIndex: number;
  /** Index of the last paragraph in the chain. */
  endIndex: number;
  /** All paragraph indices in the chain. */
  memberIndices: number[];
  /** Index of the anchor paragraph (first non-keepNext after chain), or -1 if none. */
  anchorIndex: number;
};

/**
 * Pre-scan blocks to find all keepNext chains.
 *
 * A keepNext chain is a sequence of consecutive paragraphs with keepNext=true,
 * followed by an anchor paragraph (the first non-keepNext paragraph).
 * The entire chain must stay on the same page as the anchor's first line.
 *
 * Returns a map from chain start index to chain info.
 */
export function computeKeepNextChains(blocks: FlowBlock[]): Map<number, KeepNextChain> {
  const chains = new Map<number, KeepNextChain>();

  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    // Only paragraphs can have keepNext
    if (block.kind !== 'paragraph') continue;

    const para = block as ParagraphBlock;
    // Skip paragraphs without keepNext
    if (!para.attrs?.keepNext) continue;

    // Found a keepNext paragraph — create a pair-wise chain (this paragraph
    // + the immediately following paragraph). Word implements keepNext as a
    // pair-wise relationship, NOT an extended chain. If headings A, B, C all
    // have keepNext, Word checks A+B fit, B+C fit, etc. individually — it
    // does NOT require A+B+C+anchor to fit together on one page.
    const memberIndices: number[] = [i];
    const endIndex = i;

    // Find the anchor (first paragraph after the chain)
    const potentialAnchor = endIndex + 1;
    let anchorIndex = -1;

    if (potentialAnchor < blocks.length) {
      const anchorBlock = blocks[potentialAnchor];
      // Anchor must not be a break
      if (
        anchorBlock.kind !== 'sectionBreak' &&
        anchorBlock.kind !== 'pageBreak' &&
        anchorBlock.kind !== 'columnBreak'
      ) {
        anchorIndex = potentialAnchor;
      }
    }

    // Record the chain
    chains.set(i, {
      startIndex: i,
      endIndex,
      memberIndices,
      anchorIndex,
    });
  }

  return chains;
}

/**
 * Calculate the total height needed to keep a chain together.
 *
 * Includes all chain members plus the first line of the anchor paragraph.
 */
export function calculateChainHeight(
  chain: KeepNextChain,
  blocks: FlowBlock[],
  measures: Measure[]
): number {
  let totalHeight = 0;

  // Sum heights of all chain members
  for (const memberIndex of chain.memberIndices) {
    const block = blocks[memberIndex];
    const measure = measures[memberIndex];

    if (block.kind !== 'paragraph' || measure.kind !== 'paragraph') continue;

    const para = block as ParagraphBlock;
    const paraMeasure = measure as ParagraphMeasure;

    // Add spacing before (simplified - could be more sophisticated with collapse)
    const spacingBefore = para.attrs?.spacing?.before ?? 0;
    totalHeight += spacingBefore;

    // Add paragraph height
    totalHeight += paraMeasure.totalHeight;

    // Add spacing after
    const spacingAfter = para.attrs?.spacing?.after ?? 0;
    totalHeight += spacingAfter;
  }

  // Add first line height of anchor (if any)
  if (chain.anchorIndex !== -1) {
    const anchorMeasure = measures[chain.anchorIndex];
    if (anchorMeasure?.kind === 'paragraph') {
      const anchorPara = anchorMeasure as ParagraphMeasure;
      if (anchorPara.lines.length > 0) {
        // Add just the first line height
        totalHeight += anchorPara.lines[0].lineHeight;
      }
    }
  }

  return totalHeight;
}

/**
 * Get the set of indices that are mid-chain (not chain starters).
 * These should skip the keepNext check since their chain starter already decided.
 */
export function getMidChainIndices(chains: Map<number, KeepNextChain>): Set<number> {
  const midChain = new Set<number>();

  for (const chain of chains.values()) {
    // All members except the first are mid-chain
    for (let i = 1; i < chain.memberIndices.length; i++) {
      midChain.add(chain.memberIndices[i]);
    }
  }

  return midChain;
}

/**
 * Check if a paragraph has keepLines property (all lines must stay together).
 */
export function hasKeepLines(block: FlowBlock): boolean {
  if (block.kind !== 'paragraph') return false;
  const para = block as ParagraphBlock;
  return para.attrs?.keepLines === true;
}

/**
 * Check if a paragraph should start on a new page (pageBreakBefore).
 */
export function hasPageBreakBefore(block: FlowBlock): boolean {
  if (block.kind !== 'paragraph') return false;
  const para = block as ParagraphBlock;
  return para.attrs?.pageBreakBefore === true;
}
