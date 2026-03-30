/**
 * Pure logic for building and flattening comment trees.
 *
 * Extracts the tree-building, sorting, and collapse logic from
 * CommentMarginPanel so it can be unit-tested independently.
 */

import type { Comment } from '../types/content';

export interface CommentTreeNode {
  comment: Comment;
  children: CommentTreeNode[];
  depth: number;
  childCount: number; // total descendants
}

/**
 * Build a sorted comment tree from a flat list of comments.
 *
 * - Top-level comments (no parentId) are sorted by document position (commentRanges)
 * - Replies are nested beneath their parent recursively
 * - Replies at the same level are sorted by date (oldest first)
 * - Comments without a range sort to the end
 * - Orphan replies (parent not found) become top-level
 */
export function buildCommentTree(
  comments: Comment[],
  commentRanges: Map<number, { from: number; to: number }>,
  _collapsed: Set<number>
): CommentTreeNode[] {
  // Index comments by ID for fast lookup
  const byId = new Map<number, Comment>();
  for (const c of comments) byId.set(c.id, c);

  // Group children by parentId
  const childrenOf = new Map<number, Comment[]>();
  const topLevel: Comment[] = [];

  for (const c of comments) {
    if (c.parentId != null && byId.has(c.parentId)) {
      // Valid parent exists — this is a reply
      const siblings = childrenOf.get(c.parentId) || [];
      siblings.push(c);
      childrenOf.set(c.parentId, siblings);
    } else {
      // No parent, or parent not found — treat as top-level
      topLevel.push(c);
    }
  }

  // Sort top-level by document position (ascending), rangeless at end
  topLevel.sort((a, b) => {
    const posA = commentRanges.get(a.id)?.from ?? Infinity;
    const posB = commentRanges.get(b.id)?.from ?? Infinity;
    return posA - posB;
  });

  // Sort each group of siblings by date (oldest first)
  for (const [, siblings] of childrenOf) {
    siblings.sort((a, b) => {
      const dateA = a.date ?? '';
      const dateB = b.date ?? '';
      return dateA < dateB ? -1 : dateA > dateB ? 1 : 0;
    });
  }

  // Recursively build tree nodes
  function buildNode(comment: Comment, depth: number): CommentTreeNode {
    const children = (childrenOf.get(comment.id) || []).map((c) => buildNode(c, depth + 1));
    const childCount = children.reduce((sum, c) => sum + 1 + c.childCount, 0);
    return { comment, children, depth, childCount };
  }

  return topLevel.map((c) => buildNode(c, 0));
}
