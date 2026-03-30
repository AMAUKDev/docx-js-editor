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

  // Resolve each comment's root ancestor (Word only supports 2 levels: parent + replies).
  // Any reply-to-reply gets reparented to the root ancestor for flat threading.
  function getRootParentId(commentId: number): number | undefined {
    let current = byId.get(commentId);
    let visited = 0;
    while (current?.parentId != null && visited < 100) {
      const parent = byId.get(current.parentId);
      if (!parent || parent.parentId == null) return current.parentId;
      current = parent;
      visited++;
    }
    return current?.parentId;
  }

  // Group children by their ROOT parent (flattened to max 2 levels)
  const childrenOf = new Map<number, Comment[]>();
  const topLevel: Comment[] = [];

  for (const c of comments) {
    const rootParent = c.parentId != null ? getRootParentId(c.id) : undefined;
    if (rootParent != null && byId.has(rootParent)) {
      const siblings = childrenOf.get(rootParent) || [];
      siblings.push(c);
      childrenOf.set(rootParent, siblings);
    } else if (c.parentId != null && byId.has(c.parentId)) {
      // Direct parent exists and IS top-level
      const siblings = childrenOf.get(c.parentId) || [];
      siblings.push(c);
      childrenOf.set(c.parentId, siblings);
    } else {
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
