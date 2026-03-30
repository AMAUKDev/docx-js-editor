/**
 * Unit tests for comment tree building, sorting, and collapse logic.
 *
 * Tests the pure computation that orders comments by document position,
 * builds recursive reply trees, and handles collapse/expand state.
 */

import { describe, test, expect } from 'bun:test';
import { buildCommentTree, type CommentTreeNode } from './commentTree';
import type { Comment } from '../types/content';

// Helper: create a minimal Comment object
function makeComment(
  id: number,
  author: string,
  text: string,
  opts?: { parentId?: number; date?: string; done?: boolean }
): Comment {
  return {
    id,
    author,
    date: opts?.date ?? '2026-01-01T00:00:00Z',
    content: [
      {
        type: 'paragraph' as const,
        content: [{ type: 'run' as const, content: [{ type: 'text' as const, text }] }],
        formatting: {},
      },
    ],
    parentId: opts?.parentId,
    done: opts?.done,
  };
}

describe('buildCommentTree', () => {
  describe('sorting by document position', () => {
    test('top-level comments are sorted by document position ascending', () => {
      const comments = [
        makeComment(3, 'Alice', 'Third in doc'),
        makeComment(1, 'Bob', 'First in doc'),
        makeComment(2, 'Carol', 'Second in doc'),
      ];
      // commentRanges maps comment ID -> PM position range
      const commentRanges = new Map<number, { from: number; to: number }>([
        [1, { from: 10, to: 20 }], // earliest in doc
        [2, { from: 50, to: 60 }], // middle
        [3, { from: 100, to: 110 }], // latest in doc
      ]);

      const tree = buildCommentTree(comments, commentRanges, new Set());

      expect(tree.map((n) => n.comment.id)).toEqual([1, 2, 3]);
    });

    test('comments without range markers sort to the end', () => {
      const comments = [
        makeComment(1, 'Alice', 'Has range'),
        makeComment(2, 'Bob', 'No range'),
        makeComment(3, 'Carol', 'Also has range'),
      ];
      const commentRanges = new Map<number, { from: number; to: number }>([
        [1, { from: 50, to: 60 }],
        [3, { from: 10, to: 20 }],
        // comment 2 has no range
      ]);

      const tree = buildCommentTree(comments, commentRanges, new Set());

      // 3 first (pos 10), then 1 (pos 50), then 2 (no range = end)
      expect(tree.map((n) => n.comment.id)).toEqual([3, 1, 2]);
    });
  });

  describe('reply threading', () => {
    test('replies are nested beneath their parent', () => {
      const comments = [
        makeComment(1, 'Alice', 'Parent comment'),
        makeComment(2, 'Bob', 'Reply to parent', { parentId: 1 }),
      ];
      const commentRanges = new Map([[1, { from: 10, to: 20 }]]);

      const tree = buildCommentTree(comments, commentRanges, new Set());

      expect(tree).toHaveLength(1); // only 1 top-level node
      expect(tree[0].comment.id).toBe(1);
      expect(tree[0].children).toHaveLength(1);
      expect(tree[0].children[0].comment.id).toBe(2);
      expect(tree[0].children[0].depth).toBe(1);
    });

    test('replies to replies nest recursively', () => {
      const comments = [
        makeComment(1, 'Alice', 'Top level'),
        makeComment(2, 'Bob', 'Reply', { parentId: 1 }),
        makeComment(3, 'Carol', 'Reply to reply', { parentId: 2 }),
        makeComment(4, 'Dave', 'Reply to reply to reply', { parentId: 3 }),
      ];
      const commentRanges = new Map([[1, { from: 10, to: 20 }]]);

      const tree = buildCommentTree(comments, commentRanges, new Set());

      expect(tree).toHaveLength(1);
      expect(tree[0].children).toHaveLength(1);
      expect(tree[0].children[0].children).toHaveLength(1);
      expect(tree[0].children[0].children[0].children).toHaveLength(1);

      // Check depths
      expect(tree[0].depth).toBe(0);
      expect(tree[0].children[0].depth).toBe(1);
      expect(tree[0].children[0].children[0].depth).toBe(2);
      expect(tree[0].children[0].children[0].children[0].depth).toBe(3);
    });

    test('replies within same level are sorted by date (oldest first)', () => {
      const comments = [
        makeComment(1, 'Alice', 'Parent'),
        makeComment(4, 'Dave', 'Latest reply', { parentId: 1, date: '2026-03-03T00:00:00Z' }),
        makeComment(2, 'Bob', 'First reply', { parentId: 1, date: '2026-01-01T00:00:00Z' }),
        makeComment(3, 'Carol', 'Middle reply', { parentId: 1, date: '2026-02-02T00:00:00Z' }),
      ];
      const commentRanges = new Map([[1, { from: 10, to: 20 }]]);

      const tree = buildCommentTree(comments, commentRanges, new Set());

      const replyIds = tree[0].children.map((c) => c.comment.id);
      expect(replyIds).toEqual([2, 3, 4]); // sorted by date ascending
    });

    test('multiple top-level comments each have their own reply trees', () => {
      const comments = [
        makeComment(1, 'Alice', 'First parent'),
        makeComment(2, 'Bob', 'Second parent'),
        makeComment(3, 'Carol', 'Reply to first', { parentId: 1 }),
        makeComment(4, 'Dave', 'Reply to second', { parentId: 2 }),
      ];
      const commentRanges = new Map([
        [1, { from: 10, to: 20 }],
        [2, { from: 50, to: 60 }],
      ]);

      const tree = buildCommentTree(comments, commentRanges, new Set());

      expect(tree).toHaveLength(2);
      expect(tree[0].comment.id).toBe(1);
      expect(tree[0].children).toHaveLength(1);
      expect(tree[0].children[0].comment.id).toBe(3);
      expect(tree[1].comment.id).toBe(2);
      expect(tree[1].children).toHaveLength(1);
      expect(tree[1].children[0].comment.id).toBe(4);
    });
  });

  describe('descendant count', () => {
    test('childCount includes all descendants, not just direct children', () => {
      const comments = [
        makeComment(1, 'Alice', 'Root'),
        makeComment(2, 'Bob', 'Child', { parentId: 1 }),
        makeComment(3, 'Carol', 'Grandchild', { parentId: 2 }),
        makeComment(4, 'Dave', 'Another child', { parentId: 1 }),
      ];
      const commentRanges = new Map([[1, { from: 10, to: 20 }]]);

      const tree = buildCommentTree(comments, commentRanges, new Set());

      expect(tree[0].childCount).toBe(3); // 2 + 3 + 4
      expect(tree[0].children[0].childCount).toBe(1); // just 3
      expect(tree[0].children[1].childCount).toBe(0); // leaf
    });
  });

  describe('collapse/expand', () => {
    test('collapsed set hides descendants in flat output', () => {
      const comments = [
        makeComment(1, 'Alice', 'Root'),
        makeComment(2, 'Bob', 'Reply', { parentId: 1 }),
        makeComment(3, 'Carol', 'Reply to reply', { parentId: 2 }),
      ];
      const commentRanges = new Map([[1, { from: 10, to: 20 }]]);
      const collapsed = new Set([1]); // collapse root

      const tree = buildCommentTree(comments, commentRanges, collapsed);

      // Flatten the tree to get visible nodes
      const flat = flattenTree(tree, collapsed);
      expect(flat).toHaveLength(1); // only root visible
      expect(flat[0].comment.id).toBe(1);
      expect(flat[0].collapsed).toBe(true);
    });

    test('collapsing a mid-level node hides only its descendants', () => {
      const comments = [
        makeComment(1, 'Alice', 'Root'),
        makeComment(2, 'Bob', 'Reply', { parentId: 1 }),
        makeComment(3, 'Carol', 'Reply to reply', { parentId: 2 }),
        makeComment(4, 'Dave', 'Another reply', { parentId: 1 }),
      ];
      const commentRanges = new Map([[1, { from: 10, to: 20 }]]);
      const collapsed = new Set([2]); // collapse mid-level

      const tree = buildCommentTree(comments, commentRanges, collapsed);
      const flat = flattenTree(tree, collapsed);

      // Root, Reply (collapsed), Another reply — but NOT Reply to reply
      expect(flat.map((n) => n.comment.id)).toEqual([1, 2, 4]);
      expect(flat[1].collapsed).toBe(true);
    });

    test('expanding shows all descendants again', () => {
      const comments = [
        makeComment(1, 'Alice', 'Root'),
        makeComment(2, 'Bob', 'Reply', { parentId: 1 }),
        makeComment(3, 'Carol', 'Reply to reply', { parentId: 2 }),
      ];
      const commentRanges = new Map([[1, { from: 10, to: 20 }]]);
      const collapsed = new Set<number>(); // nothing collapsed

      const tree = buildCommentTree(comments, commentRanges, collapsed);
      const flat = flattenTree(tree, collapsed);

      expect(flat).toHaveLength(3);
      expect(flat.map((n) => n.comment.id)).toEqual([1, 2, 3]);
    });
  });

  describe('orphan handling', () => {
    test('replies with missing parent become top-level', () => {
      const comments = [
        makeComment(1, 'Alice', 'Normal comment'),
        makeComment(5, 'Eve', 'Reply to deleted parent', { parentId: 99 }),
      ];
      const commentRanges = new Map([[1, { from: 10, to: 20 }]]);

      const tree = buildCommentTree(comments, commentRanges, new Set());

      // Both should be top-level since parent 99 doesn't exist
      expect(tree).toHaveLength(2);
    });
  });
});

// Helper to flatten tree respecting collapsed state — this will also be exported
// from commentTree.ts, but we import it here to test
function flattenTree(
  tree: CommentTreeNode[],
  collapsed: Set<number>
): (CommentTreeNode & { collapsed: boolean })[] {
  const result: (CommentTreeNode & { collapsed: boolean })[] = [];
  function walk(nodes: CommentTreeNode[]) {
    for (const node of nodes) {
      const isCollapsed = collapsed.has(node.comment.id) && node.children.length > 0;
      result.push({ ...node, collapsed: isCollapsed });
      if (!isCollapsed) {
        walk(node.children);
      }
    }
  }
  walk(tree);
  return result;
}
