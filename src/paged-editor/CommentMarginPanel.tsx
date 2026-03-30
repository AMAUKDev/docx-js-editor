/**
 * CommentMarginPanel — renders comment cards alongside pages at the Y-position
 * of their anchored text, similar to Microsoft Word's comment margin.
 *
 * Uses buildCommentTree for sorting by document position and recursive
 * reply threading. Supports collapse/expand of reply threads.
 */
import React, { useEffect, useState, useCallback, useRef, useLayoutEffect } from 'react';
import type { Comment } from '../types/content';
import type { EditorView } from 'prosemirror-view';
import { buildCommentTree, type CommentTreeNode } from './commentTree';

export interface CommentCardData {
  id: number;
  author: string;
  date?: string;
  text: string;
  anchorText: string;
  top: number; // Y position relative to pages container
  done?: boolean;
  parentId?: number;
  isReply?: boolean;
  depth: number;
  childCount: number; // total descendants
  collapsed: boolean; // whether this node's children are hidden
}

export interface CommentMarginPanelProps {
  /** The pages container element for coordinate mapping */
  pagesContainer: HTMLElement | null;
  /** ProseMirror EditorView for querying comment marks */
  view: EditorView | null;
  /** Comment objects from the document model */
  comments: Comment[];
  /** Whether to show the panel */
  visible: boolean;
  /** Callback when user clicks Reply on a comment */
  onReply?: (commentId: number) => void;
  /** Callback when user clicks Resolve on a comment */
  onResolve?: (commentId: number) => void;
  /** Callback when user clicks Delete on a comment */
  onDelete?: (commentId: number) => void;
  /** Callback when user clicks Edit on a comment */
  onEdit?: (commentId: number) => void;
}

/**
 * Extract text from a Comment's content paragraphs.
 */
function getCommentText(comment: Comment): string {
  if (!comment.content) return '';
  const parts: string[] = [];
  for (const para of comment.content) {
    for (const item of para.content || []) {
      if ('content' in item) {
        for (const rc of (item as { content: Array<{ type: string; text?: string }> }).content) {
          if (rc.type === 'text' && rc.text) parts.push(rc.text);
        }
      }
    }
  }
  return parts.join('');
}

/**
 * Flatten a comment tree into a list of cards, respecting collapsed state.
 */
function flattenTreeToCards(
  tree: CommentTreeNode[],
  collapsed: Set<number>,
  anchorMap: Map<number, string>
): Omit<CommentCardData, 'top'>[] {
  const result: Omit<CommentCardData, 'top'>[] = [];

  function walk(nodes: CommentTreeNode[]) {
    for (const node of nodes) {
      const isCollapsed = collapsed.has(node.comment.id) && node.children.length > 0;
      result.push({
        id: node.comment.id,
        author: node.comment.author || 'Unknown',
        date: node.comment.date,
        text: getCommentText(node.comment),
        anchorText: node.depth === 0 ? (anchorMap.get(node.comment.id) || '').slice(0, 60) : '',
        done: node.comment.done,
        parentId: node.comment.parentId,
        isReply: node.depth > 0,
        depth: node.depth,
        childCount: node.childCount,
        collapsed: isCollapsed,
      });
      if (!isCollapsed) {
        walk(node.children);
      }
    }
  }

  walk(tree);
  return result;
}

export const CommentMarginPanel: React.FC<CommentMarginPanelProps> = ({
  pagesContainer,
  view,
  comments,
  visible,
  onReply,
  onResolve,
  onDelete,
  onEdit,
}) => {
  const [cards, setCards] = useState<CommentCardData[]>([]);
  const [collapsedIds, setCollapsedIds] = useState<Set<number>>(new Set());
  const panelRef = useRef<HTMLDivElement>(null);

  const toggleCollapse = useCallback((commentId: number) => {
    setCollapsedIds((prev) => {
      const next = new Set(prev);
      if (next.has(commentId)) {
        next.delete(commentId);
      } else {
        next.add(commentId);
      }
      return next;
    });
  }, []);

  const computePositions = useCallback(() => {
    if (!pagesContainer || !view || !comments || comments.length === 0) {
      setCards([]);
      return;
    }

    const scrollParent = pagesContainer.closest('.paged-editor') as HTMLElement;
    const containerRect = scrollParent
      ? scrollParent.getBoundingClientRect()
      : pagesContainer.getBoundingClientRect();

    // Step 1: Find PM position ranges for each comment mark
    const commentRanges = new Map<number, { from: number; to: number }>();
    view.state.doc.descendants((node, pos) => {
      for (const mark of node.marks) {
        if (mark.type.name === 'comment') {
          const cid = mark.attrs.commentId as number;
          const existing = commentRanges.get(cid);
          if (!existing) {
            commentRanges.set(cid, { from: pos, to: pos + node.nodeSize });
          } else {
            existing.to = Math.max(existing.to, pos + node.nodeSize);
          }
        }
      }
    });

    // Step 2: Build anchor text map
    const anchorMap = new Map<number, string>();
    view.state.doc.descendants((node) => {
      if (node.isText) {
        for (const mark of node.marks) {
          if (mark.type.name === 'comment') {
            const cid = mark.attrs.commentId as number;
            anchorMap.set(cid, (anchorMap.get(cid) || '') + (node.text || ''));
          }
        }
      }
    });

    // Step 3: Build sorted tree and flatten
    const tree = buildCommentTree(comments, commentRanges, collapsedIds);
    const flatCards = flattenTreeToCards(tree, collapsedIds, anchorMap);

    // Step 4: Assign Y positions
    const spans = pagesContainer.querySelectorAll('span[data-pm-start][data-pm-end]');
    const results: CommentCardData[] = [];

    for (const card of flatCards) {
      let top = results.length * 80; // fallback

      if (card.depth === 0) {
        // Top-level: position at the anchored text
        const range = commentRanges.get(card.id);
        if (range) {
          for (const span of Array.from(spans)) {
            const pmStart = Number((span as HTMLElement).dataset.pmStart);
            const pmEnd = Number((span as HTMLElement).dataset.pmEnd);
            if (pmStart < range.to && pmEnd > range.from) {
              const rect = (span as HTMLElement).getBoundingClientRect();
              const scrollTop = scrollParent ? scrollParent.scrollTop : 0;
              top = rect.top - containerRect.top + scrollTop;
              break;
            }
          }
        }
      } else {
        // Reply: stack beneath previous card
        if (results.length > 0) {
          top = results[results.length - 1].top + 70;
        }
      }

      // Prevent overlapping
      if (results.length > 0) {
        const prevBottom = results[results.length - 1].top + 70;
        if (top < prevBottom) top = prevBottom;
      }

      results.push({ ...card, top });
    }

    setCards(results);
  }, [pagesContainer, view, comments, collapsedIds]);

  // Recompute positions when comments change or panel becomes visible
  useEffect(() => {
    if (!visible) return;
    const t1 = setTimeout(computePositions, 200);
    const t2 = setTimeout(computePositions, 800);
    const t3 = setTimeout(computePositions, 2000);
    return () => {
      clearTimeout(t1);
      clearTimeout(t2);
      clearTimeout(t3);
    };
  }, [visible, comments, computePositions, view]);

  // After render: measure actual card heights and reflow to prevent overlap
  useLayoutEffect(() => {
    if (!panelRef.current || cards.length === 0) return;
    const cardEls = panelRef.current.querySelectorAll<HTMLElement>('[data-testid="comment-card"]');
    if (cardEls.length !== cards.length) return;

    const GAP = 6; // px between cards
    let needsUpdate = false;
    const newCards = [...cards];

    for (let i = 1; i < newCards.length; i++) {
      const prevEl = cardEls[i - 1];
      const prevHeight = prevEl.offsetHeight;
      const prevBottom = newCards[i - 1].top + prevHeight + GAP;

      if (newCards[i].top < prevBottom) {
        newCards[i] = { ...newCards[i], top: prevBottom };
        needsUpdate = true;
      }
    }

    if (needsUpdate) {
      setCards(newCards);
    }
  }, [cards]);

  // Recompute on scroll
  useEffect(() => {
    if (!visible || !pagesContainer) return;
    const scrollParent = pagesContainer.closest('.paged-editor') as HTMLElement;
    if (!scrollParent) return;
    const onScroll = () => computePositions();
    scrollParent.addEventListener('scroll', onScroll, { passive: true });
    return () => scrollParent.removeEventListener('scroll', onScroll);
  }, [visible, pagesContainer, computePositions]);

  if (!visible || cards.length === 0) return null;

  const INDENT_PER_DEPTH = 16;

  return (
    <div
      ref={panelRef}
      className="comment-margin-panel"
      style={{
        position: 'absolute',
        top: 0,
        right: 8,
        width: 230,
        pointerEvents: 'auto',
        zIndex: 20,
      }}
    >
      {cards.map((card) => {
        const indent = card.depth * INDENT_PER_DEPTH;
        const cardWidth = 220 - indent;

        return (
          <div
            key={card.id}
            data-testid="comment-card"
            data-depth={card.depth}
            style={{
              position: 'absolute',
              top: card.top,
              left: indent,
              width: cardWidth,
              padding: '6px 8px',
              background: card.done ? '#e8e8e8' : card.depth > 0 ? '#fff8e1' : '#fffde7',
              border: card.depth > 0 ? '1px solid #e8d87a' : '1px solid #e0d68a',
              borderLeft: card.depth > 0 ? '3px solid #ffc107' : '1px solid #e0d68a',
              borderRadius: 4,
              fontSize: '11px',
              lineHeight: '1.3',
              boxShadow: '0 1px 3px rgba(0,0,0,0.1)',
              opacity: card.done ? 0.6 : 1,
            }}
          >
            {/* Author + date + collapse toggle */}
            <div
              style={{
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center',
                marginBottom: 2,
              }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: 3 }}>
                {card.childCount > 0 && (
                  <button
                    type="button"
                    data-testid="collapse-toggle"
                    aria-label={card.collapsed ? 'expand replies' : 'collapse replies'}
                    onClick={(e) => {
                      e.stopPropagation();
                      toggleCollapse(card.id);
                    }}
                    style={{
                      background: 'none',
                      border: 'none',
                      cursor: 'pointer',
                      fontSize: '10px',
                      padding: 0,
                      lineHeight: 1,
                      color: '#666',
                    }}
                  >
                    {card.collapsed ? '▶' : '▼'}
                  </button>
                )}
                <strong style={{ fontSize: '10px', color: '#333' }}>{card.author}</strong>
              </div>
              {card.date && (
                <span style={{ fontSize: '9px', color: '#999' }}>
                  {new Date(card.date).toLocaleDateString()}
                </span>
              )}
            </div>

            {/* Reply count badge when collapsed */}
            {card.collapsed && card.childCount > 0 && (
              <div
                data-testid="reply-count"
                className="reply-count-badge"
                style={{
                  fontSize: '9px',
                  color: '#888',
                  fontStyle: 'italic',
                  marginBottom: 3,
                }}
              >
                {card.childCount} {card.childCount === 1 ? 'reply' : 'replies'}
              </div>
            )}

            {/* Anchor text */}
            {card.anchorText && (
              <div
                style={{
                  borderLeft: '2px solid #ffd700',
                  paddingLeft: 4,
                  fontSize: '10px',
                  color: '#666',
                  fontStyle: 'italic',
                  marginBottom: 3,
                  whiteSpace: 'nowrap',
                  overflow: 'hidden',
                  textOverflow: 'ellipsis',
                }}
              >
                &ldquo;{card.anchorText}&rdquo;
              </div>
            )}

            {/* Comment text */}
            <div style={{ color: '#333', marginBottom: 4 }}>{card.text}</div>

            {/* Action buttons */}
            {
              <div style={{ display: 'flex', gap: 4 }}>
                {!card.done && onResolve && (
                  <button
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      onResolve(card.id);
                    }}
                    style={{
                      fontSize: '9px',
                      padding: '1px 5px',
                      border: '1px solid #28a745',
                      borderRadius: 3,
                      background: 'white',
                      color: '#28a745',
                      cursor: 'pointer',
                    }}
                  >
                    Resolve
                  </button>
                )}
                {onEdit && (
                  <button
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      onEdit(card.id);
                    }}
                    style={{
                      fontSize: '9px',
                      padding: '1px 5px',
                      border: '1px solid #007bff',
                      borderRadius: 3,
                      background: 'white',
                      color: '#007bff',
                      cursor: 'pointer',
                    }}
                  >
                    Edit
                  </button>
                )}
                {onReply && (
                  <button
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      onReply(card.id);
                    }}
                    style={{
                      fontSize: '9px',
                      padding: '1px 5px',
                      border: '1px solid #6c757d',
                      borderRadius: 3,
                      background: 'white',
                      color: '#6c757d',
                      cursor: 'pointer',
                    }}
                  >
                    Reply
                  </button>
                )}
                {onDelete && (
                  <button
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      onDelete(card.id);
                    }}
                    style={{
                      fontSize: '9px',
                      padding: '1px 5px',
                      border: '1px solid #dc3545',
                      borderRadius: 3,
                      background: 'white',
                      color: '#dc3545',
                      cursor: 'pointer',
                    }}
                  >
                    Delete
                  </button>
                )}
              </div>
            }
          </div>
        );
      })}
    </div>
  );
};
