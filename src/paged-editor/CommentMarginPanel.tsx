/**
 * CommentMarginPanel — renders comment cards alongside pages at the Y-position
 * of their anchored text, similar to Microsoft Word's comment margin.
 */
import React, { useEffect, useState, useCallback } from 'react';
import type { Comment } from '../types/content';
import type { EditorView } from 'prosemirror-view';

export interface CommentCardData {
  id: number;
  author: string;
  date?: string;
  text: string;
  anchorText: string;
  top: number; // Y position relative to pages container
  done?: boolean;
  parentId?: number; // For reply threading — replies display indented beneath parent
  isReply?: boolean; // Convenience flag for rendering
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
  /** Callback when user edits a comment's text */
  onEdit?: (commentId: number, newText: string) => void;
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
  // Editing state: which comment ID is being edited, and the draft text
  const [editingId, setEditingId] = useState<number | null>(null);
  const [editText, setEditText] = useState('');

  const computePositions = useCallback(() => {
    if (!pagesContainer || !view || !comments || comments.length === 0) {
      setCards([]);
      return;
    }

    // Use the scroll parent (the .paged-editor div) as coordinate reference
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

    // Step 3: Find visible spans matching comment ranges
    const spans = pagesContainer.querySelectorAll('span[data-pm-start][data-pm-end]');
    const results: CommentCardData[] = [];

    // Separate top-level comments from replies
    const topLevel = comments.filter((c) => !c.parentId);
    const replies = comments.filter((c) => c.parentId);
    const repliesByParent = new Map<number, Comment[]>();
    for (const reply of replies) {
      const list = repliesByParent.get(reply.parentId!) || [];
      list.push(reply);
      repliesByParent.set(reply.parentId!, list);
    }

    // Process top-level comments with their replies
    for (const comment of topLevel) {
      const range = commentRanges.get(comment.id);
      if (!range) continue;

      let bestEl: HTMLElement | null = null;
      for (const span of Array.from(spans)) {
        const pmStart = Number((span as HTMLElement).dataset.pmStart);
        const pmEnd = Number((span as HTMLElement).dataset.pmEnd);
        if (pmStart < range.to && pmEnd > range.from) {
          bestEl = span as HTMLElement;
          break;
        }
      }

      let top = results.length * 80; // fallback: stack vertically
      if (bestEl) {
        const rect = bestEl.getBoundingClientRect();
        const scrollTop = scrollParent ? scrollParent.scrollTop : 0;
        top = rect.top - containerRect.top + scrollTop;
      }

      // Prevent overlapping: ensure minimum spacing from ALL previous cards
      if (results.length > 0) {
        const prevBottom = results[results.length - 1].top + 70;
        if (top < prevBottom) top = prevBottom;
      }

      results.push({
        id: comment.id,
        author: comment.author || 'Unknown',
        date: comment.date,
        text: getCommentText(comment),
        anchorText: (anchorMap.get(comment.id) || '').slice(0, 60),
        top,
        done: comment.done,
      });

      // Add replies indented beneath parent
      const childReplies = repliesByParent.get(comment.id) || [];
      for (const reply of childReplies) {
        const lastCard = results[results.length - 1];
        const replyTop = lastCard.top + 70; // stack beneath parent/prev reply

        results.push({
          id: reply.id,
          author: reply.author || 'Unknown',
          date: reply.date,
          text: getCommentText(reply),
          anchorText: '', // replies don't need anchor text
          top: replyTop,
          done: reply.done,
          parentId: reply.parentId,
          isReply: true,
        });
      }
    }

    setCards(results);
  }, [pagesContainer, view, comments]);

  // Recompute positions when comments change or panel becomes visible
  useEffect(() => {
    if (!visible) return;
    // Multiple attempts: layout may not be ready immediately
    const t1 = setTimeout(computePositions, 200);
    const t2 = setTimeout(computePositions, 800);
    const t3 = setTimeout(computePositions, 2000);
    return () => {
      clearTimeout(t1);
      clearTimeout(t2);
      clearTimeout(t3);
    };
  }, [visible, comments, computePositions, view]);

  // Also recompute on scroll (comments move with pages but positions are absolute)
  useEffect(() => {
    if (!visible || !pagesContainer) return;
    const scrollParent = pagesContainer.closest('.paged-editor') as HTMLElement;
    if (!scrollParent) return;

    const onScroll = () => computePositions();
    scrollParent.addEventListener('scroll', onScroll, { passive: true });
    return () => scrollParent.removeEventListener('scroll', onScroll);
  }, [visible, pagesContainer, computePositions]);

  if (!visible || cards.length === 0) return null;

  return (
    <div
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
      {cards.map((card) => (
        <div
          key={card.id}
          data-testid="comment-card"
          style={{
            position: 'absolute',
            top: card.top,
            left: card.isReply ? 16 : 0,
            width: card.isReply ? 204 : 220,
            padding: '6px 8px',
            background: card.done ? '#e8e8e8' : card.isReply ? '#fff8e1' : '#fffde7',
            border: card.isReply ? '1px solid #e8d87a' : '1px solid #e0d68a',
            borderLeft: card.isReply ? '3px solid #ffc107' : '1px solid #e0d68a',
            borderRadius: 4,
            fontSize: '11px',
            lineHeight: '1.3',
            boxShadow: '0 1px 3px rgba(0,0,0,0.1)',
            opacity: card.done ? 0.6 : 1,
          }}
        >
          {/* Author + date */}
          <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 2 }}>
            <strong style={{ fontSize: '10px', color: '#333' }}>{card.author}</strong>
            {card.date && (
              <span style={{ fontSize: '9px', color: '#999' }}>
                {new Date(card.date).toLocaleDateString()}
              </span>
            )}
          </div>

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

          {/* Comment text — editable when editing */}
          {editingId === card.id ? (
            <div style={{ marginBottom: 4 }}>
              <textarea
                value={editText}
                onChange={(e) => setEditText(e.target.value)}
                onMouseDown={(e) => e.stopPropagation()}
                style={{
                  width: '100%',
                  minHeight: 40,
                  fontSize: '11px',
                  border: '1px solid #ccc',
                  borderRadius: 3,
                  padding: '3px 4px',
                  resize: 'vertical',
                }}
              />
              <div style={{ display: 'flex', gap: 4, marginTop: 3 }}>
                <button
                  type="button"
                  onClick={(e) => {
                    e.stopPropagation();
                    onEdit?.(card.id, editText);
                    // Update local card text immediately so UI reflects the change
                    setCards((prev) =>
                      prev.map((c) => (c.id === card.id ? { ...c, text: editText } : c))
                    );
                    setEditingId(null);
                  }}
                  style={{
                    fontSize: '9px',
                    padding: '1px 5px',
                    border: '1px solid #28a745',
                    borderRadius: 3,
                    background: '#28a745',
                    color: 'white',
                    cursor: 'pointer',
                  }}
                >
                  Save
                </button>
                <button
                  type="button"
                  onClick={(e) => {
                    e.stopPropagation();
                    setEditingId(null);
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
                  Cancel
                </button>
              </div>
            </div>
          ) : (
            <div style={{ color: '#333', marginBottom: 4 }}>{card.text}</div>
          )}

          {/* Action buttons */}
          {editingId !== card.id && (
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
                    setEditingId(card.id);
                    setEditText(card.text);
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
          )}
        </div>
      ))}
    </div>
  );
};
