/**
 * Text boxes read from a file: where each sits in the paragraph that holds it, and whether
 * its words were edited, so that saving puts Word's own XML for the box back unchanged.
 *
 * The editor shows a text box as its own block straight after the paragraph that holds it
 * (toProseDoc). On save (fromProseDoc) the box goes back into that paragraph, at the same
 * place in the paragraph's text.
 */

import type { Fragment } from 'prosemirror-model';
import type { Paragraph, ParagraphContent, Run } from '../../types/document';
import { getParagraphText } from '../../docx/paragraphParser';

/** Characters of text in these paragraph parts, counted the way getParagraphText counts them. */
export function textLengthOf(content: ParagraphContent[]): number {
  return getParagraphText({ type: 'paragraph', content }).length;
}

/** A fingerprint of a text box's content, to tell on save whether its words were edited. */
export function textBoxContentKey(content: Fragment): string {
  return JSON.stringify(content.toJSON());
}

/** The run cut in two after `at` characters of its text, or null when `at` falls inside a non-text part. */
function splitRun(run: Run, at: number): [Run, Run] | null {
  let seen = 0;
  for (let i = 0; i < run.content.length; i++) {
    const part = run.content[i];
    if (seen === at) {
      return [
        { ...run, content: run.content.slice(0, i) },
        { ...run, content: run.content.slice(i) },
      ];
    }
    const length = textLengthOf([{ ...run, content: [part] }]);
    if (seen + length <= at) {
      seen += length;
      continue;
    }
    if (part.type !== 'text') return null;
    const cut = at - seen;
    return [
      { ...run, content: [...run.content.slice(0, i), { ...part, text: part.text.slice(0, cut) }] },
      { ...run, content: [{ ...part, text: part.text.slice(cut) }, ...run.content.slice(i + 1)] },
    ];
  }
  return null;
}

/**
 * Put `run` into the paragraph with `textOffset` characters of the paragraph's text before it.
 * An offset past the end (the paragraph was shortened) puts it last.
 */
export function insertRunAtTextOffset(paragraph: Paragraph, run: Run, textOffset: number): void {
  let seen = 0;
  for (let i = 0; i < paragraph.content.length; i++) {
    const item = paragraph.content[i];
    const length = textLengthOf([item]);
    if (length === 0 || seen + length <= textOffset) {
      seen += length;
      continue;
    }
    if (seen === textOffset) {
      paragraph.content.splice(i, 0, run);
      return;
    }
    // The offset falls inside this part: a run is cut in two; a link or field keeps the box after it.
    const halves = item.type === 'run' ? splitRun(item, textOffset - seen) : null;
    paragraph.content.splice(i, 1, ...(halves ? [halves[0], run, halves[1]] : [item, run]));
    return;
  }
  paragraph.content.push(run);
}
