// Vertical space between paragraphs, as Word adds it: the first paragraph's space after plus the second's
// space before (no collapsing), each dropped when its paragraph has contextual spacing and both share a
// style.

export interface ParagraphSpacing {
	readonly before: number
	readonly after: number
	readonly contextual: boolean
	// What "the same style" compares (Word: the paragraph style id).
	readonly styleKey: string | number
}

export function paragraphGap(previous: ParagraphSpacing, next: ParagraphSpacing): number {
	const sameStyle = previous.styleKey === next.styleKey
	const after = previous.contextual && sameStyle ? 0 : previous.after
	const before = next.contextual && sameStyle ? 0 : next.before

	return after + before
}
