import { LINE_FLOATS } from "../../core/text/index"
import { findMatches, lineOf } from "../../dom/docx/find"
import { FIND_INTS, type DocxFindOptions, type DocxPageData } from "../../dom/docx/protocol"

// Find matches as the page list shows them. The engine searches every page (DocxSessionApi.find); this
// also runs over pages already on this side, e.g. in tests or without a session.

export interface DocxFindMatch {
	readonly page: number
	// The page's source (paragraph) and the match's offsets in its text.
	readonly source: number
	readonly start: number
	readonly end: number
}

export type { DocxFindOptions }

export function findMatchList(matches: Int32Array): DocxFindMatch[] {
	const out: DocxFindMatch[] = []

	for (let at = 0; at + FIND_INTS <= matches.length; at += FIND_INTS) {
		out.push({ page: matches[at] ?? 0, source: matches[at + 1] ?? 0, start: matches[at + 2] ?? 0, end: matches[at + 3] ?? 0 })
	}

	return out
}

export function findInDocument(
	pages: readonly (DocxPageData | undefined)[],
	query: string,
	options: DocxFindOptions = {}
): DocxFindMatch[] {
	return findMatchList(findMatches(pages, query, options))
}

export { lineOf }

// Top of the line holding a match (px from the page's top), or 0.
export function matchTop(page: DocxPageData, match: DocxFindMatch): number {
	const line = lineOf(page, match.source, match.start)

	return line < 0 ? 0 : (page.lineFloats[line * LINE_FLOATS] ?? 0)
}
