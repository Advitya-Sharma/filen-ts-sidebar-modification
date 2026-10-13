import { Fragment, type Node } from "prosemirror-model"
import { TextSelection, type EditorState, type Transaction } from "prosemirror-state"
import { nfcWithMap, type NfcText } from "../../text/nfc"
import { controlOf, typeable } from "../forms/controls"
import { controlSpans } from "../forms/positions"
import { sessionOfDoc } from "../forms/sessions"
import type { StoryPart } from "../model/index"
import { docxSchema } from "./schema"

// Find and replace in a story being edited, on the editing document itself (the viewer's find runs on laid-out
// pages, whose offsets a replace cannot use). Text is searched paragraph by paragraph: an inline atom counts as
// one character that matches nothing, so a match never spans a field, a tab or a drawing, and string offsets
// are document positions. Text is compared in NFC (as the viewer's find does), matches widened back to
// whole characters. Locked blocks are not searched (nothing in them can be replaced). Replacing leaves matches
// in content controls whose content cannot be typed into (locked, a checkbox...) as they are, as Word does.

export interface StoryFindOptions {
	readonly caseSensitive?: boolean
}

export interface StoryMatch {
	readonly from: number
	readonly to: number
}

// Stands for an inline atom in a paragraph's text: one position, matched by no query (a lone surrogate).
const ATOM = "\uDFFF"
const SPECIAL = /[.*+?^${}()|[\]\\]/g
const nodes = docxSchema.nodes

// Paragraph text by node (nodes are immutable: an edited paragraph is a new node), shared by every search.
const texts = new WeakMap<Node, NfcText>()

function paragraphText(node: Node): NfcText {
	let text = texts.get(node)

	if (text === undefined) {
		text = nfcWithMap(node.textBetween(0, node.content.size, undefined, ATOM))
		texts.set(node, text)
	}

	return text
}

// The query as a pattern: literal, Unicode-aware, case-insensitive unless asked; null for an empty query.
export function findPattern(query: string, options: StoryFindOptions = {}): RegExp | null {
	const normalised = query.normalize("NFC")

	if (normalised.length === 0) return null

	return new RegExp(normalised.replace(SPECIAL, "\\$&"), options.caseSensitive === true ? "gu" : "giu")
}

// Matches a find shows at most (a counter shows "10,000+" beyond).
export const STORY_FIND_LIMIT = 10_000

// Every match of `query` in `doc`, in document order. `limit` bounds the matches returned.
export function findInStory(doc: Node, query: string, options: StoryFindOptions = {}, limit = STORY_FIND_LIMIT): StoryMatch[] {
	const pattern = findPattern(query, options)
	const out: StoryMatch[] = []

	if (pattern === null) return out

	doc.descendants((node, pos) => {
		if (out.length >= limit) return false
		if (node.type === nodes.locked_block) return false
		if (!node.isTextblock) return true

		const { text, starts, ends } = paragraphText(node)

		pattern.lastIndex = 0

		for (let match = pattern.exec(text); match !== null && out.length < limit; match = pattern.exec(text)) {
			const found = match.index
			const foundEnd = found + match[0].length

			if (foundEnd === found) {
				pattern.lastIndex++

				continue
			}

			// Back to the original text, widened to whole characters.
			const start = starts === null ? found : (starts[found] ?? found)
			const end = ends === null ? foundEnd : (ends[foundEnd - 1] ?? foundEnd)

			out.push({ from: pos + 1 + start, to: pos + 1 + end })
		}

		return false
	})

	return out
}

// Whether a match still covers text equal to the query (the document may have changed since it was found).
export function matchHolds(doc: Node, match: StoryMatch, query: string, options: StoryFindOptions = {}): boolean {
	if (match.to > doc.content.size || match.from >= match.to) return false

	const $from = doc.resolve(match.from)

	if (!$from.parent.isTextblock || $from.end() < match.to) return false

	const pattern = findPattern(query, options)

	if (pattern === null) return false

	const text = nfcWithMap(doc.textBetween(match.from, match.to, undefined, ATOM)).text

	pattern.lastIndex = 0

	const found = pattern.exec(text)

	return found !== null && found.index === 0 && found[0].length === text.length
}

// Whether a match lies outside every content control whose content typing cannot change (`part`: the story's,
// undefined when the state is no session's: nothing is known to be locked).
function replaceable(part: StoryPart | undefined, doc: Node, match: StoryMatch): boolean {
	if (part === undefined) return true

	const $from = doc.resolve(match.from)

	for (let depth = $from.depth; depth > 0; depth--) {
		const node = $from.node(depth)

		if (node.type === nodes.sdt) {
			const info = controlOf(part, node)

			if (info !== undefined && typeable(info) !== undefined) return false
		}
	}

	for (const span of controlSpans($from.parent, $from.start())) {
		if (match.from >= span.to || match.to <= span.from) continue

		const info = controlOf(part, span.mark)

		if (info !== undefined && typeable(info) !== undefined) return false
	}

	return true
}

// Whether [from, to] of `doc` holds part of a run-level content control (a replace over it would take it out).
function controlBetween(doc: Node, from: number, to: number): boolean {
	let found = false

	doc.nodesBetween(from, to, node => {
		if (!found && node.isInline && node.marks.some(mark => mark.type === docxSchema.marks.content_control)) found = true

		return !found
	})

	return found
}

// Replaces matches (each keeping the formatting where it starts) in one step of the undo history; null when
// none still holds. Matches are applied from the end, so earlier positions stay valid.
export function replaceMatches(
	state: EditorState,
	matches: readonly StoryMatch[],
	query: string,
	replacement: string,
	options: StoryFindOptions = {}
): Transaction | null {
	const tr = state.tr
	const sorted = [...matches].sort((a, b) => b.from - a.from)
	let last = Number.POSITIVE_INFINITY

	const part = sessionOfDoc(state.doc)?.story.part

	for (const match of sorted) {
		// Overlapping matches (a query found again inside itself) are replaced once.
		if (match.to > last || !matchHolds(state.doc, match, query, options) || !replaceable(part, state.doc, match)) continue

		const marks = state.doc.resolve(match.from + 1).marks()

		if (replacement.length === 0) tr.delete(match.from, match.to)
		else tr.replaceWith(match.from, match.to, docxSchema.text(replacement, marks))

		last = match.from
	}

	return tr.docChanged ? tr : null
}

// Replaces every match of `query` in `state`, however many (find shows at most STORY_FIND_LIMIT), in one step
// of the undo history: one replace per paragraph, from its first match to its last, each replacement keeping
// the formatting where its match starts. Null when nothing matches.
export function replaceAll(
	state: EditorState,
	query: string,
	replacement: string,
	options: StoryFindOptions = {}
): { readonly tr: Transaction; readonly count: number } | null {
	const doc = state.doc
	const part = sessionOfDoc(doc)?.story.part
	const matches = findInStory(doc, query, options, Number.POSITIVE_INFINITY).filter(match => replaceable(part, doc, match))

	if (matches.length === 0) return null

	const tr = state.tr
	let end = matches.length

	// Paragraph by paragraph from the end, so earlier positions stay valid.
	while (end > 0) {
		const last = matches[end - 1]

		if (last === undefined) break

		const paragraph = doc.resolve(last.from).start()
		let start = end - 1

		// A control between two matches is not taken out and put back (its locks stay out of the way).
		while (
			start > 0 &&
			(matches[start - 1]?.from ?? -1) >= paragraph &&
			!controlBetween(doc, matches[start - 1]?.to ?? 0, matches[start]?.from ?? 0)
		)
			start--

		const first = matches[start]

		if (first === undefined) break

		let content = Fragment.empty
		let at = first.from

		for (let k = start; k < end; k++) {
			const match = matches[k]

			if (match === undefined) continue

			if (match.from > at) content = content.append(doc.slice(at, match.from).content)
			if (replacement.length > 0)
				content = content.append(Fragment.from(docxSchema.text(replacement, doc.resolve(match.from + 1).marks())))

			at = match.to
		}

		tr.replaceWith(first.from, last.to, content)
		end = start
	}

	return tr.docChanged ? { tr, count: matches.length } : null
}

// Selects a match (as Word's find does), or null when it no longer lies in the document.
export function selectMatch(state: EditorState, match: StoryMatch): Transaction | null {
	if (match.from < 0 || match.to > state.doc.content.size || match.from > match.to) return null

	return state.tr.setSelection(TextSelection.create(state.doc, match.from, match.to))
}
