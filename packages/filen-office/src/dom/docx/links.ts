import { Item, RevisionView, type DocxModel, type StoryPart } from "../../core/docx/index"
import type { DocxLinkTarget } from "./protocol"

// Hyperlinks of a laid-out paragraph as ranges of its text. Layout reports which paragraph a line comes
// from, not which items its characters come from, so each link's own text (w:hyperlink content, or a
// HYPERLINK field's result) is found in the paragraph's text in document order. Text that layout draws
// differently from what the model holds (list markers, computed fields) only shifts where the search
// starts; a link whose text is not found is left out rather than placed on the wrong words.

export interface LinkRange {
	readonly start: number
	readonly end: number
	readonly target: DocxLinkTarget
}

const QUOTED = /"([^"]*)"|(\S+)/g

// HYPERLINK "url" \l "anchor" \o "tooltip" (ECMA-376 §17.16.5.25).
export function parseHyperlinkField(instruction: string): DocxLinkTarget | undefined {
	const body = /^\s*HYPERLINK\b(.*)$/is.exec(instruction)?.[1]

	if (body === undefined) return undefined

	let href: string | undefined
	let anchor: string | undefined
	let tooltip: string | undefined
	let pending: "l" | "o" | "skip" | undefined

	for (const match of body.matchAll(QUOTED)) {
		const value = match[1] ?? match[2] ?? ""

		if (match[1] === undefined && value.startsWith("\\")) {
			const name = value.slice(1, 2).toLowerCase()

			pending = name === "l" ? "l" : name === "o" ? "o" : name === "m" || name === "n" ? undefined : "skip"

			continue
		}

		if (pending === "l") anchor = value
		else if (pending === "o") tooltip = value
		else if (pending === undefined && href === undefined) href = value

		pending = undefined
	}

	if (anchor !== undefined && anchor.length > 0) return { kind: "internal", anchor, tooltip }
	if (href !== undefined && href.length > 0) return { kind: "external", href, tooltip }

	return undefined
}

function hyperlinkTarget(part: StoryPart, item: number): DocxLinkTarget | undefined {
	const link = part.hyperlink(item)

	if (link.anchor !== undefined && link.anchor.length > 0) return { kind: "internal", anchor: link.anchor, tooltip: link.tooltip }

	// Only external relationships are links; an internal part target is never navigated to.
	if (link.external && link.target !== undefined && link.target.length > 0) {
		const href = link.docLocation === undefined || link.docLocation.length === 0 ? link.target : `${link.target}#${link.docLocation}`

		return { kind: "external", href, tooltip: link.tooltip }
	}

	return undefined
}

// The text of a field's result (between its separate and end characters) as layout shows it.
function fieldResult(part: StoryPart, separate: number, end: number, view: RevisionView): string {
	let out = ""

	for (let item = separate + 1; item < end; item++) {
		const kind = part.kinds[item] ?? 0

		if (part.hidden(item, view)) {
			item = part.end(item) - 1

			continue
		}

		if (kind === Item.Text) out += part.text(item)
		else if (kind === Item.Tab) out += "\t"
	}

	return out
}

interface Candidate {
	readonly at: number
	readonly text: string
	readonly target: DocxLinkTarget
}

// Fields of a part indexed by the paragraph their begin character sits in, built once per part.
const fieldIndex = new WeakMap<StoryPart, Map<number, number[]>>()

function paragraphOf(part: StoryPart, item: number): number {
	for (let at = item; at >= 0; at = part.parent(at)) {
		if (part.kinds[at] === Item.Paragraph) return at
	}

	return -1
}

function fieldsByParagraph(part: StoryPart): Map<number, number[]> {
	let index = fieldIndex.get(part)

	if (index === undefined) {
		index = new Map()

		const fields = part.fields

		for (let field = 0; field < fields.count; field++) {
			const begin = fields.begin[field] ?? -1

			if (begin < 0) continue

			const paragraph = paragraphOf(part, begin)
			const list = index.get(paragraph)

			if (list === undefined) index.set(paragraph, [field])
			else list.push(field)
		}

		fieldIndex.set(part, index)
	}

	return index
}

export function paragraphLinks(part: StoryPart, paragraph: number, text: string, view: RevisionView): LinkRange[] {
	const candidates: Candidate[] = []
	const end = part.end(paragraph)

	for (let item = paragraph + 1; item < end; item++) {
		if (part.kinds[item] !== Item.Hyperlink) continue

		const target = hyperlinkTarget(part, item)

		if (target !== undefined) candidates.push({ at: item, text: part.paragraphText(item, view), target })
	}

	for (const field of fieldsByParagraph(part).get(paragraph) ?? []) {
		if (part.fieldType(field, view) !== "HYPERLINK") continue

		const target = parseHyperlinkField(part.fieldInstruction(field, view))
		const separate = part.fields.separate[field] ?? -1
		const fieldEnd = part.fields.end[field] ?? -1

		if (target === undefined) continue

		const result =
			part.fields.simple[field] === 1
				? part.paragraphText(part.fields.begin[field] ?? -1, view)
				: separate >= 0 && fieldEnd > separate
					? fieldResult(part, separate, fieldEnd, view)
					: ""

		candidates.push({ at: part.fields.begin[field] ?? 0, text: result, target })
	}

	if (candidates.length === 0) return []

	candidates.sort((a, b) => a.at - b.at)

	const out: LinkRange[] = []
	let cursor = 0

	for (const candidate of candidates) {
		const needle = candidate.text

		if (needle.trim().length === 0) continue

		const at = text.indexOf(needle, cursor)

		if (at < 0) continue

		out.push({ start: at, end: at + needle.length, target: candidate.target })
		cursor = at + needle.length
	}

	return out
}

// Bookmark name -> where it starts, built once per model.
export function bookmarkIndex(model: DocxModel): Map<string, { part: number; item: number }> {
	const out = new Map<string, { part: number; item: number }>()

	model.parts.forEach((part, partIndex) => {
		for (let row = 0; row < part.bookmarks.count; row++) {
			const item = part.bookmarks.start[row] ?? -1
			const name = part.bookmarkName(row)

			if (item >= 0 && name !== undefined && !out.has(name)) out.set(name, { part: partIndex, item })
		}
	})

	return out
}

export { paragraphOf }
