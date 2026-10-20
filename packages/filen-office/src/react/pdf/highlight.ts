import { PDF_FIND_INTS } from "../../dom/pdf/index"

// Find matches painted into a rendered text layer, as pdf.js's own viewer does: the matched part of each span
// is wrapped in a `highlight appended` span (`begin`, `middle` and `end` for a match over several spans,
// `selected` for the current one). The spans' own text is put back by the returned restore.

export interface PaintedMatches {
	readonly selected: HTMLElement | null
	readonly restore: () => void
}

interface Segment {
	readonly start: number
	readonly end: number
	readonly part: string
	readonly selected: boolean
}

export function paintMatches(
	textDivs: readonly HTMLElement[],
	strs: readonly string[],
	matches: Int32Array,
	from: number,
	to: number,
	active: number
): PaintedMatches {
	const segments = new Map<number, Segment[]>()

	for (let match = from; match < to; match++) {
		const at = match * PDF_FIND_INTS
		const first = matches[at + 1] ?? 0
		const firstOffset = matches[at + 2] ?? 0
		const last = matches[at + 3] ?? 0
		const lastEnd = matches[at + 4] ?? 0

		for (let item = first; item <= last; item++) {
			const text = strs[item]

			if (text === undefined || item >= textDivs.length) break

			const start = item === first ? firstOffset : 0
			const end = item === last ? lastEnd : text.length

			if (end <= start) continue

			const part = first === last ? "" : item === first ? "begin" : item === last ? "end" : "middle"
			const list = segments.get(item) ?? []

			list.push({ start, end, part, selected: match === active })
			segments.set(item, list)
		}
	}

	let selected: HTMLElement | null = null

	for (const [item, list] of segments) {
		const div = textDivs[item]
		const text = strs[item]

		if (div === undefined || text === undefined) continue

		const document = div.ownerDocument
		const nodes: Node[] = []
		let at = 0

		for (const segment of list) {
			if (segment.start > at) nodes.push(document.createTextNode(text.slice(at, segment.start)))

			const span = document.createElement("span")

			// `appended` keeps it in the span's flow (pdf.js's text layer styles position every span absolutely).
			span.className = ["highlight", segment.part, segment.selected ? "selected" : "", "appended"]
				.filter(name => name.length > 0)
				.join(" ")
			span.textContent = text.slice(segment.start, segment.end)
			nodes.push(span)

			if (segment.selected && selected === null) selected = span

			at = segment.end
		}

		if (at < text.length) nodes.push(document.createTextNode(text.slice(at)))

		div.replaceChildren(...nodes)
	}

	return {
		selected,
		restore: () => {
			for (const item of segments.keys()) {
				const div = textDivs[item]

				if (div !== undefined) div.textContent = strs[item] ?? ""
			}
		}
	}
}
