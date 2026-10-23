import type { DrawingText } from "./displayList"

// Block text (DrawingText "block", metafile text) laid out into lines, as every renderer of a display list draws it.

type Trimming = Extract<DrawingText, { kind: "block" }>["trimming"]

// Lines of block text: paragraphs at "\n", wrapped greedily at spaces (a word wider than the line breaks
// between characters), cut to the height and trimmed where text is cut. Each word and each character is
// measured once; a line's width is the sum of its pieces.
export function layoutBlock(
	source: string,
	width: number,
	height: number,
	lineHeight: number,
	wrap: boolean,
	lineLimit: boolean,
	trimming: Trimming,
	measureText: (text: string) => number
): string[] {
	const widths = new Map<string, number>()
	const measure = (value: string): number => {
		let known = widths.get(value)

		if (known === undefined) {
			known = measureText(value)
			widths.set(value, known)
		}

		return known
	}
	const lines: string[] = []
	const maxLines = height <= 0 ? Infinity : lineLimit ? Math.floor(height / lineHeight + 1e-6) : Math.ceil(height / lineHeight - 1e-6)
	let cut = false

	for (const paragraph of source.split("\n")) {
		if (lines.length >= maxLines) {
			cut = true
			break
		}

		if (!wrap || width <= 0) {
			lines.push(paragraph)
			continue
		}

		let line = ""
		let lineWidth = 0
		const flush = (): void => {
			lines.push(line.trimEnd())
			line = ""
			lineWidth = 0
		}
		// A word wider than the line, character by character.
		const breakWord = (token: string): void => {
			for (const character of token) {
				const step = measure(character)

				if (line !== "" && lineWidth + step > width) flush()

				line += character
				lineWidth += step
			}
		}

		for (const token of paragraph.split(/(\s+)/)) {
			if (token === "") continue

			const space = /^\s+$/.test(token)
			const step = measure(token)

			if (lineWidth + step <= width) {
				line += token
				lineWidth += step
				continue
			}

			// Spaces at a break vanish; a word starts the next line.
			if (space) {
				if (line.trim() !== "") flush()

				continue
			}

			if (line.trim() !== "") flush()

			if (step <= width) {
				line = token
				lineWidth = step
			} else {
				line = ""
				lineWidth = 0
				breakWord(token)
			}
		}

		lines.push(line.trimEnd())
	}

	if (lines.length > maxLines) {
		cut = true
		lines.length = Number.isFinite(maxLines) ? maxLines : lines.length
	}

	if (lines.length === 0) return lines

	// Trimming applies where text is cut: the last kept line, or a line wider than an unwrapped box.
	const last = lines.length - 1

	for (let k = 0; k < lines.length; k++) {
		const line = lines[k] ?? ""
		const forced = k === last && cut

		if (forced || (width > 0 && measure(line) > width)) lines[k] = trim(line, width, trimming, measure, forced)
	}

	return lines
}

// One line cut to the width: characters or words dropped from the end (or the middle, for a path), with an
// ellipsis for the ellipsis kinds. Character prefix widths are summed once.
function trim(line: string, width: number, trimming: Trimming, measure: (text: string) => number, forced: boolean): string {
	const characters = Array.from(line.trimEnd())
	const ellipsis = trimming === "ellipsisCharacter" || trimming === "ellipsisWord" || trimming === "ellipsisPath"
	const prefix = [0]

	for (const character of characters) prefix.push((prefix[prefix.length - 1] ?? 0) + measure(character))

	const total = prefix[prefix.length - 1] ?? 0
	const mark = ellipsis ? measure("…") : 0
	// The last space before each position (-1 for none).
	const spaces = [-1]

	for (let k = 0; k < characters.length; k++) spaces.push(characters[k] === " " ? k : (spaces[k] ?? -1))
	const fits = (value: number): boolean => width <= 0 || value <= width

	if (trimming === "none" || (!forced && fits(total))) return characters.join("")

	if (trimming === "ellipsisPath") {
		// The middle gives way, keeping both ends.
		for (let keep = characters.length - 1; keep > 0; keep--) {
			const head = Math.ceil(keep / 2)
			const tail = keep - head
			const used = (prefix[head] ?? 0) + (total - (prefix[characters.length - tail] ?? 0)) + mark

			if (fits(used)) return `${characters.slice(0, head).join("")}…${characters.slice(characters.length - tail).join("")}`
		}

		return "…"
	}

	for (let keep = characters.length - (forced && ellipsis ? 0 : 1); keep >= 0; keep--) {
		let end = keep

		if ((trimming === "word" || trimming === "ellipsisWord") && keep < characters.length) {
			const space = spaces[keep] ?? -1

			if (space > 0) end = space
		}

		while (end > 0 && characters[end - 1] === " ") end--

		if (fits((prefix[end] ?? 0) + mark)) return characters.slice(0, end).join("") + (ellipsis ? "…" : "")
	}

	return ellipsis ? "…" : ""
}
