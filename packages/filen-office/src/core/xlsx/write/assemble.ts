import type { XlsxSection } from "../model/types"
import { Pieces } from "./markup"

// A part rebuilt from its root's children: the source bytes between and around them copied, each child copied
// or replaced, new children inserted at their schema position (before the first known child that follows them
// in `order`). A replacement stands for every child of its name (conditionalFormatting repeats): it takes the
// first one's place and the others go. The result is pieces, mostly views of the source.

export interface AssembleInput {
	readonly bytes: Uint8Array
	readonly sections: readonly XlsxSection[]
	readonly order: readonly string[]
	// New bytes by section name; an empty array removes the section.
	readonly replacements: ReadonlyMap<string, Uint8Array>
	// Writes a section in place (sheetData) instead of copying or replacing it.
	readonly inPlace?: ReadonlyMap<string, (pieces: Pieces) => void>
}

export interface Assembled {
	readonly pieces: Pieces
	// The root's children as written, in order.
	readonly names: readonly string[]
}

export function assembleSections(input: AssembleInput): Assembled {
	const { bytes, sections, order, replacements } = input
	const pieces = new Pieces()
	const names: string[] = []
	const index = (name: string): number => order.indexOf(name)
	const present = new Set(sections.map(section => section.name))
	const replaced = new Set<string>()
	const added = [...replacements]
		.filter(([name, markup]) => !present.has(name) && markup.length > 0 && index(name) >= 0)
		.sort((a, b) => index(a[0]) - index(b[0]))
	let next = 0
	let at = 0

	const insertBefore = (limit: number): void => {
		while (next < added.length && index((added[next] as [string, Uint8Array])[0]) < limit) {
			const [name, markup] = added[next] as [string, Uint8Array]

			pieces.push(markup)
			names.push(name)
			next++
		}
	}

	for (const section of sections) {
		const position = index(section.name)

		pieces.push(bytes.subarray(at, section.start))

		if (position >= 0) insertBefore(position)

		const write = input.inPlace?.get(section.name)
		const replacement = replacements.get(section.name)

		if (write !== undefined) {
			write(pieces)
			names.push(section.name)
		} else if (replacement !== undefined) {
			if (!replaced.has(section.name)) {
				replaced.add(section.name)
				pieces.push(replacement)

				if (replacement.length > 0) names.push(section.name)
			}
		} else {
			pieces.push(bytes.subarray(section.start, section.end))
			names.push(section.name)
		}

		at = section.end
	}

	// The rest goes before the root's close tag.
	let close = bytes.lastIndexOf(60, bytes.length - 1)

	while (close > at && bytes[close + 1] !== 47) close = bytes.lastIndexOf(60, close - 1)

	const end = close >= at ? close : at

	pieces.push(bytes.subarray(at, end))
	insertBefore(Number.POSITIVE_INFINITY)
	pieces.push(bytes.subarray(end))

	return { pieces, names }
}

// How many neighbouring known children of `names` stand against `order` (what the gate compares with the
// source's own count).
export function orderInversions(names: readonly string[], order: readonly string[]): number {
	let inversions = 0
	let last = -1

	for (const name of names) {
		const position = order.indexOf(name)

		if (position < 0) continue
		if (position < last) inversions++

		last = Math.max(last, position)
	}

	return inversions
}
