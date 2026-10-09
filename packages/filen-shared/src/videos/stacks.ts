// Multi-part files of one title ("Movie cd1", "Movie cd2"): same name around a numbered or lettered part marker.

import { trimTrailingSeparators } from "./naming"

export interface StackPart {
	// Name text before the marker; becomes the title of the stacked item.
	prefix: string
	// Files with equal keys stack: same prefix, marker word and trailing text, all case-insensitive.
	key: string
	part: number
}

const PART_MARKER = /(?<![\p{L}\p{N}])(cd|dvd|part|pt|disc|disk)[ ._-]?(\d{1,2}|[a-d])(?![\p{L}\p{N}])/giu
const LETTER_ONLY = /[-_.]([a-d])$/iu
const LEADING_CLOSE = /^[)\]}]+/u

function partNumber(token: string): number {
	const code = token.toLowerCase().charCodeAt(0)

	// a-d → 1-4
	return code >= 97 && code <= 100 ? code - 96 : Number(token)
}

// Part marker in a stem: the last "cd/dvd/part/pt/disc/disk" + number or a-d, else a trailing "-a" … "-d".
export function parseStackPart(stem: string): StackPart | null {
	let marker: RegExpExecArray | null = null

	if (stem.length > 3) {
		PART_MARKER.lastIndex = 0

		for (let match = PART_MARKER.exec(stem); match !== null; match = PART_MARKER.exec(stem)) {
			marker = match
		}
	}

	if (marker !== null) {
		const prefix = trimTrailingSeparators(stem.slice(0, marker.index))

		if (prefix.length === 0) {
			return null
		}

		const word = (marker[1] ?? "").toLowerCase()
		const token = marker[2] ?? ""
		const suffix = stem
			.slice(marker.index + marker[0].length)
			.replace(LEADING_CLOSE, "")
			.toLowerCase()
		// Lettered and numbered parts never stack together.
		const style = token.charCodeAt(0) > 57 ? "a" : "1"

		return {
			prefix,
			key: `${prefix.toLowerCase()}\u0000${word === "disk" ? "disc" : word}${style}\u0000${suffix}`,
			part: partNumber(token)
		}
	}

	const letter = LETTER_ONLY.exec(stem)

	if (letter === null || letter.index === 0) {
		return null
	}

	const prefix = stem.slice(0, letter.index)

	return {
		prefix,
		key: `${prefix.toLowerCase()}\u0000letter\u0000`,
		part: partNumber(letter[1] ?? "")
	}
}

export interface StackGroup {
	// Indices into the input, in part order; a single entry when the file does not stack.
	members: number[]
	// Shared name before the part markers, or null for a single file.
	prefix: string | null
}

// Groups stems into stacks. A stack needs at least two parts and no repeated part number; gaps are allowed.
// Groups come out in input order of their first member.
export function groupStacks(stems: readonly string[]): StackGroup[] {
	const parts: (StackPart | null)[] = []
	const byKey = new Map<string, number[]>()

	for (let i = 0; i < stems.length; i++) {
		const part = parseStackPart(stems[i] ?? "")

		parts.push(part)

		if (part !== null) {
			const members = byKey.get(part.key)

			if (members === undefined) {
				byKey.set(part.key, [i])
			} else {
				members.push(i)
			}
		}
	}

	const groupOf = new Int32Array(stems.length).fill(-1)
	const groups: StackGroup[] = []

	for (let i = 0; i < stems.length; i++) {
		if (groupOf[i] !== -1) {
			continue
		}

		const part = parts[i]
		const members = part === null || part === undefined ? undefined : byKey.get(part.key)

		if (part === null || part === undefined || members === undefined || members.length < 2 || !distinctParts(members, parts)) {
			groupOf[i] = groups.length
			groups.push({
				members: [i],
				prefix: null
			})

			continue
		}

		const ordered = [...members].sort((a, b) => (parts[a]?.part ?? 0) - (parts[b]?.part ?? 0))

		for (const member of ordered) {
			groupOf[member] = groups.length
		}

		groups.push({
			members: ordered,
			prefix: part.prefix
		})
	}

	return groups
}

function distinctParts(members: readonly number[], parts: readonly (StackPart | null)[]): boolean {
	const seen = new Set<number>()

	for (const member of members) {
		const part = parts[member]?.part ?? -1

		if (seen.has(part)) {
			return false
		}

		seen.add(part)
	}

	return true
}
