// Multi-part files of one title ("Movie cd1", "Movie cd2"): same name around a numbered or lettered part marker.

import { trimTrailingSeparators } from "./naming"

export interface StackPart {
	// Name text before the marker; becomes the title of the stacked item.
	prefix: string
	// Files with equal keys stack: same prefix, marker word and trailing text, all case-insensitive.
	key: string
	part: number
}

// A marker word with a part number (one or two digits) or a letter a-d, optionally one separator in between. It must
// follow a separator or bracket, so it is never the tail of a longer word, and must not run into further letters or
// digits. Everything after the word has a fixed maximum length, which keeps the scan linear on any input.
const MARKER = /(?<=[\s._\-()[\]{}（）［］【】])(cd|dvd|part|pt|disc|disk)[ ._-]?(?:(\d{1,2})|([a-d]))(?![\p{L}\p{N}])/giu
const PART_LETTER = /^[a-d]$/iu
const LETTER_SEPARATORS = "-_."
const LETTER_CODE_A = "a".charCodeAt(0)

// Spellings of one marker type, so "disc1" and "disk2" stack.
const MARKER_TYPES = new Map([
	["disk", "disc"],
	["pt", "part"]
])

function letterPart(letter: string): number {
	return letter.toLowerCase().charCodeAt(0) - LETTER_CODE_A + 1
}

function stackPart(stem: string, markerStart: number, markerEnd: number, type: string, part: number): StackPart | null {
	const prefix = trimTrailingSeparators(stem.slice(0, markerStart))

	if (prefix.length === 0) {
		return null
	}

	return {
		prefix,
		key: `${prefix.toLowerCase()}\u0000${type}\u0000${stem.slice(markerEnd).toLowerCase()}`,
		part
	}
}

// Part marker of a stem: "Movie cd1", "Movie (Part 2)", "Movie pt.3 1080p", "Movie disc b", or a lone trailing letter
// after "-", "_" or "." ("Movie-a"). Null when there is none or no title precedes it. The last marker counts.
export function parseStackPart(stem: string): StackPart | null {
	let last: RegExpExecArray | null = null

	for (const match of stem.matchAll(MARKER)) {
		last = match
	}

	const lastEnd = last === null ? -1 : last.index + last[0].length
	const letter = stem.charAt(stem.length - 1)

	if (
		lastEnd !== stem.length &&
		stem.length >= 3 &&
		LETTER_SEPARATORS.includes(stem.charAt(stem.length - 2)) &&
		PART_LETTER.test(letter)
	) {
		return stackPart(stem, stem.length - 2, stem.length, "", letterPart(letter))
	}

	if (last === null) {
		return null
	}

	const word = (last[1] ?? "").toLowerCase()
	const type = MARKER_TYPES.get(word) ?? word
	const digits = last[2]

	return digits === undefined
		? stackPart(stem, last.index, lastEnd, `${type}:letter`, letterPart(last[3] ?? "a"))
		: stackPart(stem, last.index, lastEnd, `${type}:number`, Number(digits))
}

export interface StackGroup {
	// Indices into the input, in part order; a single entry when the file does not stack.
	members: number[]
	// Shared name before the part markers, or null for a single file.
	prefix: string | null
}

function hasDistinctParts(members: readonly number[], parts: readonly (StackPart | null)[]): boolean {
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

// Groups stems into stacks. A stack needs at least two parts and no repeated part number; gaps are allowed. Files whose
// key repeats a part number all stay single, since they are more likely versions than parts.
// Groups come out in input order of their first member.
export function groupStacks(stems: readonly string[]): StackGroup[] {
	const parts = stems.map(stem => parseStackPart(stem))
	const byKey = new Map<string, number[]>()

	for (let i = 0; i < parts.length; i++) {
		const key = parts[i]?.key

		if (key === undefined) {
			continue
		}

		const members = byKey.get(key)

		if (members === undefined) {
			byKey.set(key, [i])
		} else {
			members.push(i)
		}
	}

	const stacked = new Map<string, number[]>()

	for (const [key, members] of byKey) {
		if (members.length >= 2 && hasDistinctParts(members, parts)) {
			stacked.set(
				key,
				members.sort((a, b) => (parts[a]?.part ?? 0) - (parts[b]?.part ?? 0))
			)
		}
	}

	const groups: StackGroup[] = []
	const emitted = new Set<string>()

	for (let i = 0; i < parts.length; i++) {
		const part = parts[i]
		const members = part === null || part === undefined ? undefined : stacked.get(part.key)

		if (part === null || part === undefined || members === undefined) {
			groups.push({
				members: [i],
				prefix: null
			})

			continue
		}

		// Emitted once, at the member that comes first in the input.
		if (emitted.has(part.key)) {
			continue
		}

		emitted.add(part.key)
		groups.push({
			members,
			prefix: parts[members[0] ?? i]?.prefix ?? part.prefix
		})
	}

	return groups
}
