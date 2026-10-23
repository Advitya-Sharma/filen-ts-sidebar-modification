import type { FontEntry, ListDef, ListOverrideDef, StyleDef } from "./document"
import type { CharProps, ParaProps } from "./props"

// Re-reading part of a file (the RTF writer's proof of a save): the reader's state at some top-level paragraph
// ends of the first read, taken where every group open is plain body text. A span of a written file that starts
// on bytes the source shares can be read from the checkpoint before it; once the reader's state equals the
// source's at a checkpoint whose bytes the written file shares again, the rest of the file reads as the source.

// Between checkpoints at least this many source bytes: a span is read from one before it to one after it.
export const CHECKPOINT_SPACING = 16 * 1024

// One open group's state (its destination is the body's text).
export interface FrameState {
	readonly char: CharProps
	readonly para: ParaProps
	readonly uc: number
	readonly nestProps: boolean
	readonly start: number
	// Whether the writer's openGroups counts the group as an opener (the document's, a destination's).
	readonly opener: boolean
}

export interface Checkpoint {
	// Where the reader stands: just after the word that ended a paragraph (and that paragraph's end), or before the
	// body's first plain token (paragraphEnd -1).
	readonly offset: number
	readonly paragraphEnd: number
	// The body's top-level blocks finished by then.
	readonly blocks: number
	// Tokens read so far that were not plain (see the reader): equal counts mean only plain tokens between.
	readonly unplain: number
	// The groups open: the state outside the document group first, the innermost last.
	readonly frames: readonly FrameState[]
	readonly depth: number
	readonly skip: number
	readonly noteMark: boolean
	readonly documentCodepage: number
	readonly defaultFont: number
	readonly groupFloor: number
	// The document group's end (the state after it closed), not a paragraph's.
	readonly end: boolean
}

// What a span is read with: the source's tables (fixed before the first usable checkpoint) and its checkpoints.
export interface RtfCheckpoints {
	readonly list: readonly Checkpoint[]
	// Tokens not plain before the last table destination began: a checkpoint at or before it is unusable.
	readonly tablesUntil: number
	readonly fonts: ReadonlyMap<number, FontEntry>
	readonly colors: readonly (string | undefined)[]
	readonly styles: readonly StyleDef[]
	readonly lists: readonly ListDef[]
	readonly overrides: readonly ListOverrideDef[]
	readonly revisionText: string
	// Each control word's handler and whether it is plain, as the first read learnt them.
	readonly routes: ReadonlyMap<string, number>
	readonly plainWords: ReadonlyMap<string, number>
}

// A copy of plain data and property objects (their class kept), sharing what two references shared; memoized
// markup (`cached`) is left out.
export function deepCopy<T>(value: T, seen = new Map<object, unknown>()): T {
	if (typeof value !== "object" || value === null) return value

	const known = seen.get(value)

	if (known !== undefined) return known as T
	if (value instanceof Map || value instanceof Set) return value

	if (Array.isArray(value)) {
		const out: unknown[] = []

		seen.set(value, out)

		for (const item of value) out.push(deepCopy(item, seen))

		return out as T
	}

	const out = Object.create(Object.getPrototypeOf(value) as object | null) as Record<string, unknown>

	seen.set(value, out)

	for (const [key, item] of Object.entries(value)) out[key] = key === "cached" ? undefined : deepCopy(item, seen)

	return out as T
}

// Whether two values hold the same data (deepCopy's sense: `cached` aside, maps and sets by identity).
export function deepEqual(a: unknown, b: unknown): boolean {
	if (a === b) return true
	if (typeof a !== "object" || typeof b !== "object" || a === null || b === null) return false
	if (a instanceof Map || a instanceof Set || b instanceof Map || b instanceof Set) return false
	if (Object.getPrototypeOf(a) !== Object.getPrototypeOf(b)) return false

	if (Array.isArray(a)) {
		if (!Array.isArray(b) || a.length !== b.length) return false

		for (let k = 0; k < a.length; k++) if (!deepEqual(a[k], b[k])) return false

		return true
	}

	const x = a as Record<string, unknown>
	const y = b as Record<string, unknown>
	const keys = Object.keys(x)

	if (keys.length !== Object.keys(y).length) return false

	for (const key of keys) {
		if (key === "cached") continue
		if (!Object.hasOwn(y, key) || !deepEqual(x[key], y[key])) return false
	}

	return true
}
