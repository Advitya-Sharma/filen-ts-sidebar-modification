// Text deltas: what an edit added and removed, as LibreOffice reads the files and as our model reads them.
// Comparing the two deltas (not the texts) cancels whatever LibreOffice renders differently in both files.
// Three tiers, strictest first: whole lines, words, and characters (whitespace left out), so a residual
// shows whether lines were split or joined differently or text really differs.

export type Tier = "lines" | "words" | "chars"

export const TIERS: readonly Tier[] = ["lines", "words", "chars"]

// Format characters and symbol-font characters (each side maps them its own way: private-use code points, and C1
// controls, which LibreOffice reads a w:sym w:char below F000 as, ECMA-376 Part 1 §17.3.3.30) are left out, the
// no-break hyphen and space read as their plain forms.
const INVISIBLE = /[\u0080-\u009f\u00ad\u200b-\u200f\u2060\ufeff\ue000-\uf8ff]/g
// A list label LibreOffice's plain-text export puts before a paragraph: a bullet, or a number or letter
// ending in "." or ")" ("1.", "a)", "1.2.", "iv."), followed by blanks.
const LABEL =
	/^\s*(?:[\u2022\u25e6\u25aa\u25cf\u25cb\u25a0\u25a1\u2013\u2014\u2212\u2023\u2043\u27a2\u2794\u2713\u2714\u00b7o\u00a7*-]|\(?[0-9A-Za-z]{1,6}(?:[.)][0-9A-Za-z]{1,6})*[.)])[ \t]+/

export interface NormalizeOptions {
	// Takes list labels off line starts (LibreOffice's plain-text export).
	readonly labels?: boolean
}

export function normalizeLines(text: string, options: NormalizeOptions = {}): string[] {
	const out: string[] = []

	for (const raw of text
		.normalize("NFC")
		.replace(INVISIBLE, "")
		.replace(/\u2011/g, "-")
		.split(/\r\n|\r|\n|\u2028|\u2029|\v|\f/)) {
		let line = raw.replace(/[\s\u00a0\u202f]+/g, " ").trim()

		if (options.labels === true) line = line.replace(LABEL, "").trim()
		if (line !== "") out.push(line)
	}

	return out
}

function units(lines: readonly string[], tier: Tier): string[] {
	if (tier === "lines") return [...lines]
	if (tier === "words") return lines.flatMap(line => line.split(" "))

	return lines.flatMap(line => [...line.replace(/ /g, "")])
}

function bag(items: readonly string[]): Map<string, number> {
	const out = new Map<string, number>()

	for (const item of items) out.set(item, (out.get(item) ?? 0) + 1)

	return out
}

// Signed differences after - before, without zeros.
function difference(before: Map<string, number>, after: Map<string, number>): Map<string, number> {
	const out = new Map<string, number>()

	for (const [key, times] of after) if (times !== (before.get(key) ?? 0)) out.set(key, times - (before.get(key) ?? 0))
	for (const [key, times] of before) if (!after.has(key)) out.set(key, -times)

	return out
}

export interface DeltaComparison {
	readonly equal: boolean
	// What LibreOffice's delta has beyond ours (positive: LO added more), a few entries.
	readonly residual: readonly (readonly [string, number])[]
}

// Whether the edit's delta reads the same both ways, at `tier`.
export function compareDeltas(
	theirsBefore: readonly string[],
	theirsAfter: readonly string[],
	oursBefore: readonly string[],
	oursAfter: readonly string[],
	tier: Tier
): DeltaComparison {
	const theirs = difference(bag(units(theirsBefore, tier)), bag(units(theirsAfter, tier)))
	const ours = difference(bag(units(oursBefore, tier)), bag(units(oursAfter, tier)))
	const residual = difference(ours, theirs)

	return { equal: residual.size === 0, residual: [...residual].slice(0, 8) }
}
