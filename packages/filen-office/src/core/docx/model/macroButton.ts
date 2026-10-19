const MACROBUTTON = /^\s*MACROBUTTON(?:\s+|$)/i

// What a MACROBUTTON field shows where it has no result of its own: its DisplayText (ECMA-376 Part 1
// §17.16.5.36, the field's value is its second argument): everything after the macro name, one quoted
// argument unquoted. Undefined for other fields. Shown only: the macro is never run.
export function macroButtonText(instruction: string): string | undefined {
	const head = MACROBUTTON.exec(instruction)

	if (head === null) return undefined

	const rest = instruction.slice(head[0].length)
	// The macro name: one argument, quoted or not.
	const name = /^(?:"(?:[^"\\]|\\.)*"|\S+)\s*/.exec(rest)
	// General switches (\* MERGEFORMAT) format a result; they are not shown.
	const text = withoutSwitches(name === null ? "" : rest.slice(name[0].length)).trim()
	const quoted = /^"((?:[^"\\]|\\.)*)"$/.exec(text)

	return quoted === null ? text : (quoted[1] ?? "").replace(/\\(.)/g, "$1")
}

// `text` without the general switches it ends with: from the first position where the rest is "\\*" switches
// (each "\\*", optional spaces and an argument) and spaces, as /(?:\s*\\\*\s*\S+)+\s*$/ would find it, but in
// one pass from the end: that expression backtracks exponentially on "\\*a" repeated before a plain word.
function withoutSwitches(text: string): string {
	const n = text.length
	const space = new Uint8Array(n)

	for (let i = 0; i < n; i++) space[i] = SPACE.test(text.charAt(i)) ? 1 : 0

	// The first non-space at or after each position, and the end of the run of non-spaces starting there.
	const nextWord = new Int32Array(n + 1)
	const wordEnd = new Int32Array(n + 1)

	nextWord[n] = n
	wordEnd[n] = n

	for (let i = n - 1; i >= 0; i--) {
		nextWord[i] = space[i] === 1 ? (nextWord[i + 1] ?? n) : i
		wordEnd[i] = space[i] === 1 ? i : i + 1 < n && space[i + 1] === 0 ? (wordEnd[i + 1] ?? n) : i + 1
	}

	// rests[i]: the positions at or after i where the rest is switches and spaces (zero switches allowed), counted.
	const rests = new Int32Array(n + 2)
	let first = -1

	for (let i = n; i >= 0; i--) {
		const at = nextWord[i] ?? n
		let switches = false

		if (at + 1 < n && text.charCodeAt(at) === 92 && text.charCodeAt(at + 1) === 42) {
			const argument = nextWord[at + 2] ?? n

			// An argument ending anywhere in its word where switches and spaces follow.
			if (argument < n) switches = (rests[argument + 1] ?? 0) - (rests[(wordEnd[argument] ?? n) + 1] ?? 0) > 0
		}

		if (switches) first = i

		rests[i] = (rests[i + 1] ?? 0) + (switches || at === n ? 1 : 0)
	}

	return first < 0 ? text : text.slice(0, first)
}

const SPACE = /\s/
