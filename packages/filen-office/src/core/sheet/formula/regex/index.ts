import { parse } from "./parse"
import { Program } from "./program"
import { RegexSyntaxError, RegexUnsupportedError, type Parsed } from "./syntax"

export { RegexBudgetError } from "./program"
export { RegexSyntaxError, RegexUnsupportedError } from "./syntax"

// A PCRE2 pattern compiled for this engine. Throws RegexSyntaxError for a pattern PCRE2 rejects and
// RegexUnsupportedError for one this engine does not run.
export class Regex {
	private readonly program: Program
	readonly parsed: Parsed

	constructor(parsed: Parsed, captures: boolean) {
		this.parsed = parsed
		this.program = new Program(parsed, captures)
	}

	static compile(pattern: string, options: { readonly caseless?: boolean; readonly captures?: boolean } = {}): Regex {
		return new Regex(parse(pattern, options.caseless === true), options.captures === true)
	}

	// Whether a pattern is one this engine runs: false for one PCRE2 takes that it does not (RegexUnsupportedError); a
	// pattern PCRE2 rejects is the formula's #VALUE!, so it runs.
	static runs(pattern: string): boolean {
		try {
			Regex.compile(pattern)
		} catch (error) {
			if (error instanceof RegexUnsupportedError) return false
			if (error instanceof RegexSyntaxError) return true

			throw error
		}

		return true
	}

	// The program's instructions: what a search visits at most per character, once each.
	get size(): number {
		return this.program.size
	}

	get groups(): number {
		return this.parsed.groups
	}

	// Starts a new step budget of `steps` (one per function call).
	resetBudget(steps?: number): void {
		this.program.resetBudget(steps)
	}

	test(text: string): boolean {
		return this.program.test(text)
	}

	// The first match from `start`: capture slots in UTF-16 offsets, -1 for a group that took no part.
	exec(text: string, start = 0): Int32Array | null {
		return this.program.exec(text, { start })
	}

	// Every match, as PCRE2's global matching finds them: after an empty match, one more non-empty match is tried at
	// the same place before moving on by a character (\G still holding only where the last match ended).
	*matches(text: string): Generator<Int32Array, void> {
		let start = 0
		let previousEnd = 0
		let retry = false

		while (start <= text.length) {
			const found = this.program.exec(text, { start, previousEnd, anchored: retry, notEmptyAtStart: retry })

			if (found === null) {
				if (!retry) return

				retry = false

				const code = text.codePointAt(start) ?? 0

				start += code > 0xffff ? 2 : 1
				continue
			}

			yield found

			start = found[1] ?? text.length
			previousEnd = start
			retry = found[0] === found[1]
		}
	}
}
