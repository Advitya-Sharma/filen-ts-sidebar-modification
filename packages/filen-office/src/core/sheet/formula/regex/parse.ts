import { CharClass, complement, DIGIT, HORIZONTAL, normalize, POSIX, SPACE, VERTICAL, WORD, type Ranges } from "./charset"
import { RegexSyntaxError, RegexUnsupportedError, type AssertKind, type Node, type Parsed } from "./syntax"

// A PCRE2 pattern (UTF mode, newline LF, no Unicode properties) read into a tree. What PCRE2 rejects throws
// RegexSyntaxError; what it takes but a linear engine cannot run (or this one cannot be sure it runs alike) throws
// RegexUnsupportedError.

// PCRE2's defaults: parentheses nested at most 250 deep, counted repeats at most 65535.
const MAX_NESTING = 250
const MAX_REPEAT = 65535
const MAX_GROUPS = 65535
const MAX_NAME = 128

interface Flags {
	caseless: boolean
	multiline: boolean
	dotAll: boolean
	extended: boolean
	extendedMore: boolean
	noAutoCapture: boolean
	ungreedy: boolean
	duplicateNames: boolean
}

interface Atom {
	readonly node: Node
	// Whether a quantifier may follow.
	readonly repeatable: boolean
	// \Q...\E: a quantifier applies to its last character.
	readonly quoted?: boolean
}

function syntax(message: string): never {
	throw new RegexSyntaxError(message)
}

function unsupported(message: string): never {
	throw new RegexUnsupportedError(message)
}

function isDigit(code: number): boolean {
	return code >= 0x30 && code <= 0x39
}

function hexValue(code: number): number {
	if (code >= 0x30 && code <= 0x39) return code - 0x30
	if (code >= 0x41 && code <= 0x46) return code - 0x37
	if (code >= 0x61 && code <= 0x66) return code - 0x57

	return -1
}

function isAsciiAlphanumeric(code: number): boolean {
	return isDigit(code) || (code >= 0x41 && code <= 0x5a) || (code >= 0x61 && code <= 0x7a)
}

function isNameStart(code: number): boolean {
	return (code >= 0x41 && code <= 0x5a) || (code >= 0x61 && code <= 0x7a) || code === 0x5f
}

// White space skipped under (?x): Pattern_White_Space.
function isPatternSpace(code: number): boolean {
	return (
		(code >= 0x09 && code <= 0x0d) ||
		code === 0x20 ||
		code === 0x85 ||
		code === 0x200e ||
		code === 0x200f ||
		code === 0x2028 ||
		code === 0x2029
	)
}

const ESCAPE_CLASSES: Readonly<Record<string, readonly [Ranges, boolean]>> = {
	d: [DIGIT, false],
	D: [DIGIT, true],
	w: [WORD, false],
	W: [WORD, true],
	s: [SPACE, false],
	S: [SPACE, true],
	h: [HORIZONTAL, false],
	H: [HORIZONTAL, true],
	v: [VERTICAL, false],
	V: [VERTICAL, true]
}

const SIMPLE_ESCAPES: Readonly<Record<string, number>> = { a: 0x07, e: 0x1b, f: 0x0c, n: 0x0a, r: 0x0d, t: 0x09 }

// Escapes PCRE2 knows but this engine does not run: backreferences and subroutine calls (g k), Unicode properties
// (p P X), \R (an atomic group), \K, \C, and the Perl case escapes PCRE2 takes only under options Excel's are unknown for.
const HELD_ESCAPES = new Set(["g", "k", "p", "P", "X", "R", "K", "C", "u", "U", "L", "l", "F"])

const NEWLINE = 0x0a

class Parser {
	private pos = 0
	private depth = 0
	private groups = 0
	private readonly names = new Map<string, number>()
	private readonly duplicates = new Set<string>()
	private flags: Flags
	// Whether the last item read was an option setting such as (?i).
	private optionSetting = false
	private readonly pattern: string

	constructor(pattern: string, caseless: boolean) {
		this.pattern = pattern
		this.flags = {
			caseless,
			multiline: false,
			dotAll: false,
			extended: false,
			extendedMore: false,
			noAutoCapture: false,
			ungreedy: false,
			duplicateNames: false
		}
	}

	parse(): Parsed {
		const root = this.alternation()

		if (this.pos < this.pattern.length) syntax("unmatched closing parenthesis")

		return { root, groups: this.groups, names: this.names, duplicateNames: this.duplicates }
	}

	private peek(offset = 0): number {
		return this.pattern.codePointAt(this.pos + offset) ?? -1
	}

	private next(): number {
		const code = this.pattern.codePointAt(this.pos)

		if (code === undefined) return -1

		this.pos += code > 0xffff ? 2 : 1

		return code
	}

	private eat(text: string): boolean {
		if (!this.pattern.startsWith(text, this.pos)) return false

		this.pos += text.length

		return true
	}

	private skipExtended(): void {
		if (!this.flags.extended) return

		for (;;) {
			const code = this.peek()

			if (code >= 0 && isPatternSpace(code)) this.pos++
			else if (code === 0x23) {
				while (this.pos < this.pattern.length && this.peek() !== NEWLINE) this.pos++
			} else return
		}
	}

	private alternation(): Node {
		const alternatives = [this.concat()]

		while (this.peek() === 0x7c) {
			this.pos++
			alternatives.push(this.concat())
		}

		return alternatives.length === 1 ? (alternatives[0] ?? { type: "empty" }) : { type: "alternate", alternatives }
	}

	private concat(): Node {
		const items: Node[] = []

		for (;;) {
			this.skipExtended()

			const code = this.peek()

			if (code < 0 || code === 0x7c || code === 0x29) break

			const atom = this.atom()

			this.skipExtended()

			const quantified = this.quantifier(atom)

			if (quantified === null) continue
			if (quantified.type === "concat") items.push(...quantified.items)
			else items.push(quantified)
		}

		return items.length === 1 ? (items[0] ?? { type: "empty" }) : items.length === 0 ? { type: "empty" } : { type: "concat", items }
	}

	// A quantifier at the position: [min, max] (max -1 for none) and its length, or null.
	private braces(): [number, number, number] | null {
		const rest = this.pattern.slice(this.pos, this.pos + 32)
		const strict = /^\{(\d+)(,(\d*))?\}/.exec(rest)

		if (strict === null) {
			// Forms newer PCRE2 releases read as quantifiers and older ones as text.
			if (/^\{\s*\d*\s*,?\s*\d*\s*\}/.test(rest) && /\d/.test(/^\{[^}]*\}/.exec(rest)?.[0] ?? "")) unsupported("quantifier form")

			return null
		}

		const min = Number(strict[1])
		const max = strict[2] === undefined ? min : strict[3] === "" ? -1 : Number(strict[3])

		if (min > MAX_REPEAT || max > MAX_REPEAT) syntax("number too big in {} quantifier")
		if (max >= 0 && max < min) syntax("numbers out of order in {} quantifier")

		return [min, max, strict[0].length]
	}

	private isQuantifier(): boolean {
		const code = this.peek()

		return code === 0x2a || code === 0x2b || code === 0x3f || (code === 0x7b && this.braces() !== null)
	}

	private quantifier(atom: Atom | null): Node | null {
		if (!this.isQuantifier()) return atom?.node ?? null
		if (atom === null || !atom.repeatable) {
			// After a comment or \E a quantifier applies to the item before them.
			if (atom === null && !this.optionSetting) unsupported("quantifier after a comment or \\E")

			syntax("quantifier does not follow a repeatable item")
		}

		const code = this.peek()
		let min = code === 0x2b ? 1 : 0
		let max = code === 0x3f ? 1 : -1

		if (code === 0x7b) {
			const [low, high, length] = this.braces() ?? [1, 1, 1]

			min = low
			max = high
			this.pos += length
		} else this.pos++

		let greedy = !this.flags.ungreedy
		let possessive = false

		if (this.peek() === 0x3f) {
			this.pos++
			greedy = !greedy
		} else if (this.peek() === 0x2b) {
			this.pos++
			possessive = true
			greedy = true
		}

		this.skipExtended()

		if (this.isQuantifier()) syntax("quantifier does not follow a repeatable item")

		const node = atom.node

		if (atom.quoted === true && node.type === "concat") {
			const last = node.items[node.items.length - 1]

			if (last === undefined) unsupported("quantified empty quote")

			return {
				type: "concat",
				items: [...node.items.slice(0, -1), { type: "repeat", body: last, min, max, greedy, possessive }]
			}
		}

		return { type: "repeat", body: node, min, max, greedy, possessive }
	}

	private literal(code: number): Node {
		const lower = code | 0x20
		const caseless = this.flags.caseless && (code >= 0x80 || (lower >= 0x61 && lower <= 0x7a))

		return { type: "char", code, caseless }
	}

	private atom(): Atom | null {
		const code = this.next()

		this.optionSetting = false

		switch (code) {
			case 0x28:
				return this.group()
			case 0x5b:
				return { node: this.characterClass(), repeatable: true }
			case 0x2e:
				return { node: { type: "any", newline: this.flags.dotAll }, repeatable: true }
			case 0x5e:
				return this.assertion(this.flags.multiline ? "bolMulti" : "bol")
			case 0x24:
				return this.assertion(this.flags.multiline ? "eolMulti" : "eol")
			case 0x5c:
				return this.escape()
			case 0x2a:
			case 0x2b:
			case 0x3f:
				return syntax("quantifier does not follow a repeatable item")
			case 0x7b:
				this.pos--

				if (this.braces() !== null) syntax("quantifier does not follow a repeatable item")

				this.pos++

				return { node: this.literal(code), repeatable: true }
			default:
				return { node: this.literal(code), repeatable: true }
		}
	}

	private assertion(kind: AssertKind): Atom {
		return { node: { type: "assert", kind }, repeatable: false }
	}

	private name(terminator: number): string {
		const start = this.pos

		if (!isNameStart(this.peek())) syntax("group name must start with a non-digit")

		while (isNameStart(this.peek()) || isDigit(this.peek())) this.pos++

		const name = this.pattern.slice(start, this.pos)

		if (name.length > MAX_NAME) syntax("group name is too long")
		if (this.next() !== terminator) syntax("syntax error in subpattern name (missing terminator?)")

		return name
	}

	private group(): Atom | null {
		let capture = !this.flags.noAutoCapture
		let name: string | undefined
		let scoped: Flags | undefined

		if (this.peek() === 0x2a) unsupported("verb or start-of-pattern option")

		if (this.eat("?")) {
			const code = this.next()

			switch (code) {
				case 0x23: {
					const end = this.pattern.indexOf(")", this.pos)

					if (end < 0) syntax("missing ) after (?# comment")

					this.pos = end + 1

					return null
				}
				case 0x3a:
					capture = false
					break
				case 0x3c:
					if (this.peek() === 0x3d || this.peek() === 0x21 || this.peek() === 0x2a) unsupported("lookbehind assertion")

					name = this.name(0x3e)
					break
				case 0x27:
					name = this.name(0x27)
					break
				case 0x50:
					if (this.eat("<")) name = this.name(0x3e)
					else if (this.peek() === 0x3d || this.peek() === 0x3e) unsupported("named backreference or subroutine call")
					else syntax("unrecognized character after (?P")
					break
				case 0x7c:
				case 0x3e:
				case 0x3d:
				case 0x21:
				case 0x28:
				case 0x26:
				case 0x52:
				case 0x43:
				case 0x2a:
				case 0x2b:
					return unsupported("branch reset, atomic group, lookaround, conditional, recursion or callout")
				default: {
					if (isDigit(code) || (code === 0x2d && isDigit(this.peek()))) unsupported("subroutine call")

					this.pos -= code > 0xffff ? 2 : code < 0 ? 0 : 1

					const flags = this.optionFlags()

					if (this.eat(")")) {
						this.flags = flags
						this.optionSetting = true

						return null
					}

					if (!this.eat(":")) syntax("unrecognized character after (? or (?-")

					scoped = flags
					capture = false
				}
			}
		}

		if (++this.depth > MAX_NESTING) syntax("parentheses are too deeply nested")

		let index = -1

		if (name !== undefined) capture = true

		if (capture) {
			if (++this.groups > MAX_GROUPS) syntax("too many capturing groups")

			index = this.groups
		}

		if (name !== undefined) {
			if (this.names.has(name)) {
				if (!this.flags.duplicateNames) syntax("two named subpatterns have the same name")

				this.duplicates.add(name)
			} else this.names.set(name, index)
		}

		const outer = this.flags

		if (scoped !== undefined) this.flags = scoped
		else this.flags = { ...outer }

		const body = this.alternation()

		if (!this.eat(")")) syntax("missing closing parenthesis")

		this.flags = outer
		this.depth--

		return { node: index < 0 ? body : { type: "group", index, body }, repeatable: true }
	}

	// The option letters of (?imnsxJU-imnsxJU) or (?^...), applied to a copy of the current flags.
	private optionFlags(): Flags {
		const flags = { ...this.flags }
		let on = true

		if (this.eat("^")) {
			flags.caseless = false
			flags.multiline = false
			flags.noAutoCapture = false
			flags.dotAll = false
			flags.extended = false
			flags.extendedMore = false
		}

		for (;;) {
			const code = this.peek()

			if (code === 0x29 || code === 0x3a) return flags

			this.pos++

			switch (code) {
				case 0x2d:
					if (!on) syntax("invalid hyphen in option setting")

					on = false
					break
				case 0x69:
					flags.caseless = on
					break
				case 0x6d:
					flags.multiline = on
					break
				case 0x73:
					flags.dotAll = on
					break
				case 0x6e:
					flags.noAutoCapture = on
					break
				case 0x55:
					flags.ungreedy = on
					break
				case 0x4a:
					flags.duplicateNames = on
					break
				case 0x78:
					if (this.peek() === 0x78) {
						this.pos++
						flags.extendedMore = on
					} else if (!on) flags.extendedMore = false

					flags.extended = on
					break
				default:
					syntax("unrecognized character after (? or (?-")
			}
		}
	}

	// A code point given by digits: up to `max` of them in `radix`.
	private digits(radix: 8 | 16, max: number): number {
		let value = 0
		let count = 0

		while (count < max) {
			const digit = hexValue(this.peek())

			if (digit < 0 || digit >= radix) break

			value = value * radix + digit
			this.pos++
			count++
		}

		return value
	}

	// The code of \x, \o, \c and the octal escapes; the escape letter is read.
	private codeEscape(letter: number): number {
		if (letter === 0x78) {
			if (!this.eat("{")) return this.digits(16, 2)

			const start = this.pos
			const value = this.digits(16, 8)

			if (this.pos === start || !this.eat("}")) syntax("\\x{ is not followed by hexadecimal digits and }")

			return this.checkCode(value)
		}

		if (letter === 0x6f) {
			if (!this.eat("{")) syntax("missing opening brace after \\o")

			const start = this.pos
			const value = this.digits(8, 11)

			if (this.pos === start || !this.eat("}")) syntax("non-octal character in \\o{}")

			return this.checkCode(value)
		}

		if (letter === 0x63) {
			const code = this.next()

			if (code < 0x20 || code > 0x7e) syntax("\\c must be followed by a printable ASCII character")

			return (code >= 0x61 && code <= 0x7a ? code - 0x20 : code) ^ 0x40
		}

		// \0 and (in a class) \1-\7: up to three octal digits.
		this.pos--

		return this.digits(8, 3)
	}

	private checkCode(value: number): number {
		if (value > 0x10ffff) syntax("character code point value is too large")
		if (value >= 0xd800 && value <= 0xdfff) syntax("disallowed Unicode code point (>= 0xd800 && <= 0xdfff)")

		return value
	}

	private classNode(ranges: Ranges, negated: boolean): Node {
		return { type: "class", set: new CharClass(negated ? complement(ranges) : ranges, [], false, false) }
	}

	private escape(): Atom | null {
		const code = this.next()

		if (code < 0) syntax("\\ at end of pattern")

		const letter = String.fromCodePoint(code)
		const escaped = ESCAPE_CLASSES[letter]

		if (escaped !== undefined) return { node: this.classNode(escaped[0], escaped[1]), repeatable: true }

		const simple = SIMPLE_ESCAPES[letter]

		if (simple !== undefined) return { node: this.literal(simple), repeatable: true }

		switch (letter) {
			case "b":
				return this.assertion("wordBoundary")
			case "B":
				return this.assertion("notWordBoundary")
			case "A":
				return this.assertion("start")
			case "z":
				return this.assertion("end")
			case "Z":
				return this.assertion("eol")
			case "G":
				return this.assertion("searchStart")
			case "N":
				if (this.peek() === 0x7b && this.braces() === null) unsupported("\\N{name}")

				return { node: { type: "any", newline: false }, repeatable: true }
			case "Q":
				return { node: this.quoted(), repeatable: true, quoted: true }
			case "E":
				return null
			case "0":
			case "x":
			case "o":
			case "c":
				return { node: this.literal(this.codeEscape(code)), repeatable: true }
		}

		if (isDigit(code)) return { node: this.literal(this.numbered(code)), repeatable: true }
		if (HELD_ESCAPES.has(letter)) unsupported(`\\${letter}`)
		if (isAsciiAlphanumeric(code)) syntax("unrecognized character follows \\")

		return { node: this.literal(code), repeatable: true }
	}

	// \ and a digit 1-9 outside a class: a backreference when the number is below 10, starts with 8 or 9, or names a
	// group opened before it; otherwise up to three octal digits.
	private numbered(first: number): number {
		const start = this.pos - 1
		let end = this.pos

		while (isDigit(this.pattern.charCodeAt(end))) end++

		const number = Number(this.pattern.slice(start, end))

		if (number < 10 || first >= 0x38 || number <= this.groups) unsupported("backreference")

		return this.codeEscape(first)
	}

	// The text of \Q...\E (or to the end of the pattern) as literal characters.
	private quoted(): Node {
		const items: Node[] = []

		while (this.pos < this.pattern.length && !this.eat("\\E")) items.push(this.literal(this.next()))

		return { type: "concat", items }
	}

	private characterClass(): Node {
		const negated = this.eat("^")
		const plain: number[] = []
		const cased: number[] = []
		let first = true

		for (;;) {
			if (this.flags.extendedMore) {
				while (this.peek() === 0x20 || this.peek() === 0x09) this.pos++
			}

			const code = this.peek()

			if (code < 0) syntax("missing terminating ] for character class")
			if (code === 0x5d && !first) {
				this.pos++
				break
			}

			first = false

			const item = this.classItem(plain)

			if (item === null) continue

			// A range: a-b, unless the hyphen ends the class.
			if (this.peek() === 0x2d && this.peek(1) !== 0x5d && this.peek(1) >= 0) {
				const save = this.pos

				this.pos++

				const high = this.classItem(null)

				if (high === null) {
					this.pos = save
					cased.push(item, item)
					continue
				}

				if (item === -1 || high === -1) syntax("invalid range in character class")
				if (high < item) syntax("range out of order in character class")

				cased.push(item, high)
				continue
			}

			if (item >= 0) cased.push(item, item)
		}

		return {
			type: "class",
			set: new CharClass(normalize(plain), normalize(cased), negated, this.flags.caseless)
		}
	}

	// One member of a class: its code point, -1 for a set it added to `sets` (\d, [:alpha:]), or null for nothing (\E,
	// an empty \Q\E). `sets` null: only a code point may follow (the high end of a range; a set there is an error).
	private classItem(sets: number[] | null): number | null {
		const code = this.next()

		if (code === 0x5b && (this.peek() === 0x3a || this.peek() === 0x2e || this.peek() === 0x3d)) {
			const posix = /^([:.=])(\^?)([a-z]+)([:.=])\]/.exec(this.pattern.slice(this.pos, this.pos + 16))

			if (posix !== null && posix[1] === posix[4]) {
				if (posix[1] !== ":") syntax("POSIX collating elements are not supported")

				let name = posix[3] ?? ""

				if (this.flags.caseless && (name === "upper" || name === "lower")) name = "alpha"

				const ranges = POSIX[name]

				if (ranges === undefined) syntax("unknown POSIX class name")

				this.pos += posix[0].length

				if (sets === null) syntax("invalid range in character class")

				sets.push(...(posix[2] === "^" ? complement(ranges) : ranges))

				return -1
			}
		}

		if (code !== 0x5c) return code

		const letter = this.next()

		if (letter < 0) syntax("\\ at end of pattern")

		const name = String.fromCodePoint(letter)
		const escaped = ESCAPE_CLASSES[name]

		if (escaped !== undefined) {
			if (sets === null) syntax("invalid range in character class")

			sets.push(...(escaped[1] ? complement(escaped[0]) : escaped[0]))

			return -1
		}

		const simple = SIMPLE_ESCAPES[name]

		if (simple !== undefined) return simple

		switch (name) {
			case "b":
				return 0x08
			case "E":
				return null
			case "Q": {
				// \Q...\E in a class: each character a member; the last one returned (a range may start there).
				let last: number | null = null

				while (this.pos < this.pattern.length && !this.eat("\\E")) {
					if (last !== null && sets !== null) sets.push(last, last)

					last = this.next()
				}

				return last
			}
			case "x":
			case "o":
			case "c":
				return this.codeEscape(letter)
		}

		if (letter >= 0x30 && letter <= 0x37) return this.codeEscape(letter)
		if (letter === 0x38 || letter === 0x39) return letter
		if (name === "p" || name === "P" || name === "u" || name === "U" || name === "L" || name === "l" || name === "F")
			unsupported(`\\${name}`)
		if (isAsciiAlphanumeric(letter)) syntax("escape sequence is invalid in character class")

		return letter
	}
}

// Checks possessive quantifiers: one is run as a greedy one where that matches the same, which this engine proves
// only for a single-character body followed by a disjoint single character, an end anchor it cannot reach past, or
// the end of the pattern. Any other is held.
function checkPossessive(node: Node, atEnd: boolean): void {
	switch (node.type) {
		case "group":
			checkPossessive(node.body, atEnd)
			return
		case "alternate":
			for (const alternative of node.alternatives) checkPossessive(alternative, atEnd)

			return
		case "repeat":
			if (node.possessive && node.min !== node.max && !(atEnd && single(node.body))) unsupported("possessive quantifier")

			checkPossessive(node.body, false)
			return
		case "concat":
			node.items.forEach((item, index) => {
				const after = node.items[index + 1]

				if (item.type === "repeat" && item.possessive && item.min !== item.max && after !== undefined && single(item.body)) {
					if (!disjoint(item.body, after)) unsupported("possessive quantifier")

					checkPossessive(item.body, false)
					return
				}

				checkPossessive(item, atEnd && index === node.items.length - 1)
			})
			return
		default:
			return
	}
}

function single(node: Node): boolean {
	return node.type === "char" || node.type === "class" || node.type === "any"
}

// Whether no character `body` matches lets `after` match where it stands.
function disjoint(body: Node, after: Node): boolean {
	const matches = (code: number): boolean => {
		if (body.type === "char") return body.code === code || (body.caseless && caselessEqual(body.code, code))
		if (body.type === "class") return body.set.matches(code)

		return body.type === "any" && (body.newline || code !== NEWLINE)
	}

	if (after.type === "assert") {
		if (after.kind === "end") return true
		if (after.kind === "eol") return !matches(NEWLINE)

		return false
	}

	if (after.type !== "char" || after.caseless) return false

	return !matches(after.code)
}

function caselessEqual(a: number, b: number): boolean {
	return new CharClass([], [a, a], false, true).matches(b)
}

export function parse(pattern: string, caseless: boolean): Parsed {
	const parsed = new Parser(pattern, caseless).parse()

	checkPossessive(parsed.root, true)

	return parsed
}
