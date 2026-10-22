import { CharClass } from "./charset"

// The parsed pattern.

export type AssertKind =
	| "bol" // ^
	| "bolMulti" // ^ under (?m)
	| "eol" // $ and \Z
	| "eolMulti" // $ under (?m)
	| "start" // \A
	| "end" // \z
	| "searchStart" // \G
	| "wordBoundary" // \b
	| "notWordBoundary" // \B

export type Node =
	| { readonly type: "empty" }
	| { readonly type: "char"; readonly code: number; readonly caseless: boolean }
	| { readonly type: "class"; readonly set: CharClass }
	| { readonly type: "any"; readonly newline: boolean }
	| { readonly type: "assert"; readonly kind: AssertKind }
	| { readonly type: "group"; readonly index: number; readonly body: Node }
	| { readonly type: "concat"; readonly items: readonly Node[] }
	| { readonly type: "alternate"; readonly alternatives: readonly Node[] }
	| {
			readonly type: "repeat"
			readonly body: Node
			readonly min: number
			// -1: unbounded.
			readonly max: number
			readonly greedy: boolean
			readonly possessive: boolean
	  }

export interface Parsed {
	readonly root: Node
	// Capturing groups, not counting the whole match.
	readonly groups: number
	// Group numbers by name; a name given twice (under (?J)) is in `duplicateNames`.
	readonly names: ReadonlyMap<string, number>
	readonly duplicateNames: ReadonlySet<string>
}

// A pattern PCRE2 rejects: Excel gives #VALUE!.
export class RegexSyntaxError extends Error {}

// A pattern PCRE2 takes that this engine cannot run in linear time (backreferences, lookaround, atomic groups,
// recursion, conditionals) or cannot be sure it runs as PCRE2 does: the formula keeps the file's result.
export class RegexUnsupportedError extends Error {}

// Whether a node can match without consuming anything.
export function nullable(node: Node): boolean {
	switch (node.type) {
		case "empty":
		case "assert":
			return true
		case "char":
		case "class":
		case "any":
			return false
		case "group":
			return nullable(node.body)
		case "concat":
			return node.items.every(nullable)
		case "alternate":
			return node.alternatives.some(nullable)
		case "repeat":
			return node.min === 0 || nullable(node.body)
	}
}
