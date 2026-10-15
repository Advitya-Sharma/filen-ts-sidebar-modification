// A MathML element tree as plain data, for a renderer to build with createElementNS. Tags come from a fixed
// MathML Core allowlist and every attribute is a typed, validated value: document text only ever reaches
// the tree as the character data of token elements.

export const MATH_TAGS = [
	"math",
	"mrow",
	"mi",
	"mn",
	"mo",
	"mtext",
	"mspace",
	"mfrac",
	"msqrt",
	"mroot",
	"msub",
	"msup",
	"msubsup",
	"munder",
	"mover",
	"munderover",
	"mmultiscripts",
	"mprescripts",
	"mtable",
	"mtr",
	"mtd",
	"mpadded",
	"mphantom",
	"mstyle"
] as const

export type MathTag = (typeof MATH_TAGS)[number]

const TAG_SET: ReadonlySet<string> = new Set(MATH_TAGS)

// Elements whose children are character data (exactly one string).
const TOKEN_TAGS: ReadonlySet<string> = new Set(["mi", "mn", "mo", "mtext"])

// Borders and strikes of m:borderBox (MathML Core has no menclose): the renderer draws them, e.g. as CSS
// borders and a background line on the element carrying the hint.
export const MathFrame = {
	TOP: 1,
	BOTTOM: 2,
	LEFT: 4,
	RIGHT: 8,
	STRIKE_HORIZONTAL: 16,
	STRIKE_VERTICAL: 32,
	// Bottom-left to top-right.
	STRIKE_UP: 64,
	// Top-left to bottom-right.
	STRIKE_DOWN: 128
} as const

export const MATH_FRAME_ALL = 255

// Horizontal alignment of a display equation (m:oMathParaPr/m:jc); "centerGroup" centres the group of
// equations as a block and aligns them inside it. Not a MathML attribute: the renderer positions the
// block.
export type MathAlign = "left" | "right" | "center" | "centerGroup"

export type MathColumnAlign = "left" | "center" | "right"

export interface MathAttributes {
	// math
	readonly display?: "block" | "inline"
	readonly align?: MathAlign
	// math, mstyle
	readonly displaystyle?: boolean
	// mstyle: a relative change of scriptlevel, written with its sign ("+1", "-1").
	readonly scriptlevel?: number
	// mi: upright single letters.
	readonly mathvariant?: "normal"
	// mo
	readonly form?: "prefix" | "infix" | "postfix"
	readonly stretchy?: boolean
	readonly symmetric?: boolean
	readonly largeop?: boolean
	readonly movablelimits?: boolean
	readonly fence?: boolean
	readonly separator?: boolean
	// Lengths in em.
	readonly lspace?: number
	readonly rspace?: number
	readonly minsize?: number
	readonly maxsize?: number
	// mfrac: only ever 0 (no fraction bar); absent means the default rule.
	readonly linethickness?: number
	// munder / mover / munderover
	readonly accent?: boolean
	readonly accentunder?: boolean
	// mspace, mpadded: lengths in em.
	readonly width?: number
	readonly height?: number
	readonly depth?: number
	// mtd
	readonly columnalign?: MathColumnAlign
	// MathFrame bits.
	readonly frame?: number
	// Six upper-case hex digits.
	readonly color?: string
	// A tracked change in the markup view (omml.ts MathRevision): drawn in its author's colour, underlined
	// (inserted) or struck through (deleted).
	readonly revision?: number
}

export interface MathElement {
	readonly tag: MathTag
	readonly attributes?: MathAttributes
	readonly children: readonly MathNode[]
}

export type MathNode = MathElement | string

export function isMathTag(tag: string): tag is MathTag {
	return TAG_SET.has(tag)
}

export function isTokenTag(tag: MathTag): boolean {
	return TOKEN_TAGS.has(tag)
}

export class MathTreeError extends Error {
	override readonly name = "MathTreeError"
}

const BOOLEAN_ATTRIBUTES = new Set([
	"displaystyle",
	"stretchy",
	"symmetric",
	"largeop",
	"movablelimits",
	"fence",
	"separator",
	"accent",
	"accentunder"
])

// Em lengths: finite and within a sane range, so no attribute can blow up the layout.
const LENGTH_ATTRIBUTES = new Set(["lspace", "rspace", "minsize", "maxsize", "linethickness", "width", "height", "depth"])
const MAX_LENGTH_EM = 100

const ENUMS: Readonly<Record<string, ReadonlySet<string>>> = {
	display: new Set(["block", "inline"]),
	align: new Set(["left", "right", "center", "centerGroup"]),
	mathvariant: new Set(["normal"]),
	form: new Set(["prefix", "infix", "postfix"]),
	columnalign: new Set(["left", "center", "right"])
}

// Exact child counts of the fixed-arity layout schemata.
const ARITY: Readonly<Record<string, number>> = {
	mfrac: 2,
	mroot: 2,
	msub: 2,
	msup: 2,
	munder: 2,
	mover: 2,
	msubsup: 3,
	munderover: 3,
	mprescripts: 0,
	mspace: 0
}

// Deeper trees are refused: the converter never builds one, and a renderer recursing over it must not
// exhaust its stack.
export const MAX_MATH_DEPTH = 1024

function fail(message: string): never {
	throw new MathTreeError(message)
}

function assertAttributes(attributes: unknown): void {
	if (typeof attributes !== "object" || attributes === null || Array.isArray(attributes)) fail("Attributes must be an object")

	for (const [key, value] of Object.entries(attributes)) {
		if (value === undefined) continue

		if (BOOLEAN_ATTRIBUTES.has(key)) {
			if (typeof value !== "boolean") fail(`${key} must be a boolean`)
		} else if (LENGTH_ATTRIBUTES.has(key)) {
			if (typeof value !== "number" || !Number.isFinite(value) || Math.abs(value) > MAX_LENGTH_EM) fail(`${key} must be a length`)
		} else if (key === "scriptlevel") {
			if (typeof value !== "number" || !Number.isInteger(value) || Math.abs(value) > 8) fail("scriptlevel must be a small integer")
		} else if (key === "frame") {
			if (typeof value !== "number" || !Number.isInteger(value) || value < 0 || value > MATH_FRAME_ALL)
				fail("frame must be MathFrame bits")
		} else if (key === "revision") {
			if (typeof value !== "number" || !Number.isInteger(value) || value < 1 || value > 63 || (value & 3) === 0 || (value & 3) === 3)
				fail("revision must be a tracked look")
		} else if (key === "color") {
			if (typeof value !== "string" || !/^[0-9A-F]{6}$/.test(value)) fail("color must be six upper-case hex digits")
		} else {
			const allowed = ENUMS[key]

			if (allowed === undefined) fail(`Unknown attribute ${key}`)
			if (typeof value !== "string" || !allowed.has(value)) fail(`Invalid ${key}`)
		}
	}
}

function assertElement(node: unknown, depth: number): void {
	if (depth > MAX_MATH_DEPTH) fail("Math tree too deep")
	if (typeof node !== "object" || node === null) fail("Expected an element")

	const element = node as { tag?: unknown; attributes?: unknown; children?: unknown }
	const tag = element.tag

	if (typeof tag !== "string" || !TAG_SET.has(tag)) fail(`Tag not allowed: ${String(tag)}`)
	if (element.attributes !== undefined) assertAttributes(element.attributes)
	if (!Array.isArray(element.children)) fail("Children must be an array")

	const children = element.children as unknown[]
	const arity = ARITY[tag]

	if (arity !== undefined && children.length !== arity) fail(`${tag} takes ${arity} children`)

	if (TOKEN_TAGS.has(tag)) {
		if (children.length !== 1 || typeof children[0] !== "string") fail(`${tag} holds exactly one string`)

		return
	}

	if (tag === "mmultiscripts") {
		const prescripts = children.findIndex(child => (child as { tag?: unknown } | null)?.tag === "mprescripts")

		if (children.length === 0 || prescripts === 0) fail("mmultiscripts needs a base")
		if ((prescripts < 0 ? children.length - 1 : prescripts - 1) % 2 !== 0) fail("mmultiscripts needs script pairs")
		if (prescripts >= 0 && (children.length - prescripts - 1) % 2 !== 0) fail("mmultiscripts needs prescript pairs")
	}

	for (const child of children) {
		if (typeof child === "string") fail(`${tag} cannot hold text`)

		assertElement(child, depth + 1)
	}
}

// Throws a MathTreeError unless `node` is a math tree this module could have produced: allowlisted tags,
// typed attributes, text only inside token elements, layout schemata with their child counts.
export function assertMathTree(node: unknown): asserts node is MathElement {
	assertElement(node, 0)
}
