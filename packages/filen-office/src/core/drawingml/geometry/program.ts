// Shape geometry as a compact program (ECMA-376 Part 1 §20.1.9): adjust values, guide formulas, the text
// rectangle and paths. Preset definitions ship as strings in this encoding (presets.generated.ts); custom
// geometry compiles to the same structure, so one evaluator serves both.
//
// Encoding, tokens separated by single spaces, sections by "|":
//   adjust values   "name default name default ..."
//   guides          an op code followed by its arguments, repeated
//   text rectangle  "l t r b", or empty for the whole shape
//   paths           "P flags w h" opens a path (w, h 0 when absent), then commands:
//                   "M x y", "L x y", "A wR hR stAng swAng", "Q x1 y1 x2 y2", "C x1 y1 x2 y2 x3 y3", "Z"
// An argument is a decimal literal or a slot reference: letters, the first from [A-Za-z] and the rest from
// [A-Za-z0-9], read as a number in base 52/62. Slots number the built-in guides, then the adjust values,
// then the guides.

// Built-in guides (§20.1.9.11 "Shape Guide" presets), most used first so most references are one letter.
export const BUILTIN_GUIDES = [
	"b",
	"hc",
	"r",
	"vc",
	"l",
	"t",
	"w",
	"h",
	"ss",
	"cd4",
	"wd2",
	"cd2",
	"hd2",
	"3cd4",
	"wd32",
	"hd4",
	"wd8",
	"ssd6",
	"wd4",
	"ssd8",
	"wd6",
	"hd3",
	"wd3",
	"hd5",
	"hd6",
	"wd5",
	"ssd16",
	"ssd32",
	"hd8",
	"wd12",
	"ls",
	"cd8",
	"3cd8",
	"5cd8",
	"7cd8",
	"wd10",
	"wd16",
	"hd10",
	"hd12",
	"hd16",
	"hd32",
	"ssd2",
	"ssd4"
] as const

export const BUILTIN_COUNT = BUILTIN_GUIDES.length

const BUILTIN_INDEX: ReadonlyMap<string, number> = new Map(BUILTIN_GUIDES.map((name, index) => [name, index]))

export function builtinIndex(name: string): number {
	return BUILTIN_INDEX.get(name) ?? -1
}

// 60000ths of a degree per full turn.
const CIRCLE = 21600000

// Writes the built-in guide values for a shape of w x h into slots [0, BUILTIN_COUNT).
export function fillBuiltins(slots: Float64Array, w: number, h: number): void {
	const ss = Math.min(w, h)
	const values: Readonly<Record<string, number>> = {
		b: h,
		hc: w / 2,
		r: w,
		vc: h / 2,
		l: 0,
		t: 0,
		w,
		h,
		ss,
		ls: Math.max(w, h),
		cd2: CIRCLE / 2,
		cd4: CIRCLE / 4,
		cd8: CIRCLE / 8,
		"3cd4": (CIRCLE * 3) / 4,
		"3cd8": (CIRCLE * 3) / 8,
		"5cd8": (CIRCLE * 5) / 8,
		"7cd8": (CIRCLE * 7) / 8
	}

	for (let index = 0; index < BUILTIN_COUNT; index++) {
		const name = BUILTIN_GUIDES[index] ?? ""
		const known = values[name]

		if (known !== undefined) {
			slots[index] = known

			continue
		}

		// wdN, hdN, ssdN.
		const divisor = Number(name.slice(name.startsWith("ssd") ? 3 : 2))
		const base = name.startsWith("ssd") ? ss : name.startsWith("w") ? w : h

		slots[index] = base / divisor
	}
}

// Guide operators (§20.1.9.11) with their single-letter codes and arity.
export const GuideOp = {
	VAL: 0,
	MUL_DIV: 1,
	ADD_SUB: 2,
	ADD_DIV: 3,
	IF_ELSE: 4,
	ABS: 5,
	AT2: 6,
	CAT2: 7,
	COS: 8,
	MAX: 9,
	MIN: 10,
	MOD: 11,
	PIN: 12,
	SAT2: 13,
	SIN: 14,
	SQRT: 15,
	TAN: 16
} as const

export type GuideOp = (typeof GuideOp)[keyof typeof GuideOp]

// Formula operator, its code letter in the encoding, its arity.
export const GUIDE_OPS: readonly (readonly [string, string, number])[] = [
	["val", "v", 1],
	["*/", "m", 3],
	["+-", "a", 3],
	["+/", "d", 3],
	["?:", "q", 3],
	["abs", "b", 1],
	["at2", "t", 2],
	["cat2", "c", 3],
	["cos", "o", 2],
	["max", "x", 2],
	["min", "n", 2],
	["mod", "h", 3],
	["pin", "p", 3],
	["sat2", "s", 3],
	["sin", "i", 2],
	["sqrt", "r", 1],
	["tan", "g", 2]
]

export const OP_BY_FORMULA: ReadonlyMap<string, number> = new Map(GUIDE_OPS.map(([formula], op) => [formula, op]))

const OP_BY_CODE: ReadonlyMap<string, number> = new Map(GUIDE_OPS.map(([, code], op) => [code, op]))

export function opArity(op: number): number {
	return GUIDE_OPS[op]?.[2] ?? 0
}

// Path fill modes (ST_PathFillMode) in flags bits 0-2; bit 3 set when the path is not stroked.
export const PathFill = { NORM: 0, NONE: 1, LIGHTEN: 2, LIGHTEN_LESS: 3, DARKEN: 4, DARKEN_LESS: 5 } as const

export const PATH_FILL_MASK = 7
export const PATH_NO_STROKE = 8

export const PATH_FILL_NAMES: readonly string[] = ["norm", "none", "lighten", "lightenLess", "darken", "darkenLess"]

// Geometry path commands in a compiled program; the arcs here are still DrawingML arcs.
export const GeomCommand = { PATH: 0, MOVE: 1, LINE: 2, ARC: 3, QUAD: 4, CUBIC: 5, CLOSE: 6 } as const

export const GEOM_COMMAND_LETTERS = "PMLAQCZ"

// Arguments per command (PATH: flags, w, h).
export const GEOM_COMMAND_ARITY: readonly number[] = [3, 2, 2, 4, 4, 6, 0]

// A compiled geometry. Arguments are slot indices (>= 0) or -(1 + index) into `literals`.
export interface GeometryProgram {
	readonly adjustNames: readonly string[]
	readonly adjustDefaults: Float64Array
	readonly slotCount: number
	// Per guide: op, then three argument slots (unused ones 0).
	readonly guides: Int32Array
	readonly guideCount: number
	// l, t, r, b arguments, or null for the whole shape.
	readonly rect: Int32Array | null
	// Commands with their arguments inline: code, args...; PATH's flags/w/h are literals.
	readonly paths: Int32Array
	readonly literals: Float64Array
}

const FIRST_ALPHABET = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
const REST_ALPHABET = `${FIRST_ALPHABET}0123456789`

// Slot index -> reference token.
export function encodeSlot(index: number): string {
	if (index < FIRST_ALPHABET.length) return FIRST_ALPHABET[index] ?? "a"

	let rest = index - FIRST_ALPHABET.length
	let width = 1
	let span = FIRST_ALPHABET.length * REST_ALPHABET.length

	while (rest >= span) {
		rest -= span
		span *= REST_ALPHABET.length
		width++
	}

	let tail = ""

	for (let k = 0; k < width; k++) {
		tail = (REST_ALPHABET[rest % REST_ALPHABET.length] ?? "a") + tail
		rest = Math.floor(rest / REST_ALPHABET.length)
	}

	return (FIRST_ALPHABET[rest] ?? "a") + tail
}

function decodeSlot(token: string): number {
	let index = FIRST_ALPHABET.indexOf(token[0] ?? "")

	if (index < 0) return -1
	if (token.length === 1) return index

	let offset = FIRST_ALPHABET.length
	let span = FIRST_ALPHABET.length * REST_ALPHABET.length

	for (let width = 2; width < token.length; width++) {
		offset += span
		span *= REST_ALPHABET.length
	}

	for (let k = 1; k < token.length; k++) {
		const digit = REST_ALPHABET.indexOf(token[k] ?? "")

		if (digit < 0) return -1

		index = index * REST_ALPHABET.length + digit
	}

	return offset + index
}

// Collects literals while compiling, sharing repeated values.
export class LiteralPool {
	readonly values: number[] = []
	private readonly index = new Map<number, number>()

	ref(value: number): number {
		let at = this.index.get(value)

		if (at === undefined) {
			at = this.values.length
			this.values.push(value)
			this.index.set(value, at)
		}

		return -(1 + at)
	}
}

// Parses one encoded preset. Malformed input yields a program with whatever parsed; never throws.
export function parseProgram(encoded: string): GeometryProgram {
	const sections = encoded.split("|")
	const pool = new LiteralPool()
	const argument = (token: string | undefined): number => {
		if (token === undefined || token === "") return pool.ref(0)

		const first = token.charCodeAt(0)

		if ((first >= 48 && first <= 57) || first === 45) {
			const value = Number(token)

			return pool.ref(Number.isFinite(value) ? value : 0)
		}

		const slot = decodeSlot(token)

		return slot < 0 ? pool.ref(0) : slot
	}
	const tokens = (section: string | undefined): string[] => (section === undefined || section === "" ? [] : section.split(" "))
	const adjustTokens = tokens(sections[0])
	const adjustNames: string[] = []
	const adjustDefaults: number[] = []

	for (let k = 0; k + 1 < adjustTokens.length; k += 2) {
		adjustNames.push(adjustTokens[k] ?? "")
		adjustDefaults.push(Number(adjustTokens[k + 1]) || 0)
	}

	const guideTokens = tokens(sections[1])
	const guides: number[] = []

	for (let k = 0; k < guideTokens.length;) {
		const op = OP_BY_CODE.get(guideTokens[k] ?? "") ?? -1

		if (op < 0) break

		const arity = opArity(op)

		guides.push(
			op,
			argument(guideTokens[k + 1]),
			arity > 1 ? argument(guideTokens[k + 2]) : 0,
			arity > 2 ? argument(guideTokens[k + 3]) : 0
		)
		k += 1 + arity
	}

	const rectTokens = tokens(sections[2])
	const rect = rectTokens.length === 4 ? Int32Array.from(rectTokens.map(argument)) : null
	const pathTokens = tokens(sections[3])
	const paths: number[] = []

	for (let k = 0; k < pathTokens.length;) {
		const command = GEOM_COMMAND_LETTERS.indexOf(pathTokens[k] ?? "")

		if (command < 0) break

		paths.push(command)

		const arity = GEOM_COMMAND_ARITY[command] ?? 0

		for (let a = 1; a <= arity; a++) paths.push(argument(pathTokens[k + a]))

		k += 1 + arity
	}

	return {
		adjustNames,
		adjustDefaults: Float64Array.from(adjustDefaults),
		slotCount: BUILTIN_COUNT + adjustNames.length + guides.length / 4,
		guides: Int32Array.from(guides),
		guideCount: guides.length / 4,
		rect,
		paths: Int32Array.from(paths),
		literals: Float64Array.from(pool.values)
	}
}
