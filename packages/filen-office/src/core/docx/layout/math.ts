import type { MathElement, MathNode } from "../../math/index"

// The size an equation takes in a line, estimated from its MathML tree with TeX's layout rules in
// simplified form (script sizes and shifts, fraction and radical clearances, stretched fences, n-ary
// operators, tables): close enough for line heights and wrapping, without the math font's MATH table.

export interface MathBox {
	readonly width: number
	// Above and below the baseline.
	readonly ascent: number
	readonly descent: number
}

// Width of `text` set at `size` px.
export type MathMeasure = (text: string, size: number) => number

// Shares of the em. Script sizes (70 %, 50 %) and the medium and thick operator spaces (4/18, 5/18 em)
// are TeX's (The TeXbook, chapter 18 and appendix G); the token extents, shifts and clearances are round
// estimates of the same order as TeX's math parameters, not any font's values.
const ASCENT = 0.72
const DESCENT = 0.22
// The math axis (fraction bars, operators) above the baseline.
const AXIS = 0.25
const SCRIPT_SCALE = 0.7
const SCRIPT_SCRIPT_SCALE = 0.5
const SUPERSCRIPT_SHIFT = 0.4
const SUBSCRIPT_SHIFT = 0.2
const SCRIPT_SPACE = 0.05
const FRACTION_GAP = 0.12
const FRACTION_PADDING = 0.2
const RADICAL_WIDTH = 0.8
const RADICAL_GAP = 0.2
const LIMIT_GAP = 0.1
const ACCENT_HEIGHT = 0.25
const COLUMN_GAP = 0.8
const ROW_GAP = 0.2
const DISPLAY_OPERATOR = 1.4
const THICK_SPACE = 5 / 18
const MEDIUM_SPACE = 4 / 18
const RELATIONS = new Set([..."=<>≤≥≠≈≡≅∼≃∝→←↔⇒⇐⇔∈∉∋⊂⊃⊆⊇≪≫∣∥⊥:"])
const BINARY = new Set([..."+−-±∓×÷·∗∘∪∩⊕⊗⊖∧∨"])
// Nesting deeper than this is measured as its text only.
const MAX_DEPTH = 64

interface Style {
	readonly size: number
	readonly level: number
	readonly display: boolean
}

const EMPTY: MathBox = { width: 0, ascent: 0, descent: 0 }

function element(node: MathNode | undefined): MathElement | undefined {
	return node === undefined || typeof node === "string" ? undefined : node
}

function text(node: MathElement): string {
	return node.children.map(child => (typeof child === "string" ? child : text(child))).join("")
}

// Font size factor of a script level.
function factor(level: number): number {
	return level <= 0 ? 1 : level === 1 ? SCRIPT_SCALE : SCRIPT_SCRIPT_SCALE
}

// The style `levels` script levels down: smaller, never display style.
function scripted(style: Style, levels = 1): Style {
	const level = style.level + levels

	return { size: (style.size / factor(style.level)) * factor(level), level, display: false }
}

function isLargeOperator(node: MathElement | undefined): boolean {
	return node !== undefined && node.tag === "mo" && node.attributes?.largeop === true
}

// Children side by side; stretchy fences reach as far as the tallest of them, which they already do.
function row(nodes: readonly MathNode[], style: Style, measure: MathMeasure, depth: number): MathBox {
	let width = 0
	let ascent = 0
	let descent = 0

	for (const node of nodes) {
		const own = typeof node === "string" ? token(node, style, measure) : box(node, style, measure, depth + 1)

		width += own.width
		ascent = Math.max(ascent, own.ascent)
		descent = Math.max(descent, own.descent)
	}

	return { width, ascent, descent }
}

function token(value: string, style: Style, measure: MathMeasure): MathBox {
	return { width: measure(value, style.size), ascent: ASCENT * style.size, descent: DESCENT * style.size }
}

function operator(node: MathElement, style: Style, measure: MathMeasure): MathBox {
	const value = text(node)
	const attributes = node.attributes
	const em = style.size
	const large = attributes?.largeop === true && style.display
	const size = large ? em * DISPLAY_OPERATOR : em
	const own = token(value, { ...style, size }, measure)
	const infix = attributes?.form === undefined || attributes.form === "infix"
	const inScript = style.level > 0
	const space =
		inScript || !infix
			? 0
			: value.length === 1 && RELATIONS.has(value)
				? THICK_SPACE
				: value.length === 1 && BINARY.has(value)
					? MEDIUM_SPACE
					: 0
	const left = attributes?.lspace ?? space
	const right = attributes?.rspace ?? space

	return { width: own.width + (left + right) * em, ascent: own.ascent, descent: own.descent }
}

// Base with scripts after it (subscript, superscript, or both).
function scripts(base: MathBox, sub: MathBox | undefined, sup: MathBox | undefined, style: Style): MathBox {
	const em = style.size
	let ascent = base.ascent
	let descent = base.descent

	if (sup !== undefined) ascent = Math.max(ascent, SUPERSCRIPT_SHIFT * em + sup.ascent + sup.descent * 0.5)
	if (sub !== undefined) descent = Math.max(descent, SUBSCRIPT_SHIFT * em + sub.descent + sub.ascent * 0.5)

	return {
		width: base.width + Math.max(sub?.width ?? 0, sup?.width ?? 0) + (sub !== undefined || sup !== undefined ? SCRIPT_SPACE * em : 0),
		ascent,
		descent
	}
}

// Base with limits above and below it.
function limits(base: MathBox, under: MathBox | undefined, over: MathBox | undefined, accent: boolean, style: Style): MathBox {
	const em = style.size
	const gap = LIMIT_GAP * em

	return {
		width: Math.max(base.width, under?.width ?? 0, over?.width ?? 0),
		ascent: base.ascent + (over === undefined ? 0 : accent ? ACCENT_HEIGHT * em : over.ascent + over.descent + gap),
		descent: base.descent + (under === undefined ? 0 : under.ascent + under.descent + gap)
	}
}

function table(node: MathElement, style: Style, measure: MathMeasure, depth: number): MathBox {
	const widths: number[] = []
	let height = 0
	let rows = 0

	for (const child of node.children) {
		const tr = element(child)

		if (tr === undefined) continue

		let ascent = 0
		let descent = 0

		tr.children.forEach((cell, column) => {
			const own = typeof cell === "string" ? token(cell, style, measure) : box(cell, style, measure, depth + 1)

			widths[column] = Math.max(widths[column] ?? 0, own.width)
			ascent = Math.max(ascent, own.ascent)
			descent = Math.max(descent, own.descent)
		})

		height += ascent + descent
		rows++
	}

	const em = style.size

	height += Math.max(0, rows - 1) * ROW_GAP * em

	const axis = AXIS * em

	return {
		width: widths.reduce((sum, width) => sum + width, 0) + Math.max(0, widths.length - 1) * COLUMN_GAP * em,
		ascent: height / 2 + axis,
		descent: Math.max(0, height / 2 - axis)
	}
}

function box(node: MathElement, style: Style, measure: MathMeasure, depth: number): MathBox {
	if (depth > MAX_DEPTH) return token(text(node), style, measure)

	const attributes = node.attributes
	const em = style.size
	const children = node.children
	const at = (k: number): MathBox => {
		const child = children[k]

		return child === undefined
			? EMPTY
			: typeof child === "string"
				? token(child, style, measure)
				: box(child, style, measure, depth + 1)
	}
	const scriptAt = (k: number, levels = 1): MathBox | undefined => {
		const child = children[k]

		return child === undefined
			? undefined
			: typeof child === "string"
				? token(child, scripted(style, levels), measure)
				: box(child, scripted(style, levels), measure, depth + 1)
	}

	switch (node.tag) {
		case "mi":
		case "mn":
		case "mtext":
			return token(text(node), style, measure)
		case "mo":
			return operator(node, style, measure)
		case "mspace":
			return { width: (attributes?.width ?? 0) * em, ascent: (attributes?.height ?? 0) * em, descent: (attributes?.depth ?? 0) * em }
		case "math":
		case "mstyle": {
			const display = attributes?.displaystyle ?? (node.tag === "math" && attributes?.display === "block" ? true : style.display)
			const levels = attributes?.scriptlevel ?? 0
			const inner = levels > 0 ? { ...scripted(style, levels), display } : { ...style, display }

			return row(children, inner, measure, depth)
		}
		case "mfrac": {
			const inner = style.display ? { ...style, display: false } : scripted(style)
			const numerator = children[0]
			const denominator = children[1]
			const top =
				numerator === undefined
					? EMPTY
					: typeof numerator === "string"
						? token(numerator, inner, measure)
						: box(numerator, inner, measure, depth + 1)
			const bottom =
				denominator === undefined
					? EMPTY
					: typeof denominator === "string"
						? token(denominator, inner, measure)
						: box(denominator, inner, measure, depth + 1)
			const gap = FRACTION_GAP * em
			const axis = AXIS * em

			return {
				width: Math.max(top.width, bottom.width) + FRACTION_PADDING * em,
				ascent: axis + gap + top.descent + top.ascent,
				descent: Math.max(0, gap + bottom.ascent + bottom.descent - axis)
			}
		}
		case "msqrt": {
			const inner = row(children, style, measure, depth)

			return {
				width: inner.width + RADICAL_WIDTH * em,
				ascent: inner.ascent + RADICAL_GAP * em,
				descent: inner.descent + LIMIT_GAP * em
			}
		}
		case "mroot": {
			const inner = at(0)
			const index = scriptAt(1, 2)

			return {
				width: inner.width + RADICAL_WIDTH * em + Math.max(0, (index?.width ?? 0) - RADICAL_WIDTH * em * 0.5),
				ascent: inner.ascent + RADICAL_GAP * em,
				descent: inner.descent + LIMIT_GAP * em
			}
		}
		case "msub":
			return scripts(at(0), scriptAt(1), undefined, style)
		case "msup":
			return scripts(at(0), undefined, scriptAt(1), style)
		case "msubsup":
			return scripts(at(0), scriptAt(1), scriptAt(2), style)
		case "munder":
		case "mover":
		case "munderover": {
			const base = element(children[0])
			const operatorBase = isLargeOperator(base) || (base?.tag === "mo" && base.attributes?.movablelimits === true)

			// Limits of an inline n-ary operator sit beside it as scripts.
			if (operatorBase && !style.display && base?.attributes?.movablelimits !== false) {
				return node.tag === "munder"
					? scripts(at(0), scriptAt(1), undefined, style)
					: node.tag === "mover"
						? scripts(at(0), undefined, scriptAt(1), style)
						: scripts(at(0), scriptAt(1), scriptAt(2), style)
			}

			const accent = attributes?.accent === true
			const under = node.tag === "mover" ? undefined : scriptAt(1)
			const over = node.tag === "munder" ? undefined : scriptAt(node.tag === "mover" ? 1 : 2, accent ? 0 : 1)

			return limits(at(0), under, over, accent, style)
		}
		case "mmultiscripts": {
			let width = 0
			let ascent = 0
			let descent = 0
			const base = at(0)

			for (let k = 1; k < children.length; k++) {
				const child = element(children[k])

				if (child?.tag === "mprescripts") continue

				const own = scriptAt(k) ?? EMPTY

				width += k % 2 === 1 ? own.width : 0
				ascent = Math.max(ascent, SUPERSCRIPT_SHIFT * em + own.ascent)
				descent = Math.max(descent, SUBSCRIPT_SHIFT * em + own.descent)
			}

			return {
				width: base.width + width + SCRIPT_SPACE * em,
				ascent: Math.max(base.ascent, ascent),
				descent: Math.max(base.descent, descent)
			}
		}
		case "mtable":
			return table(node, style, measure, depth)
		case "mphantom":
		case "mpadded":
		case "mrow":
		case "mtr":
		case "mtd":
		case "mprescripts":
			return row(children, style, measure, depth)
	}
}

// The box of an equation set at `size` px; `display` for equations on a line of their own (m:oMathPara).
export function mathBox(tree: MathElement, size: number, display: boolean, measure: MathMeasure): MathBox {
	return box(tree, { size, level: 0, display }, measure, 0)
}
