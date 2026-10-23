import type { MathElement, MathNode } from "../../math/index"

// An equation laid out from its MathML tree with TeX's layout rules in simplified form (script sizes and
// shifts, fraction and radical clearances, n-ary operators, tables), without the math font's MATH table: the
// box it takes in a line (close enough for line heights and wrapping), and what to draw in it (tokens, fraction
// bars, radical signs), so a drawing of the equation (PDF export) fills exactly the box layout reserved.

export interface MathBox {
	readonly width: number
	// Above and below the baseline.
	readonly ascent: number
	readonly descent: number
}

// Width of `text` set at `size` px.
export type MathMeasure = (text: string, size: number) => number

// What to draw, in px from the box's left edge on its baseline, y down. Text runs have the width measured for
// them; rules are filled rectangles; strokes are open polylines (x, y pairs) of a line `width` thick.
export type MathItem =
	| {
			readonly kind: "text"
			readonly text: string
			readonly x: number
			readonly y: number
			readonly size: number
			readonly width: number
			readonly italic: boolean
	  }
	| { readonly kind: "rule"; readonly x: number; readonly y: number; readonly width: number; readonly height: number }
	| { readonly kind: "stroke"; readonly points: readonly number[]; readonly width: number }

export interface MathLayout {
	readonly box: MathBox
	readonly items: readonly MathItem[]
}

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
const NOTHING: MathLayout = { box: EMPTY, items: [] }
// Fraction bars and radical strokes, shares of the em (with a floor of one CSS pixel).
const RULE = 0.06

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

// Items of `laid` moved by (dx, dy).
function moved(items: readonly MathItem[], dx: number, dy: number, out: MathItem[]): void {
	for (const item of items) {
		if (item.kind === "stroke") {
			out.push({ ...item, points: item.points.map((value, k) => value + (k % 2 === 0 ? dx : dy)) })
		} else {
			out.push({ ...item, x: item.x + dx, y: item.y + dy })
		}
	}
}

function rule(style: Style): number {
	return Math.max(1, RULE * style.size)
}

// Children side by side.
function row(nodes: readonly MathNode[], style: Style, measure: MathMeasure, depth: number): MathLayout {
	let width = 0
	let ascent = 0
	let descent = 0
	const items: MathItem[] = []

	for (const node of nodes) {
		const own = typeof node === "string" ? token(node, style, measure, false) : laid(node, style, measure, depth + 1)

		moved(own.items, width, 0, items)
		width += own.box.width
		ascent = Math.max(ascent, own.box.ascent)
		descent = Math.max(descent, own.box.descent)
	}

	return { box: { width, ascent, descent }, items }
}

function token(value: string, style: Style, measure: MathMeasure, italic: boolean): MathLayout {
	const width = measure(value, style.size)

	return {
		box: { width, ascent: ASCENT * style.size, descent: DESCENT * style.size },
		items: value.length === 0 ? [] : [{ kind: "text", text: value, x: 0, y: 0, size: style.size, width, italic }]
	}
}

// mi: a single letter is italic unless it says otherwise (MathML Core §3.2.1).
function identifier(node: MathElement, style: Style, measure: MathMeasure): MathLayout {
	const value = text(node)

	return token(value, style, measure, node.tag === "mi" && [...value].length === 1 && node.attributes?.mathvariant !== "normal")
}

function operator(node: MathElement, style: Style, measure: MathMeasure): MathLayout {
	const value = text(node)
	const attributes = node.attributes
	const em = style.size
	const large = attributes?.largeop === true && style.display
	const size = large ? em * DISPLAY_OPERATOR : em
	const own = token(value, { ...style, size }, measure, false)
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
	const items: MathItem[] = []

	moved(own.items, left * em, 0, items)

	return { box: { width: own.box.width + (left + right) * em, ascent: own.box.ascent, descent: own.box.descent }, items }
}

// Base with scripts after it (subscript, superscript, or both).
function scripts(base: MathLayout, sub: MathLayout | undefined, sup: MathLayout | undefined, style: Style): MathLayout {
	const em = style.size
	let ascent = base.box.ascent
	let descent = base.box.descent
	const items: MathItem[] = [...base.items]

	if (sup !== undefined) {
		ascent = Math.max(ascent, SUPERSCRIPT_SHIFT * em + sup.box.ascent + sup.box.descent * 0.5)
		moved(sup.items, base.box.width, -(SUPERSCRIPT_SHIFT * em + sup.box.descent * 0.5), items)
	}

	if (sub !== undefined) {
		descent = Math.max(descent, SUBSCRIPT_SHIFT * em + sub.box.descent + sub.box.ascent * 0.5)
		moved(sub.items, base.box.width, SUBSCRIPT_SHIFT * em + sub.box.ascent * 0.5, items)
	}

	return {
		box: {
			width:
				base.box.width +
				Math.max(sub?.box.width ?? 0, sup?.box.width ?? 0) +
				(sub !== undefined || sup !== undefined ? SCRIPT_SPACE * em : 0),
			ascent,
			descent
		},
		items
	}
}

// Base with limits above and below it, each centred on it.
function limits(base: MathLayout, under: MathLayout | undefined, over: MathLayout | undefined, accent: boolean, style: Style): MathLayout {
	const em = style.size
	const gap = LIMIT_GAP * em
	const width = Math.max(base.box.width, under?.box.width ?? 0, over?.box.width ?? 0)
	const items: MathItem[] = []

	moved(base.items, (width - base.box.width) / 2, 0, items)

	if (over !== undefined) {
		const y = accent ? -(base.box.ascent + ACCENT_HEIGHT * em - over.box.ascent) : -(base.box.ascent + gap + over.box.descent)

		moved(over.items, (width - over.box.width) / 2, y, items)
	}

	if (under !== undefined) moved(under.items, (width - under.box.width) / 2, base.box.descent + gap + under.box.ascent, items)

	return {
		box: {
			width,
			ascent: base.box.ascent + (over === undefined ? 0 : accent ? ACCENT_HEIGHT * em : over.box.ascent + over.box.descent + gap),
			descent: base.box.descent + (under === undefined ? 0 : under.box.ascent + under.box.descent + gap)
		},
		items
	}
}

function table(node: MathElement, style: Style, measure: MathMeasure, depth: number): MathLayout {
	const widths: number[] = []
	const rows: { cells: MathLayout[]; ascent: number; descent: number }[] = []
	let height = 0

	for (const child of node.children) {
		const tr = element(child)

		if (tr === undefined) continue

		let ascent = 0
		let descent = 0
		const cells: MathLayout[] = []

		tr.children.forEach((cell, column) => {
			const own = typeof cell === "string" ? token(cell, style, measure, false) : laid(cell, style, measure, depth + 1)

			cells.push(own)
			widths[column] = Math.max(widths[column] ?? 0, own.box.width)
			ascent = Math.max(ascent, own.box.ascent)
			descent = Math.max(descent, own.box.descent)
		})

		rows.push({ cells, ascent, descent })
		height += ascent + descent
	}

	const em = style.size

	height += Math.max(0, rows.length - 1) * ROW_GAP * em

	const axis = AXIS * em
	const items: MathItem[] = []
	let top = -(height / 2 + axis)

	for (const tr of rows) {
		let x = 0

		tr.cells.forEach((cell, column) => {
			const width = widths[column] ?? 0

			moved(cell.items, x + (width - cell.box.width) / 2, top + tr.ascent, items)
			x += width + COLUMN_GAP * em
		})

		top += tr.ascent + tr.descent + ROW_GAP * em
	}

	return {
		box: {
			width: widths.reduce((sum, width) => sum + width, 0) + Math.max(0, widths.length - 1) * COLUMN_GAP * em,
			ascent: height / 2 + axis,
			descent: Math.max(0, height / 2 - axis)
		},
		items
	}
}

// The radical sign from (x, the inner box's bottom) to its overbar at `top`, the overbar over `width`.
function radical(x: number, sign: number, width: number, top: number, bottom: number, style: Style): MathItem {
	const middle = (top + bottom) / 2

	return {
		kind: "stroke",
		points: [
			x + sign * 0.08,
			middle + style.size * 0.05,
			x + sign * 0.28,
			middle - style.size * 0.02,
			x + sign * 0.55,
			bottom,
			x + sign * 0.92,
			top,
			x + sign + width,
			top
		],
		width: rule(style)
	}
}

function laid(node: MathElement, style: Style, measure: MathMeasure, depth: number): MathLayout {
	if (depth > MAX_DEPTH) return token(text(node), style, measure, false)

	const attributes = node.attributes
	const em = style.size
	const children = node.children
	const at = (k: number): MathLayout => {
		const child = children[k]

		return child === undefined
			? NOTHING
			: typeof child === "string"
				? token(child, style, measure, false)
				: laid(child, style, measure, depth + 1)
	}
	const scriptAt = (k: number, levels = 1): MathLayout | undefined => {
		const child = children[k]

		return child === undefined
			? undefined
			: typeof child === "string"
				? token(child, scripted(style, levels), measure, false)
				: laid(child, scripted(style, levels), measure, depth + 1)
	}

	switch (node.tag) {
		case "mi":
		case "mn":
		case "mtext":
			return identifier(node, style, measure)
		case "mo":
			return operator(node, style, measure)
		case "mspace":
			return {
				box: {
					width: (attributes?.width ?? 0) * em,
					ascent: (attributes?.height ?? 0) * em,
					descent: (attributes?.depth ?? 0) * em
				},
				items: []
			}
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
					? NOTHING
					: typeof numerator === "string"
						? token(numerator, inner, measure, false)
						: laid(numerator, inner, measure, depth + 1)
			const bottom =
				denominator === undefined
					? NOTHING
					: typeof denominator === "string"
						? token(denominator, inner, measure, false)
						: laid(denominator, inner, measure, depth + 1)
			const gap = FRACTION_GAP * em
			const axis = AXIS * em
			const width = Math.max(top.box.width, bottom.box.width) + FRACTION_PADDING * em
			const items: MathItem[] = []

			moved(top.items, (width - top.box.width) / 2, -(axis + gap + top.box.descent), items)
			moved(bottom.items, (width - bottom.box.width) / 2, gap + bottom.box.ascent - axis, items)

			if (attributes?.linethickness !== 0) {
				const thickness = rule(style)

				items.push({
					kind: "rule",
					x: (FRACTION_PADDING * em) / 4,
					y: -axis - thickness / 2,
					width: width - (FRACTION_PADDING * em) / 2,
					height: thickness
				})
			}

			return {
				box: {
					width,
					ascent: axis + gap + top.box.descent + top.box.ascent,
					descent: Math.max(0, gap + bottom.box.ascent + bottom.box.descent - axis)
				},
				items
			}
		}
		case "msqrt": {
			const inner = row(children, style, measure, depth)
			const sign = RADICAL_WIDTH * em
			const items: MathItem[] = [
				radical(0, sign, inner.box.width, -(inner.box.ascent + RADICAL_GAP * em * 0.5), inner.box.descent, style)
			]

			moved(inner.items, sign, 0, items)

			return {
				box: {
					width: inner.box.width + sign,
					ascent: inner.box.ascent + RADICAL_GAP * em,
					descent: inner.box.descent + LIMIT_GAP * em
				},
				items
			}
		}
		case "mroot": {
			const inner = at(0)
			const index = scriptAt(1, 2)
			const sign = RADICAL_WIDTH * em
			const shift = Math.max(0, (index?.box.width ?? 0) - sign * 0.5)
			const items: MathItem[] = [
				radical(shift, sign, inner.box.width, -(inner.box.ascent + RADICAL_GAP * em * 0.5), inner.box.descent, style)
			]

			moved(inner.items, shift + sign, 0, items)

			if (index !== undefined) moved(index.items, 0, -inner.box.ascent * 0.5, items)

			return {
				box: {
					width: inner.box.width + sign + shift,
					ascent: inner.box.ascent + RADICAL_GAP * em,
					descent: inner.box.descent + LIMIT_GAP * em
				},
				items
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
			const items: MathItem[] = [...base.items]
			let x = base.box.width

			for (let k = 1; k < children.length; k++) {
				const child = element(children[k])

				if (child?.tag === "mprescripts") continue

				const own = scriptAt(k) ?? NOTHING

				// Odd positions are subscripts (they set the column), even ones the superscripts over them.
				moved(own.items, x - (k % 2 === 0 ? own.box.width : 0), k % 2 === 1 ? SUBSCRIPT_SHIFT * em : -SUPERSCRIPT_SHIFT * em, items)

				if (k % 2 === 1) x += own.box.width

				width += k % 2 === 1 ? own.box.width : 0
				ascent = Math.max(ascent, SUPERSCRIPT_SHIFT * em + own.box.ascent)
				descent = Math.max(descent, SUBSCRIPT_SHIFT * em + own.box.descent)
			}

			return {
				box: {
					width: base.box.width + width + SCRIPT_SPACE * em,
					ascent: Math.max(base.box.ascent, ascent),
					descent: Math.max(base.box.descent, descent)
				},
				items
			}
		}
		case "mtable":
			return table(node, style, measure, depth)
		case "mphantom":
			// Takes its room, shows nothing.
			return { box: row(children, style, measure, depth).box, items: [] }
		case "mpadded":
		case "mrow":
		case "mtr":
		case "mtd":
		case "mprescripts":
			return row(children, style, measure, depth)
	}
}

// An equation set at `size` px; `display` for equations on a line of their own (m:oMathPara).
export function mathLayout(tree: MathElement, size: number, display: boolean, measure: MathMeasure): MathLayout {
	return laid(tree, { size, level: 0, display }, measure, 0)
}

// The box of an equation (mathLayout's).
export function mathBox(tree: MathElement, size: number, display: boolean, measure: MathMeasure): MathBox {
	return mathLayout(tree, size, display, measure).box
}
