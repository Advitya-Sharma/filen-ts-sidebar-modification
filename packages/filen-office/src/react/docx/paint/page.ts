import type { DecorationValue } from "../../../core/docx/layout/index"
import {
	DECORATION_FLOATS,
	DECORATION_INTS,
	DecorationKind,
	FRAGMENT_FLOATS,
	FRAGMENT_INTS,
	FRAGMENT_KIND_MASK,
	FRAGMENT_LEVEL_SHIFT,
	FragmentFlag,
	FragmentKind,
	LINE_FLOATS,
	LINE_INTS,
	LineFlag,
	TabLeader
} from "../../../core/text/index"
import { LINK_INTS, type DocxPageData, type DocxTextStyle } from "../../../dom/docx/protocol"
import type { DocxTables } from "../store"

// Paints one page's display list as DOM: real text in absolutely positioned spans (so selection, copy,
// find and screen readers work), boxes for fills, borders and decorations, and empty containers for
// objects, filled later from the session. Only numbers the layout produced and values checked here reach
// styles; document text reaches the DOM only as text nodes.

// The page's layers (Layer in the layout's output), bottom to top.
const LAYER_BEHIND = 0
const LAYER_TEXT = 1
const LAYER_FRONT = 2

const GRAPHIC_FLOATS = 4
const GRAPHIC_INTS = 3
const OBJECT_FLOATS = 4
const OBJECT_INTS = 3

const Graphic = { SHADING: 1, BORDER: 2, COLUMN_SEPARATOR: 3, NOTE_SEPARATOR: 4, REVISION_BAR: 5 } as const

const HEX = /^[0-9A-Fa-f]{6}$/
// cssFont() output: optional italic, weight, size, then an internal alias and a generic family, or a
// generic family alone.
const CSS_FONT =
	/^(?:italic )?\d{1,4} \d{1,5}(?:\.\d{1,3})? ?px (?:"[A-Za-z0-9_-]{1,128}", )?(?:serif|sans-serif|monospace|cursive|fantasy|system-ui)$/
const FALLBACK_FONT = "400 16px serif"
const MAX_PX = 1e6

export interface PaintedText {
	readonly element: HTMLElement
	readonly source: number
	readonly start: number
	readonly end: number
}

export interface PaintedObject {
	readonly element: HTMLElement
	readonly id: number
	readonly kind: number
	readonly width: number
	readonly height: number
}

export interface PaintedPage {
	readonly root: HTMLElement
	// Text spans in reading order (for find highlights and selection mapping).
	readonly texts: readonly PaintedText[]
	readonly objects: readonly PaintedObject[]
}

function px(value: number | undefined): number {
	const v = value ?? 0

	return Number.isFinite(v) ? Math.max(-MAX_PX, Math.min(MAX_PX, Math.round(v * 100) / 100)) : 0
}

function hexColor(color: string | undefined): string {
	return color !== undefined && HEX.test(color) ? `#${color}` : "var(--fo-auto-text, #000)"
}

function cssFont(font: string): string {
	return CSS_FONT.test(font) ? font : FALLBACK_FONT
}

interface StyleCss {
	// Declarations every span of the style starts with.
	readonly base: string
	readonly smallCapsFont: string
	readonly ascent: number
	readonly shift: number
	readonly color: string
}

const styleCss = new WeakMap<DocxTextStyle, StyleCss>()

function cssOf(style: DocxTextStyle | undefined): StyleCss {
	if (style === undefined) {
		return {
			base: `font:${FALLBACK_FONT};line-height:20px;`,
			smallCapsFont: FALLBACK_FONT,
			ascent: 16,
			shift: 0,
			color: hexColor(undefined)
		}
	}

	let css = styleCss.get(style)

	if (css === undefined) {
		const ascent = Math.max(0, px(style.ascent))
		const height = Math.max(0, ascent + Math.max(0, px(style.descent)))
		const color = style.color.length === 0 ? hexColor(undefined) : hexColor(style.color)
		let base = `font:${cssFont(style.font)};line-height:${height}px;color:${color};`

		if (!style.kerning) base += "font-kerning:none;"

		css = { base, smallCapsFont: cssFont(style.smallCapsFont), ascent, shift: px(style.shift), color }
		styleCss.set(style, css)
	}

	return css
}

const LEADER_STYLE: Readonly<Record<number, string>> = {
	[TabLeader.DOT]: "dotted",
	[TabLeader.MIDDLE_DOT]: "dotted",
	[TabLeader.HYPHEN]: "dashed",
	[TabLeader.UNDERSCORE]: "solid",
	[TabLeader.HEAVY]: "solid"
}

// Underline styles (ST_Underline) drawn as a box's top border; anything else is a solid bar.
function underlineBorder(style: string): string | undefined {
	if (style === "double" || style === "wavyDouble") return "double"
	if (style.startsWith("dotted") || style === "dotDotDash" || style === "dashDotDotHeavy") return "dotted"
	if (style.startsWith("dash") || style.startsWith("dotDash")) return "dashed"
	if (style.startsWith("wav")) return "solid"

	return undefined
}

function borderStyle(style: string): string {
	switch (style) {
		case "dotted":
			return "dotted"
		case "dashed":
		case "dashSmallGap":
		case "dotDash":
		case "dotDotDash":
			return "dashed"
		case "double":
		case "triple":
		case "thinThickSmallGap":
		case "thickThinSmallGap":
		case "thinThickMediumGap":
		case "thickThinMediumGap":
		case "thinThickLargeGap":
		case "thickThinLargeGap":
			return "double"
		case "none":
		case "nil":
			return "none"
		default:
			return "solid"
	}
}

export interface PagePaintContext {
	readonly document: Document
	readonly tables: DocxTables
}

function box(document: Document, className: string, x: number, y: number, width: number, height: number, extra: string): HTMLElement {
	const element = document.createElement("div")

	element.className = className
	element.style.cssText = `left:${px(x)}px;top:${px(y)}px;width:${px(width)}px;height:${px(height)}px;${extra}`

	return element
}

function decorationColor(value: DecorationValue | undefined): string | undefined {
	if (value === undefined) return undefined

	switch (value.kind) {
		case "color":
			return HEX.test(value.color) ? `#${value.color}` : undefined
		case "underline":
		case "border":
			return value.color !== undefined && HEX.test(value.color) ? `#${value.color}` : undefined
		case "revision":
			return "var(--fo-revision-mark, #000)"
	}
}

// Paints a graphic (page fill, border line, rule) into its layer.
function paintGraphic(
	document: Document,
	layers: readonly HTMLElement[],
	page: DocxPageData,
	decorations: readonly DecorationValue[],
	index: number
): void {
	const f = index * GRAPHIC_FLOATS
	const i = index * GRAPHIC_INTS
	const x = page.graphicFloats[f] ?? 0
	const y = page.graphicFloats[f + 1] ?? 0
	const width = page.graphicFloats[f + 2] ?? 0
	const height = page.graphicFloats[f + 3] ?? 0
	const kind = page.graphicInts[i] ?? 0
	const value = decorations[(page.graphicInts[i + 1] ?? 0) - 1]
	const layer = layers[page.graphicInts[i + 2] ?? LAYER_TEXT] ?? layers[LAYER_TEXT]

	if (layer === undefined || !(width > 0 && height > 0)) return

	let extra: string

	switch (kind) {
		case Graphic.SHADING: {
			const color = decorationColor(value)

			if (color === undefined) return

			extra = `background:${color};`

			break
		}
		case Graphic.BORDER: {
			const color = decorationColor(value) ?? "var(--fo-auto-text, #000)"
			const style = value?.kind === "border" ? borderStyle(value.style) : "solid"

			if (style === "none") return

			// The box is the line itself: a thin side carries the stroke style.
			extra =
				style === "solid"
					? `background:${color};`
					: width >= height
						? `border-top:${px(height)}px ${style} ${color};height:0;`
						: `border-left:${px(width)}px ${style} ${color};width:0;`

			break
		}
		case Graphic.REVISION_BAR:
			extra = "background:var(--fo-revision-bar, #000);"

			break
		default:
			extra = "background:var(--fo-auto-text, #000);"
	}

	layer.appendChild(box(document, "fo-docx-box", x, y, width, height, extra))
}

// The style colour of the fragment under x on a line (for decorations drawn in the text's colour).
function colorAt(page: DocxPageData, styles: readonly DocxTextStyle[], line: number, x: number): string {
	const first = page.lineInts[line * LINE_INTS] ?? 0
	const count = page.lineInts[line * LINE_INTS + 1] ?? 0
	let style = -1

	for (let k = first; k < first + count; k++) {
		const fx = page.fragmentFloats[k * FRAGMENT_FLOATS] ?? 0

		if (fx > x + 0.5) break

		style = page.fragmentInts[k * FRAGMENT_INTS + 2] ?? -1
	}

	return cssOf(styles[style]).color
}

function paintDecoration(
	document: Document,
	layer: HTMLElement,
	page: DocxPageData,
	tables: DocxTables,
	index: number,
	under: boolean
): void {
	const f = index * DECORATION_FLOATS
	const i = index * DECORATION_INTS
	const kind = page.decorationInts[i] ?? 0
	const isFill = kind === DecorationKind.SHADING || kind === DecorationKind.HIGHLIGHT

	// Fills go under the text, lines over it.
	if (isFill !== under) return

	const x = page.decorationFloats[f] ?? 0
	const y = page.decorationFloats[f + 1] ?? 0
	let width = page.decorationFloats[f + 2] ?? 0
	const height = page.decorationFloats[f + 3] ?? 0
	const key = page.decorationInts[i + 1] ?? 0
	const line = page.decorationInts[i + 2] ?? 0
	const value = tables.decorations[key - 1]
	let extra: string

	switch (kind) {
		case DecorationKind.SHADING:
		case DecorationKind.HIGHLIGHT: {
			const color = decorationColor(value)

			if (color === undefined) return

			extra = `background:${color};`

			break
		}
		case DecorationKind.BORDER: {
			if (value?.kind !== "border") return

			const style = borderStyle(value.style)

			if (style === "none") return

			extra = `border:${px(Math.max(1, value.width))}px ${style} ${decorationColor(value) ?? colorAt(page, tables.styles, line, x)};box-sizing:border-box;`

			break
		}
		case DecorationKind.UNDERLINE: {
			const color = decorationColor(value) ?? colorAt(page, tables.styles, line, x)
			const style = value?.kind === "underline" ? underlineBorder(value.style) : undefined

			extra =
				style === undefined
					? `background:${color};`
					: style === "double"
						? `border-top:${px(Math.max(1, height))}px solid ${color};border-bottom:${px(Math.max(1, height))}px solid ${color};height:${px(height)}px;box-sizing:content-box;`
						: `border-top:${px(Math.max(1, height))}px ${style} ${color};height:0;`

			break
		}
		case DecorationKind.STRIKE:
			extra = `background:${colorAt(page, tables.styles, line, x)};`

			break
		case DecorationKind.DOUBLE_STRIKE: {
			const color = colorAt(page, tables.styles, line, x)

			extra = `border-top:${px(height)}px solid ${color};border-bottom:${px(height)}px solid ${color};height:${px(height)}px;box-sizing:content-box;`

			break
		}
		case DecorationKind.BAR:
			width = Math.max(width, 1)
			extra = "background:var(--fo-auto-text, #000);"

			break
		default:
			return
	}

	layer.appendChild(box(document, "fo-docx-box", x, y, width, height, extra))
}

function isLowerCase(char: string): boolean {
	return char !== char.toUpperCase()
}

// Small caps: lower-case letters as capitals in the smaller face, the rest as is.
function appendSmallCaps(document: Document, span: HTMLElement, text: string, smallFont: string): void {
	let at = 0

	while (at < text.length) {
		const lower = isLowerCase(text.charAt(at))
		let end = at + 1

		while (end < text.length && isLowerCase(text.charAt(end)) === lower) end++

		const piece = text.slice(at, end)

		if (lower) {
			const small = document.createElement("span")

			small.style.font = smallFont
			small.style.textTransform = "uppercase"
			small.textContent = piece
			span.appendChild(small)
		} else {
			span.appendChild(document.createTextNode(piece))
		}

		at = end
	}
}

// Link ranges of the page by source, for marking fragments.
function linksBySource(page: DocxPageData): Map<number, number[]> {
	const out = new Map<number, number[]>()

	for (let k = 0; k < page.linkCount; k++) {
		const source = page.linkInts[k * LINK_INTS] ?? -1
		const list = out.get(source)

		if (list === undefined) out.set(source, [k])
		else list.push(k)
	}

	return out
}

// The fragment's text split at link boundaries: [start, end, link index or -1] pieces in order.
function linkPieces(page: DocxPageData, ranges: readonly number[] | undefined, start: number, end: number): [number, number, number][] {
	if (ranges === undefined) return [[start, end, -1]]

	const pieces: [number, number, number][] = []
	let at = start

	for (const k of ranges) {
		const from = Math.max(start, page.linkInts[k * LINK_INTS + 1] ?? 0)
		const to = Math.min(end, page.linkInts[k * LINK_INTS + 2] ?? 0)

		if (to <= from || from < at) continue
		if (from > at) pieces.push([at, from, -1])

		pieces.push([from, to, page.linkInts[k * LINK_INTS + 3] ?? -1])
		at = to
	}

	if (at < end) pieces.push([at, end, -1])

	return pieces
}

function markLink(element: HTMLElement, link: number, tables: DocxTables): void {
	const target = tables.links[link]

	if (target === undefined) return

	element.dataset["foLink"] = String(link)
	element.setAttribute("role", "link")
	element.tabIndex = 0

	if (target.tooltip !== undefined) element.title = target.tooltip
}

const SPACES = /^\s+$/

// What copying puts after a line: a newline at a paragraph end or break, nothing after a hyphenated line,
// else the spaces between this line's last character and the next line's first (a space at least).
function lineSeparator(page: DocxPageData, line: number, flags: number, text: string, end: number): string {
	if ((flags & (LineFlag.LAST | LineFlag.LINE_BREAK | LineFlag.PAGE_BREAK | LineFlag.COLUMN_BREAK)) !== 0) return "\n"
	if ((flags & LineFlag.HYPHENATED) !== 0) return ""

	const source = page.lineInts[line * LINE_INTS + 5]
	let next = page.lineInts[line * LINE_INTS + 3] ?? end

	for (let other = line + 1; other < page.lineCount; other++) {
		if (page.lineInts[other * LINE_INTS + 5] === source) {
			next = page.lineInts[other * LINE_INTS + 2] ?? next

			break
		}
	}

	const gap = text.slice(end, next)

	return SPACES.test(gap) ? gap : " "
}

export function paintPage(page: DocxPageData, context: PagePaintContext): PaintedPage {
	const { document, tables } = context
	const root = document.createElement("div")

	root.className = "fo-docx-page-content"
	root.style.cssText = `width:${px(page.width)}px;height:${px(page.height)}px;`

	const layers = [LAYER_BEHIND, LAYER_TEXT, LAYER_FRONT].map(layer => {
		const element = document.createElement("div")

		element.className = "fo-docx-layer"
		element.dataset["foLayer"] = String(layer)

		return element
	})
	const behind = layers[LAYER_BEHIND] ?? root
	const textLayer = layers[LAYER_TEXT] ?? root
	const front = layers[LAYER_FRONT] ?? root
	const fills = document.createElement("div")
	const textElements = document.createElement("div")
	const lines = document.createElement("div")
	const objectsLayer = document.createElement("div")

	fills.className = "fo-docx-layer"
	textElements.className = "fo-docx-layer fo-docx-text"
	lines.className = "fo-docx-layer"
	objectsLayer.className = "fo-docx-layer"

	for (let g = 0; g < page.graphicCount; g++) paintGraphic(document, [behind, fills, front], page, tables.decorations, g)
	for (let d = 0; d < page.decorationCount; d++) paintDecoration(document, fills, page, tables, d, true)
	for (let d = 0; d < page.decorationCount; d++) paintDecoration(document, lines, page, tables, d, false)

	const texts: PaintedText[] = []
	const links = linksBySource(page)

	for (let line = 0; line < page.lineCount; line++) {
		const lf = line * LINE_FLOATS
		const li = line * LINE_INTS
		const baseline = page.lineFloats[lf + 2] ?? 0
		const firstFragment = page.lineInts[li] ?? 0
		const fragmentCount = page.lineInts[li + 1] ?? 0
		const flags = page.lineInts[li + 4] ?? 0
		const source = page.lineInts[li + 5] ?? -1
		const text = page.sourceTexts[source] ?? ""
		const sourceLinks = links.get(source)
		// The line's spans in logical (text) order: right-to-left runs are laid out visually, but copying and
		// screen readers follow the DOM.
		const lineTexts: PaintedText[] = []

		for (let k = firstFragment; k < firstFragment + fragmentCount; k++) {
			const ff = k * FRAGMENT_FLOATS
			const fi = k * FRAGMENT_INTS
			const start = page.fragmentInts[fi] ?? 0
			const end = page.fragmentInts[fi + 1] ?? 0
			const info = page.fragmentInts[fi + 3] ?? 0
			const kind = info & FRAGMENT_KIND_MASK

			if (kind === FragmentKind.OBJECT || end <= start) continue

			const x = page.fragmentFloats[ff] ?? 0
			const width = page.fragmentFloats[ff + 1] ?? 0
			const style = tables.styles[page.fragmentInts[fi + 2] ?? -1]
			const css = cssOf(style)
			const top = baseline - css.shift - css.ascent
			const span = document.createElement("span")

			if (kind === FragmentKind.TAB) {
				const leader = LEADER_STYLE[page.fragmentInts[fi + 4] ?? 0]

				span.className = "fo-docx-tab"
				span.style.cssText = `${css.base}left:${px(x)}px;top:${px(top)}px;width:${px(width)}px;`

				if (leader !== undefined) {
					const rule = document.createElement("span")

					rule.className = "fo-docx-leader"
					rule.style.cssText = `top:${px(css.ascent - 1)}px;border-bottom:1px ${leader} currentColor;`
					span.appendChild(rule)
				}

				span.appendChild(document.createTextNode("\t"))
			} else {
				let extra = `left:${px(x)}px;top:${px(top)}px;`
				const wordSpacing = page.fragmentFloats[ff + 2] ?? 0
				const letterSpacing = (style?.letterSpacing ?? 0) + (page.fragmentFloats[ff + 3] ?? 0)
				const scaleX = style?.scaleX ?? 1

				if (wordSpacing !== 0) extra += `word-spacing:${px(wordSpacing)}px;`
				if (letterSpacing !== 0) extra += `letter-spacing:${px(letterSpacing)}px;`
				if (scaleX !== 1 && Number.isFinite(scaleX) && scaleX > 0) extra += `transform:scaleX(${Math.min(scaleX, 10)});`
				if ((info >> FRAGMENT_LEVEL_SHIFT) % 2 === 1) extra += "direction:rtl;unicode-bidi:bidi-override;"
				if ((info & FragmentFlag.CAPS) !== 0) extra += "text-transform:uppercase;"

				span.className = "fo-docx-run"
				span.style.cssText = css.base + extra

				const pieces = linkPieces(page, sourceLinks, start, end)
				const smallCaps = (info & FragmentFlag.SMALL_CAPS) !== 0

				// A link covering the whole fragment marks the span itself; a partial one becomes an inline child,
				// so the browser places it exactly where its text runs.
				for (const [from, to, link] of pieces) {
					const content = text.slice(from, to)
					const target = link >= 0 && pieces.length > 1 ? document.createElement("span") : span

					if (smallCaps) appendSmallCaps(document, target, content, css.smallCapsFont)
					else target.appendChild(document.createTextNode(content))

					if (link >= 0) markLink(target, link, tables)
					if (target !== span) span.appendChild(target)
				}

				if ((info & FragmentFlag.HYPHEN) !== 0) span.dataset["foHyphen"] = ""
			}

			lineTexts.push({ element: span, source, start, end })
		}

		if (lineTexts.length === 0) continue

		lineTexts.sort((a, b) => a.start - b.start)

		for (const painted of lineTexts) {
			textElements.appendChild(painted.element)
			texts.push(painted)
		}

		const separator = lineSeparator(page, line, flags, text, lineTexts[lineTexts.length - 1]?.end ?? 0)

		// Copying a selection keeps paragraph breaks, line breaks and the space a soft wrap swallowed.
		if (separator.length > 0) {
			const element = document.createElement("span")

			element.className = "fo-docx-eol"
			element.textContent = separator
			textElements.appendChild(element)
		}
	}

	const objects: PaintedObject[] = []

	for (let o = 0; o < page.objectCount; o++) {
		const f = o * OBJECT_FLOATS
		const i = o * OBJECT_INTS
		const width = page.objectFloats[f + 2] ?? 0
		const height = page.objectFloats[f + 3] ?? 0
		const element = box(document, "fo-docx-object", page.objectFloats[f] ?? 0, page.objectFloats[f + 1] ?? 0, width, height, "")
		const layer = page.objectInts[i + 1] ?? LAYER_TEXT
		const id = page.objectInts[i] ?? -1

		element.dataset["foObject"] = String(id)
		;(layer === LAYER_BEHIND ? behind : layer === LAYER_FRONT ? front : objectsLayer).appendChild(element)
		objects.push({ element, id, kind: page.objectKinds[o] ?? 0, width, height })
	}

	textLayer.append(fills, textElements, lines, objectsLayer)
	root.append(behind, textLayer, front)

	return { root, texts, objects }
}
