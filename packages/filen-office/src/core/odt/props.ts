import { parseLength, parsePercent, type OdfElement, type OdfStyleSheet } from "../odf/index"
import type { OdtLossyReason } from "./types"
import { eighths, esc, twips } from "./wml"

// ODF text and paragraph properties (Part 3 §20) as WordprocessingML rPr and pPr content, in schema order.

export type Attributes = ReadonlyMap<string, string>

export interface PropsSink {
	lossy(reason: OdtLossyReason): void
}

const COLOR = /^#[0-9a-fA-F]{6}$/

function color(value: string | undefined): string | undefined {
	return value !== undefined && COLOR.test(value) ? value.slice(1).toUpperCase() : undefined
}

function unquote(value: string): string {
	const trimmed = value.trim()
	const first = trimmed[0]

	return trimmed.length >= 2 && (first === "'" || first === '"') && trimmed[trimmed.length - 1] === first ? trimmed.slice(1, -1) : trimmed
}

// The family a style:font-name* or fo:font-family* attribute names.
export function fontFamily(sheet: OdfStyleSheet, attributes: Attributes, kind: "" | "-asian" | "-complex"): string | undefined {
	const name = attributes.get(`style:font-name${kind}`)

	if (name !== undefined) return sheet.fontFaces.get(name)?.family ?? name

	const family = attributes.get(kind === "" ? "fo:font-family" : `style:font-family${kind}`)

	if (family === undefined) return undefined

	// A CSS list: the first family.
	return unquote(family.split(",")[0] ?? family)
}

function isBold(weight: string): boolean {
	return weight === "bold" || Number(weight) >= 600
}

function onOff(name: string, on: boolean): string {
	return on ? `<w:${name}/>` : `<w:${name} w:val="0"/>`
}

const UNDERLINES: Readonly<Record<string, string>> = {
	solid: "single",
	dotted: "dotted",
	dash: "dash",
	"long-dash": "dashLong",
	"dot-dash": "dotDash",
	"dot-dot-dash": "dotDotDash",
	wave: "wave"
}

function underline(attributes: Attributes): string | undefined {
	const style = attributes.get("style:text-underline-style")
	const type = attributes.get("style:text-underline-type")

	if (style === undefined && type === undefined) return undefined
	if (style === "none" || type === "none") return '<w:u w:val="none"/>'

	let value = UNDERLINES[style ?? "solid"] ?? "single"
	const bold = attributes.get("style:text-underline-width") === "bold" || attributes.get("style:text-underline-width") === "thick"

	if (type === "double") value = value === "wave" ? "wavyDouble" : "double"
	else if (bold) value = value === "single" ? "thick" : value === "wave" ? "wavyHeavy" : `${value}Heavy`
	if (attributes.get("style:text-underline-mode") === "skip-white-space" && value === "single") value = "words"

	const lineColor = attributes.get("style:text-underline-color")
	const hex = lineColor === "font-color" ? undefined : color(lineColor)

	return `<w:u w:val="${value}"${hex === undefined ? "" : ` w:color="${hex}"`}/>`
}

// style:text-position: "super"|"sub"|percent, then an optional size percent (§20.374).
function textPosition(value: string, fontSize: number | undefined): string {
	const [raise = "0%", size] = value.trim().split(/\s+/)
	const percent = raise === "super" ? 33 : raise === "sub" ? -33 : (parsePercent(raise) ?? 0)
	const scale = size === undefined ? (raise === "super" || raise === "sub" ? 58 : 100) : (parsePercent(size) ?? 100)

	if (percent === 0) return ""
	if (scale < 100 && (raise === "super" || raise === "sub" || Math.abs(percent) === 33)) {
		return `<w:vertAlign w:val="${percent > 0 ? "superscript" : "subscript"}"/>`
	}

	// A raise without shrinking: a position in half-points of the font size.
	return fontSize === undefined ? "" : `<w:position w:val="${Math.round((fontSize * percent) / 50)}"/>`
}

function language(attributes: Attributes, kind: "" | "-asian" | "-complex"): string | undefined {
	const lang = attributes.get(kind === "" ? "fo:language" : `style:language${kind}`)
	const country = attributes.get(kind === "" ? "fo:country" : `style:country${kind}`)

	if (lang === undefined || lang === "none" || lang === "zxx") return undefined

	return country === undefined || country === "none" ? lang : `${lang}-${country}`
}

// rPr content (without rStyle) for text properties. `fontSize`: the size in effect, for raised text.
export function runProperties(sheet: OdfStyleSheet, attributes: Attributes, sink: PropsSink, fontSize?: number): string {
	if (attributes.size === 0) return ""

	let out = ""
	const ascii = fontFamily(sheet, attributes, "")
	const eastAsia = fontFamily(sheet, attributes, "-asian")
	const complex = fontFamily(sheet, attributes, "-complex")

	if (ascii !== undefined || eastAsia !== undefined || complex !== undefined) {
		out += "<w:rFonts"
		if (ascii !== undefined) out += ` w:ascii="${esc(ascii)}" w:hAnsi="${esc(ascii)}"`
		if (eastAsia !== undefined) out += ` w:eastAsia="${esc(eastAsia)}"`
		if (complex !== undefined) out += ` w:cs="${esc(complex)}"`
		out += "/>"
	}

	const weight = attributes.get("fo:font-weight")
	const weightComplex = attributes.get("style:font-weight-complex")
	const style = attributes.get("fo:font-style")
	const styleComplex = attributes.get("style:font-style-complex")

	if (weight !== undefined) out += onOff("b", isBold(weight))
	if (weightComplex !== undefined) out += onOff("bCs", isBold(weightComplex))
	if (style !== undefined) out += onOff("i", style === "italic" || style === "oblique")
	if (styleComplex !== undefined) out += onOff("iCs", styleComplex === "italic" || styleComplex === "oblique")

	const transform = attributes.get("fo:text-transform")

	if (transform === "uppercase") out += "<w:caps/>"
	else if (transform === "none") out += '<w:caps w:val="0"/>'
	else if (transform === "lowercase" || transform === "capitalize") sink.lossy("textTransform")

	const variant = attributes.get("fo:font-variant")

	if (variant !== undefined) out += onOff("smallCaps", variant === "small-caps")

	const strike = attributes.get("style:text-line-through-style")
	const strikeType = attributes.get("style:text-line-through-type")

	if (strike !== undefined || strikeType !== undefined) {
		const on = strike !== "none" && strikeType !== "none"

		out += on && strikeType === "double" ? "<w:dstrike/>" : onOff("strike", on)
	}

	const outline = attributes.get("style:text-outline")
	const shadow = attributes.get("fo:text-shadow")
	const relief = attributes.get("style:font-relief")

	if (outline !== undefined) out += onOff("outline", outline === "true")
	if (shadow !== undefined) out += onOff("shadow", shadow !== "none")
	if (relief === "embossed") out += "<w:emboss/>"
	else if (relief === "engraved") out += "<w:imprint/>"

	if (attributes.get("text:display") === "none") out += "<w:vanish/>"

	const fontColor = color(attributes.get("fo:color"))

	if (attributes.get("style:use-window-font-color") === "true") out += '<w:color w:val="auto"/>'
	else if (fontColor !== undefined) out += `<w:color w:val="${fontColor}"/>`

	const spacing = attributes.get("fo:letter-spacing")

	if (spacing !== undefined) out += `<w:spacing w:val="${spacing === "normal" ? 0 : twips(parseLength(spacing) ?? 0)}"/>`

	const scale = parsePercent(attributes.get("style:text-scale"))

	if (scale !== undefined && scale !== 100) out += `<w:w w:val="${Math.max(1, Math.min(600, Math.round(scale)))}"/>`

	const size = parseLength(attributes.get("fo:font-size"))
	const position = attributes.get("style:text-position")
	const raised = position === undefined ? "" : textPosition(position, size ?? fontSize)

	if (raised.startsWith("<w:position")) out += raised
	if (size !== undefined) out += `<w:sz w:val="${Math.max(1, Math.round(size * 2))}"/>`

	const sizeComplex = parseLength(attributes.get("style:font-size-complex"))

	if (sizeComplex !== undefined) out += `<w:szCs w:val="${Math.max(1, Math.round(sizeComplex * 2))}"/>`

	const line = underline(attributes)

	if (line !== undefined) out += line

	const background = attributes.get("fo:background-color")

	if (background === "transparent") out += '<w:shd w:val="clear" w:color="auto" w:fill="auto"/>'
	else if (color(background) !== undefined) out += `<w:shd w:val="clear" w:color="auto" w:fill="${color(background)}"/>`
	if (raised.startsWith("<w:vertAlign")) out += raised

	const emphasis = attributes.get("style:text-emphasize")

	if (emphasis !== undefined && emphasis !== "none") out += `<w:em w:val="${emphasis.startsWith("circle") ? "circle" : "dot"}"/>`

	const lang = language(attributes, "")
	const langAsian = language(attributes, "-asian")
	const langComplex = language(attributes, "-complex")

	if (lang !== undefined || langAsian !== undefined || langComplex !== undefined) {
		out += "<w:lang"
		if (lang !== undefined) out += ` w:val="${esc(lang)}"`
		if (langAsian !== undefined) out += ` w:eastAsia="${esc(langAsian)}"`
		if (langComplex !== undefined) out += ` w:bidi="${esc(langComplex)}"`
		out += "/>"
	}

	if (
		(attributes.get("style:text-rotation-angle") ?? "0") !== "0" ||
		(attributes.get("style:text-combine") ?? "none") !== "none" ||
		(attributes.get("style:text-blinking") ?? "false") === "true"
	) {
		sink.lossy("textEffect")
	}

	return out
}

// A border value "width style color" (§20.183 fo:border, XSL) as a WordprocessingML border.
export function border(tag: string, value: string | undefined, padding: string | undefined, lineWidths?: string): string {
	if (value === undefined) return ""

	const parts = value.trim().split(/\s+/)
	let width: number | undefined
	let style = "solid"
	let hex = "000000"

	for (const part of parts) {
		const length = parseLength(part)

		if (length !== undefined) width = length
		else if (part.startsWith("#")) hex = color(part) ?? hex
		else if (part === "thin") width = 0.75
		else if (part === "medium") width = 2.25
		else if (part === "thick") width = 4.5
		else style = part
	}

	if (style === "none" || style === "hidden" || width === 0) return `<w:${tag} w:val="nil"/>`

	let value_ = "single"

	switch (style) {
		case "double":
			value_ = "double"
			break
		case "dotted":
			value_ = "dotted"
			break
		case "dashed":
			value_ = "dashed"
			break
		case "groove":
			value_ = "threeDEngrave"
			break
		case "ridge":
			value_ = "threeDEmboss"
			break
		case "inset":
			value_ = "inset"
			break
		case "outset":
			value_ = "outset"
			break
		case "dash-dot":
			value_ = "dotDash"
			break
		case "dash-dot-dot":
			value_ = "dotDotDash"
			break
		default:
			value_ = "single"
	}

	// A double line's total width is the inner, gap and outer widths of style:border-line-width.
	let total = width ?? 0.75

	if (value_ === "double" && lineWidths !== undefined) {
		const widths = lineWidths
			.trim()
			.split(/\s+/)
			.map(part => parseLength(part) ?? 0)

		if (widths.length === 3) total = widths.reduce((sum, item) => sum + item, 0)
	} else if (value_ === "double") {
		total = Math.max(total, 2.25)
	}

	const space = Math.max(0, Math.min(31, Math.round(parseLength(padding) ?? 0)))

	return `<w:${tag} w:val="${value_}" w:sz="${Math.max(2, Math.min(96, eighths(total)))}" w:space="${space}" w:color="${hex}"/>`
}

export interface ParagraphOptions {
	// TabsRelativeToIndent (settings.xml): tab positions count from the left margin.
	readonly tabsRelativeToIndent: boolean
	// The paragraph's left margin in effect, for relative tab stops.
	readonly marginLeft: number
	// Write the indents (and nothing of them otherwise): see setsIndent.
	readonly indents: boolean
	// Take the indents from the properties' own attributes, a side left out being 0 (LibreOffice reads an
	// automatic style's indent as one item: measured over UNO, a parent's margin-left does not show through an
	// automatic style that sets only margin-right). Otherwise from the resolved ones (common styles).
	readonly ownIndents?: boolean
	// The font size in effect, for leading added to lines.
	readonly fontSize: number
	// Indents of the list level the paragraph is in, written when its own do not win.
	readonly listIndent?: { readonly left: number; readonly firstLine: number }
}

const LEADERS: Readonly<Record<string, string>> = { ".": "dot", "-": "hyphen", _: "underscore", "·": "middleDot" }

function tabs(stops: OdfElement | undefined, options: ParagraphOptions): string {
	if (stops === undefined) return ""

	let out = ""

	for (const stop of stops.children) {
		if (stop.name !== "style:tab-stop") continue

		const type = stop.attributes.get("style:type") ?? "left"
		const alignment = type === "center" ? "center" : type === "right" ? "right" : type === "char" ? "decimal" : "left"
		const position = parseLength(stop.attributes.get("style:position")) ?? 0
		const leaderText = stop.attributes.get("style:leader-text")
		const leaderStyle = stop.attributes.get("style:leader-style")
		const leader =
			leaderText !== undefined && leaderText !== " "
				? LEADERS[leaderText]
				: leaderStyle === "dotted"
					? "dot"
					: leaderStyle === "dash"
						? "hyphen"
						: leaderStyle === "solid"
							? "underscore"
							: undefined
		const at = twips(position + (options.tabsRelativeToIndent ? options.marginLeft : 0))

		out += `<w:tab w:val="${alignment}"${leader === undefined ? "" : ` w:leader="${leader}"`} w:pos="${at}"/>`
	}

	return out === "" ? "" : `<w:tabs>${out}</w:tabs>`
}

// Whether the properties set an indent (LibreOffice's LR-space item): such a paragraph's indents win over a
// list level's (SwTextNode::AreListLevelIndentsApplicable).
export function setsIndent(attributes: Attributes): boolean {
	return attributes.has("fo:margin-left") || attributes.has("fo:margin-right") || attributes.has("fo:text-indent")
}

export interface ParagraphXml {
	readonly head: string
	readonly tail: string
}

// pPr content in two halves around numPr (head: keepNext ... widowControl; tail: pBdr ... outlineLvl).
// `own`: attributes to write; `resolved`: the same paragraph's attributes with inheritance, for values that
// combine (indents, tab positions).
export function paragraphProperties(
	own: Attributes,
	children: readonly OdfElement[],
	resolved: Attributes,
	options: ParagraphOptions,
	sink: PropsSink
): ParagraphXml {
	let head = ""
	const keepNext = own.get("fo:keep-with-next")
	const keepTogether = own.get("fo:keep-together")
	const breakBefore = own.get("fo:break-before")

	if (keepNext !== undefined) head += onOff("keepNext", keepNext === "always")
	if (keepTogether !== undefined) head += onOff("keepLines", keepTogether === "always")
	if (breakBefore !== undefined) head += onOff("pageBreakBefore", breakBefore === "page")

	const widows = own.get("fo:widows")
	const orphans = own.get("fo:orphans")

	if (widows !== undefined || orphans !== undefined) {
		head += onOff("widowControl", Number(resolved.get("fo:widows") ?? "0") >= 2 || Number(resolved.get("fo:orphans") ?? "0") >= 2)
	}

	let tail = ""
	const sides = ["top", "left", "bottom", "right"] as const

	if (sides.some(side => own.has(`fo:border-${side}`))) {
		const borders = sides
			.map(side =>
				border(
					side,
					resolved.get(`fo:border-${side}`),
					resolved.get(`fo:padding-${side}`),
					resolved.get(`style:border-line-width-${side}`)
				)
			)
			.join("")

		if (borders !== "") tail += `<w:pBdr>${borders}</w:pBdr>`
	}

	const background = own.get("fo:background-color")

	if (background === "transparent") tail += '<w:shd w:val="clear" w:color="auto" w:fill="auto"/>'
	else if (color(background) !== undefined) tail += `<w:shd w:val="clear" w:color="auto" w:fill="${color(background)}"/>`

	const stops = children.find(child => child.name === "style:tab-stops")

	if (stops !== undefined) tail += tabs(stops, options)

	const hyphenate = own.get("fo:hyphenate")

	if (hyphenate !== undefined) tail += onOff("suppressAutoHyphens", hyphenate !== "true")

	const mode = own.get("style:writing-mode")

	if (mode !== undefined) tail += onOff("bidi", mode.startsWith("rl"))

	const before = parseLength(own.get("fo:margin-top"))
	const after = parseLength(own.get("fo:margin-bottom"))
	const lineHeight = own.get("fo:line-height")
	const atLeast = parseLength(own.get("style:line-height-at-least"))
	const leading = parseLength(own.get("style:line-spacing"))
	let line = ""

	if (lineHeight !== undefined) {
		const percent = parsePercent(lineHeight)
		const length = parseLength(lineHeight)

		if (percent !== undefined) line = ` w:line="${Math.round(2.4 * percent)}" w:lineRule="auto"`
		else if (length !== undefined) line = ` w:line="${twips(length)}" w:lineRule="exact"`
		else if (lineHeight === "normal") line = ' w:line="240" w:lineRule="auto"'
	} else if (atLeast !== undefined) {
		line = ` w:line="${twips(atLeast)}" w:lineRule="atLeast"`
	} else if (leading !== undefined) {
		// Leading added under each line: a multiple of the font's line height (about 1.15 em for the bundled faces).
		sink.lossy("lineSpacingLeading")

		const normal = options.fontSize * 1.15

		line = ` w:line="${Math.round(240 * ((normal + leading) / normal))}" w:lineRule="auto"`
	}

	if (before !== undefined || after !== undefined || line !== "") {
		tail += `<w:spacing${before === undefined ? "" : ` w:before="${twips(Math.max(0, before))}"`}${after === undefined ? "" : ` w:after="${twips(Math.max(0, after))}"`}${line}/>`
	}

	if (options.indents) {
		const source = options.ownIndents === true ? own : resolved
		const left = parseLength(source.get("fo:margin-left")) ?? 0
		const right = parseLength(source.get("fo:margin-right")) ?? 0
		const indent = parseLength(source.get("fo:text-indent")) ?? 0

		tail += `<w:ind w:left="${twips(left)}" w:right="${twips(right)}" ${indent >= 0 ? `w:firstLine="${twips(indent)}"` : `w:hanging="${twips(-indent)}"`}/>`
	} else if (options.listIndent !== undefined) {
		const { left, firstLine } = options.listIndent

		tail += `<w:ind w:left="${twips(left)}" ${firstLine >= 0 ? `w:firstLine="${twips(firstLine)}"` : `w:hanging="${twips(-firstLine)}"`}/>`
	}

	const contextual = own.get("style:contextual-spacing") ?? own.get("loext:contextual-spacing")

	if (contextual !== undefined) tail += onOff("contextualSpacing", contextual === "true")

	const align = own.get("fo:text-align")

	if (align !== undefined) {
		const rtl = (resolved.get("style:writing-mode") ?? "").startsWith("rl")
		const jc =
			align === "center"
				? "center"
				: align === "justify"
					? "both"
					: align === "end"
						? rtl
							? "left"
							: "right"
						: align === "right"
							? "right"
							: align === "start"
								? rtl
									? "right"
									: "left"
								: "left"

		tail += `<w:jc w:val="${jc}"/>`
	}

	const verticalAlign = own.get("style:vertical-align")

	if (verticalAlign !== undefined && verticalAlign !== "auto") {
		tail += `<w:textAlignment w:val="${verticalAlign === "middle" ? "center" : verticalAlign === "bottom" ? "bottom" : verticalAlign === "top" ? "top" : "baseline"}"/>`
	}

	return { head, tail }
}
