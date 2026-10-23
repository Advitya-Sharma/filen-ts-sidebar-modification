import type { OdfStyle, OdfStyleSheet } from "../../odf/styles"
import type { SheetBorderLine, SheetRenderStyle } from "../../sheet/protocol"
import type { SheetStyles } from "../../sheet/style/resolve"
import type { XlsxFont } from "../../xlsx/model/types"
import { escapeAttribute } from "../../xml/index"
import type { OdsCellStyles } from "../styles/cell"
import { dataStyles } from "./dataStyle"

// The automatic styles a save adds to content.xml (ODF 1.3 Part 3 §16.2): a cell format, a row height, a column
// width, a sheet's look or a run's font the source has no style for gets one, each written once. A cell format the
// source named keeps that name. What a style cannot say is refused (StyleRefused) with the property.

export class StyleRefused extends Error {
	readonly property: string

	constructor(property: string) {
		super(`A style cannot hold ${property}`)
		this.property = property
	}
}

function hex(rgb: number): string {
	return `#${(rgb & 0xffffff).toString(16).padStart(6, "0")}`
}

// A length in points as written: exact, as a JavaScript number prints (a length takes no exponent).
function points(value: number): string {
	const text = String(value)

	return `${text.includes("e") ? value.toFixed(12) : text}pt`
}

// fo:font-family (§20.192, XSL's font-family): a name with anything but letters, digits and "-" quoted.
function family(name: string): string {
	return /^[A-Za-z][A-Za-z0-9-]*$/.test(name) ? name : `'${name.replaceAll("'", "")}'`
}

// A border as fo:border-* (§20.183) that styles/cell.ts reads back as the same line.
function border(line: SheetBorderLine | undefined): string {
	if (line === undefined) return "none"

	const color = hex(line.color)

	switch (line.style) {
		case "thin":
			return `0.74pt solid ${color}`
		case "medium":
			return `1.75pt solid ${color}`
		case "thick":
			return `2.5pt solid ${color}`
		case "double":
			return `2.6pt double ${color}`
		case "dotted":
			return `0.74pt dotted ${color}`
		case "dashed":
			return `0.74pt dashed ${color}`
		case "mediumDashed":
			return `1.75pt dashed ${color}`
		case "dashDot":
			return `0.74pt dash-dot ${color}`
		case "mediumDashDot":
			return `1.75pt dash-dot ${color}`
		case "dashDotDot":
			return `0.74pt dash-dot-dot ${color}`
		case "mediumDashDotDot":
			return `1.75pt dash-dot-dot ${color}`
		default:
			throw new StyleRefused(`a ${line.style} border`)
	}
}

function attributes(list: readonly (readonly [string, string | undefined])[]): string {
	return list
		.filter((entry): entry is readonly [string, string] => entry[1] !== undefined)
		.map(([name, value]) => ` ${name}="${escapeAttribute(value)}"`)
		.join("")
}

interface TextLook {
	readonly name?: string
	readonly size?: number
	readonly bold?: boolean
	readonly italic?: boolean
	readonly strike?: boolean
	readonly underline?: string
	readonly script?: "superscript" | "subscript"
	// 0xRRGGBB.
	readonly color?: number
}

function renderLook(style: SheetRenderStyle): TextLook {
	return {
		...(style.font !== undefined && { name: style.font }),
		...(style.size !== undefined && { size: style.size }),
		...(style.bold !== undefined && { bold: style.bold }),
		...(style.italic !== undefined && { italic: style.italic }),
		...(style.strike !== undefined && { strike: style.strike }),
		...(style.underline !== undefined && { underline: style.underline }),
		...(style.script !== undefined && { script: style.script }),
		...(style.color !== undefined && { color: style.color })
	}
}

function fontLook(font: XlsxFont, styles: SheetStyles): TextLook {
	const color = font.color === undefined || font.color.auto === true ? -1 : styles.color(font.color)

	return {
		...(font.name !== undefined && { name: font.name }),
		...(font.size !== undefined && { size: font.size }),
		...(font.bold !== undefined && { bold: font.bold }),
		...(font.italic !== undefined && { italic: font.italic }),
		...(font.strike !== undefined && { strike: font.strike }),
		...(font.underline !== undefined && font.underline !== "none" && { underline: font.underline }),
		...((font.vertAlign === "superscript" || font.vertAlign === "subscript") && { script: font.vertAlign }),
		...(color >= 0 && { color })
	}
}

// style:text-properties of a look: the whole of it for a cell's style, what it sets for a run's. A font goes by its
// font face (style:font-name, which wins over fo:font-family where both are, §20.191).
function textProperties(look: TextLook, whole: boolean, face: (family: string) => string): string {
	const underline = look.underline

	if (underline === "singleAccounting" || underline === "doubleAccounting") throw new StyleRefused("an accounting underline")

	const underlined = underline !== undefined
	const set = (on: boolean | undefined, value: string, off: string): string | undefined => (on === true ? value : whole ? off : undefined)

	return `<style:text-properties${attributes([
		["style:font-name", look.name === undefined ? undefined : face(look.name)],
		["fo:font-size", look.size === undefined ? undefined : points(look.size)],
		["fo:font-weight", set(look.bold, "bold", "normal")],
		["fo:font-style", set(look.italic, "italic", "normal")],
		["style:text-underline-style", set(underlined, "solid", "none")],
		["style:text-underline-type", underline === "double" ? "double" : undefined],
		["style:text-underline-width", underlined ? "auto" : undefined],
		["style:text-underline-color", underlined ? "font-color" : undefined],
		["style:text-line-through-style", set(look.strike, "solid", "none")],
		[
			"style:text-position",
			look.script === "superscript" ? "super 58%" : look.script === "subscript" ? "sub 58%" : whole ? "0% 100%" : undefined
		],
		["fo:color", look.color === undefined ? undefined : hex(look.color)]
	])}/>`
}

// A cell style's properties for format `xf`: its cell, paragraph and whole text properties; a font face goes by the
// name `face` gives it.
export function cellStyleProperties(styles: SheetStyles, xf: number, face: (family: string) => string): string {
	// A font without a name or size of its own has the workbook's base font's (an ODF style always has both).
	const base = styles.renderStyle(0)
	const own = styles.renderStyle(xf)
	const render: SheetRenderStyle = {
		...own,
		...(own.font === undefined && base.font !== undefined && { font: base.font }),
		...(own.size === undefined && base.size !== undefined && { size: base.size }),
		...(own.color === undefined && base.color !== undefined && { color: base.color })
	}
	const protection = styles.protection(xf)
	const locked = protection?.locked !== false
	const hidden = protection?.hidden === true
	const protect =
		protection?.hideAll === true
			? "hidden-and-protected"
			: locked && hidden
				? "protected formula-hidden"
				: locked
					? "protected"
					: hidden
						? "formula-hidden"
						: "none"

	if (render.pattern !== undefined) throw new StyleRefused("a pattern fill")
	if (render.align === "centerContinuous" || render.align === "distributed") throw new StyleRefused(`${render.align} alignment`)
	if (render.valign === "justify" || render.valign === "distributed") throw new StyleRefused(`${render.valign} alignment`)

	const rotation = render.rotation
	const angle = rotation === undefined || rotation === 255 ? "0" : String(rotation < 0 ? 360 + rotation : rotation)
	const diagonal = render.diagonalUp ?? render.diagonalDown
	const cell = `<style:table-cell-properties${attributes([
		["fo:background-color", render.fill === undefined ? "transparent" : hex(render.fill)],
		["fo:border-top", border(render.top)],
		["fo:border-bottom", border(render.bottom)],
		["fo:border-left", border(render.left)],
		["fo:border-right", border(render.right)],
		["style:diagonal-bl-tr", render.diagonalUp === undefined ? undefined : border(diagonal)],
		["style:diagonal-tl-br", render.diagonalDown === undefined ? undefined : border(diagonal)],
		["style:text-align-source", render.align === undefined ? "value-type" : "fix"],
		["style:repeat-content", render.align === "fill" ? "true" : "false"],
		["style:vertical-align", render.valign === "top" ? "top" : render.valign === "center" ? "middle" : "bottom"],
		["fo:wrap-option", render.wrap === true ? "wrap" : "no-wrap"],
		["style:shrink-to-fit", render.shrink === true ? "true" : "false"],
		["style:rotation-angle", angle],
		["style:direction", rotation === 255 ? "ttb" : "ltr"],
		["style:writing-mode", render.rightToLeft === true ? "rl-tb" : undefined],
		["style:cell-protect", protect]
	])}/>`
	const align = render.align
	const paragraph = `<style:paragraph-properties${attributes([
		[
			"fo:text-align",
			align === "left"
				? "start"
				: align === "right"
					? "end"
					: align === "center"
						? "center"
						: align === "justify"
							? "justify"
							: undefined
		],
		["fo:margin-left", render.indent === undefined ? undefined : points(render.indent * 15.75)]
	])}/>`

	return `${cell}${paragraph}${textProperties(renderLook(render), true, face)}`
}

export class OdsStyleWriter {
	private readonly sheet: OdfStyleSheet
	private readonly styles: SheetStyles
	private readonly cellStyles: OdsCellStyles
	private readonly taken = new Set<string>()
	private readonly cellNames = new Map<number, string>()
	private readonly dataNames = new Map<string, string>()
	private readonly rowNames = new Map<string, string>()
	private readonly columnNames = new Map<string, string>()
	private readonly tableNames = new Map<string, string>()
	private readonly textNames = new Map<string, string>()
	private readonly elements: string[] = []
	// Font faces added (style:font-face elements, §16.23) and the face each family goes by.
	private readonly faces = new Map<string, string>()
	private readonly faceElements: string[] = []

	constructor(cellStyles: OdsCellStyles, styles: SheetStyles) {
		this.cellStyles = cellStyles
		this.sheet = cellStyles.styleSheet
		this.styles = styles
	}

	get written(): number {
		return this.elements.length
	}

	// The automatic styles added, in the order written (a data style before the cell style naming it).
	markup(): string {
		return this.elements.join("")
	}

	// The font faces added, for office:font-face-decls.
	fontFaces(): string {
		return this.faceElements.join("")
	}

	// The font face a family goes by: one the package declares, else one added.
	private face(name: string): string {
		let face = this.faces.get(name)

		if (face === undefined) {
			for (const [declared, entry] of this.sheet.fontFaces) {
				if (entry.family === name) {
					face = declared

					break
				}
			}

			if (face === undefined) {
				face = name

				for (let n = 2; this.sheet.fontFaces.has(face); n++) face = `${name}${String(n)}`

				this.faceElements.push(
					`<style:font-face style:name="${escapeAttribute(face)}" svg:font-family="${escapeAttribute(family(name))}"/>`
				)
			}

			this.faces.set(name, face)
		}

		return face
	}

	private name(prefix: string, family: string | null): string {
		for (let n = 1; ; n++) {
			const name = `${prefix}${String(n)}`

			if (this.taken.has(name)) continue
			if (
				family === null
					? this.sheet.dataStyle(name, "content") !== undefined
					: this.sheet.style(family, name, "content") !== undefined
			)
				continue

			this.taken.add(name)

			return name
		}
	}

	// The name a cell of format `xf` is written with: the source's for a format read from it, else a new one's.
	cell(xf: number): string | undefined {
		const source = xf === 0 ? this.cellStyles.defaultStyle : this.cellStyles.sourceName(xf)

		if (source !== undefined) return source

		let name = this.cellNames.get(xf)

		if (name === undefined) {
			const markup = this.cellMarkup(xf)

			name = this.name("ceF", "table-cell")
			this.elements.push(markup.replace("\u0000", escapeAttribute(name)))
			this.cellNames.set(xf, name)
		}

		return name
	}

	private cellMarkup(xf: number): string {
		const code = this.styles.numFmtCode(xf)
		const data = code === "General" ? undefined : this.data(code, xf)

		return `<style:style style:name="\u0000" style:family="table-cell"${attributes([
			["style:parent-style-name", this.cellStyles.defaultStyle],
			["style:data-style-name", data]
		])}>${cellStyleProperties(this.styles, xf, name => this.face(name))}</style:style>`
	}

	// The data style of a number format: the source's for a code read from it, else new ones.
	private data(code: string, xf: number): string {
		const id = this.cellStyles.table.cellXfs[xf]?.numFmtId ?? -1
		const source = id < 0 ? undefined : this.cellStyles.dataStyleName(id)

		if (source !== undefined) return source

		let name = this.dataNames.get(code)

		if (name === undefined) {
			name = this.name("NF", null)

			const written = dataStyles(code, name)

			if ("refused" in written) throw new StyleRefused(`the number format ${code} (${written.refused})`)

			for (const style of written.styles) {
				if (style.name !== name) this.taken.add(style.name)

				this.elements.push(style.markup)
			}

			this.dataNames.set(code, name)
		}

		return name
	}

	// A row style of a height in points (NaN: the optimal one).
	row(heightPt: number, custom: boolean, pageBreak: boolean): string {
		const key = `${String(heightPt)}|${String(custom)}|${String(pageBreak)}`
		let name = this.rowNames.get(key)

		if (name === undefined) {
			name = this.name("roF", "table-row")
			this.elements.push(
				`<style:style style:name="${escapeAttribute(name)}" style:family="table-row"><style:table-row-properties${attributes([
					["style:row-height", Number.isNaN(heightPt) ? undefined : points(heightPt)],
					["fo:break-before", pageBreak ? "page" : "auto"],
					["style:use-optimal-row-height", custom || Number.isNaN(heightPt) ? String(!custom) : "true"]
				])}/></style:style>`
			)
			this.rowNames.set(key, name)
		}

		return name
	}

	// A column style of a width in pixels (NaN: none).
	column(widthPx: number, pageBreak: boolean): string {
		const key = `${String(widthPx)}|${String(pageBreak)}`
		let name = this.columnNames.get(key)

		if (name === undefined) {
			name = this.name("coF", "table-column")
			this.elements.push(
				`<style:style style:name="${escapeAttribute(name)}" style:family="table-column"><style:table-column-properties${attributes([
					["fo:break-before", pageBreak ? "page" : "auto"],
					["style:column-width", Number.isNaN(widthPx) ? undefined : points(widthPx * 0.75)]
				])}/></style:style>`
			)
			this.columnNames.set(key, name)
		}

		return name
	}

	// A sheet's style: the source's (`base`) with whether it shows, its direction and its tab colour (-1: none), and a
	// new one's master page.
	table(base: string | undefined, display: boolean, rightToLeft: boolean, tabColor: number, masterPage?: string): string {
		const key = `${base ?? ""}|${String(display)}|${String(rightToLeft)}|${String(tabColor)}|${masterPage ?? ""}`
		let name = this.tableNames.get(key)

		if (name === undefined) {
			const style: OdfStyle | undefined = base === undefined ? undefined : this.sheet.style("table", base, "content")
			const own = new Map(style?.properties.get("table")?.attributes ?? [])

			own.set("table:display", String(display))
			own.set("style:writing-mode", rightToLeft ? "rl-tb" : "lr-tb")
			own.delete("tableooo:tab-color")

			if (tabColor < 0) own.delete("table:tab-color")
			else own.set("table:tab-color", hex(tabColor))

			name = this.name("taF", "table")

			const element = new Map(style?.attributes ?? [])

			element.delete("style:name")
			element.delete("style:family")

			if (style?.parent !== undefined) element.set("style:parent-style-name", style.parent)
			if (masterPage !== undefined) element.set("style:master-page-name", masterPage)
			else if (!element.has("style:master-page-name")) element.set("style:master-page-name", "Default")

			this.elements.push(
				`<style:style style:name="${escapeAttribute(name)}" style:family="table"${attributes([...element])}><style:table-properties${attributes([...own])}/></style:style>`
			)
			this.tableNames.set(key, name)
		}

		return name
	}

	// A text style for a run's font (family text): only what the font sets.
	text(font: XlsxFont): string {
		const key = JSON.stringify(font)
		let name = this.textNames.get(key)

		if (name === undefined) {
			name = this.name("TF", "text")
			this.elements.push(
				`<style:style style:name="${escapeAttribute(name)}" style:family="text">${textProperties(fontLook(font, this.styles), false, name => this.face(name))}</style:style>`
			)
			this.textNames.set(key, name)
		}

		return name
	}
}
