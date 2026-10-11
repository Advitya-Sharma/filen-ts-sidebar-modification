import { readDrawingColor } from "../docx/theme/index"
import type { XmlDocument } from "../xml/index"
import { rgba, rgbaAlpha, rgbaHex, type Rgba } from "./displayList"
import type { ImageResolver } from "./picture"
import { D, drawingReader, parseFraction, type DrawingReader } from "./reader"

// What the drawing layer needs from a theme: scheme colours (after the host's colour map), the major and
// minor fonts and the format scheme for style references.
export interface DrawingTheme {
	// a:schemeClr names ("accent1", "tx1", "bg2"...) -> six hex digits, through the host's colour mapping.
	scheme(name: string): string | undefined
	readonly majorFont: ThemeFonts
	readonly minorFont: ThemeFonts
	readonly formats: FormatScheme | undefined
}

export interface ThemeFonts {
	readonly latin: string
	readonly eastAsia: string
	readonly complex: string
}

// a:fmtScheme (§20.1.4.1.14): nodes of the theme part, read when a style reference needs one.
export interface FormatScheme {
	readonly reader: DrawingReader
	readonly fills: readonly number[]
	readonly lines: readonly number[]
	readonly effects: readonly number[]
	readonly backgroundFills: readonly number[]
	// Pictures of theme fills (a:blipFill in a fill style) resolve against the theme part's relationships.
	readonly images: ImageResolver
}

export const COLOR_ELEMENTS: ReadonlySet<string> = new Set(["srgbClr", "sysClr", "schemeClr", "scrgbClr", "hslClr", "prstClr"])

const NO_FONTS: ThemeFonts = { latin: "", eastAsia: "", complex: "" }

function fontSet(reader: DrawingReader, node: number): ThemeFonts {
	if (node < 0) return NO_FONTS

	return {
		latin: reader.attr(reader.child(node, D.A, "latin"), "typeface") ?? "",
		eastAsia: reader.attr(reader.child(node, D.A, "ea"), "typeface") ?? "",
		complex: reader.attr(reader.child(node, D.A, "cs"), "typeface") ?? ""
	}
}

// A DrawingTheme from a theme part (a:theme). `scheme` resolves a:schemeClr names; the host supplies it
// so its colour mapping (w:clrSchemeMapping, p:clrMap) applies. `images` resolves the theme part's own
// relationships (picture fills in its fill styles); without it they draw as placeholders.
export function drawingTheme(
	document: XmlDocument | undefined,
	scheme: (name: string) => string | undefined,
	images: ImageResolver = () => undefined
): DrawingTheme {
	if (document === undefined) return { scheme, majorFont: NO_FONTS, minorFont: NO_FONTS, formats: undefined }

	const reader = drawingReader(document)
	const root = document.documentElement()
	const elements = reader.child(root, D.A, "themeElements")
	const fonts = reader.child(elements, D.A, "fontScheme")
	const format = reader.child(elements, D.A, "fmtScheme")
	const list = (name: string): number[] => {
		const container = reader.child(format, D.A, name)
		const out: number[] = []

		for (
			let child = container < 0 ? -1 : document.firstChildElement(container);
			child >= 0;
			child = document.nextSiblingElement(child)
		) {
			out.push(child)
		}

		return out
	}

	return {
		scheme,
		majorFont: fontSet(reader, reader.child(fonts, D.A, "majorFont")),
		minorFont: fontSet(reader, reader.child(fonts, D.A, "minorFont")),
		formats:
			format < 0
				? undefined
				: {
						reader,
						fills: list("fillStyleLst"),
						lines: list("lnStyleLst"),
						effects: list("effectStyleLst"),
						backgroundFills: list("bgFillStyleLst"),
						images
					}
	}
}

// The colour element inside a container (a:solidFill, a:gs, a:fgClr...), or -1.
export function colorChild(reader: DrawingReader, node: number): number {
	return reader.childOf(node, D.A, COLOR_ELEMENTS)
}

// Reads a colour element with its transforms, alpha included. `placeholder` is the colour a:schemeClr
// val="phClr" stands for (the style reference's colour inside a theme format).
export function readColor(reader: DrawingReader, node: number, theme: DrawingTheme, placeholder?: Rgba): Rgba | undefined {
	if (node < 0) return undefined

	const document = reader.document
	const isPlaceholder = reader.local(node) === "schemeClr" && reader.attr(node, "val") === "phClr"
	const hex = readDrawingColor(document, node, name =>
		name === "phClr" ? (placeholder === undefined ? undefined : rgbaHex(placeholder)) : theme.scheme(name)
	)

	if (hex === undefined) return undefined

	let alpha = isPlaceholder && placeholder !== undefined ? rgbaAlpha(placeholder) : 1

	for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) {
		if (reader.group(child) !== D.A) continue

		const local = document.localName(child)

		if (local !== "alpha" && local !== "alphaMod" && local !== "alphaOff") continue

		const value = parseFraction(reader.attr(child, "val") ?? "", Number.NaN)

		if (Number.isNaN(value)) continue

		if (local === "alpha") alpha = value
		else if (local === "alphaMod") alpha *= value
		else alpha += value
	}

	return rgba(hex, alpha)
}

// The colour inside a container element.
export function readColorIn(reader: DrawingReader, container: number, theme: DrawingTheme, placeholder?: Rgba): Rgba | undefined {
	return readColor(reader, colorChild(reader, container), theme, placeholder)
}
