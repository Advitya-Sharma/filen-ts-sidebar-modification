import { readDocxFontTable } from "../../fonts/index"
import { OpcError, ZipError } from "../../opc/index"
import { XmlError } from "../../xml/index"
import { DOCX_THEME_BODY_FONT, DOCX_THEME_HEADINGS_FONT } from "./commands/formatting"
import type { DocxEditContext } from "./context"

// The fonts a font picker offers, from the document alone (nothing is fetched, installed fonts are not
// enumerated: that asks for a permission and tells sites what is installed): the theme's heading and body
// fonts, the fonts the document declares, and the common fonts the engine sets with bundled substitutes.

export type DocxFontKind = "theme" | "document" | "common"

export interface DocxFontChoice {
	// What font.family takes (DOCX_THEME_HEADINGS_FONT / DOCX_THEME_BODY_FONT for the theme's fonts).
	readonly family: string
	// The typeface, as shown.
	readonly name: string
	readonly kind: DocxFontKind
	// For the theme's fonts.
	readonly role?: "headings" | "body"
}

// The originals of the bundled substitutes (core/fonts substitutes.ts): they lay out alike everywhere.
export const DOCX_COMMON_FONTS: readonly string[] = [
	"Aptos",
	"Arial",
	"Calibri",
	"Cambria",
	"Courier New",
	"Georgia",
	"Segoe UI",
	"Times New Roman"
]

function isDocumentError(error: unknown): boolean {
	return error instanceof XmlError || error instanceof ZipError || error instanceof OpcError
}

const LISTS = new WeakMap<DocxEditContext, Promise<readonly DocxFontChoice[]>>()

async function declaredFonts(context: DocxEditContext): Promise<string[]> {
	const pkg = context.model.package
	const part = pkg.part("fontTable")

	if (part === undefined || !pkg.has(part)) return []

	try {
		// Read past the package's shared parses: only the names are kept.
		return readDocxFontTable(await pkg.opc.xml(part)).map(declaration => declaration.name)
	} catch (error) {
		if (isDocumentError(error)) return []

		throw error
	}
}

async function buildList(context: DocxEditContext): Promise<readonly DocxFontChoice[]> {
	const out: DocxFontChoice[] = []
	const theme = context.resolver.theme
	const languages = context.resolver.settings.themeFontLanguage
	const seen = new Set<string>()

	for (const [role, family, ref] of [
		["headings", DOCX_THEME_HEADINGS_FONT, "majorHAnsi"],
		["body", DOCX_THEME_BODY_FONT, "minorHAnsi"]
	] as const) {
		const name = theme?.font(ref, languages)

		if (name !== undefined) out.push(Object.freeze({ family, name, kind: "theme", role }))
	}

	const add = (names: readonly string[], kind: DocxFontKind): void => {
		const fresh = names.filter(name => {
			const key = name.toLowerCase()

			if (name.trim() === "" || seen.has(key)) return false

			seen.add(key)

			return true
		})

		for (const name of fresh.sort((a, b) => a.localeCompare(b))) out.push(Object.freeze({ family: name, name, kind }))
	}

	add(await declaredFonts(context), "document")
	add(DOCX_COMMON_FONTS, "common")

	return Object.freeze(out)
}

// The font list of a document (made once per edit context).
export function docxFontList(context: DocxEditContext): Promise<readonly DocxFontChoice[]> {
	let list = LISTS.get(context)

	if (list === undefined) {
		list = buildList(context)
		LISTS.set(context, list)
	}

	return list
}
