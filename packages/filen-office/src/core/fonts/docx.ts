import type { OpcPackage } from "../opc/index"
import { Namespaces, type XmlDocument } from "../xml/index"
import { FontError } from "./errors"
import type { FontDeclaration, FontManager } from "./manager"
import { deobfuscateFont } from "./obfuscation"
import type { FontClass, FontPitch } from "./substitutes"

// word/fontTable.xml: per font its alternative names, classification hints and embedded faces.

const W = ["http://schemas.openxmlformats.org/wordprocessingml/2006/main", "http://purl.oclc.org/ooxml/wordprocessingml/main"]
const R = ["http://schemas.openxmlformats.org/officeDocument/2006/relationships", "http://purl.oclc.org/ooxml/officeDocument/relationships"]

const ZIP_STORED = 0

export type EmbeddedStyle = "regular" | "bold" | "italic" | "boldItalic"

const EMBED_ELEMENTS: readonly (readonly [string, EmbeddedStyle])[] = [
	["embedRegular", "regular"],
	["embedBold", "bold"],
	["embedItalic", "italic"],
	["embedBoldItalic", "boldItalic"]
]

export interface EmbeddedFontRef {
	readonly style: EmbeddedStyle
	readonly relationshipId: string
	readonly fontKey: string | undefined
	// Only the characters the document used are in the font.
	readonly subsetted: boolean
}

export interface DocxFontDeclaration extends FontDeclaration {
	readonly embeds: readonly EmbeddedFontRef[]
}

const FONT_CLASSES = new Set<string>(["roman", "swiss", "modern", "script", "decorative", "auto"])
const PITCHES = new Set<string>(["fixed", "variable", "default"])

function isFontClass(value: string | undefined): value is FontClass {
	return value !== undefined && FONT_CLASSES.has(value)
}

function isPitch(value: string | undefined): value is FontPitch {
	return value !== undefined && PITCHES.has(value)
}

function panose(value: string | undefined): Uint8Array | undefined {
	if (value === undefined || !/^[0-9a-fA-F]{20}$/.test(value)) return undefined

	const bytes = new Uint8Array(10)

	for (let i = 0; i < 10; i++) bytes[i] = Number.parseInt(value.slice(i * 2, i * 2 + 2), 16)

	return bytes
}

function onOff(value: string): boolean {
	return value === "1" || value === "true" || value === "on"
}

export function readDocxFontTable(document: XmlDocument): DocxFontDeclaration[] {
	const namespaces = new Namespaces(document)
	const root = document.documentElement()
	const w = W.find(uri => namespaces.is(root, uri, "fonts"))

	if (w === undefined) return []

	const attribute = (node: number, name: string): string | undefined => namespaces.attribute(node, w, name)
	const relationship = (node: number): string | undefined => {
		for (const uri of R) {
			const id = namespaces.attribute(node, uri, "id")

			if (id !== undefined) return id
		}

		return undefined
	}
	const out: DocxFontDeclaration[] = []

	for (let font = document.firstChildElement(root); font >= 0; font = document.nextSiblingElement(font)) {
		const name = namespaces.is(font, w, "font") ? attribute(font, "name") : undefined

		if (name === undefined || name.length === 0) continue

		const altNames: string[] = []
		const embeds: EmbeddedFontRef[] = []
		let fontClass: FontClass | undefined
		let pitch: FontPitch | undefined
		let panoseBytes: Uint8Array | undefined
		let charset: number | undefined

		for (let child = document.firstChildElement(font); child >= 0; child = document.nextSiblingElement(child)) {
			if (namespaces.elementUri(child) !== w) continue

			const local = document.localName(child)
			const value = attribute(child, "val")

			if (local === "altName" && value !== undefined) {
				// Word writes one name; other producers list several.
				for (const alt of value.split(",")) {
					const trimmed = alt.trim()

					if (trimmed.length > 0 && trimmed !== name) altNames.push(trimmed)
				}
			} else if (local === "family" && isFontClass(value)) {
				fontClass = value
			} else if (local === "pitch" && isPitch(value)) {
				pitch = value
			} else if (local === "panose1") {
				panoseBytes = panose(value)
			} else if (local === "charset" && value !== undefined && /^[0-9a-fA-F]{2}$/.test(value)) {
				charset = Number.parseInt(value, 16)
			} else {
				const style = EMBED_ELEMENTS.find(([element]) => element === local)?.[1]
				const id = style === undefined ? undefined : relationship(child)

				if (style !== undefined && id !== undefined) {
					const subsetted = attribute(child, "subsetted")

					embeds.push({
						style,
						relationshipId: id,
						fontKey: attribute(child, "fontKey"),
						subsetted: subsetted !== undefined && onOff(subsetted)
					})
				}
			}
		}

		out.push({ name, altNames, hints: { fontClass, pitch, panose: panoseBytes, charset }, embeds })
	}

	return out
}

// The font bytes of one embedded face, de-obfuscated and checked against the size limit before inflating.
export async function readDocxEmbeddedFont(
	pkg: OpcPackage,
	fontTablePart: string,
	ref: EmbeddedFontRef,
	maxBytes: number
): Promise<Uint8Array> {
	const target = (await pkg.relationships(fontTablePart)).byId(ref.relationshipId)
	const entry = target?.partName === undefined ? undefined : pkg.entry(target.partName)

	if (target === undefined || target.partName === undefined || entry === undefined) {
		throw new FontError("corrupt", `No font part for ${ref.relationshipId}`)
	}

	if (entry.uncompressedSize > maxBytes) throw new FontError("tooLarge", `${entry.name} exceeds the font size limit`)

	const bytes = await pkg.read(target.partName)

	// Fonts without a key are stored in the clear, as some producers do.
	if (ref.fontKey === undefined) return bytes

	return deobfuscateFont(bytes, ref.fontKey, entry.method !== ZIP_STORED)
}

// Declares every embedded face of a font table to the manager. Nothing is read until text needs a face;
// faces that then fail are reported by manager.failures() and their text falls back to a substitute.
export function declareDocxEmbeddedFonts(
	manager: FontManager,
	pkg: OpcPackage,
	fontTablePart: string,
	declarations: readonly DocxFontDeclaration[]
): void {
	for (const declaration of declarations) {
		for (const ref of declaration.embeds) {
			const bold = ref.style === "bold" || ref.style === "boldItalic"
			const italic = ref.style === "italic" || ref.style === "boldItalic"

			manager.embed(declaration.name, bold, italic, maxBytes => readDocxEmbeddedFont(pkg, fontTablePart, ref, maxBytes))
		}
	}
}
