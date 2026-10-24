import { buildZip, text } from "./zipBuilder"
import { bytesSource } from "@filen/office/core/opc"
import { openOdtAsDocx, translateOdt, type OdtTranslation } from "../core/odt/index"
import {
	type DocumentLayout,
	buildDocxModel,
	layoutDocument,
	parseNumbering,
	parseStyles,
	PropertyResolver,
	PropsStore,
	styleSettings
} from "../core/docx/index"
import { FontManager } from "../core/fonts/index"
import { HmtxMeasurer, nodeFontAssets, nodeFontHost } from "./layoutFixtures"
import { decodeUtf8 } from "../core/xml/index"

// Small ODT packages written by hand, and their translations.

export const ODF_NS = [
	'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
	'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
	'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
	'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
	'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"',
	'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
	'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
	'xmlns:xlink="http://www.w3.org/1999/xlink"',
	'xmlns:dc="http://purl.org/dc/elements/1.1/"',
	'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"',
	'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"',
	'xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"'
].join(" ")

export interface OdtParts {
	// office:automatic-styles content of content.xml.
	readonly automatic?: string
	// office:text content.
	readonly body: string
	// office:styles, office:automatic-styles and office:master-styles content of styles.xml.
	readonly styles?: string
	readonly stylesAutomatic?: string
	readonly masters?: string
	readonly settings?: string
	readonly extra?: readonly { readonly name: string; readonly data: Uint8Array }[]
	readonly manifestExtra?: string
}

export function odtBytes(parts: OdtParts): Uint8Array {
	const content = `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${ODF_NS} office:version="1.3"><office:automatic-styles>${parts.automatic ?? ""}</office:automatic-styles><office:body><office:text>${parts.body}</office:text></office:body></office:document-content>`
	const styles = `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${ODF_NS} office:version="1.3"><office:styles>${parts.styles ?? ""}</office:styles><office:automatic-styles>${parts.stylesAutomatic ?? '<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm" fo:margin-right="2cm"/></style:page-layout>'}</office:automatic-styles><office:master-styles>${parts.masters ?? '<style:master-page style:name="Standard" style:page-layout-name="pm1"/>'}</office:master-styles></office:document-styles>`
	const manifest = `<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>${parts.settings === undefined ? "" : '<manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>'}${(parts.extra ?? []).map(entry => `<manifest:file-entry manifest:full-path="${entry.name}" manifest:media-type=""/>`).join("")}${parts.manifestExtra ?? ""}</manifest:manifest>`

	return buildZip([
		{ name: "mimetype", data: text("application/vnd.oasis.opendocument.text"), method: 0 },
		{ name: "content.xml", data: text(content) },
		{ name: "styles.xml", data: text(styles) },
		...(parts.settings === undefined
			? []
			: [
					{
						name: "settings.xml",
						data: text(
							`<?xml version="1.0" encoding="UTF-8"?><office:document-settings ${ODF_NS} office:version="1.3"><office:settings><config:config-item-set config:name="ooo:configuration-settings">${parts.settings}</config:config-item-set></office:settings></office:document-settings>`
						)
					}
				]),
		...(parts.extra ?? []).map(entry => ({ name: entry.name, data: entry.data })),
		{ name: "META-INF/manifest.xml", data: text(manifest) }
	])
}

export async function translated(parts: OdtParts): Promise<OdtTranslation> {
	const result = await translateOdt(bytesSource(odtBytes(parts)))

	if (!result.ok) throw new Error(`${result.reason}: ${result.message}`)

	return result.translation
}

export function partText(translation: OdtTranslation, name: string): string {
	const part = translation.parts.get(name)

	if (!(part instanceof Uint8Array)) throw new Error(`No part ${name}`)

	return decodeUtf8(part)
}

// The w:body of the translated document.
export async function documentXml(parts: OdtParts): Promise<string> {
	return partText(await translated(parts), "word/document.xml")
}

// Lays out a translated document as the viewer does (bundled faces, hmtx advances, no browser).
export async function layOutOdt(bytes: Uint8Array): Promise<{ readonly pages: number; readonly translation: OdtTranslation }> {
	const { layout, translation } = await odtLayout(bytes)

	return { pages: layout.pages.length, translation }
}

// The layout itself, for line positions (layoutFixtures pageLines).
export async function odtLayout(bytes: Uint8Array): Promise<{ readonly layout: DocumentLayout; readonly translation: OdtTranslation }> {
	const opened = await openOdtAsDocx(bytesSource(bytes))

	if (!opened.ok) throw new Error(`${opened.reason}: ${opened.message}`)

	const pkg = opened.package
	const model = await buildDocxModel(pkg)
	const store = new PropsStore()
	const optional = async (kind: string) => {
		const name = pkg.part(kind)

		return name === undefined ? undefined : pkg.xml(name)
	}
	const stylesXml = await optional("styles")
	const styles = stylesXml === undefined ? undefined : parseStyles(stylesXml, store)
	const numberingXml = await optional("numbering")
	const numbering = numberingXml === undefined ? undefined : parseNumbering(numberingXml, store, styles)
	const fonts = new FontManager(nodeFontHost, nodeFontAssets, undefined, "odttest")
	const resolver = new PropertyResolver({ styles, numbering, settings: styleSettings(model.settings) })
	const layout = layoutDocument(model, resolver, fonts, new HmtxMeasurer(), { sliceMs: Infinity, store })

	await layout.done

	return { layout, translation: opened.translation }
}
