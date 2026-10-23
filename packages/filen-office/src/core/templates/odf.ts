import type { TemplateLocale } from "./locale"
import type { TemplatePart } from "./zip"

// What the blank OpenDocument packages share: the package frame (ODF 1.3 Part 2 §3.3 mimetype, §4 manifest), the
// Calibri font declaration (Part 3 §16.23) and the default text properties: Calibri 11 pt in the user's language
// (§20.185 fo:font-size, §20.190 fo:language, §20.182 fo:country), set for Latin, Asian and complex text alike.

export const OFFICE = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
export const STYLE = "urn:oasis:names:tc:opendocument:xmlns:style:1.0"
export const FO = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
export const SVG = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
export const TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
export const TABLE = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"

export const FONT_FACES =
	"<office:font-face-decls>" +
	'<style:font-face style:name="Calibri" svg:font-family="Calibri" style:font-family-generic="swiss" style:font-pitch="variable"/>' +
	"</office:font-face-decls>"

// The page: Letter or A4 (as LibreOffice writes them), portrait, LibreOffice's 2 cm margins.
export function pageLayout(paper: TemplateLocale["paper"]): string {
	const size = paper === "letter" ? 'fo:page-width="8.5in" fo:page-height="11in"' : 'fo:page-width="21cm" fo:page-height="29.7cm"'

	return (
		`<style:page-layout style:name="pm1"><style:page-layout-properties ${size} style:print-orientation="portrait" ` +
		'fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm" fo:margin-right="2cm"/></style:page-layout>'
	)
}

// templateLocale's tag split into fo:language and fo:country (no country when the region is unknown).
export function textProperties({ tag }: TemplateLocale): string {
	const [language, country] = tag.split("-")
	const lang = `fo:language="${language ?? "en"}"${country === undefined ? "" : ` fo:country="${country}"`}`

	return (
		`<style:text-properties style:font-name="Calibri" fo:font-size="11pt" ${lang} ` +
		'style:font-name-asian="Calibri" style:font-size-asian="11pt" style:font-name-complex="Calibri" style:font-size-complex="11pt"/>'
	)
}

// mimetype first and stored, then the streams, then the manifest listing them.
export function odfPackage(mediaType: string, content: string, styles: string): TemplatePart[] {
	return [
		["mimetype", mediaType, "stored"],
		["content.xml", content],
		["styles.xml", styles],
		[
			"META-INF/manifest.xml",
			'<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">' +
				`<manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="${mediaType}"/>` +
				'<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>' +
				'<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>' +
				"</manifest:manifest>"
		]
	]
}
