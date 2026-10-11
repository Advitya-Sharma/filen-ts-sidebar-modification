import { XmlEditor, encodeUtf8, escapeAttribute, type XmlDocument } from "../../xml/index"
import type { DocxStyles } from "../styles/index"
import { partReader } from "../model/reader"
import { Names } from "./markup"
import { W_NAMESPACE } from "./story"

// The built-in paragraph styles commands apply: [uiPriority, w:pPr content, w:rPr content], prefix "w:"
// replaced by the part's own. The values are Word's built-in defaults (current Office theme) as observed in a
// document Word 16 wrote, in the public test corpus (corpus/manifest.json, SHA-256
// 599533082406a29a6238b5a1799c5395d884a8506feafd455fe91d4554218b62); no third-party code is copied.
const DEFINITIONS: Readonly<Record<string, readonly [number, string, string]>> = {
	title: [
		10,
		'<w:spacing w:after="80" w:line="240" w:lineRule="auto"/><w:contextualSpacing/>',
		'<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/>'
	],
	subtitle: [
		11,
		'<w:numPr><w:ilvl w:val="1"/></w:numPr>',
		'<w:rFonts w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/><w:color w:val="595959" w:themeColor="text1" w:themeTint="A6"/><w:spacing w:val="15"/><w:sz w:val="28"/><w:szCs w:val="28"/>'
	],
	quote: [
		29,
		'<w:spacing w:before="160"/><w:jc w:val="center"/>',
		'<w:i/><w:iCs/><w:color w:val="404040" w:themeColor="text1" w:themeTint="BF"/>'
	],
	"heading 1": heading(0, '<w:spacing w:before="360" w:after="80"/>', "major", "", "0F4761", 40),
	"heading 2": heading(1, '<w:spacing w:before="160" w:after="80"/>', "major", "", "0F4761", 32),
	"heading 3": heading(2, '<w:spacing w:before="160" w:after="80"/>', "", "", "0F4761", 28),
	"heading 4": heading(3, '<w:spacing w:before="80" w:after="40"/>', "", "<w:i/><w:iCs/>", "0F4761", 0),
	"heading 5": heading(4, '<w:spacing w:before="80" w:after="40"/>', "", "", "0F4761", 0),
	"heading 6": heading(5, '<w:spacing w:before="40" w:after="0"/>', "", "<w:i/><w:iCs/>", "595959", 0),
	"heading 7": heading(6, '<w:spacing w:before="40" w:after="0"/>', "", "", "595959", 0),
	"heading 8": heading(7, '<w:spacing w:after="0"/>', "", "<w:i/><w:iCs/>", "272727", 0),
	"heading 9": heading(8, '<w:spacing w:after="0"/>', "", "", "272727", 0)
}

function heading(
	level: number,
	spacing: string,
	fonts: "major" | "",
	toggles: string,
	color: string,
	size: number
): readonly [number, string, string] {
	const theme =
		color === "0F4761"
			? 'w:themeColor="accent1" w:themeShade="BF"'
			: color === "595959"
				? 'w:themeColor="text1" w:themeTint="A6"'
				: 'w:themeColor="text1" w:themeTint="D8"'
	const rFonts =
		fonts === "major"
			? '<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>'
			: '<w:rFonts w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/>'

	return [
		9,
		`<w:keepNext/><w:keepLines/>${spacing}<w:outlineLvl w:val="${level}"/>`,
		`${rFonts}${toggles}<w:color w:val="${color}" ${theme}/>${size > 0 ? `<w:sz w:val="${size}"/><w:szCs w:val="${size}"/>` : ""}`
	]
}

export function hasStyleDefinition(name: string): boolean {
	return name in DEFINITIONS
}

function withPrefix(markup: string, names: Names): string {
	return names.w === "w"
		? markup
		: markup.replace(/(<\/?)w:/g, `$1${names.w === "" ? "" : `${names.w}:`}`).replace(/ w:/g, names.w === "" ? " " : ` ${names.w}:`)
}

// The w:style element for a built-in style a command used, based on the document's default paragraph style.
export function styleDefinition(names: Names, id: string, name: string, normal: string | undefined): string {
	const definition = DEFINITIONS[name]

	if (definition === undefined) throw new RangeError(`No definition for ${name}`)

	const [priority, pPr, rPr] = definition
	const based = normal === undefined ? "" : `${names.val("basedOn", normal)}${names.val("next", normal)}`

	return (
		`<${names.el("style")} ${names.attr("type")}="paragraph" ${names.attr("styleId")}="${escapeAttribute(id)}">${names.val("name", name)}${based}` +
		`${names.val("uiPriority", String(priority))}${names.empty("qFormat")}<${names.el("pPr")}>${withPrefix(pPr, names)}</${names.el("pPr")}>` +
		`<${names.el("rPr")}>${withPrefix(rPr, names)}</${names.el("rPr")}></${names.el("style")}>`
	)
}

export const STYLES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"

// styles.xml with the definitions of `missing` (style id -> built-in name) appended after the last style,
// or a new part holding only them.
export function stylesWithDefinitions(
	document: XmlDocument | undefined,
	styles: DocxStyles,
	missing: ReadonlyMap<string, string>
): Uint8Array | null {
	if (missing.size === 0) return null

	const normal = styles.defaultOf("paragraph")?.id

	if (document === undefined) {
		const names = new Names("w", "")
		const body = [...missing].map(([id, name]) => styleDefinition(names, id, name, normal)).join("")

		return encodeUtf8(
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:styles xmlns:w="${W_NAMESPACE}">${body}</w:styles>`
		)
	}

	const reader = partReader(document)
	const root = document.documentElement()
	const names = new Names(reader.namespaces.prefixFor(root, W_NAMESPACE) ?? "", "")
	const editor = new XmlEditor(document)

	editor.append(root, [...missing].map(([id, name]) => styleDefinition(names, id, name, normal)).join(""))

	return editor.serialize()
}
