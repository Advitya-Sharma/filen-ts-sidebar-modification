import { XmlEditor, encodeUtf8, escapeAttribute, type XmlDocument } from "../../xml/index"
import { builtInStyleType } from "../edit/formats"
import type { DocxStyles } from "../styles/index"
import { partReader } from "../model/reader"
import { Names } from "./markup"
import { W_NAMESPACE } from "./story"

// A built-in style commands add: w:pPr and w:rPr content (prefix "w:" replaced by the part's own), its
// w:uiPriority and the flags Word gives it. Paragraph styles are based on the document's default paragraph
// style (w:next it too unless `ownNext`), character styles on its default character style.
interface Definition {
	readonly priority: number
	readonly pPr: string
	readonly rPr: string
	readonly qFormat: boolean
	readonly hidden?: boolean
	readonly ownNext?: boolean
}

// The values are Word's built-in defaults (current Office theme) as observed in documents Word 16 wrote, in the
// public test corpus (corpus/manifest.json): the paragraph styles from SHA-256
// 599533082406a29a6238b5a1799c5395d884a8506feafd455fe91d4554218b62, the comment styles from
// c27159ed00e2bcaf3d15f4ff3b3df2db55e8590869aad93303984fb652d1ce86; no third-party code is copied.
const DEFINITIONS: Readonly<Record<string, Definition>> = {
	title: {
		priority: 10,
		pPr: '<w:spacing w:after="80" w:line="240" w:lineRule="auto"/><w:contextualSpacing/>',
		rPr: '<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/><w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/>',
		qFormat: true
	},
	subtitle: {
		priority: 11,
		pPr: '<w:numPr><w:ilvl w:val="1"/></w:numPr>',
		rPr: '<w:rFonts w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/><w:color w:val="595959" w:themeColor="text1" w:themeTint="A6"/><w:spacing w:val="15"/><w:sz w:val="28"/><w:szCs w:val="28"/>',
		qFormat: true
	},
	quote: {
		priority: 29,
		pPr: '<w:spacing w:before="160"/><w:jc w:val="center"/>',
		rPr: '<w:i/><w:iCs/><w:color w:val="404040" w:themeColor="text1" w:themeTint="BF"/>',
		qFormat: true
	},
	"heading 1": heading(0, '<w:spacing w:before="360" w:after="80"/>', "major", "", "0F4761", 40),
	"heading 2": heading(1, '<w:spacing w:before="160" w:after="80"/>', "major", "", "0F4761", 32),
	"heading 3": heading(2, '<w:spacing w:before="160" w:after="80"/>', "", "", "0F4761", 28),
	"heading 4": heading(3, '<w:spacing w:before="80" w:after="40"/>', "", "<w:i/><w:iCs/>", "0F4761", 0),
	"heading 5": heading(4, '<w:spacing w:before="80" w:after="40"/>', "", "", "0F4761", 0),
	"heading 6": heading(5, '<w:spacing w:before="40" w:after="0"/>', "", "<w:i/><w:iCs/>", "595959", 0),
	"heading 7": heading(6, '<w:spacing w:before="40" w:after="0"/>', "", "", "595959", 0),
	"heading 8": heading(7, '<w:spacing w:after="0"/>', "", "<w:i/><w:iCs/>", "272727", 0),
	"heading 9": heading(8, '<w:spacing w:after="0"/>', "", "", "272727", 0),
	"annotation text": {
		priority: 99,
		pPr: '<w:spacing w:line="240" w:lineRule="auto"/>',
		rPr: '<w:sz w:val="20"/><w:szCs w:val="20"/>',
		qFormat: false,
		ownNext: true
	},
	"annotation reference": { priority: 99, pPr: "", rPr: '<w:sz w:val="16"/><w:szCs w:val="16"/>', qFormat: false, hidden: true },
	header: headerFooter(),
	footer: headerFooter(),
	"footnote text": noteText(),
	"endnote text": noteText(),
	"footnote reference": noteReference(),
	"endnote reference": noteReference(),
	...Object.fromEntries(
		[1, 2, 3, 4, 5, 6, 7, 8, 9].map(level => [
			`toc ${String(level)}`,
			{
				priority: 39,
				pPr: `<w:spacing w:after="100"/>${level > 1 ? `<w:ind w:left="${String(220 * (level - 1))}"/>` : ""}`,
				rPr: "",
				qFormat: false
			}
		])
	)
}

// Header and footer (center and right tabs of a US Letter page with 1" margins), notes' text and number
// marks: as Word 16 defines them in docx4j__loadAndSave.docx (SHA-256
// 122f53311702e32affbc0ec99ced129990e6728a9b4b6df30f5684f73a20247f).
function headerFooter(): Definition {
	return {
		priority: 99,
		pPr: '<w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>',
		rPr: "",
		qFormat: false,
		ownNext: true
	}
}

function noteText(): Definition {
	return {
		priority: 99,
		pPr: '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>',
		rPr: '<w:sz w:val="20"/><w:szCs w:val="20"/>',
		qFormat: false,
		hidden: true,
		ownNext: true
	}
}

function noteReference(): Definition {
	return { priority: 99, pPr: "", rPr: '<w:vertAlign w:val="superscript"/>', qFormat: false, hidden: true }
}

function heading(level: number, spacing: string, fonts: "major" | "", toggles: string, color: string, size: number): Definition {
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

	return {
		priority: 9,
		pPr: `<w:keepNext/><w:keepLines/>${spacing}<w:outlineLvl w:val="${level}"/>`,
		rPr: `${rFonts}${toggles}<w:color w:val="${color}" ${theme}/>${size > 0 ? `<w:sz w:val="${size}"/><w:szCs w:val="${size}"/>` : ""}`,
		qFormat: true
	}
}

export function hasStyleDefinition(name: string): boolean {
	return name in DEFINITIONS
}

function withPrefix(markup: string, names: Names): string {
	return names.w === "w"
		? markup
		: markup.replace(/(<\/?)w:/g, `$1${names.w === "" ? "" : `${names.w}:`}`).replace(/ w:/g, names.w === "" ? " " : ` ${names.w}:`)
}

// The w:style element for a built-in style content uses, based on the document's default style of its type
// (`based`: their ids).
export function styleDefinition(
	names: Names,
	id: string,
	name: string,
	based: { readonly paragraph: string | undefined; readonly character: string | undefined }
): string {
	const definition = DEFINITIONS[name]

	if (definition === undefined) throw new RangeError(`No definition for ${name}`)

	const type = builtInStyleType(name)
	const { priority, pPr, rPr, qFormat, hidden, ownNext } = definition
	const basedOn = type === "paragraph" ? based.paragraph : based.character
	const pPrMarkup = type === "paragraph" && pPr !== "" ? `<${names.el("pPr")}>${withPrefix(pPr, names)}</${names.el("pPr")}>` : ""

	return (
		`<${names.el("style")} ${names.attr("type")}="${type}" ${names.attr("styleId")}="${escapeAttribute(id)}">${names.val("name", name)}` +
		(basedOn === undefined ? "" : names.val("basedOn", basedOn)) +
		(type === "paragraph" && based.paragraph !== undefined && ownNext !== true ? names.val("next", based.paragraph) : "") +
		names.val("uiPriority", String(priority)) +
		(hidden === true ? names.empty("semiHidden") : "") +
		(qFormat ? names.empty("qFormat") : names.empty("unhideWhenUsed")) +
		`${pPrMarkup}<${names.el("rPr")}>${withPrefix(rPr, names)}</${names.el("rPr")}></${names.el("style")}>`
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

	const based = { paragraph: styles.defaultOf("paragraph")?.id, character: styles.defaultOf("character")?.id }

	if (document === undefined) {
		const names = new Names("w", "")
		const body = [...missing].map(([id, name]) => styleDefinition(names, id, name, based)).join("")

		return encodeUtf8(
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:styles xmlns:w="${W_NAMESPACE}">${body}</w:styles>`
		)
	}

	const reader = partReader(document)
	const root = document.documentElement()
	const names = new Names(reader.namespaces.prefixFor(root, W_NAMESPACE) ?? "", "")
	const editor = new XmlEditor(document)

	editor.append(root, [...missing].map(([id, name]) => styleDefinition(names, id, name, based)).join(""))

	return editor.serialize()
}
