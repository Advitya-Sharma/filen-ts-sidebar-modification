import { encodeUtf8, parseXml, type XmlDocument } from "../core/xml/index"

// Builders for small WordprocessingML parts in tests.

export const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
export const W_STRICT = "http://purl.oclc.org/ooxml/wordprocessingml/main"
export const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
export const A = "http://schemas.openxmlformats.org/drawingml/2006/main"

export function xml(source: string): XmlDocument {
	return parseXml(encodeUtf8(source))
}

// A part whose root is `root` with the usual w: and r: bindings.
export function wordPart(root: string, body: string, w: string = W, prefix: string = "w"): XmlDocument {
	return xml(
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><${prefix}:${root} xmlns:${prefix}="${w}" xmlns:r="${R}">${body}</${prefix}:${root}>`
	)
}

// The first element with this qualified name, in document order.
export function find(document: XmlDocument, qualifiedName: string, from: number = 0): number {
	for (let node = from; node < document.count; node++) {
		if (document.name(node) === qualifiedName) return node
	}

	return -1
}

export function findAll(document: XmlDocument, qualifiedName: string): number[] {
	const out: number[] = []

	for (let node = 0; node < document.count; node++) {
		if (document.name(node) === qualifiedName) out.push(node)
	}

	return out
}

export function stylesXml(body: string, defaults: string = ""): XmlDocument {
	return wordPart("styles", `${defaults === "" ? "" : `<w:docDefaults>${defaults}</w:docDefaults>`}${body}`)
}

export function themeXml(colors: string, fonts: string): XmlDocument {
	return xml(
		`<a:theme xmlns:a="${A}" name="Test"><a:themeElements><a:clrScheme name="Test">${colors}</a:clrScheme>` +
			`<a:fontScheme name="Test">${fonts}</a:fontScheme></a:themeElements></a:theme>`
	)
}

// The Office 2013-2022 colour scheme and font scheme, as Word writes them.
export const OFFICE_THEME = themeXml(
	'<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
		'<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>' +
		'<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
		'<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>' +
		'<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>' +
		'<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>',
	'<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/>' +
		'<a:font script="Jpan" typeface="游ゴシック Light"/><a:font script="Hans" typeface="等线 Light"/><a:font script="Arab" typeface="Times New Roman"/></a:majorFont>' +
		'<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/>' +
		'<a:font script="Jpan" typeface="游明朝"/><a:font script="Hans" typeface="等线"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/></a:minorFont>'
)
