import { strToU8, zipSync, type Zippable } from "fflate"
import { bytesSource } from "@filen/office/core/opc"
import { buildDocxModel, openDocxPackage, type DocxModel, type DocxPackage } from "@filen/office/core/docx"
import { ZIP_ENTRY_TIME } from "./zipTime"

// Minimal word-processing packages assembled from XML snippets, for tests of single constructs.

const T = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const S = "http://purl.oclc.org/ooxml/officeDocument/relationships"
const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml"

export const NAMESPACES =
	'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
	'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
	'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
	'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" ' +
	'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
	'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
	'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
	'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" ' +
	'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
	'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
	'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" ' +
	'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
	'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" ' +
	'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
	'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
	'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
	'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" ' +
	'xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" ' +
	'mc:Ignorable="w14 w15 wp14"'

export interface ExtraPart {
	readonly name: string
	readonly contentType?: string
	readonly data: string | Uint8Array
}

export interface ExtraRelationship {
	readonly id: string
	// A relationship kind ("image", "hyperlink", "chart"...) or a full type URI.
	readonly type: string
	readonly target: string
	readonly external?: boolean
}

export interface SyntheticDocx {
	// Inner XML of w:body.
	readonly body?: string
	// Full XML of word/document.xml; overrides `body`.
	readonly document?: string
	// rId -> inner XML of a w:hdr / w:ftr part (named header<n>.xml / footer<n>.xml in order).
	readonly headers?: Record<string, string>
	readonly footers?: Record<string, string>
	// Inner XML of w:footnotes / w:endnotes / w:comments.
	readonly footnotes?: string
	readonly endnotes?: string
	readonly comments?: string
	// Full XML of the comment extension parts.
	readonly commentsExtended?: string
	readonly commentsIds?: string
	readonly commentsExtensible?: string
	readonly people?: string
	// Inner XML of w:settings.
	readonly settings?: string
	readonly relationships?: readonly ExtraRelationship[]
	readonly packageRelationships?: readonly ExtraRelationship[]
	readonly parts?: readonly ExtraPart[]
	readonly mainContentType?: string
	readonly strict?: boolean
}

function relationshipType(type: string, strict: boolean): string {
	if (type.includes("://")) return type

	switch (type) {
		case "commentsExtended":
		case "people":
			return `http://schemas.microsoft.com/office/2011/relationships/${type}`
		case "commentsIds":
			return `http://schemas.microsoft.com/office/2016/09/relationships/${type}`
		case "commentsExtensible":
			return `http://schemas.microsoft.com/office/2018/08/relationships/${type}`
		case "vbaProject":
			return "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
		default:
			return `${strict ? S : T}/${type}`
	}
}

function relationshipsXml(relationships: readonly ExtraRelationship[], strict: boolean): string {
	return (
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
		relationships
			.map(
				relationship =>
					`<Relationship Id="${relationship.id}" Type="${relationshipType(relationship.type, strict)}" Target="${relationship.target}"${relationship.external === true ? ' TargetMode="External"' : ""}/>`
			)
			.join("") +
		"</Relationships>"
	)
}

export function buildDocx(spec: SyntheticDocx): Uint8Array {
	const strict = spec.strict === true
	const ns = strict
		? NAMESPACES.replace(
				"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
				"http://purl.oclc.org/ooxml/wordprocessingml/main"
			).replace(`xmlns:r="${T}"`, `xmlns:r="${S}"`)
		: NAMESPACES
	const files: Zippable = {}
	const overrides: string[] = []
	const relationships: ExtraRelationship[] = [...(spec.relationships ?? [])]
	const add = (name: string, contentType: string | undefined, data: string | Uint8Array): void => {
		files[name] = typeof data === "string" ? strToU8(data) : data

		if (contentType !== undefined) overrides.push(`<Override PartName="/${name}" ContentType="${contentType}"/>`)
	}
	const xml = (root: string, inner: string): string =>
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${root} ${ns}>${inner}</w:${root}>`

	add(
		"word/document.xml",
		spec.mainContentType ?? `${WML}.document.main+xml`,
		spec.document ?? xml("document", `<w:body>${spec.body ?? ""}</w:body>`)
	)

	let headerIndex = 0

	for (const [id, inner] of Object.entries(spec.headers ?? {})) {
		const name = `header${++headerIndex}.xml`

		add(`word/${name}`, `${WML}.header+xml`, xml("hdr", inner))
		relationships.push({ id, type: "header", target: name })
	}

	let footerIndex = 0

	for (const [id, inner] of Object.entries(spec.footers ?? {})) {
		const name = `footer${++footerIndex}.xml`

		add(`word/${name}`, `${WML}.footer+xml`, xml("ftr", inner))
		relationships.push({ id, type: "footer", target: name })
	}

	const optional: [keyof SyntheticDocx, string, string, string | undefined][] = [
		["footnotes", "footnotes", "footnotes.xml", "footnotes"],
		["endnotes", "endnotes", "endnotes.xml", "endnotes"],
		["comments", "comments", "comments.xml", "comments"],
		["settings", "settings", "settings.xml", "settings"],
		["commentsExtended", "commentsExtended", "commentsExtended.xml", undefined],
		["commentsIds", "commentsIds", "commentsIds.xml", undefined],
		["commentsExtensible", "commentsExtensible", "commentsExtensible.xml", undefined],
		["people", "people", "people.xml", undefined]
	]
	let relationshipIndex = 900

	for (const [key, type, name, root] of optional) {
		const value = spec[key]

		if (typeof value !== "string") continue

		add(`word/${name}`, `${WML}.${type}+xml`, root === undefined ? value : xml(root, value))
		relationships.push({ id: `rId${relationshipIndex++}`, type, target: name })
	}

	for (const part of spec.parts ?? []) add(part.name, part.contentType, part.data)

	files["word/_rels/document.xml.rels"] = strToU8(relationshipsXml(relationships, strict))
	files["_rels/.rels"] = strToU8(
		relationshipsXml(
			[{ id: "rId1", type: "officeDocument", target: "word/document.xml" }, ...(spec.packageRelationships ?? [])],
			strict
		)
	)
	files["[Content_Types].xml"] = strToU8(
		'<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
			'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
			'<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>' +
			overrides.join("") +
			"</Types>"
	)

	return zipSync(files, { level: 1, mtime: ZIP_ENTRY_TIME })
}

export async function openSynthetic(spec: SyntheticDocx): Promise<DocxPackage> {
	const opened = await openDocxPackage(bytesSource(buildDocx(spec)))

	if (!opened.ok) throw new Error(`${opened.reason}: ${opened.message}`)

	return opened.package
}

export async function modelOf(spec: SyntheticDocx): Promise<DocxModel> {
	return buildDocxModel(await openSynthetic(spec))
}

// A paragraph of one run with plain text.
export function p(text: string, pPr: string = ""): string {
	return `<w:p>${pPr === "" ? "" : `<w:pPr>${pPr}</w:pPr>`}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
}
