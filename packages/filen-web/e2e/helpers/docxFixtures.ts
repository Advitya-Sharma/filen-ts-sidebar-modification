import { crc32, deflateSync } from "node:zlib"
import { buildZip } from "./archiveFixtures"

// The DOCX editing specs' documents and pictures, built here byte for byte (no binary checked in): one small
// synthetic .docx per feature the editor's web UI works on (comments, tracked changes, tables, a right-to-left
// table, a table of contents, content controls, equations, headers and footers), each carrying a line of
// text a spec finds it by, and the pictures an insert leg feeds the image pipeline (one over the 25 MiB
// limit, a WebP that is converted, an SVG that gets a PNG fallback).
//
// The specs that read them run only with DOCX editing built in (VITE_DOCX_EDITING=1), so the fixture setup
// builds their scenarios only then. Specs that SAVE a document upload these bytes into their own scratch
// directory instead: the shared tree is read-only.

export const DOCX_EDITING_E2E = process.env["VITE_DOCX_EDITING"] === "1"

export const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

const NS = [
	'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
	'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
	'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"',
	'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
	'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
	'mc:Ignorable="w14"'
].join(" ")
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const CT = "application/vnd.openxmlformats-officedocument.wordprocessingml"
const DATE = "2026-01-01T00:00:00Z"
const AUTHOR = 'w:author="E2E Reviewer" w:date="2026-01-01T00:00:00Z"'

// The line of text each document is found by, by file name.
export const DOCX_FIXTURE_TEXT = {
	"w3-comments.docx": "W3 comments fixture",
	"w3-tracked.docx": "W3 tracked fixture",
	"w3-tables.docx": "W3 tables fixture",
	"w3-rtl-table.docx": "W3 right-to-left table fixture",
	"w3-toc.docx": "W3 contents fixture",
	"w3-forms.docx": "W3 forms fixture",
	"w3-equations.docx": "W3 equations fixture",
	"w3-headers.docx": "W3 headers fixture"
} as const

export type DocxFixtureName = keyof typeof DOCX_FIXTURE_TEXT

function escape(text: string): string {
	return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
}

function run(text: string): string {
	return `<w:r><w:t xml:space="preserve">${escape(text)}</w:t></w:r>`
}

function paragraph(content: string, style?: string): string {
	return `<w:p>${style === undefined ? "" : `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>`}${content}</w:p>`
}

const SECTION =
	'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'

interface DocxParts {
	body: string
	// Section properties ending the body (Letter, 1" margins by default).
	section?: string
	// Further parts of word/: file name, content type (after the wordprocessingml prefix), relationship type
	// (after the relationships namespace) and XML; each gets rId<n> in order, from rId1.
	parts?: readonly { name: string; type: string; relationship: string; xml: string }[]
}

function docx({ body, section = SECTION, parts = [] }: DocxParts): Buffer {
	const overrides = parts.map(part => `<Override PartName="/word/${part.name}" ContentType="${CT}.${part.type}+xml"/>`).join("")
	const relationships = parts
		.map((part, index) => `<Relationship Id="rId${String(index + 1)}" Type="${REL}/${part.relationship}" Target="${part.name}"/>`)
		.join("")

	return buildZip([
		{
			name: "[Content_Types].xml",
			data: `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT}.document.main+xml"/>${overrides}</Types>`
		},
		{
			name: "_rels/.rels",
			data: `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`
		},
		{ name: "word/document.xml", data: `${XML}<w:document ${NS}><w:body>${body}${section}</w:body></w:document>` },
		{
			name: "word/_rels/document.xml.rels",
			data: `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`
		},
		...parts.map(part => ({ name: `word/${part.name}`, data: `${XML}${part.xml}` }))
	])
}

const BORDER = 'w:val="single" w:sz="4" w:space="0" w:color="000000"'

function table(rtl: boolean): string {
	const cell = (text: string): string => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${paragraph(run(text))}</w:tc>`
	const row = (cells: readonly string[]): string => `<w:tr>${cells.map(cell).join("")}</w:tr>`

	return (
		`<w:tbl><w:tblPr>${rtl ? "<w:bidiVisual/>" : ""}<w:tblW w:w="0" w:type="auto"/><w:tblBorders>` +
		`<w:top ${BORDER}/><w:left ${BORDER}/><w:bottom ${BORDER}/><w:right ${BORDER}/><w:insideH ${BORDER}/><w:insideV ${BORDER}/>` +
		`</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>` +
		row(["A1", "B1", "C1"]) +
		row(["A2", "B2", "C2"]) +
		"</w:tbl>"
	)
}

function control(properties: string, text: string): string {
	return `<w:sdt><w:sdtPr>${properties}</w:sdtPr><w:sdtContent>${run(text)}</w:sdtContent></w:sdt>`
}

const BUILDERS: Readonly<Record<DocxFixtureName, () => Buffer>> = {
	"w3-comments.docx": () =>
		docx({
			body:
				paragraph(
					`<w:commentRangeStart w:id="0"/>${run(DOCX_FIXTURE_TEXT["w3-comments.docx"])}<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>`
				) + paragraph(run("A second paragraph without comments.")),
			parts: [
				{
					name: "comments.xml",
					type: "comments",
					relationship: "comments",
					xml: `<w:comments ${NS}><w:comment w:id="0" ${AUTHOR} w:initials="ER">${paragraph(run("A fixture comment."))}</w:comment></w:comments>`
				}
			]
		}),
	"w3-tracked.docx": () =>
		docx({
			body:
				paragraph(
					`${run(`${DOCX_FIXTURE_TEXT["w3-tracked.docx"]} `)}<w:ins w:id="1" ${AUTHOR}>${run("inserted ")}</w:ins><w:del w:id="2" ${AUTHOR}><w:r><w:delText>deleted</w:delText></w:r></w:del>`
				) + paragraph(run("An untracked paragraph."))
		}),
	"w3-tables.docx": () =>
		docx({ body: paragraph(run(DOCX_FIXTURE_TEXT["w3-tables.docx"])) + table(false) + paragraph(run("After the table.")) }),
	"w3-rtl-table.docx": () =>
		docx({ body: paragraph(run(DOCX_FIXTURE_TEXT["w3-rtl-table.docx"])) + table(true) + paragraph(run("After the table.")) }),
	"w3-toc.docx": () =>
		docx({
			body:
				paragraph(run(DOCX_FIXTURE_TEXT["w3-toc.docx"])) +
				paragraph(
					'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>' +
						`<w:r><w:fldChar w:fldCharType="separate"/></w:r>${run("No table of contents entries found.")}<w:r><w:fldChar w:fldCharType="end"/></w:r>`
				) +
				paragraph(run("Chapter one"), "Heading1") +
				paragraph(run("Body text of chapter one.")) +
				paragraph(run("Section one point one"), "Heading2") +
				paragraph(run("Chapter two"), "Heading1"),
			parts: [
				{
					name: "styles.xml",
					type: "styles",
					relationship: "styles",
					xml:
						`<w:styles ${NS}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
						[1, 2]
							.map(
								level =>
									`<w:style w:type="paragraph" w:styleId="Heading${String(level)}"><w:name w:val="heading ${String(level)}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:outlineLvl w:val="${String(level - 1)}"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>`
							)
							.join("") +
						"</w:styles>"
				}
			]
		}),
	"w3-forms.docx": () =>
		docx({
			body:
				paragraph(run(DOCX_FIXTURE_TEXT["w3-forms.docx"])) +
				paragraph(
					run("Agree: ") +
						control(
							'<w:id w:val="101"/><w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>',
							"☐"
						)
				) +
				paragraph(
					run("Colour: ") +
						control(
							'<w:id w:val="102"/><w:dropDownList><w:listItem w:displayText="Red" w:value="Red"/><w:listItem w:displayText="Green" w:value="Green"/></w:dropDownList>',
							"Red"
						)
				) +
				paragraph(
					run("Date: ") +
						control(
							`<w:id w:val="103"/><w:date w:fullDate="${DATE}"><w:dateFormat w:val="M/d/yyyy"/><w:lid w:val="en-US"/><w:storeMappedDataAs w:val="dateTime"/><w:calendar w:val="gregorian"/></w:date>`,
							"1/1/2026"
						)
				) +
				paragraph(run("Name: ") + control('<w:id w:val="104"/><w:text/>', "Plain text control"))
		}),
	"w3-equations.docx": () =>
		docx({
			body:
				paragraph(
					run(`${DOCX_FIXTURE_TEXT["w3-equations.docx"]}: `) +
						"<m:oMath><m:sSup><m:e><m:r><m:t>a</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup><m:r><m:t>+</m:t></m:r>" +
						"<m:sSup><m:e><m:r><m:t>b</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>"
				) +
				paragraph(
					"<m:oMathPara><m:oMath><m:r><m:t>x=</m:t></m:r><m:f><m:num><m:r><m:t>1</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath></m:oMathPara>"
				)
		}),
	"w3-headers.docx": () =>
		docx({
			body: paragraph(run(DOCX_FIXTURE_TEXT["w3-headers.docx"])),
			section:
				'<w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:footerReference w:type="default" r:id="rId2"/>' +
				SECTION.slice("<w:sectPr>".length),
			parts: [
				{
					name: "header1.xml",
					type: "header",
					relationship: "header",
					xml: `<w:hdr ${NS}>${paragraph(run("W3 header text"))}</w:hdr>`
				},
				{
					name: "footer1.xml",
					type: "footer",
					relationship: "footer",
					xml: `<w:ftr ${NS}>${paragraph(run("W3 footer text"))}</w:ftr>`
				}
			]
		})
}

export function docxFixtureBytes(name: DocxFixtureName): Buffer {
	return BUILDERS[name]()
}

function pngChunk(type: string, data: Buffer): Buffer {
	const length = Buffer.alloc(4)
	const typed = Buffer.concat([Buffer.from(type, "ascii"), data])
	const crc = Buffer.alloc(4)

	length.writeUInt32BE(data.length)
	crc.writeUInt32BE(crc32(typed))

	return Buffer.concat([length, typed, crc])
}

// A valid 3000 x 3000 RGB PNG of about 27 MB (stored, not compressed): over the 25 MiB a picture may have,
// under its 20 MP, so only its size refuses it.
export function largePictureBytes(): Buffer {
	const side = 3000
	const header = Buffer.alloc(13)

	header.writeUInt32BE(side, 0)
	header.writeUInt32BE(side, 4)
	header[8] = 8
	header[9] = 2

	return Buffer.concat([
		Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
		pngChunk("IHDR", header),
		pngChunk("IDAT", deflateSync(Buffer.alloc((side * 3 + 1) * side), { level: 0 })),
		pngChunk("IEND", Buffer.alloc(0))
	])
}

// A 1 x 1 lossless WebP (VP8L).
export function webpPictureBytes(): Buffer {
	return Buffer.from("UklGRhoAAABXRUJQVlA4TA0AAAAvAAAAEAcQERGIiP4HAA==", "base64")
}

// A 120 x 80 SVG.
export function svgPictureBytes(): Buffer {
	return Buffer.from(
		'<svg xmlns="http://www.w3.org/2000/svg" width="120" height="80" viewBox="0 0 120 80"><rect width="120" height="80" fill="#3b82f6"/></svg>',
		"utf8"
	)
}
