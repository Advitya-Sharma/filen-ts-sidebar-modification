import { unzipSync, zipSync } from "fflate"
import { bytesSource } from "../core/opc/index"
import { XLSX_FEATURES, openXlsxSession } from "../core/sheet/commands"
import type { SheetEvent, SheetSessionApi } from "../core/sheet/protocol"
import type { SheetFeature } from "../core/sheet/registry"
import type { XlsxDocument } from "../core/xlsx/edit/document"
import { drawingsFeature } from "../core/xlsx/drawings/feature"
import { decodeUtf8, encodeUtf8 } from "../core/xml/index"
import { IMMEDIATE, buildFixture } from "./sheetSessionFixtures"
import { settle } from "./sheetNotesFixtures"
import type { Fixture } from "./xlsxWriteFixtures"
import { ZIP_ENTRY_TIME } from "./zipTime"
import { dib24, emf } from "./metafileFixtures"
import { TYPES, relationship } from "./xlsxRefsFixtures"

// Workbooks with a drawing part (pictures, a text box, a group, a chart frame, an absolute connector) and sessions
// over them.

export const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
export const A = "http://schemas.openxmlformats.org/drawingml/2006/main"
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

// A 1×1 PNG.
export const PNG = Uint8Array.from(
	atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="),
	character => character.charCodeAt(0)
)

export const SVG = encodeUtf8(
	'<svg xmlns="http://www.w3.org/2000/svg" width="20" height="10"><rect width="20" height="10" fill="red"/></svg>'
)

function marker(name: string, col: number, colOff: number, row: number, rowOff: number): string {
	return `<xdr:${name}><xdr:col>${String(col)}</xdr:col><xdr:colOff>${String(colOff)}</xdr:colOff><xdr:row>${String(row)}</xdr:row><xdr:rowOff>${String(rowOff)}</xdr:rowOff></xdr:${name}>`
}

export const PICTURE =
	`<xdr:twoCellAnchor editAs="oneCell">${marker("from", 1, 0, 1, 0)}${marker("to", 3, 0, 5, 0)}` +
	'<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1" descr="A red dot"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>' +
	'<xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
	'<xdr:spPr><a:xfrm><a:off x="609600" y="190500"/><a:ext cx="1219200" cy="762000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>' +
	"<xdr:clientData/></xdr:twoCellAnchor>"

export const TEXT_BOX =
	`<xdr:twoCellAnchor>${marker("from", 5, 0, 1, 0)}${marker("to", 7, 0, 3, 0)}` +
	'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="3" name="TextBox 2"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>' +
	'<xdr:spPr><a:xfrm><a:off x="3048000" y="190500"/><a:ext cx="1219200" cy="381000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
	'<a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></xdr:spPr>' +
	'<xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1100"/><a:t>Hello</a:t></a:r></a:p></xdr:txBody></xdr:sp>' +
	"<xdr:clientData/></xdr:twoCellAnchor>"

export const CHART_FRAME =
	`<xdr:twoCellAnchor>${marker("from", 1, 0, 8, 0)}${marker("to", 6, 0, 18, 0)}` +
	'<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="4" name="Chart 3"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>' +
	'<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
	'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId2"/></a:graphicData></a:graphic></xdr:graphicFrame>' +
	"<xdr:clientData/></xdr:twoCellAnchor>"

export const CONNECTOR =
	'<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="0"/>' +
	'<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="5" name="Straight Connector 4"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>' +
	'<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></xdr:spPr></xdr:cxnSp>' +
	"<xdr:clientData/></xdr:absoluteAnchor>"

export function wsDr(...objects: string[]): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}" xmlns:r="${REL}">${objects.join("")}</xdr:wsDr>`
}

const CHART_PART =
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:layout/></c:plotArea></c:chart></c:chartSpace>'

// A sheet with a picture, a text box, a chart frame and a connector, and a second sheet without a drawing.
export const DRAWINGS: Fixture = {
	sheets: [
		{
			name: "One",
			body: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><drawing r:id="rId1"/>',
			rels: relationship("rId1", "drawing", "../drawings/drawing1.xml")
		},
		{ name: "Two", body: '<sheetData><row r="1"><c r="A1"><v>5</v></c></row></sheetData>' }
	],
	parts: {
		"xl/drawings/drawing1.xml": { text: wsDr(PICTURE, TEXT_BOX, CHART_FRAME, CONNECTOR), type: TYPES.drawing },
		"xl/drawings/_rels/drawing1.xml.rels": {
			text:
				'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
				relationship("rId1", "image", "../media/image1.png") +
				relationship("rId2", "chart", "../charts/chart1.xml") +
				"</Relationships>",
			type: TYPES.relationships
		},
		"xl/charts/chart1.xml": { text: CHART_PART, type: TYPES.chart }
	}
}

// The fixture with binary media added (and a Default content type for them).
export function withMedia(bytes: Uint8Array, media: Readonly<Record<string, Uint8Array>>): Uint8Array {
	const files: Record<string, Uint8Array> = unzipSync(bytes)
	const types = decodeUtf8(files["[Content_Types].xml"] ?? new Uint8Array(0))

	files["[Content_Types].xml"] = encodeUtf8(
		types.includes('Extension="png"')
			? types
			: types.replace("<Default ", '<Default Extension="png" ContentType="image/png"/><Default ')
	)

	for (const [name, data] of Object.entries(media)) files[name] = data

	return zipSync(files, { mtime: ZIP_ENTRY_TIME })
}

export const DRAWINGS_BYTES = withMedia(buildFixture(DRAWINGS), { "xl/media/image1.png": PNG })

// The feature joins the session's list once the commands module registers it.
const EXTRA: readonly SheetFeature<XlsxDocument>[] = XLSX_FEATURES.some(feature => feature.id === drawingsFeature.id)
	? []
	: [drawingsFeature]

export async function openDrawings(
	bytes: Uint8Array = DRAWINGS_BYTES
): Promise<{ session: SheetSessionApi; sheets: number[]; events: SheetEvent[] }> {
	const events: SheetEvent[] = []
	const opened = await openXlsxSession(
		bytesSource(bytes.slice()),
		{ kind: "xlsx", slicing: IMMEDIATE, settings: { author: "Jan" } },
		event => events.push(event),
		EXTRA
	)

	if (opened.kind !== "opened") throw new Error(opened.kind)

	await settle()

	return { session: opened.session, sheets: opened.info.sheets.map(sheet => sheet.id), events }
}

// The same package written as Strict OOXML (ECMA-376 Part 1 §A.1 namespaces, the purl.oclc.org relationship types).
export function strictOf(bytes: Uint8Array): Uint8Array {
	const pairs: readonly (readonly [string, string])[] = [
		["http://schemas.openxmlformats.org/spreadsheetml/2006/main", "http://purl.oclc.org/ooxml/spreadsheetml/main"],
		["http://schemas.openxmlformats.org/officeDocument/2006/relationships", "http://purl.oclc.org/ooxml/officeDocument/relationships"],
		["http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing", "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"],
		["http://schemas.openxmlformats.org/drawingml/2006/main", "http://purl.oclc.org/ooxml/drawingml/main"],
		["http://schemas.openxmlformats.org/drawingml/2006/chart", "http://purl.oclc.org/ooxml/drawingml/chart"]
	]
	const files: Record<string, Uint8Array> = unzipSync(bytes)

	for (const [name, data] of Object.entries(files)) {
		if (!name.endsWith(".xml") && !name.endsWith(".rels")) continue

		let text = decodeUtf8(data)

		for (const [from, to] of pairs) text = text.replaceAll(from, to)

		files[name] = encodeUtf8(text.replace("<workbook ", '<workbook conformance="strict" '))
	}

	return zipSync(files, { mtime: ZIP_ENTRY_TIME })
}

// An EMF whose only content is a 2 x 2 red bitmap (EMR_STRETCHDIBITS, MS-EMF §2.3.1.7), drawn over half an inch.
export function bitmapEmf(): Uint8Array {
	const { info, bits } = dib24(2, 2, 0xff0000)

	return emf()
		.record(81, writer =>
			writer
				.i32(0)
				.i32(0)
				.i32(0)
				.i32(0)
				.i32(0)
				.i32(0)
				.i32(0)
				.i32(0)
				.i32(2)
				.i32(2)
				.u32(80)
				.u32(info.length)
				.u32(80 + info.length)
				.u32(bits.length)
				.u32(0)
				.u32(0x00cc0020)
				.i32(500)
				.i32(500)
				.raw(info)
				.raw(bits)
		)
		.build()
}
