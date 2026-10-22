import { unzipSync, zipSync } from "fflate"
import { bytesSource } from "../core/opc/index"
import { XLSX_FEATURES, openXlsxSession } from "../core/sheet/commands"
import type { SheetEvent, SheetSessionApi } from "../core/sheet/protocol"
import type { SheetFeature } from "../core/sheet/registry"
import { chartsFeature } from "../core/xlsx/charts/feature"
import type { XlsxDocument } from "../core/xlsx/edit/document"
import { drawingsFeature } from "../core/xlsx/drawings/feature"
import { decodeUtf8, encodeUtf8 } from "../core/xml/index"
import { IMMEDIATE, buildFixture } from "./sheetSessionFixtures"
import { settle } from "./sheetNotesFixtures"
import { CHART_FRAME, wsDr } from "./sheetDrawingsFixtures"
import type { Fixture } from "./xlsxWriteFixtures"
import { TYPES, relationship } from "./xlsxRefsFixtures"

// Workbooks with charts reading their cells, and sessions over them.

export const C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
export const A = "http://schemas.openxmlformats.org/drawingml/2006/main"

function pts(values: readonly (string | number)[]): string {
	return values.map((value, idx) => `<c:pt idx="${String(idx)}"><c:v>${String(value)}</c:v></c:pt>`).join("")
}

export function strCache(values: readonly string[]): string {
	return `<c:strCache><c:ptCount val="${String(values.length)}"/>${pts(values)}</c:strCache>`
}

export function numCache(values: readonly number[], format = "General"): string {
	return `<c:numCache><c:formatCode>${format}</c:formatCode><c:ptCount val="${String(values.length)}"/>${pts(values)}</c:numCache>`
}

export interface SeriesSpec {
	readonly name: string
	readonly nameCache?: string
	readonly cat: string
	readonly catCache?: readonly string[]
	readonly val: string
	readonly valCache?: readonly number[]
	readonly extra?: string
}

export function series(spec: SeriesSpec, index: number): string {
	return (
		`<c:ser><c:idx val="${String(index)}"/><c:order val="${String(index)}"/>` +
		`<c:tx><c:strRef><c:f>${spec.name}</c:f>${spec.nameCache === undefined ? "" : strCache([spec.nameCache])}</c:strRef></c:tx>` +
		'<c:invertIfNegative val="0"/>' +
		`<c:cat><c:strRef><c:f>${spec.cat}</c:f>${spec.catCache === undefined ? "" : strCache(spec.catCache)}</c:strRef></c:cat>` +
		`<c:val><c:numRef><c:f>${spec.val}</c:f>${spec.valCache === undefined ? "" : numCache(spec.valCache)}</c:numRef></c:val>` +
		`${spec.extra ?? ""}</c:ser>`
	)
}

export function barChart(specs: readonly SeriesSpec[], options: { readonly title?: string; readonly plotExtra?: string } = {}): string {
	const title =
		options.title === undefined
			? '<c:autoTitleDeleted val="1"/>'
			: `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr b="1"/></a:pPr><a:r><a:rPr lang="en-GB" b="1"/><a:t>${options.title}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`

	return (
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
		`<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>` +
		`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${specs.map(series).join("")}` +
		`<c:gapWidth val="150"/><c:axId val="111"/><c:axId val="222"/>${options.plotExtra ?? ""}</c:barChart>` +
		'<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="222"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>' +
		'<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="111"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>' +
		'</c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>'
	)
}

// Data!A1:C4 (a header row; fruit names; two columns of numbers) and a column chart of B and C on Data, plus a second
// sheet Other with one number.
export const DATA_BODY =
	'<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Fruit</t></is></c><c r="B1" t="inlineStr"><is><t>Sold</t></is></c><c r="C1" t="inlineStr"><is><t>Kept</t></is></c></row>' +
	'<row r="2"><c r="A2" t="inlineStr"><is><t>Apple</t></is></c><c r="B2"><v>3</v></c><c r="C2"><v>1</v></c></row>' +
	'<row r="3"><c r="A3" t="inlineStr"><is><t>Pear</t></is></c><c r="B3"><v>5</v></c><c r="C3"><v>2</v></c></row>' +
	'<row r="4"><c r="A4" t="inlineStr"><is><t>Plum</t></is></c><c r="B4"><v>7</v></c><c r="C4"><v>4</v></c></row></sheetData>'

export const SOLD: SeriesSpec = {
	name: "Data!$B$1",
	nameCache: "Sold",
	cat: "Data!$A$2:$A$4",
	catCache: ["Apple", "Pear", "Plum"],
	val: "Data!$B$2:$B$4",
	valCache: [3, 5, 7]
}

export const KEPT: SeriesSpec = {
	name: "Data!$C$1",
	nameCache: "Kept",
	cat: "Data!$A$2:$A$4",
	catCache: ["Apple", "Pear", "Plum"],
	val: "Data!$C$2:$C$4",
	valCache: [1, 2, 4]
}

export function chartFixture(chart: string, extraParts: Fixture["parts"] = {}, drawingRels = ""): Fixture {
	return {
		sheets: [
			{ name: "Data", body: `${DATA_BODY}<drawing r:id="rId1"/>`, rels: relationship("rId1", "drawing", "../drawings/drawing1.xml") },
			{ name: "Other", body: '<sheetData><row r="1"><c r="A1"><v>9</v></c></row></sheetData>' }
		],
		parts: {
			"xl/drawings/drawing1.xml": { text: wsDr(CHART_FRAME), type: TYPES.drawing },
			"xl/drawings/_rels/drawing1.xml.rels": {
				text:
					'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
					relationship("rId2", "chart", "../charts/chart1.xml") +
					drawingRels +
					"</Relationships>",
				type: TYPES.relationships
			},
			"xl/charts/chart1.xml": { text: chart, type: TYPES.chart },
			...extraParts
		}
	}
}

export const CHART_PART = "xl/charts/chart1.xml"

// The features join the session's list once the commands module registers them.
const EXTRA: readonly SheetFeature<XlsxDocument>[] = [drawingsFeature, chartsFeature].filter(
	feature => !XLSX_FEATURES.some(registered => registered.id === feature.id)
)

export async function openCharts(
	fixture: Fixture | Uint8Array
): Promise<{ session: SheetSessionApi; sheets: number[]; events: SheetEvent[]; bytes: Uint8Array }> {
	const bytes = fixture instanceof Uint8Array ? fixture : buildFixture(fixture)
	const events: SheetEvent[] = []
	const opened = await openXlsxSession(
		bytesSource(bytes.slice()),
		{ kind: "xlsx", slicing: IMMEDIATE, settings: { author: "Jan" } },
		event => events.push(event),
		EXTRA
	)

	if (opened.kind !== "opened") throw new Error(opened.kind)

	await settle()

	return { session: opened.session, sheets: opened.info.sheets.map(sheet => sheet.id), events, bytes }
}

export function partText(zip: Uint8Array, name: string): string {
	const data = unzipSync(zip)[name]

	return data === undefined ? "" : decodeUtf8(data)
}

export const SHEET_CHART_PART = "xl/charts/chart9.xml"

// `bytes` with a chartsheet "Graph" (after the other sheets) showing a chart of its own, 6 x 4 inches.
export function withChartsheet(bytes: Uint8Array, chart: string): Uint8Array {
	const files: Record<string, Uint8Array> = unzipSync(bytes)
	const text = (name: string): string => decodeUtf8(files[name] ?? new Uint8Array(0))
	const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	const rels = (body: string): string =>
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`

	files["xl/workbook.xml"] = encodeUtf8(
		text("xl/workbook.xml").replace("</sheets>", '<sheet name="Graph" sheetId="9" r:id="rId90"/></sheets>')
	)
	files["xl/_rels/workbook.xml.rels"] = encodeUtf8(
		text("xl/_rels/workbook.xml.rels").replace(
			"</Relationships>",
			`<Relationship Id="rId90" Type="${R}/chartsheet" Target="chartsheets/sheet9.xml"/></Relationships>`
		)
	)
	files["[Content_Types].xml"] = encodeUtf8(
		text("[Content_Types].xml").replace(
			"</Types>",
			'<Override PartName="/xl/chartsheets/sheet9.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>' +
				`<Override PartName="/xl/drawings/drawing9.xml" ContentType="${TYPES.drawing}"/><Override PartName="/${SHEET_CHART_PART}" ContentType="${TYPES.chart}"/></Types>`
		)
	)
	files["xl/chartsheets/sheet9.xml"] = encodeUtf8(
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}"><sheetPr/><sheetViews><sheetView workbookViewId="0"/></sheetViews><drawing r:id="rId1"/></chartsheet>`
	)
	files["xl/chartsheets/_rels/sheet9.xml.rels"] = encodeUtf8(
		rels(`<Relationship Id="rId1" Type="${R}/drawing" Target="../drawings/drawing9.xml"/>`)
	)
	files["xl/drawings/drawing9.xml"] = encodeUtf8(
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="${A}"><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="5486400" cy="3657600"/>` +
			'<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
			`<a:graphic><a:graphicData uri="${C}"><c:chart xmlns:c="${C}" xmlns:r="${R}" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>`
	)
	files["xl/drawings/_rels/drawing9.xml.rels"] = encodeUtf8(
		rels(`<Relationship Id="rId1" Type="${R}/chart" Target="../charts/chart9.xml"/>`)
	)
	files[SHEET_CHART_PART] = encodeUtf8(chart)

	return zipSync(files)
}
