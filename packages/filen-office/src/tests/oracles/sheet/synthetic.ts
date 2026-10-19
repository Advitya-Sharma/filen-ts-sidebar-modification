import { zipSync } from "fflate"
import { encodeUtf8 } from "@filen/office/core/xml"

// Minimal SpreadsheetML packages for the oracles' self-tests.

export const SML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const PACKAGE_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
const TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

export interface SyntheticSheet {
	readonly name: string
	// The worksheet's children (sheetData and what follows it).
	readonly body: string
}

export interface SyntheticWorkbook {
	readonly sheets: readonly SyntheticSheet[]
	// The stylesheet's children; a default one when omitted.
	readonly styles?: string
	// app.xml's Application and AppVersion; no app.xml when omitted.
	readonly application?: { readonly name: string; readonly version: string }
	// Attributes of calcPr; none when omitted.
	readonly calcPr?: string
	readonly calcChain?: boolean
}

const DEFAULT_STYLES =
	'<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
	'<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
	'<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
	'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
	'<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'

export function buildXlsx(workbook: SyntheticWorkbook): Uint8Array {
	const overrides = [
		'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
		'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
		...workbook.sheets.map(
			(_, index) =>
				`<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
		),
		...(workbook.application === undefined
			? []
			: [
					'<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
				]),
		...(workbook.calcChain === true
			? [
					'<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>'
				]
			: [])
	]
	const files: Record<string, string> = {
		"[Content_Types].xml":
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
			'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
			`${overrides.join("")}</Types>`,
		"_rels/.rels":
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PACKAGE_RELS}">` +
			`<Relationship Id="rId1" Type="${TYPE}/officeDocument" Target="xl/workbook.xml"/>` +
			(workbook.application === undefined
				? ""
				: `<Relationship Id="rId2" Type="${TYPE}/extended-properties" Target="docProps/app.xml"/>`) +
			"</Relationships>",
		"xl/workbook.xml":
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="${SML}" xmlns:r="${R}"><sheets>` +
			workbook.sheets.map((sheet, index) => `<sheet name="${sheet.name}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join("") +
			`</sheets>${workbook.calcPr === undefined ? "" : `<calcPr ${workbook.calcPr}/>`}</workbook>`,
		"xl/_rels/workbook.xml.rels":
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PACKAGE_RELS}">` +
			workbook.sheets
				.map(
					(_, index) => `<Relationship Id="rId${index + 1}" Type="${TYPE}/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`
				)
				.join("") +
			`<Relationship Id="rIdStyles" Type="${TYPE}/styles" Target="styles.xml"/>` +
			(workbook.calcChain === true ? `<Relationship Id="rIdCalc" Type="${TYPE}/calcChain" Target="calcChain.xml"/>` : "") +
			"</Relationships>",
		"xl/styles.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="${SML}">${workbook.styles ?? DEFAULT_STYLES}</styleSheet>`
	}

	workbook.sheets.forEach((sheet, index) => {
		files[`xl/worksheets/sheet${index + 1}.xml`] =
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="${SML}" xmlns:r="${R}">${sheet.body}</worksheet>`
	})

	if (workbook.application !== undefined)
		files["docProps/app.xml"] =
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">' +
			`<Application>${workbook.application.name}</Application><AppVersion>${workbook.application.version}</AppVersion></Properties>`

	if (workbook.calcChain === true)
		files["xl/calcChain.xml"] =
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<calcChain xmlns="${SML}"><c r="A1" i="1"/></calcChain>`

	return zipSync(Object.fromEntries(Object.entries(files).map(([name, text]) => [name, encodeUtf8(text)])))
}
