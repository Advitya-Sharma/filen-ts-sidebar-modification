import { zipSync } from "fflate"
import { encodeUtf8 } from "@filen/office/core/xml"
import { ZIP_ENTRY_TIME } from "./zipTime"

// Small workbook packages for the reader's tests: every part given as text (or bytes), the rest defaulted.

export const SML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
export const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const PACKAGE_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
const TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

export interface TestSheet {
	readonly name?: string
	// The worksheet part as a whole, or its root's content (wrapped in <worksheet>).
	readonly xml?: string
	readonly body?: string
	readonly bytes?: Uint8Array
	readonly state?: string
	// The relationship type's last segment (worksheet, chartsheet, ...).
	readonly type?: string
}

export interface TestWorkbook {
	readonly sheets: readonly TestSheet[]
	readonly sharedStrings?: readonly string[] | string
	readonly styles?: string
	// workbook.xml's children after <sheets>.
	readonly workbookTail?: string
	readonly workbookHead?: string
	readonly contentType?: string
	readonly extra?: Readonly<Record<string, string | Uint8Array>>
	readonly workbookRelationships?: string
}

export const DEFAULT_STYLES =
	'<numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>' +
	'<fonts count="2"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>' +
	'<font><b/><i val="0"/><u/><sz val="14"/><color rgb="FFFF0000"/><name val="Arial"/></font></fonts>' +
	'<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
	'<fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/><bgColor indexed="64"/></patternFill></fill></fills>' +
	'<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
	'<border diagonalUp="1"><left style="thin"><color auto="1"/></left><right style="medium"/><top/><bottom style="double"><color theme="4" tint="-0.25"/></bottom><diagonal style="hair"/></border></borders>' +
	'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
	'<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
	'<xf numFmtId="164" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" quotePrefix="1"><alignment horizontal="center" vertical="top" wrapText="1" indent="2" textRotation="45"/><protection locked="0"/></xf>' +
	'<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
	'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
	'<dxfs count="1"><dxf><font><b/></font><fill><patternFill><bgColor rgb="FFFFFF00"/></patternFill></fill></dxf></dxfs>'

const CONTENT_TYPES: Readonly<Record<string, string>> = {
	worksheet: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
	chartsheet: "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
	dialogsheet: "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml",
	xlMacrosheet: "application/vnd.ms-excel.macrosheet+xml"
}

function sharedStringsXml(strings: readonly string[] | string): string {
	if (typeof strings === "string") return strings

	return (
		`${DECLARATION}<sst xmlns="${SML}" count="${strings.length}" uniqueCount="${strings.length}">` +
		strings.map(text => (text.startsWith("<") ? `<si>${text}</si>` : `<si><t xml:space="preserve">${text}</t></si>`)).join("") +
		"</sst>"
	)
}

export function workbookFiles(workbook: TestWorkbook): Record<string, Uint8Array> {
	const sheets = workbook.sheets
	const overrides = [
		`<Override PartName="/xl/workbook.xml" ContentType="${workbook.contentType ?? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"}"/>`,
		'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
		'<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>',
		...sheets.map(
			(sheet, index) =>
				`<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="${CONTENT_TYPES[sheet.type ?? "worksheet"] ?? ""}"/>`
		)
	]
	const files: Record<string, string | Uint8Array> = {
		"[Content_Types].xml":
			`${DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
			'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
			`${overrides.join("")}</Types>`,
		"_rels/.rels": `${DECLARATION}<Relationships xmlns="${PACKAGE_RELS}"><Relationship Id="rId1" Type="${TYPE}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
		"xl/workbook.xml":
			`${DECLARATION}<workbook xmlns="${SML}" xmlns:r="${R}">${workbook.workbookHead ?? ""}<sheets>` +
			sheets
				.map(
					(sheet, index) =>
						`<sheet name="${sheet.name ?? `Sheet${index + 1}`}" sheetId="${index + 1}"${sheet.state === undefined ? "" : ` state="${sheet.state}"`} r:id="rId${index + 1}"/>`
				)
				.join("") +
			`</sheets>${workbook.workbookTail ?? ""}</workbook>`,
		"xl/_rels/workbook.xml.rels":
			`${DECLARATION}<Relationships xmlns="${PACKAGE_RELS}">` +
			sheets
				.map(
					(sheet, index) =>
						`<Relationship Id="rId${index + 1}" Type="${TYPE}/${sheet.type ?? "worksheet"}" Target="worksheets/sheet${index + 1}.xml"/>`
				)
				.join("") +
			'<Relationship Id="rIdStyles" Type="' +
			TYPE +
			'/styles" Target="styles.xml"/>' +
			(workbook.sharedStrings === undefined
				? ""
				: `<Relationship Id="rIdStrings" Type="${TYPE}/sharedStrings" Target="sharedStrings.xml"/>`) +
			(workbook.workbookRelationships ?? "") +
			"</Relationships>",
		"xl/styles.xml": `${DECLARATION}<styleSheet xmlns="${SML}">${workbook.styles ?? DEFAULT_STYLES}</styleSheet>`
	}

	if (workbook.sharedStrings !== undefined) files["xl/sharedStrings.xml"] = sharedStringsXml(workbook.sharedStrings)

	sheets.forEach((sheet, index) => {
		files[`xl/worksheets/sheet${index + 1}.xml`] =
			sheet.bytes ?? sheet.xml ?? `${DECLARATION}<worksheet xmlns="${SML}" xmlns:r="${R}">${sheet.body ?? "<sheetData/>"}</worksheet>`
	})

	Object.assign(files, workbook.extra ?? {})

	return Object.fromEntries(
		Object.entries(files).map(([name, content]) => [name, typeof content === "string" ? encodeUtf8(content) : content])
	)
}

export function buildWorkbook(workbook: TestWorkbook): Uint8Array {
	return zipSync(workbookFiles(workbook), { mtime: ZIP_ENTRY_TIME })
}
