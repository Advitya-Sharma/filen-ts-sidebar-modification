import { unzipSync, zipSync } from "fflate"
import { OpcPackage, bytesSource } from "../core/opc/index"
import { decodeUtf8, encodeUtf8 } from "../core/xml/index"
import { openXlsx, type XlsxWorkbook } from "../core/xlsx/read/index"
import { saveXlsx, type XlsxSaveOptions, type XlsxSaved } from "../core/xlsx/write/save"
import { ZIP_ENTRY_TIME } from "./zipTime"

// Workbooks for the splice writer's tests: packages built from part text, opened with the model's reader, saved
// and compared entry by entry.

export const SML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
export const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const PACKAGE_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

export const STYLES =
	'<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
	'<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
	'<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
	'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
	'<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>' +
	'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'

export interface FixtureSheet {
	readonly name: string
	// The whole part, or the worksheet's children.
	readonly xml?: string
	readonly body?: string
	readonly rels?: string
}

export interface Fixture {
	readonly sheets: readonly FixtureSheet[]
	// sharedStrings.xml's text (no part when omitted).
	readonly sharedStrings?: string
	// styles.xml's children (no part when null).
	readonly styles?: string | null
	// workbook.xml's children after <sheets>.
	readonly workbookTail?: string
	readonly calcChain?: boolean
	// More parts, by name, with their content type.
	readonly parts?: Readonly<Record<string, { readonly text: string; readonly type: string }>>
}

export function buildFixture(fixture: Fixture): Uint8Array {
	const overrides: string[] = [
		'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
	]
	const relationships: string[] = []
	const files: Record<string, string> = {}

	fixture.sheets.forEach((sheet, index) => {
		const name = `xl/worksheets/sheet${index + 1}.xml`

		overrides.push(
			`<Override PartName="/${name}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
		)
		relationships.push(`<Relationship Id="rId${index + 1}" Type="${R}/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`)
		files[name] = sheet.xml ?? `${DECLARATION}<worksheet xmlns="${SML}" xmlns:r="${R}">${sheet.body ?? "<sheetData/>"}</worksheet>`

		if (sheet.rels !== undefined)
			files[`xl/worksheets/_rels/sheet${index + 1}.xml.rels`] =
				`${DECLARATION}<Relationships xmlns="${PACKAGE_RELS}">${sheet.rels}</Relationships>`
	})

	if (fixture.styles !== null) {
		overrides.push(
			'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
		)
		relationships.push(`<Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/>`)
		files["xl/styles.xml"] = `${DECLARATION}<styleSheet xmlns="${SML}">${fixture.styles ?? STYLES}</styleSheet>`
	}

	if (fixture.sharedStrings !== undefined) {
		overrides.push(
			'<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
		)
		relationships.push(`<Relationship Id="rIdStrings" Type="${R}/sharedStrings" Target="sharedStrings.xml"/>`)
		files["xl/sharedStrings.xml"] = fixture.sharedStrings
	}

	if (fixture.calcChain === true) {
		overrides.push(
			'<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>'
		)
		relationships.push(`<Relationship Id="rIdCalc" Type="${R}/calcChain" Target="calcChain.xml"/>`)
		files["xl/calcChain.xml"] = `${DECLARATION}<calcChain xmlns="${SML}"><c r="A1" i="1"/></calcChain>`
	}

	for (const [name, part] of Object.entries(fixture.parts ?? {})) {
		overrides.push(`<Override PartName="/${name}" ContentType="${part.type}"/>`)
		files[name] = part.text
	}

	files["[Content_Types].xml"] =
		`${DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
		'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
		`${overrides.join("")}</Types>`
	files["_rels/.rels"] =
		`${DECLARATION}<Relationships xmlns="${PACKAGE_RELS}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`
	files["xl/workbook.xml"] =
		`${DECLARATION}<workbook xmlns="${SML}" xmlns:r="${R}"><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="100" windowHeight="100"/></bookViews><sheets>` +
		fixture.sheets.map((sheet, index) => `<sheet name="${sheet.name}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join("") +
		`</sheets>${fixture.workbookTail ?? ""}</workbook>`
	files["xl/_rels/workbook.xml.rels"] = `${DECLARATION}<Relationships xmlns="${PACKAGE_RELS}">${relationships.join("")}</Relationships>`

	return zipSync(Object.fromEntries(Object.entries(files).map(([name, text]) => [name, encodeUtf8(text)])), { mtime: ZIP_ENTRY_TIME })
}

export async function openWorkbook(bytes: Uint8Array): Promise<XlsxWorkbook> {
	const opened = await openXlsx(bytesSource(bytes))

	if (!opened.ok) throw new Error(`${opened.reason}: ${opened.message}`)

	return opened.workbook
}

export async function saved(workbook: XlsxWorkbook, options?: XlsxSaveOptions): Promise<Uint8Array> {
	const result = await saveXlsx({ opc: workbook.opc, workbook }, options)

	if (!result.ok) throw new Error(`Refused (${result.reason}): ${result.message}`)

	return (result as XlsxSaved).bytes as Uint8Array
}

export function parts(bytes: Uint8Array): Record<string, string> {
	return Object.fromEntries(Object.entries(unzipSync(bytes)).map(([name, data]) => [name, decodeUtf8(data)]))
}

// Entry names whose stored (compressed) bytes differ between two archives, and those only one has.
export async function changedEntries(
	before: Uint8Array,
	after: Uint8Array
): Promise<{ changed: string[]; added: string[]; removed: string[] }> {
	const a = await OpcPackage.open(bytesSource(before))
	const b = await OpcPackage.open(bytesSource(after))
	const raw = async (pkg: OpcPackage, name: string): Promise<Uint8Array> => {
		const entry = pkg.zip.entries.find(candidate => candidate.name === name)

		if (entry === undefined) return new Uint8Array(0)

		const layout = await pkg.zip.layout(entry)

		return pkg.zip.source.read(layout.dataOffset, entry.compressedSize)
	}
	const namesA = new Set(a.zip.entries.map(entry => entry.name))
	const namesB = new Set(b.zip.entries.map(entry => entry.name))
	const changed: string[] = []

	for (const name of namesA) {
		if (!namesB.has(name)) continue

		const left = await raw(a, name)
		const right = await raw(b, name)

		if (left.length !== right.length || left.some((byte, k) => byte !== right[k])) changed.push(name)
	}

	return {
		changed: changed.sort(),
		added: [...namesB].filter(name => !namesA.has(name)).sort(),
		removed: [...namesA].filter(name => !namesB.has(name)).sort()
	}
}

// The differing middle of two texts: the source's [from, to) replaced by the output's text there.
export function difference(before: string, after: string): { readonly from: number; readonly to: number; readonly inserted: string } {
	let start = 0

	while (start < before.length && start < after.length && before[start] === after[start]) start++

	let end = 0

	while (end < before.length - start && end < after.length - start && before[before.length - 1 - end] === after[after.length - 1 - end])
		end++

	return { from: start, to: before.length - end, inserted: after.slice(start, after.length - end) }
}

export function equalBytes(a: Uint8Array, b: Uint8Array): boolean {
	return a.length === b.length && a.every((byte, k) => byte === b[k])
}
