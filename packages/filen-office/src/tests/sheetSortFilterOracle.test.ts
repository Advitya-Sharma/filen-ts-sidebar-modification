import { mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { afterAll, beforeAll, describe, expect, it } from "vitest"
import { bytesSource } from "../core/opc/index"
import type { SheetCommand, SheetSessionApi } from "../core/sheet/protocol"
import { readAutoFilter } from "../core/sheet/filter/autoFilter"
import { RowFlag } from "../core/xlsx/model/types"
import { openXlsx } from "../core/xlsx/read/index"
import { decodeUtf8, encodeUtf8 } from "../core/xml/index"
import { sheetOffices, unoOffice } from "./oracles/sheet/libreOffice"
import { UnoSheets, type UnoCell, type UnoJob, type UnoResult } from "./oracles/sheet/unoSheets"
import { openFixture, savedBytes } from "./sheetSessionFixtures"
import { buildFixture, parts, type Fixture } from "./xlsxWriteFixtures"
import { newValidityErrors, XsdOracle, xsdOracleMissing } from "./oracles/xsd"
import { TYPES, relationship } from "./xlsxRefsFixtures"

// Sort and autofilter against LibreOffice (UNO, LibreOffice's bundled Python). A sort: LibreOffice sorts the source
// (XSortable.sort) and loads our sorted save; both dumps must hold the same cells (value, shown text, formula). A
// filter: LibreOffice loads our filtered save and stores it again; its re-save must hide the same rows and read back
// the same criteria. Where LibreOffice orders values otherwise than Excel documents ("Sort data in a workbook in the
// browser", support.microsoft.com: logical values after text, errors after logical values; LibreOffice keeps
// booleans as numbers and puts errors before text), the difference is listed and Excel's order is the one asserted.
// Opt-in: FILEN_OFFICE_SOFFICE=<soffice>.

const UNO = unoOffice(sheetOffices())
const SCRATCH = join(tmpdir(), "filen-office-sort-filter-lo")

const STRINGS = (values: readonly string[]): string =>
	`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${values.length}" uniqueCount="${values.length}">${values.map(value => `<si><t>${value}</t></si>`).join("")}</sst>`

// Twelve data rows under a header: names (case and duplicates), scores (negatives, decimals, a blank), a group key
// for a second sort key, and per row a formula reading its own row (relative) and the header row (absolute).
const WORDS = [
	"Name",
	"Score",
	"Group",
	"Calc",
	"pear",
	"Apple",
	"fig",
	"apple",
	"Banana",
	"kiwi",
	"cherry",
	"date",
	"Elder",
	"grape",
	"lime",
	"Mango"
]
const DATA: readonly (readonly [number, number | null, number])[] = [
	[4, 12.5, 2],
	[5, -3, 1],
	[6, 40, 2],
	[7, 12.5, 1],
	[8, null, 3],
	[9, 7, 1],
	[10, 0, 3],
	[11, 99, 2],
	[12, -3, 2],
	[13, 7.25, 3],
	[14, 18, 1],
	[15, 3, 3]
]

function sortFixture(): Fixture {
	let rows =
		'<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c><c r="D1" t="s"><v>3</v></c></row>'

	DATA.forEach(([word, score, group], k) => {
		const r = k + 2
		const scoreCell = score === null ? "" : `<c r="B${r}"${k % 3 === 0 ? ' s="1"' : ""}><v>${score}</v></c>`

		rows += `<row r="${r}"><c r="A${r}" t="s"><v>${word}</v></c>${scoreCell}<c r="C${r}"><v>${group}</v></c><c r="D${r}"><f>B${r}*2+C$1</f></c></row>`
	})

	return { sheets: [{ name: "Data", body: `<sheetData>${rows}</sheetData>` }], sharedStrings: STRINGS(WORDS) }
}

function normalFormula(formula: string | undefined): string {
	return (formula ?? "").replace(/^=/, "").toUpperCase()
}

function cellsOf(result: UnoResult, which: "before" | "after"): string[] {
	const sheet = (which === "before" ? result.before : result.after)?.sheets[0]

	return (sheet?.cells ?? []).map((cell: UnoCell) => `${cell[0]},${cell[1]} ${cell[2]} ${String(cell[3])} ${normalFormula(cell[5])}`)
}

function ok(result: Awaited<ReturnType<SheetSessionApi["command"]>>): void {
	if (!result.ok) throw new Error(`refused: ${result.refusal.reason} (${result.refusal.detail ?? ""})`)
}

async function hiddenRowsOf(bytes: Uint8Array): Promise<{ hidden: number[]; filter: string }> {
	const opened = await openXlsx(bytesSource(bytes))

	if (!opened.ok) throw new Error(`does not open: ${opened.reason}`)

	const entry = opened.workbook.sheets[0]

	if (entry === undefined) throw new Error("no sheet")

	const model = await opened.workbook.load(entry.id)
	const hidden: number[] = []

	for (const block of model.store.blocks)
		for (let i = 0; i < block.rowCount; i++)
			if (((block.rowFlags[i] ?? 0) & RowFlag.HIDDEN) !== 0) hidden.push(block.firstRow + (block.rows[i] ?? 0))

	const section = model.sections.find(candidate => candidate.name === "autoFilter")
	const filter = section?.bytes === null || section === undefined ? null : readAutoFilter(decodeUtf8(section.bytes))

	// One condition reads the same with and without "and" (LibreOffice writes and="1").
	const criteria = filter?.columns.map(column => [
		column.colId,
		column.criteria?.kind === "custom" && column.criteria.conditions.length === 1 ? { ...column.criteria, and: false } : column.criteria
	])

	return { hidden: hidden.sort((a, b) => a - b), filter: JSON.stringify(criteria ?? []) }
}

describe.skipIf(UNO === undefined)("sort and filter against LibreOffice", () => {
	let uno: UnoSheets

	beforeAll(() => {
		mkdirSync(SCRATCH, { recursive: true })
		uno = new UnoSheets(UNO as NonNullable<typeof UNO>, SCRATCH)
	})

	afterAll(async () => {
		await uno.close()
		rmSync(SCRATCH, { recursive: true, force: true })
	})

	async function run(job: UnoJob): Promise<UnoResult> {
		const result = await uno.run(job)

		if (!result.ok) throw new Error(`LibreOffice: ${result.error ?? "failed"}`)

		return result
	}

	const sorts: readonly { name: string; keys: { index: number; descending: boolean }[]; caseSensitive?: boolean }[] = [
		{ name: "names ascending", keys: [{ index: 0, descending: false }] },
		{ name: "scores descending (blank last)", keys: [{ index: 1, descending: true }] },
		{
			name: "group then score",
			keys: [
				{ index: 2, descending: false },
				{ index: 1, descending: false }
			]
		},
		{ name: "names case-sensitive", keys: [{ index: 0, descending: false }], caseSensitive: true }
	]

	for (const testCase of sorts) {
		it(`sorts as LibreOffice does: ${testCase.name}`, { timeout: 300_000 }, async () => {
			const fixture = sortFixture()
			const source = join(SCRATCH, "sort-source.xlsx")
			const ours = join(SCRATCH, "sort-ours.xlsx")

			writeFileSync(source, buildFixture(fixture))

			const opened = await openFixture(fixture)
			const sheet = opened.info.sheets[0]!.id
			const command: SheetCommand = {
				type: "sort.apply",
				sheet,
				range: { startRow: 0, startCol: 0, endRow: DATA.length, endCol: 3 },
				orientation: "rows",
				header: true,
				caseSensitive: testCase.caseSensitive === true,
				keys: testCase.keys
			}

			ok(await opened.session.command(command))
			writeFileSync(ours, await savedBytes(opened.session))
			opened.session.close()

			const theirs = await run({
				input: source,
				ops: [
					{
						op: "sort",
						sheet: "Data",
						range: `A1:D${DATA.length + 1}`,
						header: true,
						keys: testCase.keys.map(key => ({
							column: key.index,
							ascending: !key.descending,
							caseSensitive: testCase.caseSensitive === true
						}))
					}
				]
			})

			expect(theirs.ops?.every(op => op.ok)).toBe(true)
			expect(cellsOf(await run({ input: ours }), "before")).toEqual(cellsOf(theirs, "after"))
		})
	}

	it("orders logical values and errors as Excel documents, where LibreOffice orders them otherwise", { timeout: 300_000 }, async () => {
		const fixture: Fixture = {
			sheets: [
				{
					name: "Data",
					body:
						"<sheetData>" +
						'<row r="1"><c r="A1" t="b"><v>1</v></c></row><row r="2"><c r="A2" t="s"><v>1</v></c></row>' +
						'<row r="3"><c r="A3"><v>1</v></c></row><row r="4"><c r="A4" t="e"><v>#N/A</v></c></row>' +
						'<row r="5"><c r="A5" t="s"><v>0</v></c></row><row r="6"><c r="A6" t="b"><v>0</v></c></row>' +
						"</sheetData>"
				}
			],
			sharedStrings: STRINGS(["a", "b"])
		}
		const source = join(SCRATCH, "mixed-source.xlsx")
		const ours = join(SCRATCH, "mixed-ours.xlsx")

		writeFileSync(source, buildFixture(fixture))

		const opened = await openFixture(fixture)

		ok(
			await opened.session.command({
				type: "sort.apply",
				sheet: opened.info.sheets[0]!.id,
				range: { startRow: 0, startCol: 0, endRow: 5, endCol: 0 },
				orientation: "rows",
				header: false,
				caseSensitive: false,
				keys: [{ index: 0, descending: false }]
			})
		)
		writeFileSync(ours, await savedBytes(opened.session))
		opened.session.close()

		const shown = (result: UnoResult, which: "before" | "after"): string[] =>
			((which === "before" ? result.before : result.after)?.sheets[0]?.cells ?? []).map(cell => cell[4])

		// Numbers, text, logical values (FALSE first), errors.
		expect(shown(await run({ input: ours }), "before")).toEqual(["1", "a", "b", "FALSE", "TRUE", "#N/A"])

		const theirs = await run({
			input: source,
			ops: [{ op: "sort", sheet: "Data", range: "A1:A6", header: false, keys: [{ column: 0 }] }]
		})

		// LibreOffice: booleans are numbers to it, and errors come before text.
		expect(shown(theirs, "after")).not.toEqual(["1", "a", "b", "FALSE", "TRUE", "#N/A"])
	})

	const filters: readonly { name: string; col: number; criteria: Extract<SheetCommand, { type: "filter.column" }>["criteria"] }[] = [
		{ name: "values", col: 0, criteria: { kind: "values", values: ["apple", "fig", "kiwi"], blank: false, dates: [] } },
		{ name: "contains", col: 0, criteria: { kind: "custom", and: false, conditions: [{ operator: "equal", value: "*an*" }] } },
		{
			name: "between",
			col: 1,
			criteria: {
				kind: "custom",
				and: true,
				conditions: [
					{ operator: "greaterThanOrEqual", value: "0" },
					{ operator: "lessThan", value: "18" }
				]
			}
		},
		{ name: "top 3", col: 1, criteria: { kind: "top10", top: true, percent: false, value: 3 } },
		{ name: "bottom 25 percent", col: 1, criteria: { kind: "top10", top: false, percent: true, value: 25 } },
		{ name: "above average", col: 1, criteria: { kind: "dynamic", type: "aboveAverage" } },
		{ name: "blanks", col: 1, criteria: { kind: "values", values: [], blank: true, dates: [] } }
	]

	for (const testCase of filters) {
		it(`hides the rows LibreOffice keeps hidden on reopening: ${testCase.name}`, { timeout: 300_000 }, async () => {
			const fixture = sortFixture()
			const opened = await openFixture(fixture)
			const sheet = opened.info.sheets[0]!.id

			ok(
				await opened.session.command({
					type: "filter.toggle",
					sheet,
					range: { startRow: 0, startCol: 0, endRow: DATA.length, endCol: 3 },
					active: { row: 0, col: 0 }
				})
			)
			ok(await opened.session.command({ type: "filter.column", sheet, row: 0, col: testCase.col, criteria: testCase.criteria }))

			const ours = join(SCRATCH, "filter-ours.xlsx")
			const resaved = join(SCRATCH, "filter-lo.xlsx")

			writeFileSync(ours, await savedBytes(opened.session))
			opened.session.close()
			await run({ input: ours, saveAs: resaved })

			const mine = await hiddenRowsOf(new Uint8Array(readFileSync(ours)))
			const theirs = await hiddenRowsOf(new Uint8Array(readFileSync(resaved)))

			expect(mine.hidden.length).toBeGreaterThan(0)
			expect(theirs.hidden).toEqual(mine.hidden)

			// LibreOffice's xlsx export has no dynamic filter: it keeps the rows hidden and drops the criteria.
			if (testCase.criteria?.kind === "dynamic") expect(theirs.filter).toBe("[]")
			else expect(theirs.filter).toBe(mine.filter)
		})
	}
})

// Every part a filter or a sort writes validates against ECMA-376 Transitional (sml.xsd): the worksheet with its
// autofilter (each kind of criteria, a sort state), the table part's filter and sort state, the dxf a colour filter
// appends, and the workbook with its _FilterDatabase name. Opt-in: xmllint and the Part 4 archive.
describe.skipIf(xsdOracleMissing() !== undefined)("sort and filter against the SpreadsheetML schema", () => {
	it("writes valid worksheets, tables, styles and workbooks", { timeout: 120_000 }, async () => {
		const oracle = new XsdOracle()
		const fixture = sortFixture()
		const tableXml =
			'<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="F1:G5"><autoFilter ref="F1:G5"/><tableColumns count="2"><tableColumn id="1" name="Name"/><tableColumn id="2" name="Score"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>'
		// The table F1:G5 beside the data: its names, and scores 3 to 12.
		const body = (fixture.sheets[0]?.body ?? "").replace(/<\/row>/g, (_, offset: number, all: string) => {
			const row = Number(/<row r="(\d+)"[^>]*>[^]*$/.exec(all.slice(0, offset))?.[1] ?? "0")

			return row > 5
				? "</row>"
				: `<c r="F${row}" t="s"><v>${row === 1 ? 0 : 3 + row}</v></c><c r="G${row}"><v>${row * 3}</v></c></row>`
		})
		const withTable: Fixture = {
			...fixture,
			sheets: [
				{
					name: "Data",
					body: `${body}<tableParts count="1"><tablePart r:id="rId1"/></tableParts>`,
					rels: relationship("rId1", "table", "../tables/table1.xml")
				}
			],
			parts: { "xl/tables/table1.xml": { type: TYPES.table, text: tableXml } }
		}
		const opened = await openFixture(withTable)
		const sheet = opened.info.sheets[0]!.id
		const run = async (command: SheetCommand): Promise<void> => ok(await opened.session.command(command))

		await run({
			type: "filter.toggle",
			sheet,
			range: { startRow: 0, startCol: 0, endRow: DATA.length, endCol: 3 },
			active: { row: 0, col: 0 }
		})
		await run({
			type: "filter.column",
			sheet,
			row: 0,
			col: 0,
			criteria: { kind: "values", values: ["apple"], blank: true, dates: [{ grouping: "day", year: 2024, month: 2, day: 29 }] }
		})
		await run({ type: "filter.column", sheet, row: 0, col: 1, criteria: { kind: "top10", top: false, percent: true, value: 50 } })
		await run({ type: "filter.column", sheet, row: 0, col: 2, criteria: { kind: "dynamic", type: "thisMonth" } })
		await run({ type: "filter.column", sheet, row: 0, col: 3, criteria: { kind: "color", cell: true, color: 0xff0000 } })
		await run({
			type: "sort.apply",
			sheet,
			range: { startRow: 0, startCol: 0, endRow: DATA.length, endCol: 3 },
			orientation: "rows",
			header: true,
			caseSensitive: true,
			keys: [{ index: 0, descending: true }]
		})
		await run({
			type: "filter.column",
			sheet,
			row: 0,
			col: 6,
			criteria: {
				kind: "custom",
				and: true,
				conditions: [
					{ operator: "greaterThan", value: "3" },
					{ operator: "notEqual", value: "9" }
				]
			}
		})
		await run({
			type: "sort.apply",
			sheet,
			range: { startRow: 0, startCol: 5, endRow: 4, endCol: 6 },
			orientation: "rows",
			header: true,
			caseSensitive: false,
			keys: [{ index: 6, descending: true }]
		})

		const saved = parts(await savedBytes(opened.session))
		const source = parts(opened.bytes)
		const names = ["xl/worksheets/sheet1.xml", "xl/tables/table1.xml", "xl/styles.xml", "xl/workbook.xml"]

		opened.session.close()

		expect(saved["xl/worksheets/sheet1.xml"]).toContain("<sortState")
		expect(saved["xl/tables/table1.xml"]).toContain("<sortState")

		const after = oracle.validate(names.map(name => encodeUtf8(saved[name] ?? "")))
		const before = oracle.validate(names.map(name => encodeUtf8(source[name] ?? "")))

		oracle.dispose()
		names.forEach((name, k) => expect([name, newValidityErrors(before[k] ?? new Map(), after[k] ?? new Map())]).toEqual([name, []]))
	})
})
