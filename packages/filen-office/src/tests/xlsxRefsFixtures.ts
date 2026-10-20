import { expect } from "vitest"
import type { SheetRefusal } from "../core/sheet/protocol"
import type { XlsxWorkbook } from "../core/xlsx/read/index"
import { planRefEdit, type RefConstruct, type RefEdit, type RefPlan } from "../core/xlsx/refs/index"
import { buildFixture, changedEntries, openWorkbook, parts, saved, type Fixture } from "./xlsxWriteFixtures"

// Structure edits on built workbooks: planned, applied, saved through the gate, reverted and saved again.

export const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
export const X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
export const XM = "http://schemas.microsoft.com/office/excel/2006/main"
export const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

export const TYPES = {
	table: "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
	drawing: "application/vnd.openxmlformats-officedocument.drawing+xml",
	chart: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
	chartShapes: "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml",
	comments: "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
	vml: "application/vnd.openxmlformats-officedocument.vmlDrawing",
	threaded: "application/vnd.ms-excel.threadedcomments+xml",
	controlProperties: "application/vnd.ms-excel.controlproperties+xml",
	pivotCache: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
	pivotTable: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
	slicerCache: "application/vnd.ms-excel.slicerCache+xml",
	timelineCache: "application/vnd.ms-excel.timelineCache+xml",
	connections: "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml",
	volatile: "application/vnd.openxmlformats-officedocument.spreadsheetml.volatileDependencies+xml",
	singleCells: "application/vnd.openxmlformats-officedocument.spreadsheetml.tableSingleCells+xml",
	externalLink: "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml",
	queryTable: "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml",
	relationships: "application/vnd.openxmlformats-package.relationships+xml"
} as const

export function relationship(id: string, type: string, target: string): string {
	return `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`
}

export function rels(...relationships: string[]): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.join("")}</Relationships>`
}

// The constructs the tests saw a plan change, checked against the catalogue at the end.
export const covered = new Set<RefConstruct>()

export function rows(type: "insert" | "delete", at: number, count = 1, sheet = 0): RefEdit {
	return { kind: "axis", sheet, shift: { type, axis: "rows", at, count } }
}

export function cols(type: "insert" | "delete", at: number, count = 1, sheet = 0): RefEdit {
	return { kind: "axis", sheet, shift: { type, axis: "cols", at, count } }
}

export interface Edited {
	readonly workbook: XlsxWorkbook
	readonly plan: RefPlan
	// The saved package's parts as text.
	readonly files: Record<string, string>
	readonly sheet: string
}

// Plans, applies and saves (the gate checks the save); then reverts and checks that the save is the source again
// (every entry's compressed bytes as stored), but for calcChain, which a structure edit drops for good.
// `appended`: entries an append-only record (a shared string) leaves changed after the revert.
export async function edited(
	fixture: Fixture,
	edit: RefEdit | ((workbook: XlsxWorkbook) => RefEdit),
	expected: readonly RefConstruct[] = [],
	appended: readonly string[] = []
): Promise<Edited> {
	const source = buildFixture(fixture)
	const workbook = await openWorkbook(source)
	const result = await planRefEdit(workbook, typeof edit === "function" ? edit(workbook) : edit)

	if (!result.ok) throw new Error(`refused: ${JSON.stringify(result.refusal)}`)

	result.plan.apply()

	for (const construct of expected) expect(result.plan.touched, construct).toContain(construct)
	for (const construct of result.plan.touched) if (expected.includes(construct)) covered.add(construct)

	const files = parts(await saved(workbook))

	result.plan.revert()

	const back = await changedEntries(source, await saved(workbook))
	const dropped =
		fixture.calcChain === true || appended.length > 0 ? ["xl/calcChain.xml", "[Content_Types].xml", "xl/_rels/workbook.xml.rels"] : []
	const calc = (name: string) => !dropped.includes(name) && !appended.includes(name)

	expect(back.changed.filter(calc), "reverted save").toEqual([])
	expect(back.added.filter(calc), "reverted save").toEqual([])
	expect(back.removed.filter(calc), "reverted save").toEqual([])

	return { workbook, plan: result.plan, files, sheet: files["xl/worksheets/sheet1.xml"] ?? "" }
}

export async function refused(
	fixture: Fixture,
	edit: RefEdit | ((workbook: XlsxWorkbook) => RefEdit),
	construct?: RefConstruct
): Promise<SheetRefusal> {
	const workbook = await openWorkbook(buildFixture(fixture))
	const result = await planRefEdit(workbook, typeof edit === "function" ? edit(workbook) : edit)

	if (result.ok) throw new Error("planned, not refused")
	if (construct !== undefined) covered.add(construct)

	return result.refusal
}

export function sheetBody(body: string): Fixture {
	return { sheets: [{ name: "Data", body }] }
}
