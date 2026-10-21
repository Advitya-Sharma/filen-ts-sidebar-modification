import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { afterAll, beforeAll, describe, expect, it } from "vitest"
import { bytesSource } from "../core/opc/index"
import type { SheetCommand, SheetCommandResult, SheetSessionApi, SheetWorkbookInfo } from "../core/sheet/protocol"
import { FormulaKind, RowFlag, ColFlag } from "../core/xlsx/model/types"
import { openXlsx } from "../core/xlsx/read/index"
import { parseSqref } from "../core/xlsx/refs/sqref"
import { decodeUtf8 } from "../core/xml/index"
import { CORPUS_DIR } from "./corpusDir"
import { sheetOffices, unoOffice } from "./oracles/sheet/libreOffice"
import { UnoSheets, type UnoOperation } from "./oracles/sheet/unoSheets"
import { buildFixture, parts, type Fixture } from "./xlsxWriteFixtures"
import { openFixture, savedBytes } from "./sheetSessionFixtures"

// The session's structure commands against LibreOffice: the same operations replayed over UNO on the source file,
// LibreOffice's xlsx re-save read into the same normalised structures as our save (sheet order, names and states,
// merges, hidden rows and columns, frozen panes, tab colours, protection, defined names, cell formulas, tables,
// conditional formats, validations, drawing anchors and notes). Sequences edit a duplicated sheet again. Every
// difference must be one listed with its reason (Excel's documented behaviour breaks a tie). Opt-in:
// FILEN_OFFICE_SOFFICE=<soffice>.

const UNO = unoOffice(sheetOffices())
const SCRATCH = join(tmpdir(), "filen-office-structure-lo")
const MAX_ROW = 60
const MAX_COL = 20

function normal(formula: string): string {
	return (
		formula
			.replace(/^=/, "")
			.replace(/\$/g, "")
			// Excel writes a reference to a deleted sheet as #REF! followed by the cell (#REF!A1), and a deleted cell of a
			// sheet as the sheet followed by #REF! ('Sheet 2'!#REF!); LibreOffice writes #REF! alone for both.
			// LibreOffice writes a reference that lost only its row or column with the rest kept (B#REF!, #REF!2).
			.replace(/\$?[A-Za-z]{1,3}\$?#REF!/g, "#REF!")
			.replace(/#REF!\$?[0-9]+(?![0-9]*!)/g, "#REF!")
			.replace(/#REF!\$?[A-Z]{1,3}\$?[0-9]+/g, "#REF!")
			.replace(/(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!#REF!/g, "#REF!")
			.replace(/'([A-Za-z_][A-Za-z0-9_. ()]*)'!/g, (_, name: string) => (/[ ()]/.test(name) ? `'${name}'!` : `${name}!`))
			.toUpperCase()
	)
}

// Every structure of a saved package, normalised, by "sheet/kind".
async function structures(bytes: Uint8Array): Promise<Record<string, string>> {
	const opened = await openXlsx(bytesSource(bytes))

	if (!opened.ok) throw new Error(`does not open: ${opened.reason}`)

	const workbook = opened.workbook
	const out: Record<string, string> = {}

	out["workbook/sheets"] = workbook.sheets.map(sheet => `${sheet.name}:${sheet.state}`).join(" ")
	out["workbook/names"] = workbook.definedNames
		.filter(name => !name.name.startsWith("_xlnm._FilterDatabase"))
		.map(name => `${name.name}@${name.sheet === null ? "" : (workbook.sheet(name.sheet)?.name ?? "?")}=${normal(name.formula)}`)
		.sort()
		.join(" ")

	for (const entry of workbook.sheets) {
		if (entry.kind !== "worksheet") continue

		const model = await workbook.load(entry.id)
		const sheet = entry.name
		const merges: string[] = []

		for (let k = 0; k + 3 < model.merges.length; k += 4) merges.push([...model.merges.subarray(k, k + 4)].join(","))

		out[`${sheet}/merges`] = merges.sort().join(" ")

		const hiddenRows: number[] = []

		for (const block of model.store.blocks)
			for (let i = 0; i < block.rowCount; i++) {
				const row = block.firstRow + (block.rows[i] ?? 0)

				if (row < MAX_ROW && ((block.rowFlags[i] ?? 0) & RowFlag.HIDDEN) !== 0) hiddenRows.push(row)
			}

		out[`${sheet}/hiddenRows`] = hiddenRows.join(",")
		out[`${sheet}/hiddenCols`] = model.cols
			.filter(range => (range.flags & ColFlag.HIDDEN) !== 0 && range.first < MAX_COL)
			.flatMap(range => Array.from({ length: Math.min(range.last, MAX_COL - 1) - range.first + 1 }, (_, k) => range.first + k))
			.join(",")
		out[`${sheet}/pane`] = model.view.frozen === null ? "" : `${String(model.view.frozen.rows)}x${String(model.view.frozen.cols)}`
		out[`${sheet}/tabColor`] = model.tabColor?.argb === undefined ? "" : (model.tabColor.argb & 0xffffff).toString(16)
		out[`${sheet}/protected`] =
			model.protection === null
				? ""
				: `password ${(model.protection.hash.legacy ?? model.protection.hash.algorithmName ?? "none").toUpperCase()}`

		const formulas: string[] = []
		const table = model.store.formulas

		for (const block of model.store.blocks)
			for (let i = 0; i < block.rowCount; i++) {
				const row = block.firstRow + (block.rows[i] ?? 0)

				for (let k = block.rowStarts[i] ?? 0; k < (block.rowStarts[i + 1] ?? 0); k++) {
					const id = block.formulas?.[k] ?? -1

					if (id < 0) continue

					const col = block.cols[k] ?? 0
					const text = table.kind(id) === FormulaKind.SHARED_CHILD ? table.expanded(id, row, col) : table.text(id)

					formulas.push(`${String(row)},${String(col)} ${normal(text)}`)
				}
			}

		out[`${sheet}/formulas`] = formulas.sort().join(" | ")

		// Ranges written in the sheet's sections: conditional formats and validations, each as its sqref.
		const sqrefs = (section: string, element: RegExp): string =>
			model.sections
				.filter(candidate => candidate.name === section)
				.flatMap(candidate =>
					[...decodeUtf8(candidate.bytes ?? new Uint8Array(0)).matchAll(element)].map(match => sqref(match[1] ?? ""))
				)
				.sort()
				.join(" ")

		out[`${sheet}/conditionalFormats`] = sqrefs("conditionalFormatting", /<(?:\w+:)?conditionalFormatting\b[^>]*\ssqref="([^"]*)"/g)
		out[`${sheet}/validations`] = sqrefs("dataValidations", /<(?:\w+:)?dataValidation\b[^>]*\ssqref="([^"]*)"/g)

		// The parts the sheet relates to: its tables' ranges, its notes' cells and its drawings' anchors.
		const tables: string[] = []
		const notes: string[] = []
		const anchors: string[] = []

		for (const relationship of model.relationships.all) {
			if (relationship.external || relationship.partName === undefined) continue

			const type = relationship.type.slice(relationship.type.lastIndexOf("/") + 1)

			if (type !== "table" && type !== "comments" && type !== "drawing") continue

			const text = decodeUtf8(await workbook.parts.read(relationship.partName))

			if (type === "table") tables.push(/<(?:\w+:)?table\b[^>]*\sref="([^"]*)"/.exec(text)?.[1] ?? "?")
			else if (type === "comments")
				for (const match of text.matchAll(/<(?:\w+:)?comment\b[^>]*\sref="([^"]*)"/g)) notes.push(match[1] ?? "")
			else
				for (const match of text.matchAll(
					/<(?:\w+:)?from>\s*<(?:\w+:)?col>(\d+)<\/(?:\w+:)?col>[\s\S]*?<(?:\w+:)?row>(\d+)<\/(?:\w+:)?row>/g
				))
					anchors.push(`${match[2] ?? ""},${match[1] ?? ""}`)
		}

		out[`${sheet}/tables`] = tables.sort().join(" ")
		out[`${sheet}/notes`] = notes.sort().join(" ")
		out[`${sheet}/drawings`] = anchors.sort().join(" ")
	}

	return out
}

// The cells a sqref covers, as one canonical text whatever ranges they are written in (LibreOffice joins and splits
// ranges otherwise): bands of rows, each with its runs of columns.
function sqref(text: string): string {
	const ranges = parseSqref(text)

	if (ranges === null || ranges.length === 0) return text.trim()

	const rowCuts = [...new Set(ranges.flatMap(range => [range.startRow, range.endRow + 1]))].sort((a, b) => a - b)
	const colCuts = [...new Set(ranges.flatMap(range => [range.startCol, range.endCol + 1]))].sort((a, b) => a - b)
	const bands: { from: number; to: number; runs: string }[] = []

	for (let r = 0; r + 1 < rowCuts.length; r++) {
		const top = rowCuts[r] ?? 0
		const runs: string[] = []
		let open = -1

		for (let c = 0; c + 1 < colCuts.length; c++) {
			const left = colCuts[c] ?? 0
			const covered = ranges.some(
				range => range.startRow <= top && range.endRow >= top && range.startCol <= left && range.endCol >= left
			)

			if (covered && open < 0) open = left
			if (!covered && open >= 0) {
				runs.push(`${String(open)}-${String(left - 1)}`)
				open = -1
			}
		}

		if (open >= 0) runs.push(`${String(open)}-${String((colCuts[colCuts.length - 1] ?? 0) - 1)}`)

		const joined = runs.join("+")
		const last = bands[bands.length - 1]

		if (last !== undefined && last.runs === joined && last.to === top - 1) last.to = (rowCuts[r + 1] ?? top) - 1
		else bands.push({ from: top, to: (rowCuts[r + 1] ?? top) - 1, runs: joined })
	}

	return bands
		.filter(band => band.runs !== "")
		.map(band => `r${String(band.from)}-${String(band.to)}c${band.runs}`)
		.join(",")
}

// Every table of a saved package has an id and a name of its own (ECMA-376 Part 1 §18.5.1.2).
async function tablesUnique(bytes: Uint8Array): Promise<string[]> {
	const opened = await openXlsx(bytesSource(bytes))

	if (!opened.ok) throw new Error(`does not open: ${opened.reason}`)

	const ids: string[] = []
	const names: string[] = []

	for (const part of opened.workbook.parts.parts()) {
		if (!part.contentType.toLowerCase().endsWith("table+xml")) continue

		const root = /<(?:\w+:)?table\b[^>]*>/.exec(decodeUtf8(await opened.workbook.parts.read(part.name)))?.[0] ?? ""

		ids.push(/\sid="(\d+)"/.exec(root)?.[1] ?? "")
		names.push((/\sdisplayName="([^"]*)"/.exec(root)?.[1] ?? "").toUpperCase())
	}

	const out: string[] = []

	if (new Set(ids).size !== ids.length) out.push(`table ids repeat: ${ids.join(" ")}`)
	if (new Set(names).size !== names.length) out.push(`table names repeat: ${names.join(" ")}`)

	return out
}

function differences(ours: Record<string, string>, theirs: Record<string, string>): string[] {
	const keys = new Set([...Object.keys(ours), ...Object.keys(theirs)])
	const out: string[] = []

	for (const key of [...keys].sort())
		if ((ours[key] ?? "") !== (theirs[key] ?? "")) out.push(`${key}: ours [${ours[key] ?? ""}] LibreOffice [${theirs[key] ?? ""}]`)

	return out
}

// Differences LibreOffice's own re-save makes, or where it departs from Excel.
const ACCEPTED: readonly { readonly pattern: RegExp; readonly why: string }[] = [
	{
		pattern: /^[^:]*insert (?:rows|cells down)[^:]*: [^/]+\/validations: ours \[r2-7c2-2\] LibreOffice \[r0-7c2-2\]$/,
		why: "rows (or cells shifting down) inserted above a validation's range move it as they move a reference to the range (its sqref is a list of references, ECMA-376 Part 1 §18.3.1.32), as they move the conditional formats' ranges here; LibreOffice widens the validation alone"
	},
	{
		pattern: /^delete columns: [^/]+\/formulas: ours \[.*SUM\(#REF!\).*\] LibreOffice \[.*TABLE1\[\[#THIS ROW\]/,
		why: "a structured reference to a deleted table column becomes #REF! (support.microsoft.com, Using structured references with Excel tables); LibreOffice points it at another column"
	},
	{
		pattern: /^duplicate[^:]*: [^/]+(?: \(\d+\))+\/formulas: ours \[.*TABLE1(?:_\d+)+\[.*\] LibreOffice \[.*TABLE1\[/,
		why: "a copied sheet's table is a table of its own under a new name, and the copy's structured references name it; LibreOffice keeps them naming the source's table, which the copy's rows are not in"
	}
]

const SHEETS: Fixture = {
	sheets: [
		{
			name: "One",
			body:
				'<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c></row>' +
				'<row r="2"><c r="A2"><v>3</v></c><c r="B2"><f>Two!A1+A2</f><v>8</v></c></row>' +
				'<row r="5"><c r="A5"><f>SUM(A1:A4)</f><v>4</v></c></row></sheetData>' +
				'<mergeCells count="1"><mergeCell ref="D2:E3"/></mergeCells>' +
				'<conditionalFormatting sqref="A1:A5"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>2</formula></cfRule></conditionalFormatting>' +
				'<dataValidations count="1"><dataValidation type="whole" operator="between" allowBlank="1" sqref="C1:C6"><formula1>0</formula1><formula2>9</formula2></dataValidation></dataValidations>'
		},
		{ name: "Two", body: '<sheetData><row r="1"><c r="A1"><v>5</v></c><c r="B1"><f>One!A5*2</f><v>8</v></c></row></sheetData>' },
		{ name: "Three", body: '<sheetData><row r="1"><c r="A1"><f>SUM(One:Two!A1)</f><v>6</v></c></row></sheetData>' }
	],
	workbookTail:
		'<definedNames><definedName name="Total">One!$A$5</definedName><definedName name="Local" localSheetId="1">Two!$A$1</definedName></definedNames>'
}

type Sheets = SheetWorkbookInfo["sheets"]

interface Case {
	readonly name: string
	readonly ours: (info: SheetWorkbookInfo) => SheetCommand
	// Commands run after it, each on the sheets as the commands before left them (a duplicate edited again).
	readonly then?: readonly ((sheets: Sheets) => SheetCommand)[]
	// null: LibreOffice cannot do it headless (frozen panes need a view); its re-save of our save is compared.
	readonly theirs: UnoOperation | readonly UnoOperation[] | ((info: SheetWorkbookInfo) => UnoOperation | readonly UnoOperation[]) | null
}

const range = (startRow: number, startCol: number, endRow: number, endCol: number) => ({ startRow, startCol, endRow, endCol })

const CASES: readonly Case[] = [
	{
		name: "insert rows",
		ours: info => ({ type: "axis.insert", sheet: info.sheets[0]!.id, axis: "rows", at: 1, count: 2 }),
		theirs: { op: "insertRows", sheet: 0, at: 1, count: 2 }
	},
	{
		name: "insert rows at the top",
		ours: info => ({ type: "axis.insert", sheet: info.sheets[0]!.id, axis: "rows", at: 0, count: 2 }),
		theirs: { op: "insertRows", sheet: 0, at: 0, count: 2 }
	},
	{
		name: "delete rows",
		ours: info => ({ type: "axis.delete", sheet: info.sheets[0]!.id, axis: "rows", at: 0, count: 1 }),
		theirs: { op: "deleteRows", sheet: 0, at: 0, count: 1 }
	},
	{
		name: "insert columns",
		ours: info => ({ type: "axis.insert", sheet: info.sheets[0]!.id, axis: "cols", at: 0, count: 1 }),
		theirs: { op: "insertColumns", sheet: 0, at: 0, count: 1 }
	},
	{
		name: "delete columns",
		ours: info => ({ type: "axis.delete", sheet: info.sheets[0]!.id, axis: "cols", at: 3, count: 1 }),
		theirs: { op: "deleteColumns", sheet: 0, at: 3, count: 1 }
	},
	{ name: "delete sheet", ours: info => ({ type: "sheet.delete", sheet: info.sheets[1]!.id }), theirs: { op: "deleteSheet", sheet: 1 } },
	{
		name: "move sheet",
		ours: info => ({ type: "sheet.move", sheet: info.sheets[0]!.id, to: 2 }),
		theirs: { op: "moveSheet", sheet: 0, to: 2 }
	},
	{
		name: "rename sheet",
		ours: info => ({ type: "sheet.rename", sheet: info.sheets[0]!.id, name: "Main" }),
		theirs: { op: "renameSheet", sheet: 0, name: "Main" }
	},
	{
		name: "duplicate sheet",
		ours: info => ({ type: "sheet.duplicate", sheet: info.sheets[0]!.id }),
		theirs: info => ({ op: "copySheet", sheet: 0, name: `${info.sheets[0]!.name} (2)`, to: 1 })
	},
	{
		name: "hide sheet",
		ours: info => ({ type: "sheet.state", sheet: info.sheets[2]!.id, state: "hidden" }),
		theirs: { op: "hideSheet", sheet: 2 }
	},
	{
		name: "tab colour",
		ours: info => ({ type: "sheet.tabColor", sheet: info.sheets[0]!.id, color: 0x00b050 }),
		theirs: { op: "tabColor", sheet: 0, color: 0x00b050 }
	},
	{
		name: "hide rows",
		ours: info => ({ type: "axis.hide", sheet: info.sheets[0]!.id, axis: "rows", spans: [{ start: 2, end: 3 }], hidden: true }),
		theirs: { op: "hideRows", sheet: 0, at: 2, count: 2 }
	},
	{
		name: "hide columns",
		ours: info => ({ type: "axis.hide", sheet: info.sheets[0]!.id, axis: "cols", spans: [{ start: 1, end: 2 }], hidden: true }),
		theirs: { op: "hideColumns", sheet: 0, at: 1, count: 2 }
	},
	{ name: "freeze panes", ours: info => ({ type: "view.freeze", sheet: info.sheets[0]!.id, rows: 1, cols: 1 }), theirs: null },
	{
		name: "merge",
		ours: info => ({ type: "merge.cells", sheet: info.sheets[0]!.id, ranges: [range(6, 0, 7, 2)], mode: "merge" }),
		theirs: { op: "mergeCells", sheet: 0, range: "A7:C8" }
	},
	{
		name: "unmerge",
		ours: info => ({ type: "merge.unmerge", sheet: info.sheets[0]!.id, ranges: [range(1, 3, 1, 3)] }),
		theirs: { op: "unmergeCells", sheet: 0, range: "D2:E3" }
	},
	{
		name: "duplicate then insert rows on the copy",
		ours: info => ({ type: "sheet.duplicate", sheet: info.sheets[0]!.id }),
		then: [sheets => ({ type: "axis.insert", sheet: sheets[1]!.id, axis: "rows", at: 0, count: 2 })],
		theirs: info => [
			{ op: "copySheet", sheet: 0, name: `${info.sheets[0]!.name} (2)`, to: 1 },
			{ op: "insertRows", sheet: 1, at: 0, count: 2 }
		]
	},
	{
		name: "duplicate then delete columns on the copy",
		ours: info => ({ type: "sheet.duplicate", sheet: info.sheets[0]!.id }),
		then: [sheets => ({ type: "axis.delete", sheet: sheets[1]!.id, axis: "cols", at: 0, count: 1 })],
		theirs: info => [
			{ op: "copySheet", sheet: 0, name: `${info.sheets[0]!.name} (2)`, to: 1 },
			{ op: "deleteColumns", sheet: 1, at: 0, count: 1 }
		]
	},
	{
		name: "duplicate twice",
		ours: info => ({ type: "sheet.duplicate", sheet: info.sheets[0]!.id }),
		then: [sheets => ({ type: "sheet.duplicate", sheet: sheets[0]!.id })],
		theirs: info => [
			{ op: "copySheet", sheet: 0, name: `${info.sheets[0]!.name} (2)`, to: 1 },
			{ op: "copySheet", sheet: 0, name: `${info.sheets[0]!.name} (3)`, to: 1 }
		]
	},
	{
		name: "duplicate the copy, then insert rows on the source",
		ours: info => ({ type: "sheet.duplicate", sheet: info.sheets[0]!.id }),
		then: [
			sheets => ({ type: "sheet.duplicate", sheet: sheets[1]!.id }),
			sheets => ({ type: "axis.insert", sheet: sheets[0]!.id, axis: "rows", at: 1, count: 1 })
		],
		theirs: info => [
			{ op: "copySheet", sheet: 0, name: `${info.sheets[0]!.name} (2)`, to: 1 },
			{ op: "copySheet", sheet: 1, name: `${info.sheets[0]!.name} (2) (2)`, to: 2 },
			{ op: "insertRows", sheet: 0, at: 1, count: 1 }
		]
	},
	{
		name: "insert cells down",
		ours: info => ({ type: "cells.shift", sheet: info.sheets[0]!.id, range: range(0, 1, 0, 1), edit: "insert", axis: "rows" }),
		theirs: { op: "insertCells", sheet: 0, range: "B1", mode: "down" }
	},
	{
		name: "insert cells down across a range",
		ours: info => ({ type: "cells.shift", sheet: info.sheets[0]!.id, range: range(1, 0, 1, 0), edit: "insert", axis: "rows" }),
		theirs: { op: "insertCells", sheet: 0, range: "A2", mode: "down" }
	},
	{
		name: "delete cells up",
		ours: info => ({ type: "cells.shift", sheet: info.sheets[0]!.id, range: range(0, 0, 0, 0), edit: "delete", axis: "rows" }),
		theirs: { op: "removeCells", sheet: 0, range: "A1", mode: "up" }
	},
	{
		name: "insert cells right",
		ours: info => ({ type: "cells.shift", sheet: info.sheets[0]!.id, range: range(0, 0, 0, 0), edit: "insert", axis: "cols" }),
		theirs: { op: "insertCells", sheet: 0, range: "A1", mode: "right" }
	},
	{
		name: "delete cells left",
		ours: info => ({ type: "cells.shift", sheet: info.sheets[0]!.id, range: range(4, 0, 4, 0), edit: "delete", axis: "cols" }),
		theirs: { op: "removeCells", sheet: 0, range: "A5", mode: "left" }
	},
	{
		name: "insert cells down over a merge",
		ours: info => ({ type: "cells.shift", sheet: info.sheets[0]!.id, range: range(0, 3, 0, 4), edit: "insert", axis: "rows" }),
		theirs: { op: "insertCells", sheet: 0, range: "D1:E1", mode: "down" }
	},
	{
		name: "delete cells up over a merge",
		ours: info => ({ type: "cells.shift", sheet: info.sheets[0]!.id, range: range(0, 3, 0, 4), edit: "delete", axis: "rows" }),
		theirs: { op: "removeCells", sheet: 0, range: "D1:E1", mode: "up" }
	},
	{
		name: "insert cells down across the used columns",
		ours: info => ({ type: "cells.shift", sheet: info.sheets[0]!.id, range: range(0, 0, 1, 11), edit: "insert", axis: "rows" }),
		theirs: { op: "insertCells", sheet: 0, range: "A1:L2", mode: "down" }
	},
	{
		name: "delete cells left across the used rows",
		ours: info => ({ type: "cells.shift", sheet: info.sheets[0]!.id, range: range(0, 0, 29, 0), edit: "delete", axis: "cols" }),
		theirs: { op: "removeCells", sheet: 0, range: "A1:A30", mode: "left" }
	},
	{
		name: "duplicate then insert cells right on the copy",
		ours: info => ({ type: "sheet.duplicate", sheet: info.sheets[0]!.id }),
		then: [sheets => ({ type: "cells.shift", sheet: sheets[1]!.id, range: range(0, 0, 0, 0), edit: "insert", axis: "cols" })],
		theirs: info => [
			{ op: "copySheet", sheet: 0, name: `${info.sheets[0]!.name} (2)`, to: 1 },
			{ op: "insertCells", sheet: 1, range: "A1", mode: "right" }
		]
	},
	{
		name: "protect sheet",
		ours: info => ({
			type: "sheet.protect",
			sheet: info.sheets[1]!.id,
			allows: ["selectLockedCells", "selectUnlockedCells"],
			hash: { legacy: "83AF" }
		}),
		theirs: { op: "protectSheet", sheet: 1, password: "password" }
	}
]

const CORPUS: readonly string[] = [
	"xlsx/xlsxwriter__chart_pie01.xlsx",
	"xlsx/xlsxwriter__table10.xlsx",
	"xlsx/closedxml__cond_format_theme_color2.xlsx",
	"xlsx/xlsxwriter__comment01.xlsx",
	"xlsx/xlsxwriter__data_validation02.xlsx"
]

const CORPUS_CASES: ReadonlySet<string> = new Set([
	"insert rows",
	"delete rows",
	"insert columns",
	"delete columns",
	"hide rows",
	"freeze panes",
	"duplicate sheet",
	"duplicate then insert rows on the copy",
	"duplicate then delete columns on the copy",
	"duplicate twice",
	"duplicate the copy, then insert rows on the source",
	"insert cells down",
	"delete cells up",
	"insert cells right",
	"delete cells left",
	"insert cells down across the used columns",
	"delete cells left across the used rows"
])

async function ours(
	bytes: Uint8Array,
	testCase: Case
): Promise<{ readonly saved: Uint8Array | string; readonly theirs: readonly UnoOperation[] | null }> {
	const { session, info } = await openFixture(bytes)

	try {
		const given = typeof testCase.theirs === "function" ? testCase.theirs(info) : testCase.theirs
		const theirs = given === null ? null : Array.isArray(given) ? (given as readonly UnoOperation[]) : [given as UnoOperation]
		let sheets: Sheets = info.sheets
		let result: SheetCommandResult = await (session as SheetSessionApi).command(testCase.ours(info))

		for (const next of testCase.then ?? []) {
			if (!result.ok) break

			sheets = result.change.sheets ?? sheets
			result = await (session as SheetSessionApi).command(next(sheets))
		}

		return { saved: result.ok ? await savedBytes(session) : `refused ${JSON.stringify(result.refusal)}`, theirs }
	} finally {
		session.close()
	}
}

describe.skipIf(UNO === undefined)("structure commands against LibreOffice", () => {
	beforeAll(() => {
		mkdirSync(SCRATCH, { recursive: true })
	})

	afterAll(() => {
		rmSync(SCRATCH, { recursive: true, force: true })
	})

	// A LibreOffice of its own for each operation: 7.5 fails some operations on a document after others ran in the
	// same process ("illegal object given!", then the bridge goes).
	async function unoRun(job: Parameters<UnoSheets["run"]>[0]): Promise<Awaited<ReturnType<UnoSheets["run"]>>> {
		const uno = new UnoSheets(UNO as NonNullable<typeof UNO>, SCRATCH)

		try {
			return await uno.run(job)
		} finally {
			await uno.close()
		}
	}

	const sources: { name: string; bytes: () => Uint8Array | null; cases: readonly Case[] }[] = [
		{ name: "fixture", bytes: () => buildFixture(SHEETS), cases: CASES },
		...CORPUS.map(name => ({
			name,
			bytes: () => (existsSync(join(CORPUS_DIR, name)) ? new Uint8Array(readFileSync(join(CORPUS_DIR, name))) : null),
			cases: CASES.filter(testCase => CORPUS_CASES.has(testCase.name))
		}))
	]

	for (const source of sources) {
		it(`equals LibreOffice's re-save: ${source.name}`, { timeout: 900_000 }, async () => {
			const bytes = source.bytes()

			if (bytes === null) return

			const stem = source.name.replace(/[^\w]+/g, "-")
			const input = join(SCRATCH, `${stem}.xlsx`)
			const unexplained: string[] = []
			const failedOps: string[] = []

			writeFileSync(input, bytes)

			const baselineTarget = join(SCRATCH, `${stem}-base.xlsx`)
			const baseline = await unoRun({ input, saveAs: baselineTarget, maxCells: 1000 })

			expect(baseline.ok, baseline.error).toBe(true)

			const roundTrip = new Set(
				differences(await structures(bytes), await structures(new Uint8Array(readFileSync(baselineTarget)))).map(line =>
					line.slice(0, line.indexOf(":"))
				)
			)

			for (const testCase of source.cases) {
				const { saved, theirs } = await ours(bytes, testCase)

				if (typeof saved === "string") {
					console.log(`structure-lo ${source.name} ${testCase.name}: ${saved}`)

					// A duplicate edited again is refused only for what the copy cannot hold.
					if (testCase.then !== undefined) unexplained.push(`${testCase.name}: ${saved}`)

					continue
				}

				const target = join(SCRATCH, `${stem}-${testCase.name.replaceAll(" ", "-")}.xlsx`)
				const ourFile = join(SCRATCH, `${stem}-${testCase.name.replaceAll(" ", "-")}-ours.xlsx`)

				writeFileSync(ourFile, saved)

				for (const problem of await tablesUnique(saved)) unexplained.push(`${testCase.name}: ${problem}`)

				const job =
					theirs === null
						? { input: ourFile, saveAs: target, maxCells: 1000 }
						: { input, ops: theirs, saveAs: target, maxCells: 1000 }
				const failed = (result: Awaited<ReturnType<typeof unoRun>>): boolean =>
					!result.ok || (theirs !== null && (result.ops?.length !== theirs.length || result.ops.some(op => !op.ok)))
				let replay = await unoRun(job)

				// 7.5 fails some operations now and then on a fresh process too ("illegal object given!" deleting columns
				// through a table): tried again, and reported when it keeps failing (there is nothing to compare).
				for (let attempt = 0; attempt < 2 && failed(replay); attempt++) replay = await unoRun(job)

				if (failed(replay)) {
					console.log(
						`structure-lo ${source.name} ${testCase.name}: LibreOffice failed: ${replay.error ?? replay.ops?.find(op => !op.ok)?.error ?? ""}`
					)
					failedOps.push(testCase.name)

					continue
				}

				for (const line of differences(await structures(saved), await structures(new Uint8Array(readFileSync(target))))) {
					const key = line.slice(0, line.indexOf(":"))

					if (roundTrip.has(key)) continue

					const accepted = ACCEPTED.find(entry => entry.pattern.test(`${testCase.name}: ${line}`))

					if (accepted !== undefined)
						console.log(`structure-lo ${source.name} ${testCase.name}: accepted (${accepted.why}): ${line}`)
					else unexplained.push(`${testCase.name}: ${line}`)
				}
			}

			expect(unexplained).toEqual([])
			// LibreOffice performs at least all but one operation on each source.
			expect(failedOps.length).toBeLessThanOrEqual(1)
		})
	}
})

// Corpus sheets frozen while scrolled (the view's topLeftCell past A1), as LibreOffice loads their views: the frozen
// pane's first row and column, and the scrolling pane's.
function scrolledFrozenSheets(): { file: string; parts: string[] }[] {
	const directory = join(CORPUS_DIR, "xlsx")

	if (!existsSync(directory)) return []

	const out: { file: string; parts: string[] }[] = []

	for (const file of readdirSync(directory)
		.filter(name => name.endsWith(".xlsx"))
		.sort()) {
		let files: Record<string, string>

		try {
			files = parts(new Uint8Array(readFileSync(join(directory, file))))
		} catch {
			continue
		}

		const found = Object.keys(files).filter(name => {
			if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(name)) return false

			const view =
				/<(?:\w+:)?sheetView\b[^>]*>[\s\S]*?(?:<\/(?:\w+:)?sheetView>|$)/.exec((files[name] ?? "").slice(0, 20_000))?.[0] ?? ""
			const topLeft = /<(?:\w+:)?sheetView\b[^>]*\stopLeftCell="([^"]*)"/.exec(view)?.[1]

			return /state="frozen/.test(view) && topLeft !== undefined && topLeft !== "A1"
		})

		if (found.length > 0) out.push({ file, parts: found })
	}

	return out
}

describe.skipIf(UNO === undefined)("frozen views scrolled before freezing, against LibreOffice", () => {
	it("shows each corpus sheet's frozen block and scrolling pane where LibreOffice loads them", { timeout: 1_800_000 }, async () => {
		const scratch = join(SCRATCH, "views")
		const differ: string[] = []
		let compared = 0

		mkdirSync(scratch, { recursive: true })

		for (const { file, parts: sheetParts } of scrolledFrozenSheets()) {
			const bytes = new Uint8Array(readFileSync(join(CORPUS_DIR, "xlsx", file)))
			const input = join(scratch, file)

			writeFileSync(input, bytes)

			const uno = new UnoSheets(UNO as NonNullable<typeof UNO>, scratch)
			let theirs: Awaited<ReturnType<UnoSheets["run"]>>

			try {
				theirs = await uno.run({ input, views: true, maxCells: 1 })
			} finally {
				await uno.close()
			}

			if (!theirs.ok || theirs.views === undefined) {
				differ.push(`${file}: LibreOffice failed: ${theirs.error ?? ""}`)
				continue
			}

			const opened = await openXlsx(bytesSource(bytes))

			if (!opened.ok) throw new Error(`${file}: ${opened.reason}`)

			const { session } = await openFixture(bytes)

			try {
				for (const part of sheetParts) {
					const entry = opened.workbook.sheets.find(candidate => candidate.part === part)
					const view = theirs.views.find(candidate => candidate.name === entry?.name)

					if (entry === undefined || view === undefined) continue

					const pane = (await session.layout(entry.id)).pane

					if (pane === null) {
						differ.push(`${file} ${entry.name}: no frozen pane`)
						continue
					}

					// LibreOffice's panes: the frozen (top-left) one first, the scrolling (bottom-right) one last.
					const frozenPane = view.panes[0] ?? [0, 0]
					const scrolling = view.panes[view.panes.length - 1] ?? [0, 0]
					const ours = {
						frozenRow: pane.rows > 0 ? pane.firstRow : pane.topRow,
						frozenCol: pane.cols > 0 ? pane.firstCol : pane.leftCol,
						topRow: pane.topRow,
						leftCol: pane.leftCol,
						splitRow: pane.rows > 0 ? pane.firstRow + pane.rows : 0,
						splitColumn: pane.cols > 0 ? pane.firstCol + pane.cols : 0
					}
					const lo = {
						frozenRow: frozenPane[0],
						frozenCol: frozenPane[1],
						topRow: scrolling[0],
						leftCol: scrolling[1],
						splitRow: view.splitRow,
						splitColumn: view.splitColumn
					}

					const line = `${file} ${entry.name}: ours ${JSON.stringify(ours)} LibreOffice ${JSON.stringify(lo)}`
					const panesAgree =
						ours.frozenRow === lo.frozenRow &&
						ours.frozenCol === lo.frozenCol &&
						ours.topRow === lo.topRow &&
						ours.leftCol === lo.leftCol
					// LibreOffice keeps a frozen split as a position on screen, which it reads back past its own scrolling
					// pane here (rows of custom heights and borders at the split): its panes are compared, not its split.
					const selfContradicting =
						(lo.splitRow > 0 && lo.topRow < lo.splitRow) || (lo.splitColumn > 0 && lo.leftCol < lo.splitColumn)

					compared++

					if (JSON.stringify(ours) === JSON.stringify(lo)) continue
					if (panesAgree || selfContradicting)
						console.log(`frozen-view accepted (LibreOffice's split read back from screen): ${line}`)
					else differ.push(line)
				}
			} finally {
				session.close()
			}
		}

		console.log(`frozen-view sheets compared: ${String(compared)}`)
		expect(differ).toEqual([])
	})
})
