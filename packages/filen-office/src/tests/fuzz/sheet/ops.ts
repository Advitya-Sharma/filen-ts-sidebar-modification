import {
	SHEET_MAX_COLS,
	SHEET_MAX_ROWS,
	type SheetAxis,
	type SheetCommand,
	type SheetCommandResult,
	type SheetFormatPatch,
	type SheetId,
	type SheetRange,
	type SheetRefusal
} from "../../../core/sheet/protocol"
import type { SheetValidationRule } from "../../../core/sheet/dv/xlsx"
import type { SheetFilterCriteria } from "../../../core/sheet/filter/criteria"
import type { SheetSession } from "../../../core/sheet/session"
import type { XlsxDocument } from "../../../core/xlsx/edit/document"
import type { Rand } from "../config"

// The sheet fuzz's operations: every command family the xlsx session registers (cells, format, axis, shift, resize,
// sheet, sheets, hide, freeze, merge, protection, clipboard, annotations, validation, sort, filter) and undo/redo, with
// targets near what a sheet holds and inputs that reach the edge cases (formulas that do not read, text that looks
// like a number, characters XML cannot hold, the grid's edges).

// Refusal reasons the protocol names (protocol.ts SheetRefusalReasons). Anything else a command answers with is
// an untyped refusal.
export const KNOWN_REFUSALS: ReadonlySet<string> = new Set([
	"readOnly",
	"tooLarge",
	"sheetName",
	"lastSheet",
	"formula",
	"arrayFormula",
	"tableHeader",
	"structureLocked",
	"pivot",
	"merged",
	"protectedSheet",
	"protectedWorkbook",
	"encoding",
	"styleLimit",
	"unsupported",
	"mergeLosesValues",
	"notesUnavailable",
	"authorNeeded",
	"notOwnComment",
	"linkTarget",
	"filterCriteria"
])

export type Outcome =
	{ readonly kind: "applied" } | { readonly kind: "refused"; readonly refusal: SheetRefusal } | { readonly kind: "noop" }

export interface OpContext {
	readonly session: SheetSession<XlsxDocument>
	readonly rand: Rand
	// Cells at the grid's last rows and columns too.
	readonly edges: boolean
	// Formulas whose result is a range ("=A1:A3", "={1,2}"), where a single value is expected.
	readonly ranges: boolean
}

export interface Op {
	readonly name: string
	readonly family: string
	readonly weight: number
	// The call to make, decided now (from the model as it stands), run when the harness awaits it.
	plan(context: OpContext): (() => Promise<Outcome>) | null
}

function int(rand: Rand, from: number, to: number): number {
	return from + Math.floor(rand() * (to - from + 1))
}

function pick<T>(rand: Rand, items: readonly T[]): T {
	return items[Math.floor(rand() * items.length)] as T
}

function outcome(result: SheetCommandResult): Outcome {
	return result.ok ? { kind: "applied" } : { kind: "refused", refusal: result.refusal }
}

interface Target {
	readonly sheet: SheetId
	// Past the last row and column with content, at least 1.
	readonly rows: number
	readonly cols: number
}

// A parsed worksheet to edit, with its extent.
function target(context: OpContext): Target | null {
	const sheets = context.session.document.worksheets().filter(entry => entry.model !== null)

	if (sheets.length === 0) return null

	const entry = pick(context.rand, sheets)
	const store = entry.model?.store

	return {
		sheet: entry.id,
		rows: Math.max(1, Math.min(store?.rowExtent ?? 1, 400)),
		cols: Math.max(1, Math.min(store?.colExtent ?? 1, 40))
	}
}

// A cell near the content: mostly inside it, sometimes just past it, rarely at the grid's edges.
function cell(rand: Rand, at: Target, edges: boolean): { row: number; col: number } {
	const roll = edges ? rand() : 1

	if (roll < 0.03) return { row: SHEET_MAX_ROWS - 1 - int(rand, 0, 2), col: int(rand, 0, at.cols) }
	if (roll < 0.06) return { row: int(rand, 0, at.rows), col: SHEET_MAX_COLS - 1 - int(rand, 0, 2) }

	return { row: int(rand, 0, at.rows + 2), col: int(rand, 0, at.cols + 1) }
}

function range(rand: Rand, at: Target, edges: boolean): SheetRange {
	const start = cell(rand, at, edges)
	const endRow = Math.min(SHEET_MAX_ROWS - 1, start.row + (rand() < 0.7 ? int(rand, 0, 3) : int(rand, 0, 40)))
	const endCol = Math.min(SHEET_MAX_COLS - 1, start.col + int(rand, 0, 3))

	return { startRow: start.row, startCol: start.col, endRow, endCol }
}

function a1(row: number, col: number): string {
	let name = ""

	for (let index = col + 1; index > 0; index = Math.floor((index - 1) / 26)) name = String.fromCharCode(65 + ((index - 1) % 26)) + name

	return `${name}${row + 1}`
}

const FORMULAS = ["=SUM(A1", "=1/0", "=NOSUCHFUNC(1)", '=CONCAT("a","b")', "=Total+1", "=Calc!A1+1", "=Data!B2", "=ROW()"]
const RANGE_FORMULAS = ["=A1:A3", "={1,2}"]

// What a user types into a cell.
function input(rand: Rand, at: Target, ranges: boolean): string {
	const near = (): string => a1(int(rand, 0, at.rows), int(rand, 0, at.cols))
	const choices: (() => string)[] = [
		() => String(int(rand, -1000, 100_000)),
		() => (rand() * 1e6 - 5e5).toFixed(int(rand, 0, 6)),
		() => pick(rand, ["1,234.5", "50%", "-0", "1e308", "1E-5", "  12  ", "0012", "$5", "(3)", "1/2"]),
		() => pick(rand, ["hello", "Grüße", "漢字", "a\tb", "line\nbreak", "_x000D_", "<&>\"'", "\u0001ctl", "=", "'=1+1", "'007", " "]),
		() => pick(rand, ["TRUE", "false", "#N/A", "#DIV/0!", "2024-03-05", "14:30", "2024-02-29 08:15", "1900-02-29"]),
		() => `=${near()}*2`,
		() => `=SUM(${near()}:${near()})`,
		() => `=IF(${near()}>0,"pos","neg")`,
		() => `=${near()}&"x"`,
		() => (ranges && rand() < 0.2 ? pick(rand, RANGE_FORMULAS) : pick(rand, FORMULAS)),
		() => "",
		() => "x".repeat(int(rand, 200, 33_000))
	]
	const weights = [4, 2, 2, 3, 2, 3, 2, 1, 1, 2, 2, 0.1]
	let roll = rand() * weights.reduce((sum, weight) => sum + weight, 0)

	for (let i = 0; i < choices.length; i++) {
		roll -= weights[i] ?? 0

		if (roll < 0) return (choices[i] as () => string)()
	}

	return "1"
}

function color(rand: Rand): SheetFormatPatch["color"] {
	return rand() < 0.5
		? int(rand, 0, 0xffffff)
		: { theme: int(rand, 0, 11), ...(rand() < 0.5 && { tint: Math.round((rand() * 2 - 1) * 100) / 100 }) }
}

function formatPatch(rand: Rand): SheetFormatPatch {
	const patch: Record<string, unknown> = {}
	const setters: (() => void)[] = [
		() => (patch["bold"] = rand() < 0.6),
		() => (patch["italic"] = rand() < 0.6),
		() => (patch["strike"] = rand() < 0.5),
		() => (patch["underline"] = pick(rand, ["single", "double", "singleAccounting", null])),
		() => (patch["script"] = pick(rand, ["superscript", "subscript", null])),
		() => (patch["font"] = pick(rand, ["Arial", "Calibri", "Times New Roman", null])),
		() => (patch["size"] = pick(rand, [8, 11, 14.5, 36, null])),
		() => (patch["color"] = rand() < 0.2 ? null : color(rand)),
		() => (patch["fill"] = rand() < 0.2 ? null : color(rand)),
		() => (patch["align"] = pick(rand, ["left", "center", "right", "justify", null])),
		() => (patch["valign"] = pick(rand, ["top", "center", "bottom", null])),
		() => (patch["wrap"] = rand() < 0.5),
		() => (patch["indent"] = int(rand, 0, 5)),
		() => (patch["rotation"] = pick(rand, [0, 45, 90, 255])),
		() =>
			(patch["numFmt"] = pick(rand, [
				"General",
				"0.00",
				"#,##0",
				"0%",
				"m/d/yyyy",
				"yyyy-mm-dd hh:mm",
				"@",
				'[Red]"x"0;-0',
				"0.00E+00"
			])),
		() => (patch["locked"] = rand() < 0.5)
	]

	for (let i = int(rand, 1, 3); i > 0; i--) pick(rand, setters)()

	return patch as SheetFormatPatch
}

function command(context: OpContext, build: (at: Target) => SheetCommand | null): (() => Promise<Outcome>) | null {
	const at = target(context)
	const made = at === null ? null : build(at)

	return made === null ? null : async () => outcome(await context.session.command(made))
}

// A data validation as the dialog sends it: every kind and operator, bounds typed as numbers, dates, text or
// formulas (some that do not read), lists typed or referenced, messages up to and past Excel's limits.
function rand2(rand: Rand): boolean {
	return rand() < 0.2
}

function validationRule(rand: Rand): SheetValidationRule {
	const type = pick(rand, ["none", "whole", "decimal", "list", "date", "time", "textLength", "custom"] as const)
	const operator = pick(rand, [
		"between",
		"notBetween",
		"equal",
		"notEqual",
		"lessThan",
		"lessThanOrEqual",
		"greaterThan",
		"greaterThanOrEqual"
	] as const)
	const bound = (): string =>
		pick(rand, ["1", "-2.5", "1000000", "2024-02-29", "12:30", "=A1", "=$B$2*2", "=SUM(", "abc", "", "=TODAY()"])
	const value1 =
		type === "list"
			? pick(rand, ["a,b,c", "Yes,No", "=$A$1:$A$5", "=A1:B1", '="x,y"', "x".repeat(256), "=SUM(", "1,2,3"])
			: type === "custom"
				? pick(rand, ["=A1>0", "=ISNUMBER(A1)", "=COUNTIF($A:$A,A1)=1", "=SUM(", "=LEN(A1)<5"])
				: bound()
	const text = (limit: number): string => pick(rand, ["", "Pick", "a".repeat(limit), "a".repeat(limit + 1), "\u0001", "Ünï\ncode"])

	return {
		type,
		operator,
		value1,
		value2: bound(),
		allowBlank: rand() < 0.7,
		inCellDropdown: rand() < 0.8,
		showInputMessage: rand() < 0.5,
		promptTitle: text(32),
		prompt: text(255),
		showErrorMessage: rand() < 0.8,
		errorStyle: pick(rand, ["stop", "warning", "information"] as const),
		errorTitle: text(32),
		error: text(225)
	}
}

// Criteria of every kind an autofilter column takes, with values near the fixtures' cells.
function criteria(rand: Rand): SheetFilterCriteria | null {
	switch (int(rand, 0, 6)) {
		case 0:
			return null
		case 1:
			return {
				kind: "values",
				values: pick(rand, [["1"], ["a", "x"], [], ["TRUE", "#DIV/0!"]]),
				blank: rand() < 0.5,
				dates: rand() < 0.2 ? [{ grouping: "month", year: 2024, month: 2 }] : []
			}
		case 2:
			return {
				kind: "custom",
				and: rand() < 0.5,
				conditions: [
					{
						operator: pick(rand, ["equal", "notEqual", "greaterThan", "lessThanOrEqual"] as const),
						value: pick(rand, ["1", "*a*", "b?", "", " ", "45000"])
					},
					...(rand() < 0.5
						? [{ operator: pick(rand, ["greaterThanOrEqual", "lessThan"] as const), value: String(int(rand, -5, 50)) }]
						: [])
				]
			}
		case 3:
			return { kind: "top10", top: rand() < 0.5, percent: rand() < 0.3, value: int(rand, 1, 20) }
		case 4:
			return {
				kind: "dynamic",
				type: pick(rand, ["aboveAverage", "belowAverage", "today", "thisMonth", "lastYear", "Q2", "M11"] as const)
			}
		case 5:
			return { kind: "color", cell: rand() < 0.7, color: rand() < 0.3 ? null : pick(rand, [0xffff00, 0x000000, 0xff0000]) }
		default:
			return { kind: "values", values: [String(int(rand, 0, 9))], blank: false, dates: [] }
	}
}

const NOTE_TEXTS = ["Fuzz:\nnote", "", "a & <b>", "line\nbreak", "\u0001ctl", "x".repeat(300)]
const LINKS: readonly ({ readonly url: string } | { readonly location: string })[] = [
	{ url: "https://filen.io/a?b=1&c" },
	{ url: "mailto:a@b.c" },
	{ url: "javascript:alert(1)" },
	{ location: "A1" },
	{ location: "Data!B2" },
	{ location: "Nowhere!A1" }
]

// A comment of the thread on a cell (the first, or any), null when the cell has none.
async function threadComment(context: OpContext, sheet: SheetId, at: { row: number; col: number }, first: boolean): Promise<string | null> {
	const thread = await context.session.query("thread.get", { sheet, row: at.row, col: at.col })
	const comments = thread?.comments ?? []

	return (first ? comments[0] : comments[comments.length - 1])?.id ?? null
}

const SHEET_NAMES = ["New", "Data", "data", "Calc", "Sheet 9", "a/b", "", "x".repeat(32), "History", "'quoted", "Ünïcode", "Data (2)"]

export const OPS: readonly Op[] = [
	{
		name: "cells.set",
		family: "cells",
		weight: 8,
		plan: context =>
			command(context, at => {
				const count = context.rand() < 0.9 ? int(context.rand, 1, 4) : int(context.rand, 5, 30)
				const rows: number[] = []
				const cols: number[] = []
				const inputs: string[] = []

				for (let i = 0; i < count; i++) {
					const place = cell(context.rand, at, context.edges)

					rows.push(place.row)
					cols.push(place.col)
					inputs.push(input(context.rand, at, context.ranges))
				}

				return { type: "cells.set", sheet: at.sheet, cells: { rows: Uint32Array.from(rows), cols: Uint16Array.from(cols), inputs } }
			})
	},
	{
		name: "cells.clear",
		family: "cells",
		weight: 2,
		plan: context =>
			command(context, at => ({
				type: "cells.clear",
				sheet: at.sheet,
				ranges: [range(context.rand, at, context.edges)],
				what: pick(context.rand, ["contents", "formats", "all"] as const)
			}))
	},
	{
		name: "format.apply",
		family: "format",
		weight: 4,
		plan: context =>
			command(context, at => ({
				type: "format.apply",
				sheet: at.sheet,
				ranges:
					context.rand() < 0.8
						? [range(context.rand, at, context.edges)]
						: [range(context.rand, at, context.edges), range(context.rand, at, context.edges)],
				patch: formatPatch(context.rand)
			}))
	},
	{
		name: "axis.insert",
		family: "axis",
		weight: 2,
		plan: context =>
			command(context, at => {
				const axis: SheetAxis = context.rand() < 0.6 ? "rows" : "cols"

				return {
					type: "axis.insert",
					sheet: at.sheet,
					axis,
					at: int(context.rand, 0, axis === "rows" ? at.rows + 1 : at.cols + 1),
					count: context.rand() < 0.95 ? int(context.rand, 1, 3) : pick(context.rand, [0, 2_000_000])
				}
			})
	},
	{
		name: "axis.delete",
		family: "axis",
		weight: 2,
		plan: context =>
			command(context, at => {
				const axis: SheetAxis = context.rand() < 0.6 ? "rows" : "cols"

				return {
					type: "axis.delete",
					sheet: at.sheet,
					axis,
					at: int(context.rand, 0, axis === "rows" ? at.rows : at.cols),
					count: int(context.rand, 1, 3)
				}
			})
	},
	{
		name: "cells.shift",
		family: "shift",
		weight: 2,
		plan: context =>
			command(context, at => ({
				type: "cells.shift",
				sheet: at.sheet,
				range: range(context.rand, at, context.edges),
				edit: context.rand() < 0.5 ? "insert" : "delete",
				axis: context.rand() < 0.5 ? "rows" : "cols"
			}))
	},
	{
		name: "axis.resize",
		family: "resize",
		weight: 2,
		plan: context =>
			command(context, at => {
				const axis: SheetAxis = context.rand() < 0.5 ? "rows" : "cols"
				const count = int(context.rand, 1, 3)
				const indices = Array.from({ length: count }, () => int(context.rand, 0, axis === "rows" ? at.rows + 2 : at.cols + 2))
				const sizes = indices.map(() => (context.rand() < 0.2 ? Number.NaN : pick(context.rand, [0, 1, 20, 64.5, 300, 1000])))

				return { type: "axis.resize", sheet: at.sheet, axis, indices: Uint32Array.from(indices), sizes: Float32Array.from(sizes) }
			})
	},
	{
		name: "sheet.add",
		family: "sheet",
		weight: 0.6,
		plan: context => {
			const count = context.session.document.workbook.sheets.length
			const name = context.rand() < 0.4 ? undefined : pick(context.rand, SHEET_NAMES)
			const index = context.rand() < 0.5 ? undefined : int(context.rand, 0, count)
			const made: SheetCommand = { type: "sheet.add", ...(name !== undefined && { name }), ...(index !== undefined && { index }) }

			return async () => outcome(await context.session.command(made))
		}
	},
	{
		name: "sheet.rename",
		family: "sheet",
		weight: 0.6,
		plan: context =>
			command(context, at => ({
				type: "sheet.rename",
				sheet: at.sheet,
				name:
					context.rand() < 0.5 ? `${pick(context.rand, SHEET_NAMES)}${int(context.rand, 0, 9)}` : pick(context.rand, SHEET_NAMES)
			}))
	},
	{
		name: "sheet.delete",
		family: "sheets",
		weight: 0.3,
		plan: context => command(context, at => ({ type: "sheet.delete", sheet: at.sheet }))
	},
	{
		name: "sheet.move",
		family: "sheets",
		weight: 0.4,
		plan: context =>
			command(context, at => ({
				type: "sheet.move",
				sheet: at.sheet,
				to: int(context.rand, 0, context.session.document.workbook.sheets.length - 1)
			}))
	},
	{
		name: "sheet.duplicate",
		family: "sheets",
		weight: 0.3,
		plan: context => command(context, at => ({ type: "sheet.duplicate", sheet: at.sheet }))
	},
	{
		name: "sheet.state",
		family: "sheets",
		weight: 0.3,
		plan: context =>
			command(context, at => ({
				type: "sheet.state",
				sheet: at.sheet,
				state: pick(context.rand, ["visible", "hidden", "veryHidden"] as const)
			}))
	},
	{
		name: "sheet.tabColor",
		family: "sheets",
		weight: 0.2,
		plan: context =>
			command(context, at => ({
				type: "sheet.tabColor",
				sheet: at.sheet,
				color: pick(context.rand, [null, 0xff0000, { theme: 4 }, { theme: 5, tint: -0.25 }])
			}))
	},
	{
		name: "axis.hide",
		family: "hide",
		weight: 0.8,
		plan: context =>
			command(context, at => {
				const axis: SheetAxis = context.rand() < 0.6 ? "rows" : "cols"
				const start = int(context.rand, 0, axis === "rows" ? at.rows + 1 : at.cols + 1)

				return {
					type: "axis.hide",
					sheet: at.sheet,
					axis,
					spans: [{ start, end: start + int(context.rand, 0, 3) }],
					hidden: context.rand() < 0.6
				}
			})
	},
	{
		name: "view.freeze",
		family: "freeze",
		weight: 0.5,
		plan: context =>
			command(context, at =>
				context.rand() < 0.25
					? { type: "view.freeze", sheet: at.sheet, rows: 0, cols: 0 }
					: { type: "view.freeze", sheet: at.sheet, rows: int(context.rand, 0, 5), cols: int(context.rand, 0, 3) }
			)
	},
	{
		name: "validation.set",
		family: "validation",
		weight: 0.6,
		plan: context =>
			command(context, at => ({
				type: "validation.set",
				sheet: at.sheet,
				ranges: rand2(context.rand)
					? [range(context.rand, at, context.edges), range(context.rand, at, false)]
					: [range(context.rand, at, context.edges)],
				rule: context.rand() < 0.25 ? null : validationRule(context.rand),
				...(context.rand() < 0.5 && { at: cell(context.rand, at, false) })
			}))
	},
	{
		name: "merge.cells",
		family: "merge",
		weight: 0.8,
		plan: context =>
			command(context, at => ({
				type: "merge.cells",
				sheet: at.sheet,
				ranges: [range(context.rand, at, false)],
				mode: pick(context.rand, ["merge", "center", "across"] as const),
				discard: context.rand() < 0.7
			}))
	},
	{
		name: "sort.apply",
		family: "sort",
		weight: 0.6,
		plan: context =>
			command(context, at => {
				const sorted = range(context.rand, at, false)
				const rows = context.rand() < 0.85
				const keys = Array.from({ length: int(context.rand, 1, 3) }, () => ({
					index: rows ? int(context.rand, sorted.startCol, sorted.endCol) : int(context.rand, sorted.startRow, sorted.endRow),
					descending: context.rand() < 0.5
				}))

				return {
					type: "sort.apply",
					sheet: at.sheet,
					range: sorted,
					orientation: rows ? "rows" : "cols",
					header: context.rand() < 0.5,
					caseSensitive: context.rand() < 0.3,
					keys
				}
			})
	},
	{
		name: "filter.toggle",
		family: "filter",
		weight: 0.4,
		plan: context =>
			command(context, at => {
				const filtered = range(context.rand, at, false)

				return {
					type: "filter.toggle",
					sheet: at.sheet,
					range: filtered,
					active: { row: filtered.startRow, col: filtered.startCol }
				}
			})
	},
	{
		name: "filter.column",
		family: "filter",
		weight: 0.6,
		plan: context =>
			command(context, at => {
				const header = cell(context.rand, at, false)

				return { type: "filter.column", sheet: at.sheet, row: header.row, col: header.col, criteria: criteria(context.rand) }
			})
	},
	{
		name: "filter.clear",
		family: "filter",
		weight: 0.2,
		plan: context =>
			command(context, at => {
				const header = cell(context.rand, at, false)

				return { type: context.rand() < 0.5 ? "filter.clear" : "filter.reapply", sheet: at.sheet, row: header.row, col: header.col }
			})
	},
	{
		name: "merge.unmerge",
		family: "merge",
		weight: 0.4,
		plan: context => command(context, at => ({ type: "merge.unmerge", sheet: at.sheet, ranges: [range(context.rand, at, false)] }))
	},
	{
		name: "sheet.protect",
		family: "protection",
		weight: 0.2,
		plan: context =>
			command(context, at =>
				context.rand() < 0.5
					? { type: "sheet.unprotect", sheet: at.sheet }
					: {
							type: "sheet.protect",
							sheet: at.sheet,
							allows: [
								"selectLockedCells",
								"selectUnlockedCells",
								...(context.rand() < 0.5 ? (["insertRows", "formatCells"] as const) : [])
							],
							...(context.rand() < 0.5 && { hash: { legacy: "83AF" } })
						}
			)
	},
	{
		name: "clipboard.paste",
		family: "clipboard",
		weight: 3,
		plan: context => {
			const from = target(context)
			const to = target(context)

			if (from === null || to === null) return null

			const source = range(context.rand, from, context.edges)
			const cut = context.rand() < 0.3
			const at = cell(context.rand, to, context.edges)
			const foreign = context.rand() < 0.25

			return async () => {
				const copied = await context.session.query("copy", { sheet: from.sheet, range: source, cut })

				if ("refused" in copied) return { kind: "refused", refusal: { reason: copied.refused } }

				const text = foreign ? `${copied.text}\t${input(context.rand, to, context.ranges)}` : copied.text

				return outcome(
					await context.session.command({
						type: "clipboard.paste",
						sheet: to.sheet,
						at,
						text,
						...(!foreign && { clip: copied.clip })
					})
				)
			}
		}
	},
	{
		name: "note.set",
		family: "annotations",
		weight: 1,
		plan: context =>
			command(context, at => ({
				type: "note.set",
				sheet: at.sheet,
				...cell(context.rand, at, context.edges),
				text: pick(context.rand, NOTE_TEXTS)
			}))
	},
	{
		name: "note.delete",
		family: "annotations",
		weight: 0.5,
		plan: context =>
			command(context, at => ({ type: "note.delete", sheet: at.sheet, ranges: [range(context.rand, at, context.edges)] }))
	},
	{
		name: "note.show",
		family: "annotations",
		weight: 0.3,
		plan: context =>
			command(context, at => ({ type: "note.show", sheet: at.sheet, ...cell(context.rand, at, false), shown: context.rand() < 0.6 }))
	},
	{
		name: "thread.add",
		family: "annotations",
		weight: 0.8,
		plan: context =>
			command(context, at => ({
				type: "thread.add",
				sheet: at.sheet,
				...cell(context.rand, at, context.edges),
				text: pick(context.rand, NOTE_TEXTS.slice(2))
			}))
	},
	{
		name: "thread.edit",
		family: "annotations",
		weight: 0.3,
		plan: context => {
			const at = target(context)

			if (at === null) return null

			const place = cell(context.rand, at, false)
			const first = context.rand() < 0.5
			const text = pick(context.rand, NOTE_TEXTS.slice(2))

			return async () => {
				const id = await threadComment(context, at.sheet, place, first)

				return id === null
					? { kind: "noop" }
					: outcome(await context.session.command({ type: "thread.edit", sheet: at.sheet, id, text }))
			}
		}
	},
	{
		name: "thread.delete",
		family: "annotations",
		weight: 0.3,
		plan: context => {
			const at = target(context)

			if (at === null) return null

			const place = cell(context.rand, at, false)
			const first = context.rand() < 0.5

			return async () => {
				const id = await threadComment(context, at.sheet, place, first)

				return id === null
					? { kind: "noop" }
					: outcome(await context.session.command({ type: "thread.delete", sheet: at.sheet, id }))
			}
		}
	},
	{
		name: "link.set",
		family: "annotations",
		weight: 0.8,
		plan: context =>
			command(context, at => ({
				type: "link.set",
				sheet: at.sheet,
				ranges: [range(context.rand, at, context.edges)],
				target: pick(context.rand, LINKS),
				...(context.rand() < 0.5 && { text: pick(context.rand, ["Link", "=1+1", "", "123"]) }),
				...(context.rand() < 0.3 && { tooltip: pick(context.rand, ["tip", "x".repeat(300)]) })
			}))
	},
	{
		name: "link.remove",
		family: "annotations",
		weight: 0.5,
		plan: context =>
			command(context, at => ({ type: "link.remove", sheet: at.sheet, ranges: [range(context.rand, at, context.edges)] }))
	},
	{
		name: "undo",
		family: "history",
		weight: 2,
		plan: context => async () => {
			const result = await context.session.undo()

			return result.ok ? { kind: "applied" } : { kind: "refused", refusal: result.refusal }
		}
	},
	{
		name: "redo",
		family: "history",
		weight: 1,
		plan: context => async () => {
			const result = await context.session.redo()

			return result.ok ? { kind: "applied" } : { kind: "refused", refusal: result.refusal }
		}
	}
]

export const FAMILIES: readonly string[] = [...new Set(OPS.map(op => op.family))]

// Picks ops by weight, from all of them or the families and names given.
export class OpPicker {
	private readonly ops: readonly Op[]
	private readonly total: number

	constructor(families?: readonly string[], names?: readonly string[]) {
		this.ops = OPS.filter(
			op => (families === undefined || families.includes(op.family)) && (names === undefined || names.includes(op.name))
		)

		if (this.ops.length === 0) throw new Error("no sheet fuzz op matches")

		this.total = this.ops.reduce((sum, op) => sum + op.weight, 0)
	}

	pick(rand: Rand): Op {
		let roll = rand() * this.total

		for (const op of this.ops) {
			roll -= op.weight

			if (roll < 0) return op
		}

		return this.ops.at(-1) as Op
	}
}
