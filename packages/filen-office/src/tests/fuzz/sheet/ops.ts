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
import type { SheetSession } from "../../../core/sheet/session"
import type { XlsxDocument } from "../../../core/xlsx/edit/document"
import type { Rand } from "../config"

// The sheet fuzz's operations: every command family the xlsx session registers (cells, format, axis, resize,
// sheet, sheets, hide, freeze, merge, protection, clipboard) and undo/redo, with targets near what a sheet holds and inputs that reach the edge cases
// (formulas that do not read, text that looks like a number, characters XML cannot hold, the grid's edges).

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
	"mergeLosesValues"
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
