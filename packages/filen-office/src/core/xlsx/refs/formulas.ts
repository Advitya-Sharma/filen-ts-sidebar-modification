import { formulaTranslator, renameSheetInFormula, rewriteFormula, sameSheet, shiftFormula } from "../../sheet/refs/formulaRefs"
import type { AxisShift } from "./axis"
import { moveReference, type CellMove } from "./moveRefs"
import { rewriteStructured, type TableEdit } from "./structured"

// Formula text moved for one structure edit, as Excel moves it (ECMA-376 Part 1 §18.17; support.microsoft.com
// "Create a 3-D reference to the same cell range on multiple worksheets" for 3-D spans):
// - rows or columns of a sheet: references into it move; one whose area was deleted becomes #REF!. 3-D spans and
//   references into other workbooks stay as written.
// - a sheet deleted: references to it become #REF! (Sheet2!A1 -> #REF!A1); a 3-D span losing an end narrows to
//   the next sheet inside, and one losing every sheet becomes #REF!.
// - a sheet renamed: references name it anew (quoted when the name needs it).
// - a sheet moved: a 3-D span whose ends swapped order is written in sheet order.
// Structured references (Table1[Col]) name tables and columns, not cells: they become #REF! only when their
// table or a column they name is deleted.

export type FormulaOp =
	| { readonly kind: "axis"; readonly target: string; readonly shift: AxisShift }
	| { readonly kind: "deleteSheet"; readonly target: string; readonly order: readonly string[] }
	| { readonly kind: "renameSheet"; readonly target: string; readonly name: string }
	| { readonly kind: "moveSheet"; readonly order: readonly string[] }
	// Cells cut and pasted: `target` is the cut range's sheet.
	| { readonly kind: "moveCells"; readonly target: string; readonly move: CellMove }

// No sheet name holds U+0000 (XML 1.0 cannot carry it), so a reference given it as its sheet is one whose sheet
// was deleted; the rendered `'\0'!` becomes `#REF!`.
const DELETED = "\u0000"
const DELETED_PREFIX = `'${DELETED}'!`
// A sheet name no reference matches: the own sheet of a formula with none.
const NO_SHEET = "\u0000"

function includesName(upper: string, name: string): boolean {
	const target = name.toUpperCase()

	return upper.includes(target) || (target.includes("'") && upper.includes(target.replaceAll("'", "''")))
}

const ENTITY = /&(?:#x([0-9a-fA-F]+)|#([0-9]+)|(amp|lt|gt|quot|apos));/g
const NAMED: Readonly<Record<string, string>> = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" }

// Markup with its entities and character references read: a prefilter's view of the text a parser gives.
function unescaped(markup: string): string {
	if (!markup.includes("&")) return markup

	return markup.replace(ENTITY, (whole, hex: string | undefined, decimal: string | undefined, named: string | undefined) => {
		if (named !== undefined) return NAMED[named] ?? whole

		const code = hex !== undefined ? parseInt(hex, 16) : Number(decimal)

		return code <= 0x10ffff ? String.fromCodePoint(code) : whole
	})
}

function indexOfSheet(order: readonly string[], name: string): number {
	return order.findIndex(candidate => sameSheet(candidate, name))
}

export class FormulaShift {
	readonly op: FormulaOp
	private readonly tables: ReadonlyMap<string, TableEdit>
	private readonly broke: () => void

	// `tables`: tables deleted or losing columns, by upper-cased name; `broke` hears of each formula whose
	// structured references broke.
	constructor(op: FormulaOp, tables: ReadonlyMap<string, TableEdit> = new Map(), broke: () => void = () => undefined) {
		this.op = op
		this.tables = tables
		this.broke = broke
	}

	// Whether markup (sheet names in attributes or formulas, with entities) names the edited sheet.
	namesTarget(markup: string): boolean {
		const op = this.op

		if (op.kind === "moveSheet") return false

		const upper = unescaped(markup).toUpperCase()

		return includesName(upper, op.target) || (op.kind === "moveCells" && includesName(upper, op.move.to))
	}

	// mayChange for formulas held in markup, where a name or an operator may be written as an entity.
	markupMayChange(markup: string, own: string | null): boolean {
		return this.mayChange(unescaped(markup), own)
	}

	// Whether the edit can change a formula on `own` (null: a formula with no sheet of its own).
	mayChange(text: string, own: string | null): boolean {
		if (this.tables.size > 0 && text.includes("[")) return true

		const op = this.op

		if (op.kind === "axis" && own !== null && sameSheet(own, op.target)) return true
		if (op.kind === "moveCells" && own !== null && (sameSheet(own, op.target) || sameSheet(own, op.move.to))) return true
		if (!text.includes("!")) return false
		if (op.kind === "moveSheet") return text.includes(":")

		const upper = text.toUpperCase()

		return (
			includesName(upper, op.target) ||
			(op.kind === "deleteSheet" && upper.includes(":")) ||
			(op.kind === "moveCells" && includesName(upper, op.move.to))
		)
	}

	// The formula after the edit; the same string when nothing changed. `table`: the table whose formula this is
	// (unqualified structured references name it).
	shift(text: string, own: string | null, table: string | null = null): string {
		let out = this.mayChange(text, own) ? this.shiftReferences(text, own) : text

		if (this.tables.size > 0 && out.includes("[")) {
			const rewritten = rewriteStructured(out, table, this.tables)

			if (rewritten !== out) this.broke()

			out = rewritten
		}

		return out
	}

	// A formula written relative to an anchor cell (conditional formats, validations): moved from the anchor to
	// the cell `rows`/`cols` away first, then shifted.
	shiftAnchored(text: string, own: string | null, rows: number, cols: number): string {
		return this.shift(rows === 0 && cols === 0 ? text : formulaTranslator(text)(rows, cols), own)
	}

	private shiftReferences(text: string, own: string | null): string {
		const op = this.op

		switch (op.kind) {
			case "axis":
				return shiftFormula(text, own ?? NO_SHEET, op.target, op.shift).formula
			case "renameSheet":
				return renameSheetInFormula(text, op.target, op.name)
			case "deleteSheet": {
				const deleted = indexOfSheet(op.order, op.target)
				const out = rewriteFormula(text, ref => {
					if (ref.external || ref.sheet === null) return undefined

					if (ref.lastSheet === null) return sameSheet(ref.sheet, op.target) ? { ...ref, sheet: DELETED } : undefined

					const first = indexOfSheet(op.order, ref.sheet)
					const last = indexOfSheet(op.order, ref.lastSheet)

					if (first < 0 || last < 0 || (deleted !== Math.min(first, last) && deleted !== Math.max(first, last))) return undefined

					const low = Math.min(first, last)
					const high = Math.max(first, last)

					if (low === high) return { ...ref, sheet: DELETED, lastSheet: null }

					const from = deleted === low ? low + 1 : low
					const to = deleted === high ? high - 1 : high

					return {
						...ref,
						sheet: op.order[from] ?? ref.sheet,
						lastSheet: from === to ? null : (op.order[to] ?? ref.lastSheet)
					}
				})

				return out.includes(DELETED_PREFIX) ? out.replaceAll(DELETED_PREFIX, "#REF!") : out
			}
			case "moveCells":
				return rewriteFormula(text, ref => moveReference(ref, own, op.move))
			case "moveSheet":
				return rewriteFormula(text, ref => {
					if (ref.external || ref.sheet === null || ref.lastSheet === null) return undefined

					const first = indexOfSheet(op.order, ref.sheet)
					const last = indexOfSheet(op.order, ref.lastSheet)

					return first >= 0 && last >= 0 && first > last ? { ...ref, sheet: ref.lastSheet, lastSheet: ref.sheet } : undefined
				})
		}
	}
}

// Whether a defined name's formula has a reference without a sheet (relative to the sheet it is used on), which
// no structure edit can move unambiguously.
export function hasUnqualifiedArea(text: string): boolean {
	let found = false

	rewriteFormula(text, ref => {
		if (!ref.external && ref.sheet === null && ref.area !== null) found = true

		return undefined
	})

	return found
}
