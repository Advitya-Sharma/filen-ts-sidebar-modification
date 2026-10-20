import { SHEET_MAX_COLS, SHEET_MAX_ROWS, type SheetId } from "../../sheet/protocol"
import { defineSheetFeature, type SheetHandled } from "../../sheet/registry"
import type { XlsxSheetView } from "../model/types"
import type { XlsxDocument } from "./document"
import { setSheetView } from "./structure-view"

// Frozen panes (ECMA-376 Part 1 §18.3.1.66 pane, state "frozen"): the rows above and the columns left of the
// scrolling pane stay in view. Written into the sheet's first view: xSplit and ySplit count the frozen columns and
// rows from the view's first cell (MS-OI29500 §2.1.640 d), the pane's topLeftCell is the scrolling pane's first
// cell, and the active pane is the one that scrolls. The rows and columns frozen here start at the sheet's first, so
// a frozen axis starts the view there too.

// "view.freeze" is declared with the text sessions' own handler (core/sheet/text.ts), which keeps it in the
// session alone.

export function freezePanes(
	document: XlsxDocument,
	sheet: SheetId,
	rows: number,
	cols: number,
	topRow: number | undefined,
	leftCol: number | undefined,
	changes: { layout(sheet: SheetId): void }
): SheetHandled {
	const model = document.model(sheet)
	const top = topRow ?? rows
	const left = leftCol ?? cols

	if (model === null || document.entry(sheet)?.kind !== "worksheet") return { refused: { reason: "unsupported", detail: "sheet" } }
	if (
		![rows, cols, top, left].every(Number.isInteger) ||
		rows < 0 ||
		cols < 0 ||
		rows >= SHEET_MAX_ROWS ||
		cols >= SHEET_MAX_COLS ||
		top < rows ||
		left < cols ||
		top >= SHEET_MAX_ROWS ||
		left >= SHEET_MAX_COLS
	)
		return { refused: { reason: "tooLarge" } }

	const before = model.view
	const frozen = rows === 0 && cols === 0 ? null : { rows, cols, topLeftRow: top, topLeftCol: left }
	const current = before.frozen
	const viewTop = rows > 0 ? 0 : before.topRow
	const viewLeft = cols > 0 ? 0 : before.leftCol

	if (
		viewTop === before.topRow &&
		viewLeft === before.leftCol &&
		(frozen === null
			? current === null
			: current !== null &&
				current.rows === frozen.rows &&
				current.cols === frozen.cols &&
				current.topLeftRow === frozen.topLeftRow &&
				current.topLeftCol === frozen.topLeftCol)
	)
		return null

	const view: XlsxSheetView = { ...before, frozen, topRow: viewTop, leftCol: viewLeft }
	const apply = (): (() => void) => setSheetView(document.workbook, model, view)
	let restore = apply()

	changes.layout(sheet)

	return {
		bytes: 256,
		undo: sink => {
			restore()
			sink.layout(sheet)
		},
		redo: sink => {
			restore = apply()
			sink.layout(sheet)
		}
	}
}

export const freezeFeature = defineSheetFeature<XlsxDocument>({
	id: "freeze",
	commands: {
		"view.freeze": {
			info: { write: true },
			run: (context, command) =>
				freezePanes(context.doc, command.sheet, command.rows, command.cols, command.topRow, command.leftCol, context.changes)
		}
	}
})
