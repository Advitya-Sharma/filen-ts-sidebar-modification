import { SHEET_MAX_COLS, SHEET_MAX_EDIT_CELLS, SHEET_MAX_ROWS, type SheetCommandOf } from "../protocol"
import { defineSheetFeature, type SheetCommandContext, type SheetHandled } from "../registry"
import { textFeature, type TextSheetDocument } from "../text"

// Excel's Insert and Delete of cells (support.microsoft.com, "Insert or delete cells, rows, and columns"), and the
// text files' handler of it. A workbook's (core/xlsx/edit/shiftCells.ts) moves what refers to the cells too; a text
// file holds nothing that does, so its band's texts move as one cell entry, undone as one.

declare module "../protocol" {
	interface SheetCommandMap {
		// The cells of `range` inserted (the cells there and after them in its columns move down, axis "rows", or in
		// its rows right, axis "cols") or deleted (the cells after them move up or left). A range across every column
		// (row) inserts or deletes its whole rows (columns).
		"cells.shift": {
			readonly sheet: SheetId
			readonly range: SheetRange
			readonly edit: "insert" | "delete"
			readonly axis: SheetAxis
		}
	}
}

function shiftTexts(
	context: SheetCommandContext<TextSheetDocument>,
	command: SheetCommandOf<"cells.shift">
): SheetHandled | Promise<SheetHandled> {
	const document = context.doc
	const { startRow, startCol, endRow, endCol } = command.range

	if (
		![startRow, startCol, endRow, endCol].every(Number.isInteger) ||
		startRow < 0 ||
		startCol < 0 ||
		startRow > endRow ||
		startCol > endCol ||
		endRow >= SHEET_MAX_ROWS ||
		endCol >= SHEET_MAX_COLS
	)
		return { refused: { reason: "tooLarge" } }

	const rows = command.axis === "rows"
	const at = rows ? startRow : startCol
	const count = rows ? endRow - startRow + 1 : endCol - startCol + 1
	const insert = command.edit === "insert"

	// Every column (row): whole rows (columns).
	if (rows ? startCol === 0 && endCol === SHEET_MAX_COLS - 1 : startRow === 0 && endRow === SHEET_MAX_ROWS - 1) {
		const whole = { sheet: command.sheet, axis: command.axis, at, count }

		if (insert) return textFeature.commands?.["axis.insert"]?.run(context, { type: "axis.insert", ...whole }) ?? null

		return textFeature.commands?.["axis.delete"]?.run(context, { type: "axis.delete", ...whole }) ?? null
	}

	const limit = rows ? SHEET_MAX_ROWS : SHEET_MAX_COLS
	const lines = document.rows
	const changedRows: number[] = []
	const changedCols: number[] = []
	const inputs: string[] = []
	const bandEnd = rows ? Math.min(endCol, document.width - 1) : Math.min(endRow, lines.length - 1)

	for (let cross = rows ? startCol : startRow; cross <= bandEnd; cross++) {
		// The band's line: a column (cells moving down or up) or a row (right or left).
		const length = rows ? lines.length : (lines[cross]?.length ?? 0)
		const read = (index: number): string => (rows ? document.text(index, cross) : document.text(cross, index))

		if (insert)
			for (let index = Math.max(at, limit - count); index < length; index++)
				if (read(index) !== "") return { refused: { reason: "tooLarge", detail: "cells" } }

		const end = insert ? Math.min(limit, length + count) : length

		for (let index = at; index < end; index++) {
			const from = insert ? index - count : index + count
			const next = insert ? (from >= at ? read(from) : "") : from < length ? read(from) : ""

			if (next === read(index)) continue
			if (inputs.length === SHEET_MAX_EDIT_CELLS) return { refused: { reason: "tooLarge" } }

			changedRows.push(rows ? index : cross)
			changedCols.push(rows ? cross : index)
			inputs.push(next)
		}
	}

	const set = textFeature.commands?.["cells.set"]

	if (inputs.length === 0 || set === undefined) return null

	return set.run(context, {
		type: "cells.set",
		sheet: command.sheet,
		cells: { rows: Uint32Array.from(changedRows), cols: Uint16Array.from(changedCols), inputs }
	})
}

export const textShiftFeature = defineSheetFeature<TextSheetDocument>({
	id: "textShift",
	commands: {
		"cells.shift": { info: { write: true }, run: shiftTexts }
	}
})
