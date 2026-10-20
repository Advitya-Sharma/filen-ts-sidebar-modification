import type { SheetRange } from "../protocol"
import { defineSheetFeature, type SheetReadContext } from "../registry"
import { TEXT_SHEET, textFeature, type TextSheetDocument } from "../text"
import { findCells, replaceCells, type FindDocument, type FindSheet } from "./find"

// Find and Replace over a text file's one sheet: what a cell shows is what was typed.

class TextFindSheet implements FindSheet {
	private readonly document: TextSheetDocument

	constructor(document: TextSheetDocument) {
		this.document = document
	}

	scan(range: SheetRange, visit: (row: number, col: number) => boolean): boolean {
		const rows = this.document.rows
		const last = Math.min(range.endRow, rows.length - 1)

		for (let row = range.startRow; row <= last; row++) {
			const values = rows[row]

			if (values === undefined) continue

			const end = Math.min(range.endCol, values.length - 1)

			for (let col = range.startCol; col <= end; col++) if ((values[col] ?? "") !== "" && visit(row, col)) return true
		}

		return false
	}

	shown(row: number, col: number): string {
		return this.document.text(row, col)
	}

	typed(row: number, col: number): string {
		return this.document.text(row, col)
	}
}

function findDocument(context: SheetReadContext<TextSheetDocument>): FindDocument {
	const cells = new TextFindSheet(context.doc)

	return {
		searched: () => [TEXT_SHEET],
		sheet: sheet => Promise.resolve(sheet === TEXT_SHEET ? cells : null),
		write: (sheet, entries, changes) =>
			textFeature.commands?.["cells.set"]?.run({ ...context, changes }, { type: "cells.set", sheet, cells: entries }) ?? {
				refused: { reason: "unsupported" }
			}
	}
}

export const textFindFeature = defineSheetFeature<TextSheetDocument>({
	id: "find",
	commands: {
		"find.replace": {
			info: { write: true },
			run: (context, command) => replaceCells(findDocument(context), command, context)
		}
	},
	queries: {
		find: { run: (context, request) => findCells(findDocument(context), request, context) }
	}
})
