import { useEffect, useRef, useState, type RefObject } from "react"
import type { SheetId, SheetQueryRequest, SheetRange } from "@filen/office/core/sheet"
import { selectionContains, singleCell, type SheetGridHandle, type SheetGridState } from "@filen/office/react/sheet"
import type { DocxFindControls } from "@/features/preview/components/docxViewer.logic"
import { closeSheetFind } from "@/features/spreadsheet/find/findState"
import { SHEET_FIND_LIMIT, matchAt, matchNear, walkOrder, type SheetFindCell } from "@/features/spreadsheet/find/sheetFind.logic"

// Find and Replace on the grid, as the find bar drives it (DocxFindField's controls): the session finds (in its
// worker), from the active cell on as Excel's Find Next goes, and the bar keeps the matches and steps through them,
// selecting each match's cell. Past the most one find lists, the matches are a window from the cell it stands on:
// stepping off either end finds the next window from there. The matches are found again when the file changes, the
// options change or (searching one sheet) the sheet does.

export interface SheetFindOptions {
	readonly wholeCell: boolean
	readonly lookIn: "values" | "formulas"
	readonly within: "sheet" | "workbook"
}

interface Found {
	readonly query: string
	readonly caseSensitive: boolean
	readonly cells: Uint32Array
	readonly more: boolean
	readonly active: number
	// The session version the matches reflect.
	readonly version: number
	readonly key: string
	// The selection searched (one of several cells, as Excel searches it), or none: the whole sheet or workbook.
	readonly within: readonly SheetRange[] | undefined
	// The sheet shown when it was searched.
	readonly sheet: SheetId
}

// Where a find goes from: a cell, and the match to stand on: the first at it or after it, or the last before it.
interface FindStart {
	readonly cell: SheetFindCell
	readonly start: "at" | "after" | "before"
}

const NONE: Found = {
	query: "",
	caseSensitive: false,
	cells: new Uint32Array(0),
	more: false,
	active: -1,
	version: -1,
	key: "",
	within: undefined,
	sheet: 0
}

export function useSheetFind({
	grid,
	gridRef,
	canEdit,
	options
}: {
	grid: SheetGridState
	gridRef: RefObject<SheetGridHandle | null>
	canEdit: boolean
	options: SheetFindOptions
}): DocxFindControls {
	const { model } = grid
	const [found, setFound] = useState<Found>(NONE)
	// The search in flight: a newer one makes an older answer stale.
	const asked = useRef(0)
	// What the last refresh was for (the version and what to search), so one change finds again once.
	const refreshed = useRef("")
	const scope: SheetQueryRequest<"find">["scope"] = options.within === "workbook" ? "workbook" : { sheet: grid.sheet }
	const key = `${options.within}:${options.lookIn}:${String(options.wholeCell)}:${options.within === "sheet" ? String(grid.sheet) : ""}`
	const order = grid.sheets.map(sheet => sheet.id)

	const go = (cell: SheetFindCell): void => {
		const show = (): void => {
			gridRef.current?.goTo({ row: cell.row, col: cell.col })
		}
		const snapshot = model.getSnapshot()

		// Searching a selection keeps it: the active cell moves inside it.
		if (cell.sheet === snapshot.sheet && !singleCell(snapshot.selection) && selectionContains(snapshot.selection, cell.row, cell.col)) {
			model.select({ ...snapshot.selection, active: { row: cell.row, col: cell.col } })
			gridRef.current?.reveal({ row: cell.row, col: cell.col })

			return
		}

		if (cell.sheet === snapshot.sheet) {
			show()

			return
		}

		void model.commitEditing(null).then(done => {
			if (done) {
				model.setSheet(cell.sheet)
				show()
				// The sheet's view is laid out in the next frame: its cell is brought into view then.
				requestAnimationFrame(show)
			}
		})
	}

	// Finds `query` from `from` (by default at the active cell), and `jump` selects the match it stands on.
	const run = (
		query: string,
		caseSensitive: boolean,
		within: readonly SheetRange[] | undefined,
		jump: boolean,
		from?: FindStart
	): void => {
		const ticket = ++asked.current

		if (query === "") {
			setFound(NONE)

			return
		}

		const snapshot = model.getSnapshot()
		const { cell, start } = from ?? {
			cell: { sheet: snapshot.sheet, row: snapshot.selection.active.row, col: snapshot.selection.active.col },
			start: "at"
		}

		model.session
			.query("find", {
				scope,
				...(within !== undefined && options.within === "sheet" && { within }),
				text: query,
				matchCase: caseSensitive,
				wholeCell: options.wholeCell,
				lookIn: options.lookIn,
				from: cell,
				start,
				limit: SHEET_FIND_LIMIT
			})
			.then(
				result => {
					if (ticket !== asked.current) {
						return
					}

					const total = Math.floor(result.cells.length / 3)
					// The first match the walk found, or (before `cell`) the last.
					const at = start === "before" ? total - 1 : 0
					// Every match listed: in walk order from the start, as the count numbers them.
					const whole = result.more ? null : walkOrder(result.cells, order)
					const cells = whole?.cells ?? result.cells
					const active = total === 0 ? -1 : whole === null ? at : (at - whole.shift + total) % total

					setFound({
						query,
						caseSensitive,
						cells,
						more: result.more,
						active,
						version: result.version,
						key,
						within,
						sheet: snapshot.sheet
					})

					const match = matchAt(cells, active)

					if (jump && match !== null) {
						go(match)
					}
				},
				() => undefined
			)
	}

	// Found again when the file changed under the matches, or what to search did.
	const stale = found.query !== "" && (found.version < grid.state.version || found.key !== key)
	const refresh = `${String(grid.state.version)}:${key}`

	useEffect(() => {
		if (stale && refreshed.current !== refresh) {
			refreshed.current = refresh

			// Another sheet is searched within its own selection, not the one searched before.
			const { selection } = model.getSnapshot()
			const within =
				found.sheet === grid.sheet
					? found.within
					: options.within === "sheet" && !singleCell(selection)
						? selection.ranges
						: undefined

			run(found.query, found.caseSensitive, within, false)
		}
	})

	const count = Math.floor(found.cells.length / 3)

	return {
		query: found.query,
		caseSensitive: found.caseSensitive,
		count,
		capped: found.more,
		active: found.active,
		search: (query, caseSensitive) => {
			const { selection } = model.getSnapshot()

			// A selection of more than one cell is what a sheet search searches, as in Excel.
			run(query, caseSensitive, options.within === "sheet" && !singleCell(selection) ? selection.ranges : undefined, true)
		},
		step: direction => {
			if (count === 0) {
				return
			}

			// Excel's Find Next goes from the active cell: on a match, to the one beside it; elsewhere (moved away, or
			// the match edited away), to the nearest the step's way. Past either end of a window, or away from its
			// matches, the next window from the active cell.
			const snapshot = model.getSnapshot()
			const here = { sheet: snapshot.sheet, row: snapshot.selection.active.row, col: snapshot.selection.active.col }
			const current = matchAt(found.cells, found.active)
			const on = current !== null && current.sheet === here.sheet && current.row === here.row && current.col === here.col
			const next = on ? found.active + direction : found.more ? -1 : matchNear(found.cells, order, here, direction)

			if (found.more && (next < 0 || next >= count)) {
				run(found.query, found.caseSensitive, found.within, true, { cell: here, start: direction > 0 ? "after" : "before" })

				return
			}

			const active = (next + count) % count
			const cell = matchAt(found.cells, active)

			setFound({ ...found, active })

			if (cell !== null) {
				go(cell)
			}
		},
		close: () => {
			asked.current++
			setFound(NONE)
			closeSheetFind(model)
			gridRef.current?.focus()
		},
		replace: {
			disabled: !canEdit,
			one: replacement => {
				const cell = matchAt(found.cells, found.active)

				if (cell === null) {
					return
				}

				void model.intent()
				void model
					.command({
						type: "find.replace",
						scope,
						text: found.query,
						replacement,
						matchCase: found.caseSensitive,
						wholeCell: options.wholeCell,
						cells: Uint32Array.of(cell.sheet, cell.row, cell.col),
						...(found.within !== undefined && { within: found.within })
					})
					.then(() => {
						run(found.query, found.caseSensitive, found.within, true, { cell, start: "after" })
					})
			},
			all: replacement => {
				void model.intent()
				void model
					.command({
						type: "find.replace",
						scope,
						text: found.query,
						replacement,
						matchCase: found.caseSensitive,
						wholeCell: options.wholeCell,
						...(found.within !== undefined && { within: found.within })
					})
					.then(() => {
						run(found.query, found.caseSensitive, found.within, false)
					})
			}
		}
	}
}
