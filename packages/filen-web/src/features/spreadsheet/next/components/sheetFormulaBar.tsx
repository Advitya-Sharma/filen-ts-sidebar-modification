import { useState, type KeyboardEvent, type ReactNode, type RefObject } from "react"
import { flushSync } from "react-dom"
import { useTranslation } from "react-i18next"
import { toast } from "sonner"
import { rangeName } from "@filen/office/core/sheet"
import type { SheetGridHandle, SheetGridState } from "@filen/office/react/sheet"
import { isImeKeydown } from "@/lib/ime"
import { parseNameBox } from "@/features/spreadsheet/next/lib/nameBox.logic"

// The name box: the selection's name, and where to go when one is typed (a cell, or a range to select).
function NameBox({ grid, gridRef }: { grid: SheetGridState; gridRef: RefObject<SheetGridHandle | null> }) {
	const { t } = useTranslation("preview")
	const { selection } = grid
	const range = selection.ranges[selection.ranges.length - 1]
	const shown = range === undefined ? "" : rangeName(range)
	// What is typed while the box has focus; null shows the selection.
	const [typed, setTyped] = useState<string | null>(null)

	function go(): void {
		const target = parseNameBox(typed ?? "")

		if (target === null) {
			toast.error(t("previewSheetNameBoxInvalid"))

			return
		}

		const start = { row: target.startRow, col: target.startCol }
		const end = { row: target.endRow, col: target.endCol }

		setTyped(null)
		gridRef.current?.goTo(start)

		if (end.row !== start.row || end.col !== start.col) {
			grid.model.select({ active: start, anchor: start, focus: end, ranges: [target] })
		}

		gridRef.current?.focus()
	}

	return (
		<input
			aria-label={t("previewSheetNameBox")}
			value={typed ?? shown}
			spellCheck={false}
			autoComplete="off"
			className="h-7 w-24 shrink-0 truncate rounded-md bg-muted px-2 text-center font-mono text-xs tabular-nums outline-none focus-visible:ring-2 focus-visible:ring-ring"
			onFocus={event => {
				setTyped(shown)
				event.currentTarget.select()
			}}
			onChange={event => {
				setTyped(event.target.value)
			}}
			onBlur={() => {
				setTyped(null)
			}}
			onKeyDown={event => {
				if (isImeKeydown(event.nativeEvent)) {
					return
				}

				if (event.key === "Enter") {
					event.preventDefault()
					go()
				} else if (event.key === "Escape") {
					event.preventDefault()
					event.stopPropagation()
					setTyped(null)
					gridRef.current?.focus()
				}
			}}
		/>
	)
}

// The bar over the grid: the name box, the active cell's contents (editable as the in-cell editor is: both edit
// one entry), and what the viewer has to say beside them.
export function SheetFormulaBar({
	grid,
	gridRef,
	canEdit,
	children
}: {
	grid: SheetGridState
	gridRef: RefObject<SheetGridHandle | null>
	canEdit: boolean
	children: ReactNode
}) {
	const { t } = useTranslation("preview")
	const { model, editing, active } = grid

	function commit(down: number, right: number): void {
		const { row, col } = grid.selection.active

		void model.commitEditing(null).then(done => {
			if (done && (down !== 0 || right !== 0)) {
				gridRef.current?.goTo({ row: Math.max(0, row + down), col: Math.max(0, col + right) })
			}
		})
		gridRef.current?.focus()
	}

	function handleKey(event: KeyboardEvent<HTMLInputElement>): void {
		if (isImeKeydown(event.nativeEvent)) {
			return
		}

		if (event.key === "Enter") {
			event.preventDefault()
			commit(event.shiftKey ? -1 : 1, 0)
		} else if (event.key === "Tab") {
			event.preventDefault()
			commit(0, event.shiftKey ? -1 : 1)
		} else if (event.key === "Escape") {
			event.preventDefault()
			event.stopPropagation()
			model.cancelEditing()
			gridRef.current?.focus()
		}
	}

	return (
		<div className="flex h-9 shrink-0 items-center gap-2 border-b border-border px-2 text-sm">
			<NameBox
				grid={grid}
				gridRef={gridRef}
			/>
			{canEdit ? (
				<input
					data-sheet-editor=""
					aria-label={t("previewSpreadsheetCellContents")}
					aria-invalid={editing?.invalid === true}
					value={editing === null ? active.input : editing.text}
					spellCheck={false}
					autoComplete="off"
					className="h-7 min-w-0 flex-1 rounded-md bg-transparent px-2 font-mono text-xs outline-none focus-visible:bg-muted/60 aria-invalid:text-destructive"
					onFocus={event => {
						if (editing !== null) {
							return
						}

						const bar = event.currentTarget

						// The in-cell editor takes focus as it appears; the entry is typed here, so focus comes back.
						flushSync(() => {
							model.startEditing(null, "edit")
						})
						bar.focus({ preventScroll: true })
					}}
					onChange={event => {
						model.setEditText(event.target.value)
					}}
					onKeyDown={handleKey}
					onBlur={event => {
						const related = event.relatedTarget

						if (related instanceof Element && related.closest("[data-sheet-editor]") !== null) {
							return
						}

						if (model.getSnapshot().editing !== null && document.hasFocus()) {
							void model.commitEditing(null)
						}
					}}
				/>
			) : (
				<span
					aria-label={t("previewSpreadsheetCellContents")}
					className="min-w-0 flex-1 truncate font-mono text-xs"
				>
					{active.input}
				</span>
			)}
			{children}
		</div>
	)
}
