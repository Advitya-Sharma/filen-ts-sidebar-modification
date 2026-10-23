import { useState, type KeyboardEvent, type ReactNode, type RefObject } from "react"
import { flushSync } from "react-dom"
import { useTranslation } from "react-i18next"
import { HistoryIcon } from "lucide-react"
import { toast } from "sonner"
import { CellFlag, rangeName } from "@filen/office/core/sheet"
import { cellsEditable, celllessSheet, type SheetGridHandle, type SheetGridState } from "@filen/office/react/sheet"
import { Button } from "@/components/ui/button"
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip"
import { cn } from "@filen/shared"
import { isImeKeydown } from "@/lib/ime"
import { parseNameBox } from "@/features/spreadsheet/lib/nameBox.logic"
import type { SheetFormulaExtension } from "@/features/spreadsheet/formulaBar/useSheetFormula"

// The name box: the selection's name, and where to go when one is typed (a cell, or a range to select). A sheet
// without cells (a chartsheet) has neither: the box is empty and disabled.
function NameBox({ grid, gridRef }: { grid: SheetGridState; gridRef: RefObject<SheetGridHandle | null> }) {
	const { t } = useTranslation("preview")
	const { selection } = grid
	const range = selection.ranges[selection.ranges.length - 1]
	const cellless = celllessSheet(grid)
	const shown = range === undefined || cellless ? "" : rangeName(range)
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
			disabled={cellless}
			spellCheck={false}
			autoComplete="off"
			className="h-7 w-24 shrink-0 truncate rounded-md bg-muted px-2 text-center font-mono text-xs tabular-nums outline-none focus-visible:ring-2 focus-visible:ring-ring disabled:opacity-50 pointer-coarse:text-base"
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

// The formula being typed with its references in their colours (useSheetFormula's coloredText), drawn under the
// bar's own text (which is clear while it shows): the same colours outline the cells they name. Kept in step with the
// input's horizontal scroll.
function ReferenceBackdrop({ colored, scroll }: { colored: ReactNode; scroll: number }) {
	return (
		<div
			aria-hidden="true"
			className="pointer-events-none absolute inset-0 overflow-hidden px-2 font-mono text-xs leading-7 whitespace-pre pointer-coarse:text-base"
		>
			<div style={{ transform: `translateX(${String(-scroll)}px)` }}>{colored}</div>
		</div>
	)
}

// The bar over the grid: the name box, AutoSum, the active cell's contents (editable as the in-cell editor is: both
// edit one entry), and what the viewer has to say beside them.
export function SheetFormulaBar({
	grid,
	gridRef,
	canEdit,
	formula,
	children
}: {
	grid: SheetGridState
	gridRef: RefObject<SheetGridHandle | null>
	canEdit: boolean
	// The formula typing extension (AutoSum's button, the references' colours); none for a view without it.
	formula?: SheetFormulaExtension
	children: ReactNode
}) {
	const { t } = useTranslation("preview")
	const { model, editing, active } = grid
	// A chartsheet's, dialog sheet's or macro sheet's cells are not edited.
	const editable = canEdit && cellsEditable(grid)
	const [scroll, setScroll] = useState(0)
	const coloredText = editing === null ? null : (formula?.coloredText ?? null)
	const colored = coloredText !== null
	// A formula whose result this app does not calculate (a function it lacks, a name over a huge area, a legacy
	// array): the cell shows what the file saved, which an edit elsewhere does not update.
	const cached = editing === null && active.known && (active.flags & CellFlag.CACHED) !== 0

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
			{editable && formula !== undefined ? (
				<Button
					variant="ghost"
					size="icon-xs"
					className="shrink-0 font-semibold"
					aria-label={t("previewSheetAutoSum")}
					title={t("previewSheetAutoSum")}
					onClick={formula.autoSum}
				>
					Σ
				</Button>
			) : null}
			{editable ? (
				<div className="relative flex h-7 min-w-0 flex-1 rounded-md focus-within:bg-muted/60">
					{colored ? (
						<ReferenceBackdrop
							colored={coloredText}
							scroll={scroll}
						/>
					) : null}
					<input
						data-sheet-editor=""
						aria-label={t("previewSpreadsheetCellContents")}
						aria-invalid={editing?.invalid === true}
						value={editing === null ? active.input : editing.text}
						spellCheck={false}
						autoComplete="off"
						className={cn(
							"relative h-7 min-w-0 flex-1 bg-transparent px-2 font-mono text-xs leading-7 outline-none aria-invalid:text-destructive pointer-coarse:text-base",
							colored && "text-transparent caret-foreground selection:text-transparent"
						)}
						onScroll={event => {
							setScroll(event.currentTarget.scrollLeft)
						}}
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
				</div>
			) : (
				<span
					aria-label={t("previewSpreadsheetCellContents")}
					className="min-w-0 flex-1 truncate font-mono text-xs"
				>
					{active.input}
				</span>
			)}
			{cached ? (
				<Tooltip>
					<TooltipTrigger
						render={
							<span
								role="img"
								tabIndex={0}
								aria-label={t("previewSheetCachedResult")}
								className="flex size-6 shrink-0 items-center justify-center rounded-md text-muted-foreground outline-none focus-visible:ring-2 focus-visible:ring-ring"
							>
								<HistoryIcon className="size-3.5" />
							</span>
						}
					/>
					<TooltipContent>{t("previewSheetCachedResult")}</TooltipContent>
				</Tooltip>
			) : null}
			{children}
		</div>
	)
}
