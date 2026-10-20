import { Fragment, type ReactNode } from "react"
import { useTranslation } from "react-i18next"
import { AxisFlag, resetTargets, type SheetAxis, type SheetAxisLayout, type SheetCommand } from "@filen/office/core/sheet"
import type { SheetGridState } from "@filen/office/react/sheet"
import { ContextMenu, ContextMenuContent, ContextMenuItem, ContextMenuSeparator, ContextMenuTrigger } from "@/components/ui/context-menu"

// The structural items, in menu order; a separator follows each axis's delete. `after` inserts past the
// selection's end rather than at its start.
const STRUCTURE_ITEMS = [
	{ key: "previewSpreadsheetInsertRowsAbove", type: "axis.insert", axis: "rows", after: false },
	{ key: "previewSpreadsheetInsertRowsBelow", type: "axis.insert", axis: "rows", after: true },
	{ key: "previewSpreadsheetDeleteRows", type: "axis.delete", axis: "rows", after: false },
	{ key: "previewSpreadsheetInsertColumnsLeft", type: "axis.insert", axis: "cols", after: false },
	{ key: "previewSpreadsheetInsertColumnsRight", type: "axis.insert", axis: "cols", after: true },
	{ key: "previewSpreadsheetDeleteColumns", type: "axis.delete", axis: "cols", after: false }
] as const

function hiddenIn(layout: SheetAxisLayout): (index: number) => boolean {
	const hidden = new Set<number>()

	layout.flags.forEach((flags, k) => {
		if ((flags & AxisFlag.HIDDEN) !== 0) {
			hidden.add(layout.indices[k] ?? -1)
		}
	})

	return index => hidden.has(index)
}

// The grid's menu: rows and columns inserted or deleted, contents cleared, sizes reset. Read-only, it holds only the
// size resets.
export function SheetGridMenu({
	grid,
	canEdit,
	run,
	children
}: {
	grid: SheetGridState
	canEdit: boolean
	run: (command: SheetCommand) => void
	children: ReactNode
}) {
	const { t } = useTranslation("preview")
	const { sheet, selection, layout } = grid
	const range = selection.ranges[selection.ranges.length - 1]

	function resets(axis: SheetAxis): number[] {
		if (range === undefined || layout === null) {
			return []
		}

		const sizes = axis === "cols" ? layout.cols : layout.rows

		return axis === "cols"
			? resetTargets(range.startCol, range.endCol, layout.colCount, hiddenIn(sizes))
			: resetTargets(range.startRow, range.endRow, layout.rowCount, hiddenIn(sizes))
	}

	// Through the grid's model, as a header drag is: sizes the file cannot hold are the app's to reset. A refusal
	// reaches the model's refusal handler; a session that died says so by its own event.
	function reset(axis: SheetAxis, indices: number[]): void {
		grid.model.resize(sheet, axis, Uint32Array.from(indices), new Float32Array(indices.length).fill(Number.NaN)).catch(() => undefined)
	}

	return (
		<ContextMenu>
			<ContextMenuTrigger className="flex min-h-0 flex-1 flex-col">{children}</ContextMenuTrigger>
			<ContextMenuContent>
				{canEdit && range !== undefined ? (
					<>
						{STRUCTURE_ITEMS.map(item => {
							const rows = item.axis === "rows"
							const count = rows ? range.endRow - range.startRow + 1 : range.endCol - range.startCol + 1
							const start = rows ? range.startRow : range.startCol

							return (
								<Fragment key={item.key}>
									<ContextMenuItem
										onClick={() => {
											run({ type: item.type, sheet, axis: item.axis, at: item.after ? start + count : start, count })
										}}
									>
										{t(item.key, { count })}
									</ContextMenuItem>
									{item.type === "axis.delete" ? <ContextMenuSeparator /> : null}
								</Fragment>
							)
						})}
						<ContextMenuItem
							onClick={() => {
								run({ type: "cells.clear", sheet, ranges: selection.ranges, what: "contents" })
							}}
						>
							{t("previewSpreadsheetClearCells")}
						</ContextMenuItem>
						<ContextMenuSeparator />
					</>
				) : null}
				<ResetItems
					resets={resets}
					onReset={reset}
				/>
			</ContextMenuContent>
		</ContextMenu>
	)
}

// Its own component so the targets are worked out only while the menu is open.
function ResetItems({ resets, onReset }: { resets: (axis: SheetAxis) => number[]; onReset: (axis: SheetAxis, indices: number[]) => void }) {
	const { t } = useTranslation("preview")
	const cols = resets("cols")
	const rows = resets("rows")

	return (
		<>
			<ContextMenuItem
				disabled={cols.length === 0}
				onClick={() => {
					onReset("cols", cols)
				}}
			>
				{t("previewSpreadsheetResetColumnWidth", { count: cols.length })}
			</ContextMenuItem>
			<ContextMenuItem
				disabled={rows.length === 0}
				onClick={() => {
					onReset("rows", rows)
				}}
			>
				{t("previewSpreadsheetResetRowHeight", { count: rows.length })}
			</ContextMenuItem>
		</>
	)
}
