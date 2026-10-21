import { useEffect, useState, type RefObject } from "react"
import { useTranslation } from "react-i18next"
import type { SheetId, SheetSessionApi } from "@filen/office/core/sheet"
import type { SheetGridHandle } from "@filen/office/react/sheet"
import { useCardPlacement } from "@/features/spreadsheet/next/comments/placement"

// Notes shown all the time (Excel's "Show note"): each in a small card beside its cell, read-only, kept there as the
// grid scrolls. Their text is read when the layout lists them, and again after each change.

export function ShownNotes({
	gridRef,
	session,
	sheet,
	cells,
	version,
	rightToLeft
}: {
	gridRef: RefObject<SheetGridHandle | null>
	session: SheetSessionApi
	sheet: SheetId
	// Row, col pairs (the layout's shownNotes).
	cells: Uint32Array | undefined
	version: number
	rightToLeft: boolean
}) {
	const points: { row: number; col: number }[] = []

	for (let k = 0; cells !== undefined && k + 1 < cells.length; k += 2) points.push({ row: cells[k] ?? 0, col: cells[k + 1] ?? 0 })

	return (
		<>
			{points.map(point => (
				<ShownNote
					key={`${String(sheet)}:${String(point.row)}:${String(point.col)}`}
					gridRef={gridRef}
					session={session}
					sheet={sheet}
					cell={point}
					version={version}
					rightToLeft={rightToLeft}
				/>
			))}
		</>
	)
}

function ShownNote({
	gridRef,
	session,
	sheet,
	cell,
	version,
	rightToLeft
}: {
	gridRef: RefObject<SheetGridHandle | null>
	session: SheetSessionApi
	sheet: SheetId
	cell: { readonly row: number; readonly col: number }
	version: number
	rightToLeft: boolean
}) {
	const { t } = useTranslation("preview")
	const ref = useCardPlacement(gridRef, cell, rightToLeft)
	const [note, setNote] = useState<{ text: string; author: string } | null>(null)
	const { row, col } = cell

	useEffect(() => {
		let current = true

		session.query("note.get", { sheet, row, col }).then(
			found => {
				if (current) {
					setNote(found)
				}
			},
			() => undefined
		)

		return () => {
			current = false
		}
	}, [session, sheet, row, col, version])

	return (
		<div
			ref={ref}
			role="note"
			aria-label={t("previewSheetNoteBy", { author: note?.author ?? "" })}
			className="pointer-events-none invisible absolute top-0 left-0 max-h-40 w-48 overflow-hidden rounded-xl bg-popover p-2 text-xs break-words whitespace-pre-wrap text-popover-foreground shadow-md ring-1 ring-foreground/5 dark:ring-foreground/10"
		>
			{note?.text ?? ""}
		</div>
	)
}
