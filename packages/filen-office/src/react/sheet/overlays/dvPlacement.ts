import { useEffect, useRef, type RefObject } from "react"
import type { SheetLayout } from "../../../core/sheet/protocol"
import type { SheetValidationLayout } from "../../../core/sheet/dv/xlsx"
import type { GridRect } from "../layers"
import type { SheetGridHandle } from "../SheetGrid"

// Data validation over the grid: which of a sheet's validations (its layout's "validations") a cell has, and where the
// overlays that show it go: the list's button against the cell's trailing edge, the input message under the cell.

export type SheetValidationEntry = SheetValidationLayout["entries"][number]

// The validation shown for a cell (a list's button, an input message), or null.
export function validationEntry(layout: SheetLayout | null, row: number, col: number): SheetValidationEntry | null {
	const validations = layout?.extensions?.validations

	if (validations === undefined) return null

	const { ranges, rules, entries } = validations

	for (let k = 0; k + 3 < ranges.length; k += 4) {
		if (row >= (ranges[k] ?? 0) && row <= (ranges[k + 2] ?? 0) && col >= (ranges[k + 1] ?? 0) && col <= (ranges[k + 3] ?? 0))
			return entries[rules[k / 4] ?? 0] ?? null
	}

	return null
}

export type ValidationPlacement = "button" | "message" | "messageBesideList"

// The list button's least width (its size in the app).
const BUTTON_WIDTH = 20

// Where an overlay goes for a cell's box (as shown): the button beside the cell's trailing edge (its left in a
// right-to-left sheet), as tall as the cell; the message below the cell, from its middle. While the list is open
// (it opens from the button's left edge, rightwards) the message ends where the button starts, so both show.
export function placeOverlay(
	rect: GridRect,
	placement: ValidationPlacement,
	rightToLeft: boolean,
	size: number
): { left: number; top: number; height?: number } {
	const height = Math.max(12, rect.height - 1)

	if (placement === "button") return { left: rightToLeft ? rect.x - Math.max(size, height) : rect.x + rect.width, top: rect.y, height }

	const top = rect.y + rect.height + 4

	if (placement === "messageBesideList") {
		const button = rightToLeft ? rect.x - Math.max(BUTTON_WIDTH, height) : rect.x + rect.width

		return { left: Math.max(0, button - size), top }
	}

	return { left: rightToLeft ? rect.x : rect.x + rect.width / 2, top }
}

// A ref for an element kept on a cell as the grid scrolls and zooms (SheetGridHandle.onFrame), hidden while the
// cell is off screen. Placed by style, never by a render per frame.
export function useCellPlacement<T extends HTMLElement>(
	handle: RefObject<SheetGridHandle | null>,
	cell: { readonly row: number; readonly col: number } | null,
	placement: ValidationPlacement,
	rightToLeft: boolean
): RefObject<T | null> {
	const element = useRef<T | null>(null)
	const row = cell?.row ?? -1
	const col = cell?.col ?? -1

	useEffect(() => {
		const grid = handle.current

		if (grid === null || row < 0) return

		const place = (): void => {
			const node = element.current

			if (node === null) return

			const rect = grid.cellRect({ row, col })

			if (rect === null) {
				node.style.visibility = "hidden"

				return
			}

			const at = placeOverlay(rect, placement, rightToLeft, node.offsetWidth)

			node.style.visibility = "visible"
			node.style.transform = `translate(${String(Math.round(at.left))}px, ${String(Math.round(at.top))}px)`

			if (at.height !== undefined) node.style.height = `${String(Math.round(at.height))}px`
		}

		place()

		return grid.onFrame(place)
	}, [handle, row, col, placement, rightToLeft])

	return element
}
