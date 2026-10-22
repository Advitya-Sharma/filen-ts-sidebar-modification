import { useLayoutEffect, useRef, type RefObject } from "react"
import type { SheetGridHandle } from "@filen/office/react/sheet"
import { cardPlace } from "@/features/spreadsheet/comments/annotations.logic"

// Keeps an overlay element (the returned ref's) beside its cell as the grid scrolls and zooms, hidden while the cell is
// off screen: placed by style from SheetGridHandle.onFrame, and again when the element changes size, never by a render
// per frame. Placed before passive effects run, so a field in the card is on show when it takes focus.
export function useCardPlacement(
	gridRef: RefObject<SheetGridHandle | null>,
	cell: { readonly row: number; readonly col: number },
	rightToLeft: boolean
): RefObject<HTMLDivElement | null> {
	const element = useRef<HTMLDivElement>(null)
	const { row, col } = cell

	useLayoutEffect(() => {
		const grid = gridRef.current
		const node = element.current

		if (grid === null || node === null) {
			return undefined
		}

		const place = (): void => {
			const rect = grid.cellRect({ row, col })

			if (rect === null) {
				node.style.visibility = "hidden"

				return
			}

			// The overlay layer the card is placed in (the window, before layout gives it one).
			const layer = node.offsetParent
			const bounds =
				layer instanceof HTMLElement
					? { width: layer.clientWidth, height: layer.clientHeight }
					: { width: window.innerWidth, height: window.innerHeight }
			const at = cardPlace(rect, node.offsetWidth, node.offsetHeight, bounds, rightToLeft)

			node.style.visibility = "visible"
			node.style.transform = `translate(${String(Math.round(at.left))}px, ${String(Math.round(at.top))}px)`
		}

		place()

		const unframe = grid.onFrame(place)
		// A card that grows (a field opening, a thread read) is placed again.
		const observer = typeof ResizeObserver === "undefined" ? null : new ResizeObserver(place)

		observer?.observe(node)

		return () => {
			unframe()
			observer?.disconnect()
		}
	}, [gridRef, row, col, rightToLeft])

	return element
}
