import { AxisFlag, type SheetAxisLayout } from "../../sheet/protocol"
import { EMU_PER_PX, type DrawingAnchor, type DrawingMarker } from "./anchor"

// Where an anchor's edges lie on the sheet, in EMU from its top left, through the sheet's row and column sizes (a
// hidden row or column has none). Offsets past a cell's size end at its far edge, as Excel draws them.

export interface SheetAxisPositions {
	// Where `index` starts, EMU.
	start(index: number): number
	size(index: number): number
}

// Positions along one axis of a sheet layout (pixels at 100 %, 9,525 EMU each).
export function axisPositions(layout: SheetAxisLayout): SheetAxisPositions {
	const count = layout.indices.length
	// Running difference from the default size before each listed index.
	const before = new Float64Array(count + 1)

	for (let k = 0; k < count; k++) {
		const hidden = ((layout.flags[k] ?? 0) & AxisFlag.HIDDEN) !== 0
		const size = hidden ? 0 : (layout.sizes[k] ?? layout.defaultSize)

		before[k + 1] = (before[k] ?? 0) + size - layout.defaultSize
	}

	// The listed entries before `index` (binary search over the ascending indices).
	const listedBefore = (index: number): number => {
		let low = 0
		let high = count

		while (low < high) {
			const mid = (low + high) >>> 1

			if ((layout.indices[mid] ?? 0) < index) low = mid + 1
			else high = mid
		}

		return low
	}

	return {
		start: index => Math.round((index * layout.defaultSize + (before[listedBefore(index)] ?? 0)) * EMU_PER_PX),
		size: index => {
			const k = listedBefore(index)

			if (layout.indices[k] !== index) return Math.round(layout.defaultSize * EMU_PER_PX)

			return ((layout.flags[k] ?? 0) & AxisFlag.HIDDEN) !== 0 ? 0 : Math.round((layout.sizes[k] ?? layout.defaultSize) * EMU_PER_PX)
		}
	}
}

function edge(axis: SheetAxisPositions, index: number, offset: number): number {
	return axis.start(index) + Math.min(offset, axis.size(index))
}

export interface AnchorBox {
	readonly x: number
	readonly y: number
	readonly cx: number
	readonly cy: number
}

// The box an anchor puts its object in, EMU.
export function anchorBox(anchor: DrawingAnchor, rows: SheetAxisPositions, cols: SheetAxisPositions): AnchorBox {
	const point = (marker: DrawingMarker): { x: number; y: number } => ({
		x: edge(cols, marker.col, marker.colOff),
		y: edge(rows, marker.row, marker.rowOff)
	})

	switch (anchor.mode) {
		case "absolute":
			return { x: anchor.x, y: anchor.y, cx: anchor.cx, cy: anchor.cy }
		case "oneCell": {
			const from = point(anchor.from)

			return { ...from, cx: anchor.cx, cy: anchor.cy }
		}
		case "twoCell": {
			const from = point(anchor.from)
			const to = point(anchor.to)

			return { ...from, cx: Math.max(0, to.x - from.x), cy: Math.max(0, to.y - from.y) }
		}
	}
}
