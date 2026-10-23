import { axisEntryAt, entriesBefore, entryCount } from "../../sheet/axisRuns"
import { AxisFlag, type SheetAxisLayout } from "../../sheet/protocol"
import { EMU_PER_PX, type DrawingAnchor, type DrawingMarker } from "./anchor"

// Where an anchor's edges lie on the sheet, in EMU from its top left, through the sheet's row and column sizes (a
// hidden row or column has none). Offsets past a cell's size end at its far edge, as Excel draws them.

export interface SheetAxisPositions {
	// Where `index` starts, EMU.
	start(index: number): number
	size(index: number): number
}

// Positions along one axis of a sheet layout (pixels at 100 %, 9,525 EMU each); an entry may stand for a run.
export function axisPositions(layout: SheetAxisLayout): SheetAxisPositions {
	const count = layout.indices.length
	const sizeOf = (k: number): number => (((layout.flags[k] ?? 0) & AxisFlag.HIDDEN) !== 0 ? 0 : (layout.sizes[k] ?? layout.defaultSize))
	// Running difference from the default size before each entry.
	const before = new Float64Array(count + 1)

	for (let k = 0; k < count; k++) before[k + 1] = (before[k] ?? 0) + (sizeOf(k) - layout.defaultSize) * entryCount(layout, k)

	return {
		start: index => {
			const k = entriesBefore(layout, index) - 1
			const covered = k < 0 ? 0 : Math.min(entryCount(layout, k), index - (layout.indices[k] ?? 0))
			const difference = k < 0 ? 0 : (before[k] ?? 0) + (sizeOf(k) - layout.defaultSize) * covered

			return Math.round((index * layout.defaultSize + difference) * EMU_PER_PX)
		},
		size: index => {
			const k = axisEntryAt(layout, index)

			return Math.round((k < 0 ? layout.defaultSize : sizeOf(k)) * EMU_PER_PX)
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
