import type { SheetRange } from "../../../core/sheet/protocol"
import { cellSelection, extendSelection, type GridPoint, type GridSelection } from "../selection"

// How a finger works the grid. The browser pans (touch-action), so a press is only known for what it was once it
// ends: a tap selects, a second tap on the same cell soon after edits it (the soft keyboard opens), and a tap on the
// cell that is already the selection, or a press held still, asks for the touch menu (Cut, Copy, Paste, Select all,
// and Edit). A selection made by touch gets a handle at its start corner to drag it wider; its end corner keeps the
// fill handle.

// A press held this long without moving is a long press (as the DOCX editor's).
export const LONG_PRESS_MS = 500
// A touch that moved less than this (px) is a tap or a long press.
export const TAP_SLOP = 10
// A press longer than this is no tap.
export const TAP_MS = 500
// A second tap this soon (ms) and this near (px) is a double tap.
export const DOUBLE_TAP_MS = 350
export const DOUBLE_TAP_SLOP = 24

export interface TouchPress {
	readonly pointerId: number
	readonly x: number
	readonly y: number
	readonly time: number
}

export interface TouchTap {
	readonly x: number
	readonly y: number
	readonly time: number
	readonly cell: GridPoint
}

// Whether a release at (x, y) at `time` ends `press` as a tap.
export function isTap(press: TouchPress, x: number, y: number, time: number): boolean {
	return Math.hypot(x - press.x, y - press.y) < TAP_SLOP && time - press.time < TAP_MS
}

// Whether a pointer that moved to (x, y) is still a long press in the making.
export function stillPressed(press: TouchPress, x: number, y: number): boolean {
	return Math.hypot(x - press.x, y - press.y) < TAP_SLOP
}

// Whether `tap` follows `previous` closely enough, on the same cell, to be its second half.
export function isDoubleTap(previous: TouchTap | null, tap: TouchTap): boolean {
	return (
		previous !== null &&
		tap.time - previous.time <= DOUBLE_TAP_MS &&
		Math.hypot(tap.x - previous.x, tap.y - previous.y) <= DOUBLE_TAP_SLOP &&
		tap.cell.row === previous.cell.row &&
		tap.cell.col === previous.cell.col
	)
}

// The selection is the single cell (or merge) at `point`: a tap there asks for the menu rather than selecting again.
export function isSoleCell(selection: GridSelection, point: GridPoint, merges: readonly SheetRange[]): boolean {
	const range = selection.ranges.length === 1 ? selection.ranges[0] : undefined
	const own = cellSelection(point, merges).ranges[0]

	return (
		range !== undefined &&
		own !== undefined &&
		selection.active.row === point.row &&
		selection.active.col === point.col &&
		range.startRow === own.startRow &&
		range.startCol === own.startCol &&
		range.endRow === own.endRow &&
		range.endCol === own.endCol
	)
}

// The selection a start-handle drag to `point` makes: the last range from its end corner (where the drag began,
// fixed) to the point; the active cell stays while the range holds it, else it is the point.
export function dragStartHandle(
	selection: GridSelection,
	range: SheetRange,
	point: GridPoint,
	merges: readonly SheetRange[]
): GridSelection {
	const extended = extendSelection({ ...selection, anchor: { row: range.endRow, col: range.endCol } }, point, merges)
	const last = extended.ranges[extended.ranges.length - 1]
	const { active } = selection
	const keeps =
		last !== undefined &&
		active.row >= last.startRow &&
		active.row <= last.endRow &&
		active.col >= last.startCol &&
		active.col <= last.endCol

	return keeps ? extended : { ...extended, active: point }
}
