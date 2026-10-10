// Portions derived from SuperDoc 1.46.3 (packages/layout-engine/style-engine/src/ooxml/index.ts:
// determineCellStyleTypes, TABLE_STYLE_PRECEDENCE, DEFAULT_TBL_LOOK), copyright Harbour and the SuperDoc
// contributors, licensed under the GNU Affero General Public License v3.0. Modified for @filen/office.

import type { TableLook } from "../props/types"
import type { TableStyleType } from "./styles"

// Word applies 0x04A0 (first row, first column, no vertical banding) to tables without w:tblLook.
export const DEFAULT_TABLE_LOOK: TableLook = {
	firstRow: true,
	lastRow: false,
	firstColumn: true,
	lastColumn: false,
	noHBand: false,
	noVBand: true
}

// Conditional formats in the order they apply (later wins). Word paints column banding over row banding and
// column edges under row edges, which differs from the order listed in §17.7.6.
export const CONDITION_ORDER: readonly TableStyleType[] = [
	"wholeTable",
	"band1Horz",
	"band2Horz",
	"band1Vert",
	"band2Vert",
	"firstCol",
	"lastCol",
	"firstRow",
	"lastRow",
	"nwCell",
	"neCell",
	"swCell",
	"seCell"
]

// One bit per condition, by its position in CONDITION_ORDER.
export function conditionBit(type: TableStyleType): number {
	return 1 << CONDITION_ORDER.indexOf(type)
}

const BIT_WHOLE = 1 << 0
const BIT_BAND1_HORZ = 1 << 1
const BIT_BAND2_HORZ = 1 << 2
const BIT_BAND1_VERT = 1 << 3
const BIT_BAND2_VERT = 1 << 4
const BIT_FIRST_COL = 1 << 5
const BIT_LAST_COL = 1 << 6
const BIT_FIRST_ROW = 1 << 7
const BIT_LAST_ROW = 1 << 8
const BIT_NW = 1 << 9
const BIT_NE = 1 << 10
const BIT_SW = 1 << 11
const BIT_SE = 1 << 12

// Conditions that describe a whole row (for w:trPr of a table style).
export const ROW_CONDITIONS = BIT_WHOLE | BIT_BAND1_HORZ | BIT_BAND2_HORZ | BIT_FIRST_ROW | BIT_LAST_ROW

// Where a cell sits. Columns are grid columns (w:tblGrid), so spans, merges and w:gridBefore do not shift
// the first/last-column and vertical banding regions. w:cnfStyle is a cache Word recomputes from the position,
// so it plays no part here.
export interface CellPosition {
	readonly rowIndex: number
	readonly rowCount: number
	readonly gridColumn: number
	readonly gridSpan: number
	readonly gridColumnCount: number
}

// The conditions applying to one cell as a bit mask over CONDITION_ORDER.
export function cellConditions(look: TableLook, position: CellPosition, rowBandSize: number, colBandSize: number): number {
	let mask = BIT_WHOLE
	const rowBand = rowBandSize > 0 ? rowBandSize : 1
	const colBand = colBandSize > 0 ? colBandSize : 1
	const columnStart = position.gridColumn
	const columnEnd = columnStart + (position.gridSpan > 0 ? position.gridSpan : 1)
	const columnCount = position.gridColumnCount
	// Banding skips the header row and first column when those are styled.
	const bandRow = Math.max(0, position.rowIndex - (look.firstRow ? 1 : 0))
	const bandColumn = Math.max(0, columnStart - (look.firstColumn ? 1 : 0))

	if (!look.noHBand) mask |= Math.floor(bandRow / rowBand) % 2 === 0 ? BIT_BAND1_HORZ : BIT_BAND2_HORZ
	if (!look.noVBand) mask |= Math.floor(bandColumn / colBand) % 2 === 0 ? BIT_BAND1_VERT : BIT_BAND2_VERT

	const firstRow = look.firstRow && position.rowIndex === 0
	const lastRow = look.lastRow && position.rowCount > 0 && position.rowIndex === position.rowCount - 1
	const firstColumn = look.firstColumn && columnStart === 0
	const lastColumn = look.lastColumn && columnCount > 0 && columnEnd >= columnCount

	if (firstRow) mask |= BIT_FIRST_ROW
	if (lastRow) mask |= BIT_LAST_ROW
	if (firstColumn) mask |= BIT_FIRST_COL
	if (lastColumn) mask |= BIT_LAST_COL

	// Corners need both their row and column flags (MS-OI29500).
	if (firstRow && firstColumn) mask |= BIT_NW
	if (firstRow && lastColumn) mask |= BIT_NE
	if (lastRow && firstColumn) mask |= BIT_SW
	if (lastRow && lastColumn) mask |= BIT_SE

	return mask
}
