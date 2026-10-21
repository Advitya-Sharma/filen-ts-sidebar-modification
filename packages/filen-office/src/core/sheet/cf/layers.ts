// What conditional formatting adds to a tile beyond a cell's render style: data bars and icons, per entry
// (CellFlag.CF_BAR and CellFlag.CF_ICON mark the entries). A grid draws them (react/sheet/overlays/cf).

declare module "../protocol" {
	interface SheetTileLayerMap {
		// One bar per entry, laid out left to right in fractions of the cell's width (a right-to-left sheet mirrors
		// it as it mirrors the cell); `axis` is where the axis line goes, NaN for none.
		readonly cfBars: {
			readonly entries: Uint16Array
			readonly starts: Float32Array
			readonly ends: Float32Array
			readonly axes: Float32Array
			// 0xRRGGBB.
			readonly fills: Uint32Array
			// 0xRRGGBB, -1 for none.
			readonly borders: Int32Array
			readonly axisColors: Int32Array
			// CfBarFlag bits.
			readonly flags: Uint8Array
		}
		// One icon per entry: its set (an index into ICON_SETS) and its index in the set, the lowest value's first.
		readonly cfIcons: {
			readonly entries: Uint16Array
			readonly sets: Uint8Array
			readonly icons: Uint8Array
		}
	}
}

export const CfBarFlag = {
	// A gradient from the colour to white, as Excel draws it; else solid.
	GRADIENT: 1 << 0
} as const

// §18.18.42 ST_IconSetType, then the sets [MS-XLSX] §2.6.81 adds. Icons in a set go from the lowest value's.
export const ICON_SETS = [
	"3Arrows",
	"3ArrowsGray",
	"3Flags",
	"3TrafficLights1",
	"3TrafficLights2",
	"3Signs",
	"3Symbols",
	"3Symbols2",
	"4Arrows",
	"4ArrowsGray",
	"4RedToBlack",
	"4Rating",
	"4TrafficLights",
	"5Arrows",
	"5ArrowsGray",
	"5Rating",
	"5Quarters",
	"3Stars",
	"3Triangles",
	"5Boxes"
] as const

export type IconSetName = (typeof ICON_SETS)[number]

// How many icons a set has, from its name.
export function iconCount(set: string): number {
	const count = Number(set.charAt(0))

	return count >= 3 && count <= 5 ? count : 3
}

export function iconSetIndex(set: string): number {
	return (ICON_SETS as readonly string[]).indexOf(set)
}
