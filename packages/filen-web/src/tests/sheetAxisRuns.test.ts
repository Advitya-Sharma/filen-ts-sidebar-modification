import { describe, expect, it } from "vitest"
import { AxisFlag, type SheetAxisLayout, type SheetLayout } from "@filen/office/core/sheet"
import { unhideSpans } from "@/features/spreadsheet/menus/structure.logic"

// An OpenDocument sheet's hidden rows can come as one layout entry for a run of them: the menus read the run.

const RUN: SheetAxisLayout = {
	defaultSize: 20,
	indices: Uint32Array.of(10),
	sizes: Float32Array.of(20),
	styles: Uint32Array.of(0),
	flags: Uint16Array.of(AxisFlag.HIDDEN),
	counts: Uint32Array.of(1_000_000)
}

const EMPTY: SheetAxisLayout = {
	defaultSize: 64,
	indices: new Uint32Array(0),
	sizes: new Float32Array(0),
	styles: new Uint32Array(0),
	flags: new Uint16Array(0)
}

describe("hidden runs in the structure menus", () => {
	it("offers unhide for a selection inside or beside a hidden run, and not before it", () => {
		const layout = { rows: RUN, cols: EMPTY } as unknown as SheetLayout

		expect(unhideSpans([{ start: 500_000, end: 500_002 }], layout, "rows")).toEqual([{ start: 499_999, end: 500_003 }])
		expect(unhideSpans([{ start: 8, end: 9 }], layout, "rows")).toEqual([{ start: 7, end: 10 }])
		expect(unhideSpans([{ start: 2, end: 5 }], layout, "rows")).toBeNull()
	})
})
