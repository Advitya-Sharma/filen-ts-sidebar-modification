import { describe, expect, expectTypeOf, it } from "vitest"
import type { EncryptionRefusal } from "../dom/encryption/sealedPackage"
import type { PackageSealer, PackageSealOptions } from "../dom/encryption/seal"
import type { OfficeKeyTransfer } from "../dom/encryption/webCrypto"
import type { DocxTakenPackage } from "../dom/docx/protection"
import {
	AxisFlag,
	CellFlag,
	SHEET_MAX_COLS,
	SHEET_MAX_ROWS,
	TILE_CELLS,
	TILE_COLS,
	TILE_ROWS,
	cellTileKey,
	tileKey,
	tileKeyCol,
	tileKeyRow,
	tileKeySheet,
	tilePosition,
	type SheetChange,
	type SheetCommand,
	type SheetCommandOf,
	type SheetCommandResult,
	type SheetEncryptionRefusal,
	type SheetQueryRequest,
	type SheetQueryResult,
	type SheetSealOptions,
	type SheetSealer,
	type SheetTakenPackage,
	type SheetTile
} from "../core/sheet/index"
import { defineSheetFeature, sheetRegistry, type SheetEditStep, type SheetFeature } from "../core/sheet/registry"
import { CellKind, ColFlag, RowFlag, SLOT_BLOCK, type XlsxSheetStore } from "../core/xlsx/index"

function bitsOf(flags: Readonly<Record<string, number>>, skip: readonly string[] = []): number[] {
	return Object.entries(flags)
		.filter(([name]) => !skip.includes(name))
		.map(([, bit]) => bit)
}

function disjoint(bits: readonly number[]): boolean {
	let seen = 0

	for (const bit of bits) {
		if ((seen & bit) !== 0) return false

		seen |= bit
	}

	return true
}

describe("tile keys", () => {
	it("round-trip at the grid's corners and for large sheet ids", () => {
		const lastTileRow = SHEET_MAX_ROWS / TILE_ROWS - 1
		const lastTileCol = SHEET_MAX_COLS / TILE_COLS - 1

		for (const sheet of [0, 1, 2 ** 20, 2 ** 30 - 1]) {
			for (const [row, col] of [
				[0, 0],
				[lastTileRow, 0],
				[0, lastTileCol],
				[lastTileRow, lastTileCol]
			] as const) {
				const key = tileKey(sheet, row, col)

				expect(Number.isSafeInteger(key)).toBe(true)
				expect([tileKeySheet(key), tileKeyRow(key), tileKeyCol(key)]).toEqual([sheet, row, col])
			}
		}
	})

	it("place every cell of a tile at its own position", () => {
		const positions = new Set<number>()

		for (let row = TILE_ROWS * 3; row < TILE_ROWS * 4; row++) {
			for (let col = TILE_COLS * 7; col < TILE_COLS * 8; col++) {
				expect(cellTileKey(5, row, col)).toBe(tileKey(5, 3, 7))
				positions.add(tilePosition(row, col))
			}
		}

		expect(positions.size).toBe(TILE_CELLS)
		expect(Math.max(...positions)).toBe(TILE_CELLS - 1)
		expect(TILE_CELLS).toBeLessThanOrEqual(0x10000)
		expect(cellTileKey(0, SHEET_MAX_ROWS - 1, SHEET_MAX_COLS - 1)).toBe(
			tileKey(0, SHEET_MAX_ROWS / TILE_ROWS - 1, SHEET_MAX_COLS / TILE_COLS - 1)
		)
	})
})

describe("flag tables", () => {
	it("give every cell flag its own bit, inside 32 bits", () => {
		const bits = bitsOf(CellFlag)

		expect(disjoint(bits)).toBe(true)
		expect(bits.every(bit => bit > 0 && bit <= 2 ** 31)).toBe(true)
	})

	it("keep axis and row flags apart from their outline levels", () => {
		expect(disjoint(bitsOf(AxisFlag, ["OUTLINE_SHIFT"]))).toBe(true)
		expect(AxisFlag.OUTLINE_MASK >> AxisFlag.OUTLINE_SHIFT).toBe(7)
		expect(disjoint(bitsOf(RowFlag, ["OUTLINE_SHIFT"]))).toBe(true)
		expect(RowFlag.OUTLINE_MASK >> RowFlag.OUTLINE_SHIFT).toBe(7)
		expect(RowFlag.OUTLINE_MASK).toBeLessThan(0x10000)
		expect(disjoint(bitsOf(ColFlag, ["OUTLINE_SHIFT"]))).toBe(true)
	})

	it("number cell kinds as one byte each", () => {
		const kinds = Object.values(CellKind)

		expect(new Set(kinds).size).toBe(kinds.length)
		expect(kinds.every(kind => kind >= 0 && kind < 256)).toBe(true)
	})

	it("fit a block's largest cell index in a slot", () => {
		// A block's cell arrays index below SLOT_BLOCK; slots of 512 blocks stay exact.
		expect(Number.isSafeInteger(SLOT_BLOCK * 512 * 1024)).toBe(true)
		expectTypeOf<ReturnType<XlsxSheetStore["locate"]>>().toEqualTypeOf<number>()
	})
})

describe("protocol messages", () => {
	it("survive a structured clone with their typed arrays", () => {
		const tile: SheetTile = {
			sheet: 1,
			tileRow: 0,
			tileCol: 0,
			version: 3,
			count: 2,
			positions: Uint16Array.of(0, 33),
			texts: Int32Array.of(0, 1),
			inputs: Int32Array.of(-1, 2),
			styles: Uint32Array.of(0, 4),
			flags: Uint32Array.of(CellFlag.NUMBER, CellFlag.FORMULA | CellFlag.RICH),
			strings: ["1", "total", "=SUM(A1:A3)"],
			layers: { runs: { runEntries: Uint16Array.of(1), runStarts: Uint32Array.of(0), runStyles: Uint32Array.of(5) } }
		}
		const change: SheetChange = {
			state: { version: 3, dirty: true, canUndo: true, canRedo: false },
			patches: [{ ...tile, mode: "merge", cleared: Uint16Array.of(64) }],
			invalidated: [{ sheet: 1, startRow: 64, startCol: 0, endRow: SHEET_MAX_ROWS - 1, endCol: SHEET_MAX_COLS - 1 }],
			layouts: [{ sheet: 1, layout: { rowCount: 70, merges: Uint32Array.of(0, 0, 1, 1) } }],
			styles: { from: 4, styles: [{ bold: true, color: 0x1f4e79 }, { top: { style: "thin", color: 0 } }] }
		}

		expect(structuredClone(change)).toEqual(change)
	})

	it("type commands by their name", () => {
		expectTypeOf<SheetCommandOf<"cells.set">>().toEqualTypeOf<
			{ readonly type: "cells.set" } & {
				readonly sheet: number
				readonly cells: { readonly rows: Uint32Array; readonly cols: Uint16Array; readonly inputs: readonly string[] }
			}
		>()
		expectTypeOf<Extract<SheetCommand, { type: "axis.insert" }>["axis"]>().toEqualTypeOf<"rows" | "cols">()
		expectTypeOf<SheetQueryRequest<"find">["lookIn"]>().toEqualTypeOf<"values" | "formulas">()
		expectTypeOf<SheetQueryResult<"dataRegion">>().toHaveProperty("endCol")
		expectTypeOf<Extract<SheetCommandResult, { ok: false }>["refusal"]["reason"]>().toExtend<string>()
	})
})

describe("encryption hooks", () => {
	it("take dom/encryption's sealer, options, refusals and taken packages as they are", () => {
		expectTypeOf<PackageSealer>().toExtend<SheetSealer>()
		expectTypeOf<SheetSealOptions>().toExtend<PackageSealOptions>()
		expectTypeOf<EncryptionRefusal>().toEqualTypeOf<SheetEncryptionRefusal>()
		expectTypeOf<DocxTakenPackage>().toExtend<SheetTakenPackage<OfficeKeyTransfer>>()
	})
})

interface Doc {
	cells: Map<string, string>
}

const step: SheetEditStep = { bytes: 0, undo: () => undefined, redo: () => undefined }

describe("sheet registry", () => {
	const cells = defineSheetFeature<Doc>({
		id: "cells",
		commands: {
			"cells.set": {
				info: { write: true },
				run: (context, command) => {
					expectTypeOf(command.cells.inputs).toEqualTypeOf<readonly string[]>()
					context.doc.cells.set("A1", command.cells.inputs[0] ?? "")
					context.changes.content(command.sheet, { startRow: 0, startCol: 0, endRow: 0, endCol: 0 })

					return step
				}
			}
		},
		queries: {
			dataRegion: {
				run: (_context, request) => ({ startRow: request.row, startCol: request.col, endRow: request.row, endCol: request.col })
			}
		}
	})

	it("finds each handler and the feature owning it", () => {
		const registry = sheetRegistry([
			cells,
			{ id: "sort", commands: { "sort.apply": { info: { write: true, protection: "sort" }, run: () => null } } }
		])

		expect(registry.command("cells.set")?.info.write).toBe(true)
		expect(registry.command("sort.apply")?.info.protection).toBe("sort")
		expect(registry.command("sheet.add")).toBeUndefined()
		expect(registry.owner("sort.apply")).toBe("sort")
		expect(registry.owner("dataRegion")).toBe("cells")
		expect(registry.query("find")).toBeUndefined()
		expect(
			registry
				.query("dataRegion")
				?.run({ doc: { cells: new Map() }, settings: {}, slicing: undefined, signal: undefined }, { sheet: 0, row: 2, col: 3 })
		).toEqual({
			startRow: 2,
			startCol: 3,
			endRow: 2,
			endCol: 3
		})
	})

	it("refuses a name two features register", () => {
		const again: SheetFeature<Doc> = { id: "again", commands: { "cells.set": { info: { write: true }, run: () => null } } }

		expect(() => sheetRegistry([cells, again])).toThrow(/"cells.set" is registered by both "cells" and "again"/)
		expect(() => sheetRegistry([cells, { ...cells, commands: {} }])).toThrow(/"cells" is registered twice/)
	})

	it("keeps tile producers in feature order", () => {
		const first = (): void => undefined
		const second = (): void => undefined
		const registry = sheetRegistry<Doc>([
			{ id: "a", tiles: first },
			{ id: "b", tiles: second }
		])

		expect(registry.tiles).toEqual([first, second])
	})
})
