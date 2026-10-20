import { unzipSync } from "fflate"
import { describe, expect, it } from "vitest"
import { bytesSource } from "../core/opc/index"
import { openXlsxSession } from "../core/sheet/commands"
import type { SheetCommandResult, SheetTile } from "../core/sheet/protocol"
import { fixtureBytes } from "./fuzz/sheet/fixtures"
import { equalBytes } from "./xlsxWriteFixtures"

// Sequences the sheet fuzz found, kept as plain tests.

const IMMEDIATE = { sliceMs: 1_000_000, now: () => 0, pause: () => Promise.resolve() }

async function open(name: string) {
	const bytes = fixtureBytes(name)
	const opened = await openXlsxSession(bytesSource(bytes.slice()), { kind: "xlsx", slicing: IMMEDIATE }, () => undefined)

	if (opened.kind !== "opened") throw new Error(`expected an open, got ${opened.kind}`)

	expect(await opened.session.intent()).toEqual({ state: "writable" })

	return { session: opened.session, bytes }
}

function ok(result: SheetCommandResult): void {
	if (!result.ok) throw new Error(`refused: ${result.refusal.reason}`)
}

function entry(tile: SheetTile, row: number, col: number): string {
	const position = ((row & 63) << 5) | (col & 31)

	for (let i = 0; i < tile.count; i++) {
		if (tile.positions[i] !== position) continue

		const text = tile.texts[i] ?? -1
		const input = tile.inputs[i] ?? -1

		return `${text < 0 ? "" : tile.strings[text]}|${input < 0 ? "" : tile.strings[input]}`
	}

	return "absent"
}

describe("sheet fuzz regressions", () => {
	it("undoes a row deletion without bringing back a later edit's cells", async () => {
		const { session, bytes } = await open("bare")
		const cells = (rows: number[], cols: number[], inputs: string[]) => ({
			rows: Uint32Array.from(rows),
			cols: Uint16Array.from(cols),
			inputs
		})

		ok(await session.command({ type: "axis.delete", sheet: 0, axis: "rows", at: 1, count: 2 }))
		// B1 written in place, then a new row: the block the deletion's undo holds is replaced.
		ok(await session.command({ type: "cells.set", sheet: 0, cells: cells([0, 3], [1, 0], ["=C1*2", "5"]) }))
		ok(await session.undo())
		ok(await session.undo())

		const { tiles } = await session.tiles(0, Uint32Array.of(0, 0))

		expect(entry(tiles[0] as SheetTile, 0, 1)).toBe("2|=A1+1")

		const saved = await session.save()

		if (!saved.ok) throw new Error(saved.message)

		const sheet = "xl/worksheets/sheet1.xml"

		expect(equalBytes(unzipSync(saved.bytes)[sheet] ?? new Uint8Array(), unzipSync(bytes)[sheet] ?? new Uint8Array())).toBe(true)
		session.close()
	})
})
