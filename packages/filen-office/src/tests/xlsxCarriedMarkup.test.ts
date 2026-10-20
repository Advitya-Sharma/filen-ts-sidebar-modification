import { describe, expect, it } from "vitest"
import type { SheetCommandResult } from "../core/sheet/protocol"
import { firstTile, openFixture, savedBytes } from "./sheetSessionFixtures"
import { changedEntries, parts } from "./xlsxWriteFixtures"

// What a cell holds beyond the model (a picture in the cell: vm) goes with it through the session's own copy and
// cut: a cut moves it, a copy shares it (several cells may name one value metadata record, as Excel writes them);
// a paste of plain text carries values only, and undo and redo give each state back.

function ok(result: SheetCommandResult): void {
	if (!result.ok) throw new Error(`refused: ${result.refusal.reason} ${result.refusal.detail ?? ""}`)
}

const PICTURE = '<sheetData><row r="1"><c r="A1" t="e" vm="1"><v>#VALUE!</v></c><c r="B1"><v>2</v></c></row></sheetData>'

async function opened(sheets = [{ name: "One", body: PICTURE }]) {
	const fixture = await openFixture({ sheets })
	const one = fixture.info.sheets[0]!.id

	for (const sheet of fixture.info.sheets) await firstTile(fixture.session, sheet.id)
	await fixture.session.intent()

	return { ...fixture, one }
}

async function copy(session: Awaited<ReturnType<typeof opened>>["session"], sheet: number, cut: boolean) {
	const copied = await session.query("copy", { sheet, range: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 }, cut })

	if (!("clip" in copied)) throw new Error("copy refused")

	return copied
}

describe("a picture in a cell", () => {
	it.each([
		["within its row", { row: 0, col: 3 }, '<c r="D1" t="e" vm="1">'],
		["to another row", { row: 4, col: 0 }, '<c r="A5" t="e" vm="1">']
	])("moves with a cut %s, and undo and redo give each state back", async (_, at, moved) => {
		const { session, one, bytes } = await opened()
		const copied = await copy(session, one, true)

		ok(await session.command({ type: "clipboard.paste", sheet: one, at, text: copied.text, clip: copied.clip }))

		const sheet = parts(await savedBytes(session))["xl/worksheets/sheet1.xml"] ?? ""

		expect(sheet).toContain(moved)
		expect(sheet.match(/vm="1"/g)).toHaveLength(1)

		ok(await session.undo())
		expect(await changedEntries(bytes, await savedBytes(session))).toEqual({ changed: [], added: [], removed: [] })

		ok(await session.redo())
		expect(parts(await savedBytes(session))["xl/worksheets/sheet1.xml"]).toContain(moved)
		session.close()
	})

	it("is shared by a copy pasted elsewhere", async () => {
		const { session, one } = await opened()
		const copied = await copy(session, one, false)

		ok(await session.command({ type: "clipboard.paste", sheet: one, at: { row: 2, col: 0 }, text: copied.text, clip: copied.clip }))

		const sheet = parts(await savedBytes(session))["xl/worksheets/sheet1.xml"] ?? ""

		expect(sheet).toContain('<c r="A1" t="e" vm="1">')
		expect(sheet).toContain('<c r="A3" t="e" vm="1">')
		session.close()
	})

	it("is not carried by a paste of plain text", async () => {
		const { session, one } = await opened()
		const copied = await copy(session, one, false)

		ok(await session.command({ type: "clipboard.paste", sheet: one, at: { row: 2, col: 0 }, text: copied.text }))
		expect(parts(await savedBytes(session))["xl/worksheets/sheet1.xml"]?.match(/vm="1"/g)).toHaveLength(1)
		session.close()
	})

	it("goes with a copy to another sheet", async () => {
		const { session, one, info } = await opened([
			{ name: "One", body: PICTURE },
			{ name: "Two", body: "<sheetData/>" }
		])
		const two = info.sheets[1]!.id
		const copied = await copy(session, one, false)

		ok(await session.command({ type: "clipboard.paste", sheet: two, at: { row: 1, col: 1 }, text: copied.text, clip: copied.clip }))
		expect(parts(await savedBytes(session))["xl/worksheets/sheet2.xml"]).toContain('<c r="B2" t="e" vm="1">')
		session.close()
	})
})

describe("an error value", () => {
	it.each(["#N/A", "#value!", "#DIV/0!"])("typed as %s is that error, as Excel takes it, and quoted is text", async input => {
		const { session, one } = await opened([{ name: "One", body: "<sheetData/>" }])

		ok(
			await session.command({
				type: "cells.set",
				sheet: one,
				cells: { rows: Uint32Array.of(0, 1), cols: Uint16Array.of(0, 0), inputs: [input, `'${input}`] }
			})
		)

		const sheet = parts(await savedBytes(session))["xl/worksheets/sheet1.xml"] ?? ""

		expect(sheet).toContain(`<c r="A1" t="e"><v>${input.toUpperCase()}</v></c>`)
		expect(sheet).not.toContain('<c r="A2" t="e"')
		session.close()
	})
})
