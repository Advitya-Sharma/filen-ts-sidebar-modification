import { describe, expect, it } from "vitest"
import { firstTile, openFixture, tileTexts } from "./sheetSessionFixtures"

// The active sheet's preview answers the first tile row only when the start of its part holds every row of it.

// Text deflate cannot shrink much: the part is large compressed too.
function noise(length: number, seed: number): string {
	let state = seed >>> 0
	let out = ""

	for (let i = 0; i < length; i++) {
		state = (Math.imul(state, 1664525) + 1013904223) >>> 0
		out += String.fromCharCode(97 + (state >>> 27))
	}

	return out
}

function rows(from: number, to: number, cell: (row: number) => string): string {
	let out = ""

	for (let row = from; row <= to; row++) out += `<row r="${row}">${cell(row)}</row>`

	return out
}

describe("the active sheet's preview", () => {
	it("is not used when the start of the part it reads holds fewer rows than the first tile row", { timeout: 60_000 }, async () => {
		// Rows of about 200 KB: the 4 MB the preview reads holds about 20 of them.
		const body = `<sheetData>${rows(1, 80, row => `<c r="A${row}"><v>${row}</v></c><c r="B${row}" t="inlineStr"><is><t>${noise(200 * 1024, row)}</t></is></c>`)}</sheetData>`
		const { session, events } = await openFixture({ sheets: [{ name: "One", body }] })
		const tile = await firstTile(session, 0)

		expect(tileTexts(tile).get("49,0")).toBe("50")
		expect(events.some(event => event.type === "change")).toBe(false)
		session.close()
	})

	it("is not used when the rows it reads are out of order", async () => {
		const descending = `<sheetData>${rows(1, 70, row => `<c r="A${71 - row}"><v>${71 - row}</v></c>`).replace(/<row r="(\d+)">/g, (_, row: string) => `<row r="${71 - Number(row)}">`)}${rows(71, 3000, row => `<c r="A${row}" t="inlineStr"><is><t>${noise(2000, row)}</t></is></c>`)}</sheetData>`
		const { session, events } = await openFixture({ sheets: [{ name: "One", body: descending }] })
		const tile = await firstTile(session, 0)

		expect(tileTexts(tile).get("0,0")).toBe("1")
		expect(tileTexts(tile).get("63,0")).toBe("64")
		expect(events.some(event => event.type === "change")).toBe(false)
		session.close()
	})

	it("answers the first tile row when it holds every row of it, sparse rows included", async () => {
		const body = `<sheetData>${rows(1, 10, row => `<c r="A${row}"><v>${row}</v></c>`)}<row r="500"><c r="A500"><v>500</v></c></row>${rows(501, 4000, row => `<c r="A${row}" t="inlineStr"><is><t>${noise(2000, row)}</t></is></c>`)}</sheetData>`
		const { session, info } = await openFixture({ sheets: [{ name: "One", body }] })

		// Answered from the preview: the sheet is not parsed yet.
		expect(info.sheets[0]?.ready).toBe(false)
		expect(tileTexts(await firstTile(session, 0)).get("9,0")).toBe("10")
		session.close()
	})
})
