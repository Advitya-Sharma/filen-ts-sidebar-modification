import { describe, expect, it } from "vitest"
import { parseCsvRecords } from "../core/sheet/csv/csvParse"
import { SHEET_MAX_ROWS } from "../core/sheet/protocol"
import { saveAsCellsFeature } from "../core/sheet/save/cells"
import { saveAsXlsxFeature } from "../core/sheet/save/xlsx"
import { openTextSession } from "../core/sheet/text"
import { cells, firstTile, tileTexts } from "./sheetSessionFixtures"
import { budget, cpuMs } from "./timing"

// A text file with more records than a sheet holds opens as Excel and LibreOffice open it: its first SHEET_MAX_ROWS
// rows, view-only, and nothing that would write the rest away (a save, a CSV, a copy in another format) is made.

describe("a CSV longer than a sheet", () => {
	it("keeps the first rows when asked, and says it stopped", () => {
		expect(parseCsvRecords("a\nb\nc\n", ",", { maxRows: 2, keepFirstRows: true })).toEqual({ rows: [["a"], ["b"]], complete: false })
		expect(parseCsvRecords("a\nb\n", ",", { maxRows: 2, keepFirstRows: true })).toEqual({ rows: [["a"], ["b"]], complete: true })
		expect(parseCsvRecords("a\nb", ",", { maxRows: 2, keepFirstRows: true })).toEqual({ rows: [["a"], ["b"]], complete: true })
		// A record past the limit is not read: its quoted field may run to the end of the file.
		expect(parseCsvRecords('a\nb\n"c\nd', ",", { maxRows: 2, keepFirstRows: true }).complete).toBe(false)
	})

	it("opens its first rows view-only, and refuses every copy that would leave the rest out", { timeout: 60_000 }, async () => {
		const rows = SHEET_MAX_ROWS + 1
		const bytes = new TextEncoder().encode(Array.from({ length: rows }, (_, row) => `r${row}`).join("\n"))
		const before = cpuMs()
		const opened = await openTextSession(bytes, { kind: "csv" }, () => undefined, [saveAsCellsFeature, saveAsXlsxFeature])
		const spent = cpuMs() - before

		if (opened.kind !== "opened") throw new Error(opened.kind)

		const { session, info } = opened

		expect(spent).toBeLessThan(budget(5_000))
		expect(info.writability).toEqual({ state: "viewOnly", reason: "incomplete" })
		expect(info.activeLayout.rowCount).toBe(SHEET_MAX_ROWS)
		expect(tileTexts(await firstTile(session, 0)).get("0,0")).toBe("r0")

		const saved = await session.save()

		expect(saved.ok).toBe(false)
		expect(await session.export("csv")).toEqual({ ok: false, refusal: { reason: "tooLarge", detail: "incomplete" } })
		expect(await session.query("saveAs.cells", {})).toEqual({ ok: false, refusal: { reason: "tooLarge", detail: "incomplete" } })
		expect(await session.query("saveAs.xlsx", { name: "big" })).toEqual({
			ok: false,
			refusal: { reason: "tooLarge", detail: "incomplete" }
		})

		const edit = await session.command({ type: "cells.set", sheet: 0, cells: cells([[0, 0, "x"]]) })

		expect(edit.ok).toBe(false)
		session.close()
	})

	it("of exactly a sheet's rows opens whole and writable", async () => {
		const bytes = new TextEncoder().encode("x\n".repeat(SHEET_MAX_ROWS))
		const opened = await openTextSession(bytes, { kind: "csv" }, () => undefined, [saveAsCellsFeature, saveAsXlsxFeature])

		if (opened.kind !== "opened") throw new Error(opened.kind)

		expect(opened.info.writability).toEqual({ state: "writable" })
		expect((await opened.session.query("saveAs.cells", {})).ok).toBe(true)
		opened.session.close()
	})
})
