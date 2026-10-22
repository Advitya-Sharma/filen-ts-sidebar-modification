import { unzipSync, zipSync } from "fflate"
import { describe, expect, it } from "vitest"
import { bytesSource } from "../core/opc/index"
import { openTextSession } from "../core/sheet/text"
import { openXlsx } from "../core/xlsx/read/index"
import type { SheetStore } from "../core/xlsx/model/store"
import { decodeUtf8, encodeUtf8 } from "../core/xml/index"
import { R, SML, buildFixture, firstTile, savedBytes, tileTexts } from "./sheetSessionFixtures"
import { changeOf, openAnnotated, range } from "./sheetNotesFixtures"

// The formats a spreadsheet opens as beyond the Transitional workbook and CSV/TSV: a .tab file (tab-separated text),
// and Strict OOXML (ECMA-376 Part 1), which is viewed and edited like Transitional once its proof passes, and saved
// in its own namespaces (§A.1: Strict and Transitional differ in them).

const STRICT_SML = "http://purl.oclc.org/ooxml/spreadsheetml/main"
const STRICT_R = "http://purl.oclc.org/ooxml/officeDocument/relationships"

// A built workbook moved to the Strict namespaces: the SpreadsheetML main namespace and the relationships one (for
// r:id attributes and relationship types alike).
function strict(bytes: Uint8Array): Uint8Array {
	const files: Record<string, Uint8Array> = unzipSync(bytes)

	for (const [name, data] of Object.entries(files)) {
		if (name.endsWith(".xml") || name.endsWith(".rels"))
			files[name] = encodeUtf8(decodeUtf8(data).replaceAll(SML, STRICT_SML).replaceAll(R, STRICT_R))
	}

	return zipSync(files)
}

// Every part naming a Transitional SpreadsheetML or relationships namespace.
function transitionalParts(bytes: Uint8Array): string[] {
	return Object.entries(unzipSync(bytes))
		.filter(
			([name, data]) =>
				(name.endsWith(".xml") || name.endsWith(".rels")) &&
				/schemas\.openxmlformats\.org\/(spreadsheetml\/2006\/main|officeDocument\/2006\/relationships)/.test(decodeUtf8(data))
		)
		.map(([name]) => name)
}

const STRICT_BOOK = strict(
	buildFixture({
		sheets: [{ name: "Data", body: '<sheetData><row r="1"><c r="A1"><v>3</v></c><c r="B1" t="s"><v>0</v></c></row></sheetData>' }],
		sharedStrings: `<sst xmlns="${SML}" count="1" uniqueCount="1"><si><t>kept</t></si></sst>`
	})
)

describe("tab-separated .tab files", () => {
	it("open with tabs as the separator, and save an edit in the file's own separator and line ends", async () => {
		const bytes = encodeUtf8("name\tvalue\r\na,b\t1\r\n")
		const opened = await openTextSession(bytes, { kind: "tab" }, () => undefined)

		if (opened.kind !== "opened") throw new Error(opened.kind)

		const { session } = opened
		const texts = tileTexts(await firstTile(session, 0))

		expect(texts.get("1,0")).toBe("a,b")
		expect(texts.get("1,1")).toBe("1")
		expect(decodeUtf8(await savedBytes(session))).toBe("name\tvalue\r\na,b\t1\r\n")

		changeOf(
			await session.command({
				type: "cells.set",
				sheet: 0,
				cells: { rows: Uint32Array.of(2), cols: Uint16Array.of(1), inputs: ["x\ty"] }
			})
		)
		expect(decodeUtf8(await savedBytes(session))).toBe('name\tvalue\r\na,b\t1\r\n\t"x\ty"\r\n')
		session.close()
	})
})

describe("Strict OOXML workbooks", () => {
	it("are proven writable, and an edit saves in the Strict namespaces and reads back", async () => {
		const { session, sheets } = await openAnnotated(STRICT_BOOK)
		const sheet = sheets[0] ?? 0

		expect(tileTexts(await firstTile(session, sheet)).get("0,1")).toBe("kept")
		expect(await session.intent()).toEqual({ state: "writable" })

		changeOf(
			await session.command({
				type: "cells.set",
				sheet,
				cells: { rows: Uint32Array.of(1, 1), cols: Uint16Array.of(0, 1), inputs: ["new text", "=A1*2"] }
			})
		)

		const saved = await savedBytes(session)

		expect(transitionalParts(saved)).toEqual([])

		const reopened = await openXlsx(bytesSource(saved))

		if (!reopened.ok) throw new Error(reopened.reason)

		const model = await reopened.workbook.load(0)
		const store = model.store as SheetStore

		expect(reopened.workbook.strict).toBe(true)
		expect(reopened.workbook.sharedStrings.text(store.value(store.locate(1, 0)))).toBe("new text")
		expect(store.locate(1, 1)).toBeGreaterThanOrEqual(0)
		session.close()
	})

	it("count dates from the 1900 date base, which has no 1900-02-29", async () => {
		// [MS-OI29500] 2.1.585 j: "If workbook@conformance equals strict, the 1900 date system is used" (ECMA-376
		// Part 1 §18.17.4.1: serial 1 is 1899-12-31, 61 is 1900-03-01), whatever dateCompatibility says.
		const { session, sheets } = await openAnnotated(STRICT_BOOK)
		const sheet = sheets[0] ?? 0

		changeOf(
			await session.command({
				type: "cells.set",
				sheet,
				cells: {
					rows: Uint32Array.of(2, 2, 2, 2),
					cols: Uint16Array.of(0, 1, 2, 3),
					inputs: ["=DATE(1900,1,1)", "=DAY(1)", "1900-01-01", "=DATE(1900,3,1)"]
				}
			})
		)

		const texts = tileTexts(await firstTile(session, sheet))

		expect([texts.get("2,0"), texts.get("2,1"), texts.get("2,3")]).toEqual(["2", "31", "61"])

		const sheetXml = decodeUtf8(unzipSync(await savedBytes(session))["xl/worksheets/sheet1.xml"] ?? new Uint8Array())

		expect(sheetXml).toMatch(/<c r="C3"[^>]*><v>2<\/v><\/c>/)
		session.close()
	})

	it("gives new parts and links Strict relationship types", async () => {
		const { session, sheets } = await openAnnotated(STRICT_BOOK)
		const sheet = sheets[0] ?? 0

		changeOf(await session.command({ type: "note.set", sheet, row: 0, col: 0, text: "Jan:\na note" }))
		changeOf(await session.command({ type: "link.set", sheet, ranges: [range(2, 0)], target: { url: "https://filen.io/" } }))

		const saved = await savedBytes(session)
		const files = Object.fromEntries(Object.entries(unzipSync(saved)).map(([name, data]) => [name, decodeUtf8(data)]))
		const rels = files["xl/worksheets/_rels/sheet1.xml.rels"] ?? ""

		expect(transitionalParts(saved)).toEqual([])
		expect(rels).toContain(`Type="${STRICT_R}/comments"`)
		expect(rels).toContain(`Type="${STRICT_R}/vmlDrawing"`)
		expect(rels).toContain(`Type="${STRICT_R}/hyperlink"`)
		session.close()
	})
})
