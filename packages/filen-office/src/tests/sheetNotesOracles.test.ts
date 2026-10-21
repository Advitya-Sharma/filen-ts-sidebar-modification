import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { unzipSync } from "fflate"
import { afterAll, beforeAll, describe, expect, it } from "vitest"
import type { SheetSessionApi } from "../core/sheet/protocol"
import { decodeUtf8 } from "../core/xml/index"
import { exportSheets, sheetOffices, sheetProfile } from "./oracles/sheet/libreOffice"
import { loadAllowlist, newSdkErrors, OpenXmlSdk, openXmlSdkMissing, sdkErrorCounts } from "./oracles/openXmlSdk"
import { newValidityErrors, XsdOracle, xsdOracleMissing } from "./oracles/xsd"
import { buildFixture, savedBytes } from "./sheetSessionFixtures"
import { NOTES, changeOf, openAnnotated, range, settle } from "./sheetNotesFixtures"

// Notes, threaded comments and hyperlinks this session writes, against the outside readers: LibreOffice (each
// install FILEN_OFFICE_SOFFICE names) opens the saved file and shows them (its ODF export holds the notes'
// annotations and the links), and the Open XML SDK validator finds no error the source had not. Opt-in; skips
// without the tools.

async function edited(): Promise<{ source: Uint8Array; saved: Uint8Array }> {
	const source = buildFixture({
		...NOTES,
		sheets: [
			NOTES.sheets[0] ?? { name: "One" },
			{ name: "Two", body: '<sheetData><row r="1"><c r="A1"><v>5</v></c></row></sheetData>' }
		]
	})
	const { session, sheets } = await openAnnotated(source, { author: "Jan" })
	const run = async (step: (session: SheetSessionApi) => ReturnType<SheetSessionApi["command"]>): Promise<void> => {
		changeOf(await step(session))
	}
	const one = sheets[0] ?? 0
	const two = sheets[1] ?? 0

	await run(s => s.command({ type: "note.set", sheet: one, row: 1, col: 1, text: "Ann:\nedited note" }))
	await run(s => s.command({ type: "note.set", sheet: two, row: 2, col: 2, text: "Jan:\nnew note" }))
	await run(s => s.command({ type: "note.show", sheet: two, row: 2, col: 2, shown: true }))
	await run(s => s.command({ type: "thread.add", sheet: two, row: 4, col: 0, text: "a thread" }))
	await run(s => s.command({ type: "thread.add", sheet: two, row: 4, col: 0, text: "a reply" }))
	await run(s => s.command({ type: "link.set", sheet: two, ranges: [range(6, 1)], target: { url: "https://filen.io/x" }, text: "Filen" }))
	await run(s => s.command({ type: "link.set", sheet: two, ranges: [range(7, 1)], target: { location: "One!B2" }, text: "back" }))

	const saved = await savedBytes(session)

	session.close()

	return { source, saved }
}

const OFFICES = sheetOffices()

describe.skipIf(OFFICES.length === 0)("notes, threads and links in LibreOffice", () => {
	let scratch: string

	beforeAll(() => {
		scratch = mkdtempSync(join(tmpdir(), "filen-office-notes-lo-"))
	})

	afterAll(() => {
		rmSync(scratch, { recursive: true, force: true })
	})

	for (const office of OFFICES) {
		it(`${office.name} opens them`, { timeout: 180_000 }, async () => {
			const { saved } = await edited()
			const input = join(scratch, `${office.name}.xlsx`)

			writeFileSync(input, saved)

			const [ods] = await exportSheets(
				office,
				sheetProfile(scratch, office.name),
				"ods",
				[input],
				join(scratch, office.name),
				120_000
			)

			expect(ods?.ok, ods?.error).toBe(true)

			const content = decodeUtf8(unzipSync(readFileSync(ods?.outputs[0] ?? ""))["content.xml"] ?? new Uint8Array(0))

			// LibreOffice writes each note as an office:annotation in its cell, a line per text:p.
			expect(content).toContain("edited note")
			expect(content).toContain("new note")
			expect(content).toMatch(/<office:annotation[^>]*office:display="true"[^>]*>(?:(?!<\/office:annotation>).)*new note/s)
			// The thread, through the placeholder note at least.
			expect(content).toContain("a thread")
			expect(content).toContain("a reply")
			expect(content).toContain('xlink:href="https://filen.io/x"')
			expect(content).toMatch(/xlink:href="#One\.B2"|xlink:href="#One!B2"/)

			// Saved again by LibreOffice as xlsx, the notes and links read back here as written.
			const [xlsx] = await exportSheets(
				office,
				sheetProfile(scratch, `${office.name}-xlsx`),
				"xlsx",
				[input],
				join(scratch, `${office.name}-xlsx`),
				120_000
			)

			expect(xlsx?.ok, xlsx?.error).toBe(true)

			const reread = await openAnnotated(new Uint8Array(readFileSync(xlsx?.outputs[0] ?? "")))
			const [one = 0, two = 0] = reread.sheets

			await settle()
			expect((await reread.session.query("note.get", { sheet: one, row: 1, col: 1 }))?.text).toBe("Ann:\nedited note")
			expect((await reread.session.query("note.get", { sheet: two, row: 2, col: 2 }))?.text).toBe("Jan:\nnew note")
			expect((await reread.session.query("link.get", { sheet: two, row: 6, col: 1 }))?.url).toBe("https://filen.io/x")
			expect((await reread.session.query("link.get", { sheet: two, row: 7, col: 1 }))?.place).toEqual({
				sheet: one,
				range: range(1, 1)
			})
			reread.session.close()
		})
	}
})

const SDK_MISSING = openXmlSdkMissing()

describe.skipIf(SDK_MISSING !== undefined)("notes, threads and links against the Open XML SDK validator", () => {
	let sdk: OpenXmlSdk
	let scratch: string

	beforeAll(() => {
		sdk = new OpenXmlSdk()
		scratch = mkdtempSync(join(tmpdir(), "filen-office-notes-sdk-"))
	})

	afterAll(async () => {
		await sdk.close()
		rmSync(scratch, { recursive: true, force: true })
	})

	it("adds no error the source did not have", { timeout: 120_000 }, async () => {
		const { source, saved } = await edited()

		writeFileSync(join(scratch, "source.xlsx"), source)
		writeFileSync(join(scratch, "saved.xlsx"), saved)

		const before = sdkErrorCounts(await sdk.validate(join(scratch, "source.xlsx")))
		const after = await sdk.validate(join(scratch, "saved.xlsx"))

		expect(after.fatal).toBeUndefined()
		expect(newSdkErrors(before, sdkErrorCounts(after), loadAllowlist())).toEqual([])
	})
})

const XSD_MISSING = xsdOracleMissing()

describe.skipIf(XSD_MISSING !== undefined)("notes, threads and links against the ECMA-376 schemas", () => {
	let oracle: XsdOracle

	beforeAll(() => {
		oracle = new XsdOracle()
	})

	afterAll(() => {
		oracle.dispose()
	})

	it("writes comments parts and worksheets the schemas accept, adding no error their source had not", { timeout: 60_000 }, async () => {
		const { source, saved } = await edited()
		const before = unzipSync(source)
		const after = unzipSync(saved)
		const names = ["xl/comments1.xml", "xl/comments2.xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"]
		const sourceErrors = oracle.validate(names.map(name => before[name] ?? new Uint8Array(0)))
		const savedErrors = oracle.validate(names.map(name => after[name] ?? new Uint8Array(0)))

		expect(
			Object.fromEntries(
				names.map((name, k) => [
					name,
					before[name] === undefined
						? [...(savedErrors[k]?.keys() ?? [])]
						: newValidityErrors(sourceErrors[k] ?? new Map(), savedErrors[k] ?? new Map(), before[name])
				])
			)
		).toEqual(Object.fromEntries(names.map(name => [name, []])))
	})
})
