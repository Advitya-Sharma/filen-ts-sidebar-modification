import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { unzipSync } from "fflate"
import { afterAll, beforeAll, describe, expect, it } from "vitest"
import type { SheetSessionApi } from "../core/sheet/protocol"
import type { SheetDrawings } from "../core/xlsx/drawings/feature"
import { decodeUtf8 } from "../core/xml/index"
import { exportSheets, sheetOffices, sheetProfile } from "./oracles/sheet/libreOffice"
import { loadAllowlist, newSdkErrors, OpenXmlSdk, openXmlSdkMissing, sdkErrorCounts } from "./oracles/openXmlSdk"
import { newValidityErrors, XsdOracle, xsdOracleMissing } from "./oracles/xsd"
import { savedBytes } from "./sheetSessionFixtures"
import { changeOf, settle } from "./sheetNotesFixtures"
import { DRAWINGS_BYTES, PNG, SVG, openDrawings } from "./sheetDrawingsFixtures"

// Drawing edits against the outside readers: LibreOffice (each install FILEN_OFFICE_SOFFICE names) opens the saved
// file, shows its pictures and shapes, and saved again as xlsx keeps each object's anchor cells; the Open XML SDK
// validator finds no error the source had not; the drawing parts and sheets validate against the ECMA-376 schemas.
// Opt-in; skips without the tools.

async function drawingsOf(session: SheetSessionApi, sheet: number): Promise<SheetDrawings | undefined> {
	await session.layout(sheet)
	await settle()

	return (await session.layout(sheet)).extensions?.drawings
}

async function edited(): Promise<{ source: Uint8Array; saved: Uint8Array }> {
	const { session, sheets } = await openDrawings()
	const one = sheets[0] ?? 0
	const two = sheets[1] ?? 0
	const objects = (await drawingsOf(session, one))?.objects ?? []
	const picture = objects[0]
	const box = objects[1]

	if (picture === undefined || box === undefined) throw new Error("fixture objects")

	changeOf(
		await session.command({
			type: "drawing.move",
			sheet: one,
			key: picture.key,
			stamp: picture.stamp,
			anchor: {
				mode: "twoCell",
				editAs: "oneCell",
				from: { row: 10, col: 2, rowOff: 0, colOff: 0 },
				to: { row: 14, col: 4, rowOff: 0, colOff: 0 }
			}
		})
	)
	changeOf(await session.command({ type: "drawing.alt", sheet: one, key: box.key, stamp: box.stamp, alt: "A yellow box" }))
	changeOf(
		await session.command({
			type: "drawing.insertImage",
			sheet: two,
			placement: { from: { row: 2, col: 1, rowOff: 0, colOff: 0 }, cx: 1219200, cy: 381000 },
			bytes: PNG,
			name: "",
			alt: "Logo"
		})
	)
	changeOf(
		await session.command({
			type: "drawing.insertImage",
			sheet: two,
			placement: { from: { row: 6, col: 3, rowOff: 0, colOff: 0 } },
			bytes: SVG,
			fallback: PNG,
			name: "icon.svg",
			alt: ""
		})
	)
	// Rows inserted above everything: the catalogue moves the anchors.
	changeOf(await session.command({ type: "axis.insert", sheet: one, axis: "rows", at: 0, count: 1 }))

	const saved = await savedBytes(session)

	session.close()

	return { source: DRAWINGS_BYTES, saved }
}

const OFFICES = sheetOffices()

describe.skipIf(OFFICES.length === 0)("drawings in LibreOffice", () => {
	let scratch: string

	beforeAll(() => {
		scratch = mkdtempSync(join(tmpdir(), "filen-office-drawings-lo-"))
	})

	afterAll(() => {
		rmSync(scratch, { recursive: true, force: true })
	})

	for (const office of OFFICES) {
		it(`${office.name} opens them where they were put`, { timeout: 180_000 }, async () => {
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

			const files = unzipSync(readFileSync(ods?.outputs[0] ?? ""))
			const content = decodeUtf8(files["content.xml"] ?? new Uint8Array(0))

			// The pictures are frames with their images, the text box keeps its text and alternative text.
			expect((content.match(/<draw:image /g) ?? []).length).toBeGreaterThanOrEqual(3)
			expect(content).toContain("Hello")
			expect(content).toMatch(/<svg:desc>A yellow box<\/svg:desc>/)
			expect(content).toMatch(/<svg:desc>Logo<\/svg:desc>/)

			// Saved again by LibreOffice as xlsx, each object reads back here on the cells it was put on.
			const [xlsx] = await exportSheets(
				office,
				sheetProfile(scratch, `${office.name}-xlsx`),
				"xlsx",
				[input],
				join(scratch, `${office.name}-xlsx`),
				120_000
			)

			expect(xlsx?.ok, xlsx?.error).toBe(true)

			const reread = await openDrawings(new Uint8Array(readFileSync(xlsx?.outputs[0] ?? "")))
			const [one = 0, two = 0] = reread.sheets
			const cells = (drawings: SheetDrawings | undefined): unknown[] =>
				(drawings?.objects ?? [])
					.filter(object => object.kind !== "chart")
					// LibreOffice saves every object on cells: the absolute connector at the sheet's top left is on A1.
					.map(object => (object.anchor.mode === "absolute" ? [0, 0] : [object.anchor.from.row, object.anchor.from.col]))
					.sort()

			const ours = await openDrawings(saved)

			expect(cells(await drawingsOf(reread.session, one))).toEqual(cells(await drawingsOf(ours.session, ours.sheets[0] ?? 0)))
			ours.session.close()
			expect(cells(await drawingsOf(reread.session, two))).toEqual([
				[2, 1],
				[6, 3]
			])
			reread.session.close()
		})
	}
})

const SDK_MISSING = openXmlSdkMissing()

describe.skipIf(SDK_MISSING !== undefined)("drawings against the Open XML SDK validator", () => {
	let sdk: OpenXmlSdk
	let scratch: string

	beforeAll(() => {
		sdk = new OpenXmlSdk()
		scratch = mkdtempSync(join(tmpdir(), "filen-office-drawings-sdk-"))
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

describe.skipIf(XSD_MISSING !== undefined)("drawings against the ECMA-376 schemas", () => {
	let oracle: XsdOracle

	beforeAll(() => {
		oracle = new XsdOracle()
	})

	afterAll(() => {
		oracle.dispose()
	})

	it("writes drawing parts and worksheets the schemas accept, adding no error their source had not", { timeout: 60_000 }, async () => {
		const { source, saved } = await edited()
		const before = unzipSync(source)
		const after = unzipSync(saved)
		const names = ["xl/drawings/drawing1.xml", "xl/drawings/drawing2.xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"]
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
