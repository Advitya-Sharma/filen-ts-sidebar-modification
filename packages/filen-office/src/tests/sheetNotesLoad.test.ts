import { describe, expect, it } from "vitest"
import { bytesSource } from "../core/opc/index"
import { openXlsxSession } from "../core/sheet/commands"
import { CellFlag, type SheetEvent } from "../core/sheet/protocol"
import { defineSheetFeature, type SheetReadyResult } from "../core/sheet/registry"
import type { XlsxDocument } from "../core/xlsx/edit/document"
import { annotationStore } from "../core/xlsx/comments/store"
import type { SheetSession } from "../core/sheet/session"
import { IMMEDIATE, buildFixture, openFixture, tileEntry } from "./sheetSessionFixtures"
import { NOTES } from "./sheetNotesFixtures"

// What a feature reads beside a sheet (SheetFeature.ready) never holds up the first tile: it lands later and patches
// the tiles already sent, and a load that fails is reported.

const FIXTURE = buildFixture({ sheets: [{ name: "One", body: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' }] })

function deferred<T>(): { promise: Promise<T>; resolve: (value: T) => void; reject: (error: Error) => void } {
	let resolve: (value: T) => void = () => undefined
	let reject: (error: Error) => void = () => undefined
	const promise = new Promise<T>((done, fail) => {
		resolve = done
		reject = fail
	})

	return { promise, resolve, reject }
}

async function settle(): Promise<void> {
	for (let k = 0; k < 20; k++) await new Promise(resolve => setTimeout(resolve, 0))
}

describe("feature loads beside a sheet", () => {
	it("sends the first tile before a slow load lands, then patches it", async () => {
		const load = deferred<SheetReadyResult>()
		let landed = false
		const feature = defineSheetFeature<XlsxDocument>({
			id: "slow",
			ready: async () => {
				const result = await load.promise

				landed = true

				return result
			},
			tiles: (_context, tile) => {
				if (landed && tile.firstRow === 0 && tile.firstCol === 0) tile.addFlags(tile.entry(0, 0), CellFlag.NOTE)
			}
		})
		const events: SheetEvent[] = []
		const opened = await openXlsxSession(
			bytesSource(FIXTURE.slice()),
			{ kind: "xlsx", slicing: IMMEDIATE },
			event => events.push(event),
			[feature]
		)

		if (opened.kind !== "opened") throw new Error(opened.kind)

		const { tiles } = await opened.session.tiles(opened.info.activeSheet, Uint32Array.of(0, 0))
		const first = tiles[0]

		if (first === undefined) throw new Error("no tile")

		expect((tileEntry(first, 0, 0)?.flags ?? 0) & CellFlag.NOTE).toBe(0)

		load.resolve({ ranges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }] })
		await settle()

		const patch = events
			.flatMap(event => (event.type === "change" ? event.change.patches : []))
			.find(candidate => candidate.tileRow === 0 && candidate.tileCol === 0)

		expect(patch).toBeDefined()
		expect((tileEntry(patch ?? first, 0, 0)?.flags ?? 0) & CellFlag.NOTE).toBe(CellFlag.NOTE)

		opened.session.close()
	})

	it("reports a load that fails, and the session goes on", async () => {
		const feature = defineSheetFeature<XlsxDocument>({
			id: "broken",
			ready: () => Promise.reject(new Error("comments part malformed"))
		})
		const events: SheetEvent[] = []
		const opened = await openXlsxSession(
			bytesSource(FIXTURE.slice()),
			{ kind: "xlsx", slicing: IMMEDIATE },
			event => events.push(event),
			[feature]
		)

		if (opened.kind !== "opened") throw new Error(opened.kind)

		await opened.session.tiles(opened.info.activeSheet, Uint32Array.of(0, 0))
		await settle()

		expect(events).toContainEqual({
			type: "featureUnavailable",
			sheet: opened.info.activeSheet,
			feature: "broken",
			message: "comments part malformed"
		})

		opened.session.close()
	})

	it("reads a sheet's notes without parsing the sheet", async () => {
		const opened = await openXlsxSession(
			bytesSource(
				buildFixture({ ...NOTES, sheets: [NOTES.sheets[1] ?? { name: "Two" }, NOTES.sheets[0] ?? { name: "One" }] }).slice()
			),
			{ kind: "xlsx" },
			() => undefined
		)

		if (opened.kind !== "opened") throw new Error(opened.kind)

		const session = opened.session as SheetSession<XlsxDocument>
		const document = session.document
		const second = opened.info.sheets[1]?.id ?? -1
		const store = annotationStore(document.workbook)

		await store.load(document, second)

		expect(document.isReady(second)).toBe(false)
		expect([...(store.commentsOf(second)?.notes.values() ?? [])].map(note => note.text)).toEqual(["Ann:\nfirst note", "Ann:\nsecond"])
		session.close()
	})

	it("marks the notes of a sheet holding no values in view mode, before any intent to edit", async () => {
		const [first, ...rest] = NOTES.sheets

		if (first === undefined) throw new Error("no sheet")

		const { session, info, events } = await openFixture({
			...NOTES,
			sheets: [{ ...first, body: '<sheetData/><legacyDrawing r:id="rId2"/>' }, ...rest]
		})

		await session.tiles(info.activeSheet, Uint32Array.of(0, 0))
		await settle()

		const patch = events
			.flatMap(event => (event.type === "change" ? event.change.patches : []))
			.findLast(candidate => candidate.tileRow === 0 && candidate.tileCol === 0)

		if (patch === undefined) throw new Error("no patch")

		expect((tileEntry(patch, 1, 1)?.flags ?? 0) & CellFlag.NOTE).toBe(CellFlag.NOTE)
		expect((tileEntry(patch, 4, 3)?.flags ?? 0) & CellFlag.NOTE).toBe(CellFlag.NOTE)
		session.close()
	})
})
