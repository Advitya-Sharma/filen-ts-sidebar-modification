import { existsSync, readdirSync, readFileSync } from "node:fs"
import { MessageChannel, type MessagePort } from "node:worker_threads"
import { dirname, join } from "node:path"
import { describe, expect, it } from "vitest"
import { CsvDocument, parseCsvFile, type EditOp } from "../core/sheet/index"
import {
	tileKey,
	tileText,
	type SheetChange,
	type SheetCommandResult,
	type SheetFileKind,
	type SheetSessionApi
} from "../core/sheet/protocol"
import { openTextSession } from "../core/sheet/text"
import { sheetWorkerHost, type SheetWorkerLike } from "../dom/sheet/client"
import { serveTextWorker } from "../dom/sheet/text"
import { CORPUS_DIR } from "./corpusDir"
import { cells, firstTile, numeric, savedBytes, tileEntry, tileTexts } from "./sheetSessionFixtures"
import { equalBytes } from "./xlsxWriteFixtures"

// CSV, TSV and .tab files as sessions: text stored as typed, the file written back in its own encoding, separator
// and line ends, an unedited (or undone) save as the source byte for byte; and the text path's imports, which
// never reach the workbook stack or the formula engine. The corpus differential holds every CSV and TSV to
// HEAD's text editing (CsvDocument).

async function open(bytes: Uint8Array, kind: SheetFileKind = "csv"): Promise<SheetSessionApi> {
	const opened = await openTextSession(bytes, { kind }, () => undefined)

	if (opened.kind !== "opened") throw new Error(opened.kind)

	return opened.session
}

function changeOf(result: SheetCommandResult): SheetChange {
	if (!result.ok) throw new Error(result.refusal.reason)

	return result.change
}

const utf8 = (text: string): Uint8Array => new TextEncoder().encode(text)

describe("text sessions", () => {
	it("shows text as stored, numbers on the right, and saves an unedited file as its source", async () => {
		const bytes = utf8('name;value\r\n"a;b";007\r\nx;1,5\r\n')
		const session = await open(bytes)
		const tile = await firstTile(session, 0)

		expect(tileTexts(tile).get("1,0")).toBe("a;b")
		expect(tileTexts(tile).get("1,1")).toBe("007")
		expect(numeric(tileEntry(tile, 2, 1)?.flags ?? 0)).toBe(true)
		expect(equalBytes(await savedBytes(session), bytes)).toBe(true)
		session.close()
	})

	it("stores a typed formula as its text, and writes the file in its own separator and line ends", async () => {
		const bytes = utf8("a\tb\nc\td\n")
		const session = await open(bytes, "tsv")
		const change = changeOf(
			await session.command({
				type: "cells.set",
				sheet: 0,
				cells: cells([
					[0, 1, "=1+1"],
					[2, 0, "'x"]
				])
			})
		)

		expect(change.state.dirty).toBe(true)
		expect(new TextDecoder().decode(await savedBytes(session))).toBe("a\t=1+1\nc\td\n'x\n")
		session.close()
	})

	it("refuses a character a windows-1252 file cannot hold", async () => {
		const bytes = Uint8Array.of(0x63, 0x61, 0x66, 0xe9, 0x2c, 0x31, 0x0a)
		const session = await open(bytes)

		expect(tileTexts(await firstTile(session, 0)).get("0,0")).toBe("café")

		const result = await session.command({ type: "cells.set", sheet: 0, cells: cells([[0, 1, "€ and ✓"]]) })

		expect(result.ok ? null : result.refusal.reason).toBe("encoding")
		expect(equalBytes(await savedBytes(session), bytes)).toBe(true)
		session.close()
	})

	it("inserts and deletes rows and columns, and undone edits save as the source", async () => {
		const bytes = utf8("a,b,c\r\nd,e,f\r\ng,h,i\r\n")
		const session = await open(bytes)

		changeOf(await session.command({ type: "axis.delete", sheet: 0, axis: "rows", at: 0, count: 1 }))
		changeOf(await session.command({ type: "axis.insert", sheet: 0, axis: "cols", at: 1, count: 2 }))
		expect(new TextDecoder().decode(await savedBytes(session))).toBe("d,,,e,f\r\ng,,,h,i\r\n")

		const past = changeOf(await session.command({ type: "axis.insert", sheet: 0, axis: "rows", at: 50, count: 1 }))

		expect(past.state.version).toBeGreaterThan(0)
		changeOf(await session.undo())
		changeOf(await session.undo())
		expect(equalBytes(await savedBytes(session), bytes)).toBe(true)
		session.close()
	})

	it("pastes text and moves a cut", async () => {
		const session = await open(utf8("a,b\r\nc,d\r\n"))
		const copied = await session.query("copy", { sheet: 0, range: { startRow: 0, startCol: 0, endRow: 0, endCol: 1 }, cut: true })

		if ("refused" in copied) throw new Error("refused")

		changeOf(await session.command({ type: "clipboard.paste", sheet: 0, at: { row: 2, col: 0 }, text: copied.text, clip: copied.clip }))
		expect(new TextDecoder().decode(await savedBytes(session))).toBe(",\r\nc,d\r\na,b\r\n")
		session.close()
	})

	it("refuses formats and sizes, which a text file does not hold", async () => {
		const session = await open(utf8("a\n"))
		const format = await session.command({
			type: "format.apply",
			sheet: 0,
			ranges: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }],
			patch: { bold: true }
		})

		expect(format.ok ? null : format.refusal.reason).toBe("unsupported")
		session.close()
	})

	it("runs in a text worker", async () => {
		const { port1, port2 } = new MessageChannel()
		const page: SheetWorkerLike = {
			postMessage: (message, transfer) => port1.postMessage(message, transfer ?? []),
			addEventListener: (type, listener) => {
				if (type === "message") port1.on("message", (data: unknown) => listener({ data }))
			}
		}

		serveTextWorker({
			post: (message, transfer) => port2.postMessage(message, [...transfer]),
			listen: onMessage => {
				port2.on("message", onMessage)
			}
		})

		const host = sheetWorkerHost(page)
		const result = await host.open(utf8("1,2\n"), { kind: "csv" }, () => undefined)

		if (result.kind !== "opened") throw new Error(result.kind)

		changeOf(await result.session.command({ type: "cells.set", sheet: 0, cells: cells([[0, 0, "3"]]) }))
		expect(new TextDecoder().decode(await savedBytes(result.session))).toBe("3,2\n")

		const workbook = await host.open(utf8("PK"), { kind: "xlsx" }, () => undefined)

		expect(workbook).toMatchObject({ kind: "failed", reason: "unsupported" })
		host.dispose()
		;(port1 as MessagePort).close()
		;(port2 as MessagePort).close()
	})
})

describe("the text path's imports", () => {
	it("never reach the workbook reader and writer, the formula engine or encryption", () => {
		const seen = new Set<string>()
		const packages = new Set<string>()
		const visit = (file: string): void => {
			if (seen.has(file)) return

			seen.add(file)

			for (const [, specifier] of readFileSync(file, "utf8").matchAll(
				/^(?:import|export)\s(?!type\b)(?:[^";]*?\sfrom\s)?"([^"]+)"/gm
			)) {
				if (specifier === undefined) continue

				if (specifier.startsWith(".")) {
					const target = join(dirname(file), specifier)

					visit(target.endsWith(".ts") ? target : existsSync(`${target}.ts`) ? `${target}.ts` : join(target, "index.ts"))
				} else {
					packages.add(specifier)
				}
			}
		}

		visit(join(import.meta.dirname, "..", "dom", "sheet", "text.ts"))

		const files = [...seen].map(file => file.slice(file.indexOf("/src/") + 5))

		expect(files).toContain("core/sheet/text.ts")
		expect(
			files.filter(
				file => file.startsWith("core/xlsx/") || file.startsWith("core/sheet/formula/") || file.startsWith("dom/encryption/")
			)
		).toEqual([])
		expect(files).not.toContain("core/sheet/commands.ts")
		expect(files).not.toContain("dom/sheet/host.ts")
		expect([...packages].sort()).toEqual([])
	})
})

const CORPUS = (["csv", "tsv"] as const).flatMap(kind =>
	existsSync(join(CORPUS_DIR, kind))
		? readdirSync(join(CORPUS_DIR, kind)).map(name => ({ kind, path: join(CORPUS_DIR, kind, name) }))
		: []
)

describe.skipIf(CORPUS.length === 0)("text sessions against HEAD's text editing, over the corpus", () => {
	it("shows the same cells, saves unedited as the source and edited as HEAD does", { timeout: 300_000 }, async () => {
		const differences: string[] = []
		let compared = 0

		for (const { kind, path } of CORPUS) {
			const bytes = new Uint8Array(readFileSync(path))
			let head: CsvDocument

			try {
				const parsed = parseCsvFile(bytes, kind === "tsv")

				head = new CsvDocument(parsed.rows, parsed.format)
			} catch {
				const opened = await openTextSession(bytes, { kind }, () => undefined)

				if (opened.kind === "opened") differences.push(`${path}: HEAD refuses it, the session opens it`)

				continue
			}

			const opened = await openTextSession(bytes.slice(), { kind }, () => undefined)

			if (opened.kind !== "opened") {
				differences.push(`${path}: ${opened.kind}`)
				continue
			}

			const { session, info } = opened
			const sheet = head.doc().sheets[0]

			compared++

			if (info.writability.state !== (head.writable ? "writable" : "viewOnly"))
				differences.push(`${path}: writability ${info.writability.state}`)
			if (info.activeLayout.rowCount !== sheet?.rowCount || info.activeLayout.colCount !== sheet.colCount)
				differences.push(`${path}: extent`)

			// Every cell HEAD shows, through the session's tiles.
			const tileRows = Math.ceil((sheet?.rowCount ?? 0) / 64)
			const tileCols = Math.ceil((sheet?.colCount ?? 0) / 32)
			const shown = new Map<number, { text: string; numeric: boolean }>()

			for (let tileRow = 0; tileRow < tileRows; tileRow++) {
				const coordinates: number[] = []

				for (let tileCol = 0; tileCol < tileCols; tileCol++) coordinates.push(tileRow, tileCol)

				const { tiles } = await session.tiles(0, Uint32Array.from(coordinates))

				for (const tile of tiles) {
					for (let i = 0; i < tile.count; i++) {
						const position = tile.positions[i] ?? 0
						const row = tile.tileRow * 64 + (position >> 5)
						const col = tile.tileCol * 32 + (position & 31)

						shown.set(row * 16_384 + col, {
							text: tile.strings[tile.texts[i] ?? -1] ?? "",
							numeric: numeric(tile.flags[i] ?? 0)
						})
					}
				}

				session.release(Float64Array.from(tiles, tile => tileKey(tile.sheet, tile.tileRow, tile.tileCol)))
			}

			for (const [key, view] of sheet?.cells ?? []) {
				const ours = shown.get(key)

				// A tile carries what a cell paints of a long text (the cell query gives the rest).
				if (ours?.text !== tileText(view.text) || ours.numeric !== (view.numeric === true)) {
					differences.push(`${path}: cell ${key} ${JSON.stringify(view)} / ${JSON.stringify(ours)}`)
					break
				}
			}

			if (shown.size !== (sheet?.cells.size ?? 0)) differences.push(`${path}: ${shown.size} cells, HEAD ${sheet?.cells.size ?? 0}`)

			if (head.writable) {
				const saved = await session.save()

				if (!saved.ok || !equalBytes(saved.bytes, bytes)) differences.push(`${path}: unedited save is not the source`)

				// The same edits, then both saves.
				const ops: EditOp[] = [
					{ type: "setCells", sheet: 0, cells: [{ row: 1, col: 1, input: 'edited, "quoted"' }] },
					{ type: "delete", sheet: 0, axis: "rows", at: 0, count: 1 },
					{ type: "insert", sheet: 0, axis: "cols", at: 0, count: 1 }
				]

				for (const op of ops) head.apply(op)

				changeOf(await session.command({ type: "cells.set", sheet: 0, cells: cells([[1, 1, 'edited, "quoted"']]) }))
				changeOf(await session.command({ type: "axis.delete", sheet: 0, axis: "rows", at: 0, count: 1 }))
				changeOf(await session.command({ type: "axis.insert", sheet: 0, axis: "cols", at: 0, count: 1 }))

				const edited = await session.save()

				if (!edited.ok || !equalBytes(edited.bytes, head.serialize().bytes))
					differences.push(`${path}: edited save differs from HEAD's`)

				for (let i = 0; i < 3; i++) changeOf(await session.undo())

				const undone = await session.save()

				if (!undone.ok || !equalBytes(undone.bytes, bytes)) differences.push(`${path}: undone save is not the source`)
			}

			session.close()
		}

		process.stdout.write(`text corpus: ${compared} of ${CORPUS.length} files compared with HEAD\n`)
		expect(compared).toBeGreaterThan(0)
		expect(differences).toEqual([])
	})
})
