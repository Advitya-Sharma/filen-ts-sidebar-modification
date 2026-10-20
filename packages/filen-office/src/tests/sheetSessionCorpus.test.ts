import { existsSync, readdirSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { unzipSync } from "fflate"
import { describe, expect, it } from "vitest"
import { bytesSource } from "../core/opc/index"
import { openXlsxSession } from "../core/sheet/commands"
import type { SheetFileKind } from "../core/sheet/protocol"
import { openXlsx } from "../core/xlsx/read/index"
import type { SheetStore } from "../core/xlsx/model/store"
import { CORPUS_DIR } from "./corpusDir"
import { IMMEDIATE } from "./sheetSessionFixtures"
import { equalBytes } from "./xlsxWriteFixtures"

// Sessions over every OOXML spreadsheet of the corpus: each opens, is proven, takes an edit of an unlocked cell and
// saves through the gate; the save reopens with the edit, and an undone edit saves the sheet part as its source.
// View-only files carry a typed reason. Opt-in (FILEN_OFFICE_SHEET_CORPUS=1): a few minutes over the corpus.

const RUN = process.env["FILEN_OFFICE_SHEET_CORPUS"] === "1"
const KINDS = ["xlsx", "xlsm", "xltx", "xltm"] as const
const FILES = KINDS.flatMap(kind =>
	existsSync(join(CORPUS_DIR, kind))
		? readdirSync(join(CORPUS_DIR, kind)).map(name => ({ kind, path: join(CORPUS_DIR, kind, name) }))
		: []
)

// A sheet part saved after an undone edit equals its source, but for a dimension that did not cover the cells: a
// rewritten sheet's dimension is widened to them (a reader drops cells outside it).
function sameSheetPart(saved: Uint8Array | undefined, source: Uint8Array | undefined): boolean {
	if (saved === undefined || source === undefined) return false
	if (equalBytes(saved, source)) return true

	const dimension = /<(?:\w+:)?dimension\b[^>]*\/>/
	const decoder = new TextDecoder()

	return decoder.decode(saved).replace(dimension, "") === decoder.decode(source).replace(dimension, "")
}

describe.skipIf(!RUN || FILES.length === 0)("sheet sessions over the corpus", () => {
	it("opens, proves, edits and saves every workbook through the gate", { timeout: 1_800_000 }, async () => {
		const tally = { opened: 0, writable: 0, saved: 0, refused: new Map<string, number>(), viewOnly: new Map<string, number>() }
		const failures: string[] = []

		for (const { kind, path } of FILES) {
			const bytes = new Uint8Array(readFileSync(path))
			const opened = await openXlsxSession(bytesSource(bytes), { kind: kind as SheetFileKind, slicing: IMMEDIATE }, () => undefined)

			if (opened.kind !== "opened") continue

			tally.opened++

			const { session } = opened
			const verdict = await session.intent()

			if (verdict.state !== "writable") {
				if (verdict.state === "viewOnly") tally.viewOnly.set(verdict.reason, (tally.viewOnly.get(verdict.reason) ?? 0) + 1)
				else failures.push(`${path}: ${verdict.state}`)

				session.close()
				continue
			}

			tally.writable++

			const document = session.document
			const target = document.worksheets().find(entry => entry.model !== null && entry.model.protection === null)

			if (target === undefined) {
				session.close()
				continue
			}

			const store = target.model?.store as SheetStore
			const row = store.rowExtent + 1
			const edit = await session.command({
				type: "cells.set",
				sheet: target.id,
				cells: { rows: Uint32Array.of(row), cols: Uint16Array.of(0), inputs: ["corpus edit"] }
			})

			if (!edit.ok) {
				tally.refused.set(edit.refusal.reason, (tally.refused.get(edit.refusal.reason) ?? 0) + 1)
				session.close()
				continue
			}

			const saved = await session.save()

			if (!saved.ok) {
				failures.push(`${path}: ${saved.reason} ${saved.message}`)
				session.close()
				continue
			}

			const back = await openXlsx(bytesSource(saved.bytes))

			if (!back.ok) {
				failures.push(`${path}: reopen ${back.reason}`)
				session.close()
				continue
			}

			const entry = back.workbook.sheets.find(candidate => candidate.name === target.name)
			const model = entry === undefined ? null : await back.workbook.load(entry.id)
			const slot = model?.store.locate(row, 0) ?? -1

			if (model === null || slot < 0 || back.workbook.sharedStrings.text((model.store as SheetStore).value(slot)) !== "corpus edit")
				failures.push(`${path}: the edit did not read back`)
			else tally.saved++

			await session.undo()

			const undone = await session.save()

			if (!undone.ok) failures.push(`${path}: undone save ${undone.reason} ${undone.message}`)
			else if (!sameSheetPart(unzipSync(undone.bytes)[target.part], unzipSync(bytes)[target.part]))
				failures.push(`${path}: the undone edit's sheet part differs from its source`)

			session.close()
		}

		process.stdout.write(
			`opened ${tally.opened}, writable ${tally.writable}, saved ${tally.saved}; view-only ${JSON.stringify([...tally.viewOnly])}; refused ${JSON.stringify([...tally.refused])}\n`
		)

		expect(failures).toEqual([])
		expect(tally.writable / tally.opened).toBeGreaterThanOrEqual(0.9)
	})
})
