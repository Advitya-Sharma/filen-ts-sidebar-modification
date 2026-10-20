import { existsSync, readFileSync, readdirSync } from "node:fs"
import { join } from "node:path"
import { unzipSync } from "fflate"
import { describe, expect, it } from "vitest"
import { bytesSource } from "../core/opc/index"
import { decodeUtf8 } from "../core/xml/index"
import { openXlsx, type XlsxWorkbook } from "../core/xlsx/read/index"
import { planRefEdit, type RefEdit } from "../core/xlsx/refs/index"
import { saveXlsx } from "../core/xlsx/write/save"
import { CORPUS_DIR } from "./corpusDir"
import { changedEntries } from "./xlsxWriteFixtures"

// Structure edits over every openable OOXML spreadsheet in the corpus: each edit is planned (refused with a typed
// reason, or applied), every applied edit saves through the gate, and its revert saves the source's bytes again
// (calcChain, which a structure edit drops for good, aside). The tally of refusals by construct is logged.

const KINDS = ["xlsx", "xlsm", "xltx", "xltm"]
const FILES = KINDS.flatMap(kind =>
	existsSync(join(CORPUS_DIR, kind)) ? readdirSync(join(CORPUS_DIR, kind)).map(name => join(kind, name)) : []
)
// What dropping calcChain changes besides calcChain.xml.
const CALC_CHAIN = new Set(["[Content_Types].xml", "xl/_rels/workbook.xml.rels", "xl/calcChain.xml"])

function rewritten(a: string, b: string): boolean {
	// A source dimension that does not cover the cells is widened by any rewrite of the sheet.
	const plain = (text: string) =>
		text
			.replace(/>\s+</g, "><")
			.replace(/ r="[A-Z]*[0-9]+"/g, "")
			.replace(/<dimension [^>]*>/, "")

	return a !== b && plain(a) === plain(b)
}

function edits(workbook: XlsxWorkbook): RefEdit[] {
	const first = workbook.sheets.find(sheet => sheet.kind === "worksheet")

	if (first === undefined) return []

	const sheet = first.id
	const out: RefEdit[] = [
		{ kind: "axis", sheet, shift: { type: "insert", axis: "rows", at: 1, count: 2 } },
		{ kind: "axis", sheet, shift: { type: "delete", axis: "rows", at: 2, count: 1 } },
		{ kind: "axis", sheet, shift: { type: "insert", axis: "cols", at: 0, count: 1 } },
		{ kind: "axis", sheet, shift: { type: "delete", axis: "cols", at: 1, count: 1 } },
		{ kind: "renameSheet", sheet, name: `${first.name.slice(0, 20)} (2)` }
	]
	const last = workbook.sheets[workbook.sheets.length - 1]

	if (workbook.sheets.length > 1 && last !== undefined)
		out.push({ kind: "deleteSheet", sheet: last.id }, { kind: "moveSheet", sheet, to: workbook.sheets.length - 1 })

	return out
}

describe.skipIf(FILES.length === 0)("structure edits over the corpus", () => {
	it("refuses with a reason, or saves through the gate and reverts to the source", { timeout: 1_800_000 }, async () => {
		const refusals = new Map<string, number>()
		const failures: string[] = []
		let applied = 0
		let planned = 0
		let reserialised = 0

		for (const file of FILES) {
			const source = new Uint8Array(readFileSync(join(CORPUS_DIR, file)))
			const opened = await openXlsx(bytesSource(source))

			if (!opened.ok) continue

			for (const edit of edits(opened.workbook)) {
				const reopened = await openXlsx(bytesSource(source))

				if (!reopened.ok) continue

				const workbook = reopened.workbook
				let result: Awaited<ReturnType<typeof planRefEdit>>

				planned++

				try {
					result = await planRefEdit(workbook, edit)
				} catch (error) {
					failures.push(`${file} ${edit.kind}: plan threw ${String(error)}`)

					continue
				}

				if (!result.ok) {
					const key = `${result.refusal.reason}: ${(result.refusal.detail ?? "").replace(/ [^ ]*\.xml$/, "")}`

					refusals.set(key, (refusals.get(key) ?? 0) + 1)

					continue
				}

				result.plan.apply()

				const saved = await saveXlsx({ opc: workbook.opc, workbook })

				if (!saved.ok) {
					failures.push(`${file} ${JSON.stringify(edit)}: save refused (${saved.reason}) ${saved.message}`)

					continue
				}

				applied++
				result.plan.revert()

				const back = await saveXlsx({ opc: workbook.opc, workbook })

				if (!back.ok || back.bytes === undefined) {
					failures.push(`${file} ${edit.kind}: reverted save refused`)

					continue
				}

				const changed = await changedEntries(source, back.bytes)
				const left = [...changed.changed, ...changed.added, ...changed.removed].filter(name => !CALC_CHAIN.has(name))
				const sourceParts = unzipSync(source)
				const backParts = unzipSync(back.bytes)
				// The writer serialises a marked row or name list anew even when it equals the source: the markup then
				// differs in white space between elements, implied row and cell numbers and a widened dimension alone.
				const real = left.filter(name => {
					const a = sourceParts[name]
					const b = backParts[name]

					if (a === undefined || b === undefined || !rewritten(decodeUtf8(a), decodeUtf8(b))) return true

					reserialised++

					return false
				})

				if (real.length > 0) failures.push(`${file} ${JSON.stringify(edit)}: reverted save differs in ${real.join(", ")}`)
			}
		}

		console.log(
			`refs corpus: ${FILES.length} files, ${planned} edits planned, ${applied} applied, ${reserialised} reverted parts written anew, refused: ${[
				...refusals
			]
				.sort((a, b) => b[1] - a[1])
				.map(([key, count]) => `${key} ×${count}`)
				.join("; ")}`
		)
		expect(failures).toEqual([])
		expect(applied).toBeGreaterThan(planned / 2)
	})
})
