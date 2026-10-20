import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { CORPUS_DIR } from "./corpusDir"
import { corpusSpreadsheets } from "./oracles/sheet/excelCached"
import { exportShown, sheetOffices, sheetProfile } from "./oracles/sheet/libreOffice"
import { buildXlsx } from "./oracles/sheet/synthetic"
import { cachedValuesOnly, compareFills, compareShown, readShownInput, type ShownReason } from "./sheetNumfmtShown"

// Display text and fills against LibreOffice's "as shown" CSV and HTML exports, over the corpus: each workbook is
// opened as constants (cachedValuesOnly: every cell its cached value), exported, and every cell's text compared
// with the formatter's; each difference is classified (sheetNumfmtShown.ts), and none may stay unclassified.
// Opt-in: FILEN_OFFICE_SHEET_SHOWN=1 with FILEN_OFFICE_SOFFICE=<soffice> (the first install is used).

const INSTALLS = sheetOffices()
const ENABLED = process.env["FILEN_OFFICE_SHEET_SHOWN"] === "1" && INSTALLS.length > 0
const BATCH = 8

describe.skipIf(INSTALLS.length === 0)("the as-shown comparison", () => {
	it("compares formats, classifies differences and keeps cached values", { timeout: 300_000 }, async () => {
		const office = INSTALLS[0] as (typeof INSTALLS)[number]
		const work = mkdtempSync(join(tmpdir(), "filen-office-numfmt-"))
		const codes = ["0.00", '"$"#,##0_);[Red]\\("$"#,##0\\)', "m/d/yyyy", "[h]:mm:ss", "# ?/?", "0.0%", "General"]
		const styles =
			`<numFmts count="${codes.length}">${codes.map((code, i) => `<numFmt numFmtId="${164 + i}" formatCode="${code.replaceAll("&", "&amp;").replaceAll('"', "&quot;")}"/>`).join("")}</numFmts>` +
			'<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
			'<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
			'<fill><patternFill patternType="solid"><fgColor rgb="FF95B3D7"/></patternFill></fill></fills>' +
			'<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
			'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
			`<cellXfs count="${codes.length + 1}"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
			codes.map((_, i) => `<xf numFmtId="${164 + i}" fontId="0" fillId="2" borderId="0" xfId="0" applyNumberFormat="1"/>`).join("") +
			"</cellXfs>"
		const values = [1234.567, -1234.5, 45000.25, 1.75, 5.25, 0.1234, 1 / 3]
		const row = values.map((value, i) => `<c r="${String.fromCharCode(65 + i)}1" s="${i + 1}"><v>${value}</v></c>`).join("")
		// A stale cached result: shown as cached, not recalculated.
		const stale = '<c r="A2" s="1"><f>1+1</f><v>7</v></c>'

		try {
			const path = join(work, "formats.xlsx")

			writeFileSync(
				path,
				cachedValuesOnly(
					buildXlsx({
						sheets: [{ name: "S", body: `<sheetData><row r="1">${row}</row><row r="2">${stale}</row></sheetData>` }],
						styles
					})
				)
			)

			const [shown] = await exportShown(office, sheetProfile(work, "numfmt"), [path], join(work, "out"), 60_000)
			const input = await readShownInput(new Uint8Array(readFileSync(path)))
			const result = compareShown(input, shown as NonNullable<typeof shown>)
			const fills = compareFills(input, shown as NonNullable<typeof shown>)

			expect(result.compared).toBe(values.length + 1)
			expect(result.mismatches.map(item => `${item.code}: ${item.ours} / ${item.theirs} (${item.reason})`)).toEqual([
				"General: 0.333333333 / 0.333333333333333 (generalDigits)"
			])
			expect(fills.compared).toBeGreaterThan(0)
			expect(fills.mismatches).toEqual([])
		} finally {
			rmSync(work, { recursive: true, force: true })
		}
	})
})

describe.skipIf(!ENABLED || !existsSync(join(CORPUS_DIR, "xlsx")))("number formats and fills against LibreOffice over the corpus", () => {
	it("shows every cell as LibreOffice does, or as Excel's documentation says where they differ", { timeout: 3_600_000 }, async () => {
		const office = INSTALLS[0] as (typeof INSTALLS)[number]
		const work = mkdtempSync(join(tmpdir(), "filen-office-numfmt-corpus-"))
		const profile = sheetProfile(work, "numfmt-corpus")
		const files = corpusSpreadsheets()
		const reasons = new Map<ShownReason, number>()
		const unclassified: string[] = []
		const fillMismatches: string[] = []
		let compared = 0
		let equal = 0
		let fillsCompared = 0
		let fillsEqual = 0
		let failed = 0
		let unreadable = 0

		try {
			for (let start = 0; start < files.length; start += BATCH) {
				// Encrypted packages (OLE2 containers) are not zips: left out.
				const batch = files.slice(start, start + BATCH).filter(relative => {
					try {
						cachedValuesOnly(new Uint8Array(readFileSync(join(CORPUS_DIR, relative))))

						return true
					} catch {
						unreadable++

						return false
					}
				})
				const inputs = batch.map((relative, index) => {
					const path = join(work, `${start + index}.${relative.slice(relative.lastIndexOf(".") + 1)}`)

					writeFileSync(path, cachedValuesOnly(new Uint8Array(readFileSync(join(CORPUS_DIR, relative)))))

					return path
				})
				const shown = await exportShown(office, profile, inputs, join(work, `out-${start}`), 120_000)

				for (let index = 0; index < batch.length; index++) {
					const result = shown[index]
					const relative = batch[index] as string

					if (result === undefined || !result.ok) {
						failed++
						continue
					}

					const input = await readShownInput(new Uint8Array(readFileSync(inputs[index] as string)))
					const comparison = compareShown(input, result)
					const fills = compareFills(input, result)

					compared += comparison.compared
					equal += comparison.equal
					fillsCompared += fills.compared
					fillsEqual += fills.equal

					for (const mismatch of comparison.mismatches) {
						reasons.set(mismatch.reason, (reasons.get(mismatch.reason) ?? 0) + 1)

						if (mismatch.reason === "unclassified" && unclassified.length < 50)
							unclassified.push(
								`${relative} ${mismatch.sheet}!R${mismatch.row + 1}C${mismatch.column + 1} ${JSON.stringify(mismatch.code)} ${JSON.stringify(mismatch.value?.value)}: ${JSON.stringify(mismatch.ours)} / ${JSON.stringify(mismatch.theirs)}`
							)
					}

					for (const mismatch of fills.mismatches.slice(0, 5))
						fillMismatches.push(
							`${relative} ${mismatch.sheet}!R${mismatch.row + 1}C${mismatch.column + 1}: ${mismatch.ours} / ${mismatch.theirs}`
						)
				}

				rmSync(join(work, `out-${start}`), { recursive: true, force: true })

				for (const path of inputs) rmSync(path, { force: true })
			}
		} finally {
			rmSync(work, { recursive: true, force: true })
		}

		console.log(
			`numfmt as shown (LibreOffice ${office.version}): ${files.length} workbooks (${unreadable} encrypted, ${failed} not exported), ${compared} cells, ${equal} equal; ` +
				`differences ${JSON.stringify(Object.fromEntries(reasons))}; fills ${fillsEqual}/${fillsCompared} equal`
		)
		console.log(unclassified.join("\n"))
		console.log(fillMismatches.join("\n"))

		expect(compared).toBeGreaterThan(0)
		expect(unclassified).toEqual([])
		// What is left of the fills are cells LibreOffice's HTML table lays out otherwise, a few in a million.
		expect(fillsEqual / fillsCompared).toBeGreaterThan(0.9999)
	})
})
