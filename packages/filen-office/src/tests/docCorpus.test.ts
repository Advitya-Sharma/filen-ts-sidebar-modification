import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { bytesSource } from "@filen/office/core/opc"
import { DocError, docText, readDoc, type DocErrorCode } from "@filen/office/core/doc"
import { CORPUS_DIR } from "./corpusDir"
import { convert, offices, outputPath } from "./interop/soffice"

// The corpus .doc/.dot files: each reads or is refused with its typed reason, in time linear in its size. With
// FILEN_OFFICE_SOFFICE set, the plain text of each equals LibreOffice's .doc → .txt export (whitespace collapsed)
// apart from the accepted differences below.

const FAMILIES = ["doc", "dot"]
const fetched = existsSync(join(CORPUS_DIR, "doc"))

function corpusFiles(): string[] {
	if (!fetched) return []

	return FAMILIES.flatMap(family =>
		existsSync(join(CORPUS_DIR, family)) ? readdirSync(join(CORPUS_DIR, family)).map(name => `${family}/${name}`) : []
	).sort()
}

const REFUSED: Readonly<Record<string, DocErrorCode>> = {
	// Not Word binary files under a Word extension: Graphviz sources and a DOCX package.
	"dot/frictionless__package.dot": "notWordBinary",
	"dot/tika__testGRAPHVIZdc.dot": "notWordBinary",
	"doc/govuk__Audit_20Committee_20minutes_208_20June_202012.doc": "notWordBinary",
	// Word 6/95 (nFib 0x65, 0x68).
	"doc/lo__tdf150197_anlv2ListFormat.doc": "legacyVersion",
	"doc/lo__tdf59896.doc": "legacyVersion",
	"doc/poi__57843.doc": "legacyVersion",
	"doc/poi__Bug60942b.doc": "legacyVersion",
	// Password protected (decryption is a separate step).
	"doc/poi__PasswordProtected.doc": "encrypted",
	"doc/poi__password_password_cryptoapi.doc": "encrypted",
	"doc/poi__password_tika_binaryrc4.doc": "encrypted"
}

// Where LibreOffice's text differs from ours, and why.
const ACCEPTED: Readonly<Record<string, string>> = {
	"doc/govuk__6720-energy-storage-comp-research-app-form.doc": "LibreOffice rebuilds the TOC field result from the headings",
	"doc/govuk__Audit_20Committee_20minutes_202_20May_202013.doc":
		"LibreOffice numbers a level with an LFO formatting override as a separate list",
	"doc/govuk__Small_Business_Survey_2007_to_2008.doc": "LibreOffice numbers a level with an LFO formatting override as a separate list",
	"doc/govuk__dh_133032.doc": "LibreOffice numbers a level with an LFO formatting override as a separate list",
	"doc/govuk__dh_133115.doc": "LibreOffice numbers a level with an LFO formatting override as a separate list",
	"doc/govuk__dh_133123.doc": "LibreOffice numbers a level with an LFO formatting override as a separate list",
	"doc/govuk__FSP_year_1_main_report.doc": "LibreOffice writes no bullet for an empty bulleted table cell",
	"doc/govuk__cscf-annual-report-template.doc": "LibreOffice leaves numbers out of empty numbered table cells",
	"doc/govuk__the-role-of-capital-infrastructure-in-services-for-young-people.doc":
		"LibreOffice neither numbers nor counts an empty numbered paragraph",
	"doc/govuk__crim-proc-rules-2020-part-05-temp.doc": "LISTNUM fields advance the list numbering in LibreOffice",
	"doc/govuk__crim-proc-rules-2025-part-06.doc": "LISTNUM fields advance the list numbering in LibreOffice",
	"doc/govuk__crim-proc-rules-2025-part-10.doc": "footnotes restart on each page: the shown number needs page layout",
	"doc/govuk__crim-proc-rules-2025-part-18.doc": "footnotes restart on each page: the shown number needs page layout",
	"doc/govuk__crim-proc-rules-2025-part-27.doc": "footnotes restart on each page: the shown number needs page layout",
	"doc/govuk__crim-proc-rules-2025-part-49.doc": "footnotes restart on each page: the shown number needs page layout",
	"doc/govuk__dh_134176.doc": "LibreOffice updates the DATE field to today",
	"doc/lo__tdf160301.doc": "LibreOffice shows a DOCVARIABLE field's name instead of its stored result",
	"doc/lo__tdf104239_chapterNumbering.doc": "LibreOffice maps heading numbering to its outline numbering and drops the level 5 bullet",
	"dot/lo__tdf120761_zOrder.dot": "LibreOffice writes the text of a text box in the header; the body is empty",
	"doc/poi__Bug61268.doc": "LibreOffice drops a bullet added by a tracked paragraph property change",
	// Refused here, read by LibreOffice.
	"dot/frictionless__package.dot": "not a Word binary file (LibreOffice imports it as plain text)",
	"dot/tika__testGRAPHVIZdc.dot": "not a Word binary file (LibreOffice imports it as plain text)",
	"doc/govuk__Audit_20Committee_20minutes_208_20June_202012.doc": "a DOCX package (LibreOffice detects the format by content)",
	"doc/lo__tdf150197_anlv2ListFormat.doc": "Word 6/95 file, refused",
	"doc/lo__tdf59896.doc": "Word 6/95 file, refused",
	"doc/poi__57843.doc": "Word 6/95 file, refused",
	"doc/poi__Bug60942b.doc": "Word 6/95 file, refused"
}

function load(path: string): Uint8Array {
	const bytes = readFileSync(join(CORPUS_DIR, path))

	return new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength)
}

type Outcome = { readonly text: string } | { readonly refused: DocErrorCode }

async function ourText(path: string): Promise<Outcome> {
	try {
		return { text: docText(await readDoc(bytesSource(load(path)))) }
	} catch (error) {
		if (error instanceof DocError) return { refused: error.code }

		throw error
	}
}

const normalize = (text: string): string =>
	text
		.replace(/^\ufeff/u, "")
		.replace(/[\s\u00a0]+/gu, " ")
		.trim()

describe.skipIf(!fetched)("Word binary corpus", () => {
	const files = corpusFiles()

	it("reads every Word 97 and later file and refuses the others with their typed reason", { timeout: 120_000 }, async () => {
		const outcomes: Record<string, string> = {}

		for (const path of files) {
			const result = await ourText(path)

			if ("refused" in result || REFUSED[path] !== undefined) outcomes[path] = "refused" in result ? result.refused : "read"
		}

		expect(outcomes).toEqual(REFUSED)
	})

	it("parses in time proportional to the file", { timeout: 120_000 }, async () => {
		const slow: string[] = []

		for (const path of files) {
			if (REFUSED[path] !== undefined) continue

			const bytes = load(path)
			const started = performance.now()

			docText(await readDoc(bytesSource(bytes)))

			const ms = performance.now() - started

			// Generous for a loaded machine: tens of milliseconds per megabyte are measured.
			if (ms > 250 + (400 * bytes.length) / 1e6) slow.push(`${path}: ${ms.toFixed(0)} ms for ${bytes.length} bytes`)
		}

		expect(slow).toEqual([])
	})
})

const OFFICES = offices()

describe.skipIf(!fetched || OFFICES.length === 0)("Word binary text against LibreOffice", () => {
	for (const office of OFFICES) {
		it(`equals the text export of ${office.name} apart from accepted differences`, { timeout: 3_600_000 }, async () => {
			const work = mkdtempSync(join(tmpdir(), "filen-office-doc-text-"))

			try {
				mkdirSync(join(work, "in"))

				const inputs = corpusFiles().map(path => {
					const copy = join(work, "in", path.replace("/", "__"))

					copyFileSync(join(CORPUS_DIR, path), copy)

					return copy
				})

				await convert(office, join(work, "profile"), "txt", inputs, join(work, "out"), 120_000)

				const unexpected: string[] = []
				let equal = 0

				for (const [k, path] of corpusFiles().entries()) {
					const output = outputPath(join(work, "out"), inputs[k] ?? "", "txt")
					const theirs = existsSync(output) ? normalize(readFileSync(output, "utf8")) : undefined
					const ours = await ourText(path)
					const same = "text" in ours ? theirs !== undefined && normalize(ours.text) === theirs : theirs === undefined

					if (same) equal++
					else if (ACCEPTED[path] === undefined) unexpected.push(path)
				}

				expect(unexpected).toEqual([])
				// Every other file is equal (the encrypted ones refused by both).
				expect(equal).toBe(corpusFiles().length - Object.keys(ACCEPTED).length)
			} finally {
				rmSync(work, { recursive: true, force: true })
			}
		})
	}
})
