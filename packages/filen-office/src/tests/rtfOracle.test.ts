import { existsSync, readFileSync, readdirSync, writeFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { LINE_INTS, LineFlag, LineInt } from "@filen/office/core/text"
import { Region, RevisionView, type DocumentLayout } from "@filen/office/core/docx"
import { symbolTable } from "@filen/office/core/fonts"
import { readRtf } from "../core/rtf/index"
import { CORPUS_DIR } from "./corpusDir"
import { layOut } from "./layoutFixtures"

// The corpus RTF files read into the docx engine and laid out, against LibreOffice 7.5's own reading:
// FILEN_OFFICE_RTF_ORACLE names a directory holding <name>.txt (soffice --convert-to "txt:Text (encoded):UTF8")
// and pages.json ({ "<name>": <pages of soffice --convert-to pdf> }) for the corpus's rtf/ files. Body text
// (hidden text and tracked deletions included, as LibreOffice exports them) must equal LibreOffice's on at
// least 95 % of the files and page counts be within one on at least 85 %. Opt-in; the report goes to
// <oracle>/report.json.

const ORACLE = process.env["FILEN_OFFICE_RTF_ORACLE"] || undefined
// The files read: the corpus's, or FILEN_OFFICE_RTF_ORACLE_INPUT's (the RTF writer's saves of them, with the oracle
// made from those).
const DIRECTORY = process.env["FILEN_OFFICE_RTF_ORACLE_INPUT"] || join(CORPUS_DIR, "rtf")
const ENABLED = ORACLE !== undefined && existsSync(join(ORACLE, "pages.json")) && existsSync(DIRECTORY)

// List bullets LibreOffice's text export writes as U+2022 whatever the glyph; symbol-font codes it writes as
// their 8-bit character.
const BULLETS = /[\u2022\u25a0\u25aa\u25cf\u25cb\u25e6\u00b7\u25ba\u27a2\u2023\u2043\u2b9a\u2981\uf0b7\uf0a7\uf0d8\uf076]/gu

// What a Wingdings code draws, back to the code (LibreOffice exports the code).
const WINGDINGS = new Map<string, string>()

for (const [index, unicode] of (symbolTable("wingdings") ?? []).entries()) {
	if (unicode > 0x7f) WINGDINGS.set(String.fromCodePoint(unicode), String.fromCharCode(index + 0x20))
}

function words(text: string): string[] {
	return [...text.replace(BULLETS, "\u2022")]
		.map(character => WINGDINGS.get(character) ?? character)
		.join("")
		.normalize("NFC")
		.replace(/\ufeff|\ufffc|\u00ad/gu, "")
		.replace(/[\uf020-\uf0ff]/gu, code => String.fromCharCode((code.codePointAt(0) ?? 0xf020) - 0xf000))
		.replace(/\u00a0/gu, " ")
		.replace(/\u2011/gu, "-")
		.split(/\s+/u)
		.filter(word => word !== "")
}

// The body's paragraphs as laid out, in the order they start: each paragraph's lines joined (lines of
// cells side by side interleave in the display list); a table's header rows repeated on later pages count once.
export function bodyText(layout: DocumentLayout): string {
	const paragraphs = new Map<string, string>()
	const ended = new Set<string>()

	for (const page of layout.pages) {
		const text = page.text

		for (let line = 0; line < text.lineCount; line++) {
			const ints = line * LINE_INTS
			const source = text.lineInts[ints + LineInt.SOURCE] ?? 0
			const region = page.sourceRegion(source)

			if (region !== Region.BODY && region !== Region.GENERATED) continue

			const key = region === Region.GENERATED ? `g${page.number}:${source}` : `${page.sourceStory(source)}:${page.sourceItem(source)}`
			if (ended.has(key)) continue

			const shown = (page.sourceTexts[source] ?? "").slice(
				text.lineInts[ints + LineInt.TEXT_START] ?? 0,
				text.lineInts[ints + LineInt.TEXT_END] ?? 0
			)

			paragraphs.set(key, (paragraphs.get(key) ?? "") + shown)

			if (((text.lineInts[ints + LineInt.FLAGS] ?? 0) & LineFlag.LAST) !== 0) ended.add(key)
		}
	}

	return [...paragraphs.values()].join("\n")
}

// Files whose text differs only where LibreOffice 7.5 reads them otherwise than Word does (Word's own \\listtext
// labels, which the files store, agree with ours) or than the specification says.
const ACCEPTED: Readonly<Record<string, string>> = {
	govuk__Cat_B3_Impact_Assessment:
		"LibreOffice leaves list numbers out of table cells (Word's \\listtext has them) and shows a footnote mark as 0",
	govuk__Draft_Gambling_Act_2005_Order_2011_Validated: "LibreOffice shows no footnote marks numbered in letters (\\ftnnalc)",
	govuk__Equality_Impact_Assessment_for_the_draft_Heritage_Protection_Bill_and_White_Paper_RTF__195kb_:
		"LibreOffice leaves one footnote mark out, numbering the rest one lower",
	govuk__Gov_Response_Broadband_for_all_Cm_8457: "LibreOffice shows a PAGE field on the third page as 0",
	govuk__Triennial_Review_of_the_TVC_Stage_Two: 'LibreOffice exports a Courier New "o" bullet as U+25E6',
	"govuk__better-workplace-pensions-putting-savers-interests-first":
		"LibreOffice numbers paragraphs through their style's \\ls; the specification numbers only paragraphs with their own \\lsN",
	"govuk__child-maintenance-client-funds-account-2012-to-2013": "LibreOffice's text export leaves framed paragraphs out",
	govuk__dh_134435: "LibreOffice continues list numbers where Word restarts them (Word's \\listtext agrees with ours)",
	"govuk__fas-regs-2011-consult-annexb":
		"LibreOffice continues list numbers where Word restarts them (Word's \\listtext agrees with ours)",
	"govuk__housing-benefit-factsheet-1-removal-spare-room-subsidy-welsh": "LibreOffice decodes a bullet as the control character U+0097",
	"govuk__single-tier-pension": "LibreOffice leaves out a bullet in a table cell that Word's \\listtext has",
	"govuk__work-capability-assessment-year-4-paul-litchfield":
		"LibreOffice continues list numbers where Word restarts them (Word's \\listtext agrees with ours)",
	lo__FWDP90_min: "LibreOffice shows the text of {\\*\\revtbl Unknown;} in the body",
	"lo__forcepoint72-1": "A fuzzed file: LibreOffice recovers its garbage differently",
	"lo__forcepoint93-2": "A fuzzed file: LibreOffice recovers its garbage differently",
	tika__testRTFHyperlinkAndStyles:
		"LibreOffice decodes \\fcharset1 text as U+FFFD; the specification reads it in the document's code page",
	tika__testRTFInvalidUnicode: 'LibreOffice shows unpaired surrogates as "?"; ours are U+FFFD'
}

// Saves (FILEN_OFFICE_RTF_ORACLE_INPUT) whose text differs only where LibreOffice 7.5's text export writes a bullet
// list's label as U+2022 whatever its level's glyph (measured: a \\levelnfc23 level of Courier New "o" exports as
// U+2022): the scripted list edit put a paragraph in an existing list of the file whose first level is "o" or "–".
// Or where the scripted section break ends an empty paragraph of a list: LibreOffice drops that paragraph (and its
// label) when it makes the section a page style.
const ACCEPTED_SAVED: Readonly<Record<string, string>> = {
	govuk__BDUKFramework_Delivery_Model_Summary: "LibreOffice drops the empty list paragraph a section break ends",
	govuk__NShott_InterimFindingsLocalTV_240910: "LibreOffice drops the empty list paragraph a section break ends",
	govuk__dh_134460: "LibreOffice drops the empty list paragraph a section break ends",
	govuk__DCMS_Annual_Report_2009: 'LibreOffice exports a "–" bullet label as U+2022',
	"govuk__benefit-cap-review-of-the-first-year": 'LibreOffice exports an "o" bullet label as U+2022',
	"govuk__bereavement-benefit-consultation-response": 'LibreOffice exports an "o" bullet label as U+2022',
	"govuk__building-understanding-main-report": 'LibreOffice exports an "o" bullet label as U+2022',
	"govuk__health-at-work": 'LibreOffice exports an "o" bullet label as U+2022'
}
const SAVES = process.env["FILEN_OFFICE_RTF_ORACLE_INPUT"] !== undefined

describe.skipIf(!ENABLED)("RTF against LibreOffice 7.5", () => {
	it("reads the corpus as LibreOffice does", { timeout: 1_800_000 }, async () => {
		const oracle = ORACLE as string
		const pages = JSON.parse(readFileSync(join(oracle, "pages.json"), "utf8")) as Record<string, number>
		const files = readdirSync(DIRECTORY)
			.filter(name => name.endsWith(".rtf"))
			.sort()
		const report: { name: string; text: boolean; pages: number; theirPages: number; readMs: number }[] = []

		for (const name of files) {
			const base = name.slice(0, -4)
			const bytes = new Uint8Array(readFileSync(join(DIRECTORY, name)))
			const started = performance.now()
			const read = await readRtf(bytes)
			const readMs = performance.now() - started

			expect(read.ok, name).toBe(true)

			if (!read.ok) continue

			const layout = await layOut(read.docx, { view: RevisionView.Markup, showHidden: true })
			const theirs = words(readFileSync(join(oracle, `${base}.txt`), "utf8"))
			const shown = bodyText(layout)
			const ours = words(shown)

			writeFileSync(join(oracle, `${base}.ours.txt`), shown)

			report.push({
				name: base,
				text: theirs.length === ours.length && theirs.every((word, k) => word === ours[k]),
				pages: layout.pages.length,
				theirPages: pages[base] ?? -1,
				readMs: Math.round(readMs)
			})
		}

		const textEqual = report.filter(entry => entry.text).length
		const pagesClose = report.filter(entry => Math.abs(entry.pages - entry.theirPages) <= 1).length

		writeFileSync(join(oracle, "report.json"), JSON.stringify({ files: report.length, textEqual, pagesClose, report }, null, 1))

		const unexpected = report
			.filter(entry => !entry.text && ACCEPTED[entry.name] === undefined && (!SAVES || ACCEPTED_SAVED[entry.name] === undefined))
			.map(entry => entry.name)

		// Every difference is one LibreOffice makes; with those counted, 95 %.
		expect(unexpected).toEqual([])
		const explained = report.filter(
			entry => entry.text || ACCEPTED[entry.name] !== undefined || (SAVES && ACCEPTED_SAVED[entry.name] !== undefined)
		).length

		expect(explained / report.length).toBeGreaterThanOrEqual(0.95)
		expect(pagesClose / report.length).toBeGreaterThanOrEqual(0.85)
	})
})
