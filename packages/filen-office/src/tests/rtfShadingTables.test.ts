import { strFromU8, unzipSync } from "fflate"
import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { GRAPHIC_INTS, PageGraphic } from "../core/docx/layout/output"
import { readRtf } from "../core/rtf/index"
import { layOut } from "./layoutFixtures"
import { editRtf, latin1, paragraphs, saveRtf, type RtfEditing } from "./rtfWriteFixtures"

// RTF shading patterns (paragraph, character, cell and row), the row and cell words of a table definition and the
// obsolete \level: read into the translation and written back, or the block locked; never dropped.

const H =
	"{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0\\froman Times New Roman;}}{\\colortbl;\\red255\\green0\\blue0;\\red0\\green0\\blue255;}"

async function bodyXml(bytes: Uint8Array): Promise<string> {
	const read = await readRtf(bytes)

	if (!read.ok) throw new Error(read.reason)

	return strFromU8(unzipSync(read.docx)["word/document.xml"] ?? new Uint8Array())
}

// The fills a translation's first page paints.
async function shadings(bytes: Uint8Array): Promise<number> {
	const read = await readRtf(bytes)

	if (!read.ok) throw new Error(read.reason)

	const page = (await layOut(read.docx)).pages[0]
	let count = 0

	for (let g = 0; g < (page?.graphicCount ?? 0); g++) if (page?.graphicInts[g * GRAPHIC_INTS] === PageGraphic.SHADING) count++

	return count
}

function find(editing: RtfEditing, text: string): number {
	let at = -1

	editing.session.state.doc.descendants((node, pos) => {
		if (at < 0 && node.isTextblock && node.textContent.includes(text)) at = pos

		return at < 0
	})

	if (at < 0) throw new Error(`no paragraph with ${text}`)

	return at
}

function type(editing: RtfEditing, text: string, inserted: string): void {
	const s = editing.session

	s.dispatch(s.state.tr.insertText(inserted, find(editing, text) + 1))
}

async function saved(editing: RtfEditing): Promise<string> {
	return strFromU8((await saveRtf(editing)).bytes, true)
}

async function insertRow(editing: RtfEditing, text: string, where: "insertAbove" | "insertBelow"): Promise<void> {
	const s = editing.session
	const cell = paragraphs(s.state.doc, true).find(({ node }) => node.textContent.includes(text))

	if (cell === undefined) throw new Error(`no cell with ${text}`)

	s.dispatch(s.state.tr.setSelection(TextSelection.create(s.state.doc, cell.pos + 1)))
	expect(await editing.editor.execute({ type: `table.row.${where}` })).toBe(true)
}

// The row definitions written (\trowd to its last \cellx), in order.
function rowDefinitions(rtf: string): string[] {
	return [...rtf.matchAll(/\\trowd(?:(?!\\trowd)[^{}])*?\\cellx\d+(?!(?:(?!\\trowd)[^{}])*?\\cellx)/gsu)].map(match => match[0])
}

const ROW =
	"\\trowd\\irow0\\irowband0\\ltrrow\\ts11\\trgaph108\\trleft720\\trcbpat1\\trcfpat2\\trshdng2500\\trftsWidth3\\trwWidth5000\\trftsWidthB3\\trwWidthB300" +
	"\\trpaddl108\\trpaddr108\\trpaddfl3\\trpaddfr3\\tblind828\\tblindtype3 \\clvertalt\\clftsWidth3\\clwWidth2400\\cellx3220\\clvertalt\\clftsWidth3\\clwWidth2600\\cellx5720"

function table(row: string, text = "A1"): string {
	return `${row}\\pard\\plain\\intbl ${text} cell\\cell B cell\\cell${row}\\row`
}

describe("RTF tables: row shading, indent, cell widths and table style", () => {
	it("reads row shading as the table's, cell widths as preferred widths and the invisible cell before a row", async () => {
		const xml = await bodyXml(latin1(`${H}${table(ROW)}\\pard After\\par}`))

		expect(xml).toContain('<w:shd w:val="pct25" w:color="0000FF" w:fill="FF0000"/><w:tblLayout')
		expect(xml).toContain('<w:tcW w:w="2400" w:type="dxa"/>')
		expect(xml).toContain('<w:tcW w:w="2600" w:type="dxa"/>')
		expect(xml).toContain('<w:wBefore w:w="300" w:type="dxa"/>')
	})

	it("writes an inserted row with its siblings' shading, indent, cell widths and dangling table style", async () => {
		const editing = await editRtf(latin1(`${H}${table(ROW)}\\pard After\\par}`))

		await insertRow(editing, "A1", "insertBelow")

		const out = await saved(editing)
		const rows = rowDefinitions(out)
		const written = rows.at(-1) ?? ""

		expect(rows).toHaveLength(3)

		for (const word of [
			"\\trcbpat1",
			"\\trcfpat2",
			"\\trshdng2500",
			"\\ts11",
			"\\trleft720",
			"\\tblind828\\tblindtype3",
			"\\trftsWidthB3\\trwWidthB300",
			"\\clftsWidth3\\clwWidth2400",
			"\\clftsWidth3\\clwWidth2600"
		])
			expect(written, word).toContain(word)

		const xml = await bodyXml(latin1(out))

		expect(xml.match(/<w:tr>/gu)).toHaveLength(2)
		expect(xml).not.toContain("tblPrEx")
	})

	it("restates \\tblind as the source measures it, to the cell's edge as well as to its text", async () => {
		const edge = ROW.replace("\\tblind828", "\\tblind720")
		const editing = await editRtf(latin1(`${H}${table(edge)}\\pard After\\par}`))

		await insertRow(editing, "A1", "insertBelow")
		expect(rowDefinitions(await saved(editing)).at(-1)).toContain("\\trleft720\\tblind720\\tblindtype3")
	})

	it("writes a table style the style sheet holds by its number", async () => {
		const styled = `${H}{\\stylesheet{\\ql \\snext0 Normal;}{\\*\\ts11\\tsrowd\\trftsWidthB3\\trpaddl108\\trpaddr108\\trpaddfl3\\trpaddfr3 \\snext11 Normal Table;}}`
		const editing = await editRtf(latin1(`${styled}${table(ROW)}\\pard After\\par}`))

		await insertRow(editing, "A1", "insertBelow")
		expect(rowDefinitions(await saved(editing)).at(-1)).toContain("\\ts11")
	})

	it("shows a row shaded unlike the first in its row's table exceptions, and writes a row modelled on it so", async () => {
		const plain = ROW.replace("\\trcbpat1\\trcfpat2\\trshdng2500", "")
		const blue = ROW.replace("\\trcbpat1\\trcfpat2\\trshdng2500", "\\trcbpat2\\trbgdkcross")
		const source = `${H}${table(ROW)}${table(plain, "A2")}${table(blue, "A3")}\\pard After\\par}`
		const xml = await bodyXml(latin1(source))

		expect(xml).toContain('<w:tr><w:tblPrEx><w:shd w:val="clear" w:color="auto" w:fill="auto"/></w:tblPrEx>')
		expect(xml).toContain('<w:tr><w:tblPrEx><w:shd w:val="horzCross" w:color="auto" w:fill="0000FF"/></w:tblPrEx>')
		// The view fills the first and third rows' two cells each, the second's none.
		expect(await shadings(latin1(source))).toBe(4)
		expect(await shadings(latin1(`${H}${table(plain)}\\pard After\\par}`))).toBe(0)

		const editing = await editRtf(latin1(source))

		expect(editing.locked).toBe(0)
		await insertRow(editing, "A3", "insertBelow")

		const out = await saved(editing)
		const rows = rowDefinitions(out)

		expect(rows).toHaveLength(7)
		expect(rows.at(-1)).toContain("\\trcbpat2")
		expect(rows.at(-1)).toContain("\\trbgdkcross")
		expect(rows.at(-1)).not.toContain("\\trcbpat1")
		expect((await bodyXml(latin1(out))).match(/<w:tblPrEx><w:shd w:val="horzCross"/gu)).toHaveLength(2)
	})

	it("locks a table holding a nested table whose definition names a row pattern by number", async () => {
		const nested =
			"\\trowd\\trgaph108\\trleft0\\cellx6000\\pard\\plain\\intbl\\itap2 N1 cell\\nestcell{\\*\\nesttableprops\\trowd\\trgaph108\\trpat2\\cellx2000\\nestrow}" +
			"\\pard\\plain\\intbl Outer cell\\cell\\row\\pard After\\par}"

		expect((await editRtf(latin1(`${H}${nested}`))).locked).toBe(1)
	})

	it("locks a table whose definition names a row pattern by number or a pattern beside a percentage", async () => {
		for (const words of ["\\trpat3", "\\trbghoriz\\trshdng2000", "\\clbghoriz\\clshdng2000"]) {
			const row = `\\trowd\\trgaph108${words}\\cellx3000\\cellx6000`
			const editing = await editRtf(latin1(`${H}${row}\\pard\\plain\\intbl a\\cell b\\cell${row}\\row\\pard After\\par}`))

			expect(editing.locked, words).toBe(1)
		}
	})
})

// RTF 1.9.1's pattern words after their prefix (cells and rows spell the dark horizontal one "dkhor"), and ST_Shd's.
const PATTERNS: readonly (readonly [string, string, string])[] = [
	["horiz", "horiz", "thinHorzStripe"],
	["vert", "vert", "thinVertStripe"],
	["fdiag", "fdiag", "thinReverseDiagStripe"],
	["bdiag", "bdiag", "thinDiagStripe"],
	["cross", "cross", "thinHorzCross"],
	["dcross", "dcross", "thinDiagCross"],
	["dkhoriz", "dkhor", "horzStripe"],
	["dkvert", "dkvert", "vertStripe"],
	["dkfdiag", "dkfdiag", "reverseDiagStripe"],
	["dkbdiag", "dkbdiag", "diagStripe"],
	["dkcross", "dkcross", "horzCross"],
	["dkdcross", "dkdcross", "diagCross"]
]

describe("RTF shading patterns", () => {
	it("reads every paragraph, character, cell and row pattern as its w:shd pattern", async () => {
		for (const [text, cells, value] of PATTERNS) {
			const row = `\\trowd\\trgaph108\\trbg${cells}\\trcfpat2\\clbg${cells}\\clcfpat1\\clcbpat2\\cellx3000`
			const xml = await bodyXml(
				latin1(
					`${H}\\pard\\bg${text}\\cfpat1\\cbpat2 Para\\par\\pard {\\chbg${text}\\chcfpat1\\chcbpat2 Run}\\par${row}\\pard\\plain\\intbl c\\cell${row}\\row\\pard End\\par}`
				)
			)
			const shd = `<w:shd w:val="${value}" w:color="FF0000" w:fill="0000FF"/>`

			expect(xml, text).toContain(`<w:pPr>${shd}</w:pPr>`)
			expect(xml, text).toContain(`<w:rPr>${shd}</w:rPr>`)
			expect(xml, text).toMatch(new RegExp(`<w:tcPr>[^]*${shd}[^]*</w:tcPr>`, "u"))
			expect(xml, text).toContain(`<w:shd w:val="${value}" w:color="0000FF" w:fill="auto"/><w:tblLayout`)
		}
	})

	it("writes every pattern back when its paragraph, run or row is rewritten", async () => {
		for (const [text, cells] of PATTERNS) {
			const row = `\\trowd\\trgaph108\\trbg${cells}\\clbg${cells}\\clcbpat2\\cellx3000`
			const editing = await editRtf(
				latin1(
					`${H}\\pard\\bg${text}\\cbpat2 Para\\par\\pard {\\chbg${text}\\chcbpat2 Run}\\par${row}\\pard\\plain\\intbl cell\\cell${row}\\row\\pard End\\par}`
				)
			)

			expect(editing.locked, text).toBe(0)
			type(editing, "Para", "X")
			type(editing, "Run", "Y")
			await insertRow(editing, "cell", "insertBelow")

			const out = await saved(editing)
			const written = rowDefinitions(out).at(-1) ?? ""

			expect(out, text).toMatch(new RegExp(`\\\\bg${text}\\\\cbpat2\\{\\\\plain XPara`, "u"))
			expect(out, text).toMatch(new RegExp(`\\\\chbg${text}\\\\chcbpat2 YRun`, "u"))
			expect(written, text).toContain(`\\trbg${cells}`)
			expect(written, text).toContain(`\\clbg${cells}\\clcbpat2`)
		}
	})

	it("locks a paragraph shown under a pattern beside a percentage, or under the obsolete \\level", async () => {
		for (const source of [
			`${H}\\pard\\bghoriz\\shading2000 One para\\par\\pard Two\\par}`,
			`${H}\\pard {\\chbghoriz\\chshdng2000 One} para\\par\\pard Two\\par}`,
			`${H}\\pard\\level3 One para\\par\\pard Two\\par}`
		]) {
			const editing = await editRtf(latin1(source))

			expect(editing.locked, source).toBe(1)
		}
	})
})
