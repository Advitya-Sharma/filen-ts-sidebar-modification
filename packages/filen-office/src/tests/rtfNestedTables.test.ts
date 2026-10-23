import { existsSync, readFileSync, readdirSync } from "node:fs"
import { join } from "node:path"
import { strFromU8, unzipSync } from "fflate"
import type { Node as ProseNode } from "prosemirror-model"
import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { readRtf } from "../core/rtf/index"
import { RtfWriteRefusal } from "../core/rtf/write/out"
import { CORPUS_DIR } from "./corpusDir"
import { editRtf, latin1, paragraphs, saveRtf, type RtfEditing } from "./rtfWriteFixtures"

// Rows of nested tables (RTF 1.9.1 "Tables": \itapN paragraphs ending in \nestcell, the row's definitions in
// {\*\nesttableprops ... \nestrow}) inserted and deleted: the new row written in that form, the rest of the outer
// row's bytes kept.

const H = "{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0\\froman Times New Roman;}}"

async function documentXml(bytes: Uint8Array): Promise<string> {
	const read = await readRtf(bytes)

	if (!read.ok) throw new Error(read.reason)

	return strFromU8(unzipSync(read.docx)["word/document.xml"] ?? new Uint8Array())
}

// The tables' row counts in document order (outer tables before those nested in them).
function rowCounts(xml: string): number[] {
	const counts: number[] = []
	const stack: number[] = []

	for (const match of xml.matchAll(/<w:tbl>|<\/w:tbl>|<w:tr[ >]/gu)) {
		if (match[0] === "<w:tbl>") {
			stack.push(counts.length)
			counts.push(0)
		} else if (match[0] === "</w:tbl>") stack.pop()
		else {
			const at = stack.at(-1)

			if (at !== undefined) counts[at] = (counts[at] ?? 0) + 1
		}
	}

	return counts
}

async function execute(editing: RtfEditing, text: string, type: string): Promise<boolean> {
	const s = editing.session
	const cell = paragraphs(s.state.doc, true).find(({ node }) => node.textContent.includes(text))

	if (cell === undefined) throw new Error(`no cell with ${text}`)

	s.dispatch(s.state.tr.setSelection(TextSelection.create(s.state.doc, cell.pos + 1)))

	return editing.editor.execute({ type } as never)
}

// As LibreOffice writes a nested table: no groups around it.
const BARE =
	`${H}\\trowd\\trgaph108\\trleft0\\cellx6000\\pard\\intbl\\itap2 N1 cell\\nestcell{\\*\\nesttableprops\\trowd\\trgaph108\\trleft100\\tblind208\\tblindtype3\\cellx3000\\nestrow}` +
	"\\pard\\intbl\\itap2 N2 cell\\nestcell{\\*\\nesttableprops\\trowd\\trgaph108\\trleft100\\tblind208\\tblindtype3\\cellx3000\\nestrow}" +
	"\\pard\\intbl Outer cell\\cell\\row\\pard After\\par}"

// As Word writes it (RTF 1.9.1's example): each nested cell and row followed by {\nonesttables \par }, in groups.
const WORD =
	`${H}\\trowd\\trgaph108\\trleft0\\cellx6000\\pard\\plain\\intbl\\itap2 {N1 cell\\nestcell{\\nonesttables\n\\par }}` +
	"\\pard\\plain\\intbl\\itap2 {{\\*\\nesttableprops\\trowd\\trgaph108\\trleft8\\clftsWidth3\\clwWidth2340\\cellx2348\\nestrow}{\\nonesttables\n\\par }}" +
	"\\pard\\plain\\intbl\\itap2 {N2 cell\\nestcell{\\nonesttables\n\\par }}" +
	"\\pard\\plain\\intbl\\itap2 {{\\*\\nesttableprops\\trowd\\trgaph108\\trleft8\\clftsWidth3\\clwWidth2340\\cellx2348\\nestrow}{\\nonesttables\n\\par }}" +
	"\\pard\\plain\\intbl {Outer cell\\cell}\\trowd\\trgaph108\\trleft0\\cellx6000\\row\\pard\\plain After\\par}"

// Three levels: a table in a cell of a table nested in the outer one.
const DEEP =
	`${H}\\trowd\\trgaph108\\trleft0\\cellx8000` +
	"\\pard\\intbl\\itap3 D1 cell\\nestcell{\\*\\nesttableprops\\trowd\\trgaph108\\trleft200\\cellx3000\\nestrow}{\\nonesttables\\par}" +
	"\\pard\\intbl\\itap2 M1 cell\\nestcell{\\*\\nesttableprops\\trowd\\trgaph108\\trleft100\\cellx6000\\nestrow}{\\nonesttables\\par}" +
	"\\pard\\intbl Outer cell\\cell\\row\\pard After\\par}"

// What `out` adds to `source` at one place (`out` is `source` with it inserted).
function inserted(source: string, out: string): string {
	let prefix = 0

	while (prefix < source.length && source[prefix] === out[prefix]) prefix++

	let suffix = 0

	while (suffix < source.length - prefix && source.at(-1 - suffix) === out.at(-1 - suffix)) suffix++

	expect(prefix + suffix).toBe(source.length)

	return out.slice(prefix, out.length - suffix)
}

const NEW_ROW_2 =
	/\{\\pard\\plain\\intbl\\itap2\\nestcell\{\\nonesttables\\par\}\}\{\\\*\\nesttableprops\\trowd[^{}]*\\nestrow\}\{\\nonesttables\\par\}$/u

describe("RTF nested tables", () => {
	for (const [name, source, kept] of [
		["as LibreOffice writes them", BARE, "\\pard\\intbl Outer cell\\cell\\row"],
		["as Word writes them", WORD, "\\pard\\plain\\intbl {Outer cell\\cell}\\trowd\\trgaph108\\trleft0\\cellx6000\\row"]
	] as const) {
		it(`inserts a row into a nested table ${name}, keeping the rest of the outer row's bytes`, async () => {
			const editing = await editRtf(latin1(source))

			expect(editing.locked).toBe(0)
			expect(await execute(editing, "N1", "table.row.insertBelow")).toBe(true)

			const written = await saveRtf(editing)
			const out = strFromU8(written.bytes, true)

			// The source as it was around one insertion: the new row, \itap2 paragraphs ending in \nestcell and its
			// definition in the nested-table destination, after the first nested row.
			expect(inserted(source, out)).toMatch(NEW_ROW_2)
			expect(out.indexOf(inserted(source, out))).toBeGreaterThan(out.indexOf("N1 cell"))
			expect(out.indexOf(inserted(source, out))).toBeLessThan(out.indexOf("N2 cell"))
			expect(out).toContain(kept)
			expect(rowCounts(await documentXml(written.bytes))).toEqual([1, 3])
		})

		it(`deletes a row of a nested table ${name}`, async () => {
			const editing = await editRtf(latin1(source))

			expect(await execute(editing, "N2", "table.row.delete")).toBe(true)

			const written = await saveRtf(editing)
			const xml = await documentXml(written.bytes)

			expect(rowCounts(xml)).toEqual([1, 1])
			expect(xml).not.toContain("N2 cell")
			expect(strFromU8(written.bytes, true)).toContain(kept)
		})
	}

	it("inserts a row into a table nested three deep", async () => {
		const editing = await editRtf(latin1(DEEP))

		expect(editing.locked).toBe(0)
		expect(await execute(editing, "D1", "table.row.insertAbove")).toBe(true)

		const written = await saveRtf(editing)
		const out = strFromU8(written.bytes, true)

		expect(inserted(DEEP, out)).toMatch(
			/^\{\\pard\\plain\\intbl\\itap3\\nestcell\{\\nonesttables\\par\}\}\{\\\*\\nesttableprops\\trowd[^{}]*\\nestrow\}\{\\nonesttables\\par\}$/u
		)
		expect(out.indexOf(inserted(DEEP, out))).toBeLessThan(out.indexOf("D1 cell"))
		expect(rowCounts(await documentXml(written.bytes))).toEqual([1, 1, 2])
	})
})

// The first paragraph of a nested table's cell outside a locked block, as a position inside it; -1 for none.
function nestedCell(doc: ProseNode): number {
	let target = -1

	doc.descendants((node, pos) => {
		if (target >= 0 || node.type.name === "locked_block") return false

		if (node.type.name === "table") {
			let nested = -1

			node.descendants((inner, at) => {
				if (nested < 0 && inner.type.name === "table")
					inner.descendants((cell, k) => {
						if (nested < 0 && cell.isTextblock) nested = at + k + 2

						return nested < 0
					})

				return nested < 0
			})

			if (nested >= 0) target = pos + nested + 1

			return false
		}

		return true
	})

	return target
}

const DIRECTORY = join(CORPUS_DIR, "rtf")

// A corpus file with what locks its nested tables for reasons of their own taken out: smart tags (\*\xmlopen), form
// fields, Word's \sectd after a field in a header, a bare \ulc.
function unlocked(text: string): string {
	return text
		.replace(/\{\\\*\\xmlopen\\xmlns\d+\{\\factoidname [^}]*\}\}|\{\\\*\\xmlclose\}/gu, "")
		.replace(/\\sectd \\ltrsect\s*\\linex0\\endnhere\\sectdefaultcl\\sftnbj /gu, "")
		.replace(/\{\\\*\\formfield\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\}\}/gu, "")
		.replace(/\{\\\*\\datafield [0-9a-f\s]*\}/gu, "")
		.replace(/\\ulc2\n/gu, "\n")
}

// Inserts a row below the first nested cell outside a locked block (in the body, else in a header or footer), or
// deletes its row, and saves: "saved", "locked" (no such cell) or the refusal.
async function nestedEdit(bytes: Uint8Array, type: string, untracked = false): Promise<string> {
	const editing = await editRtf(bytes)

	// RTF 1.9.1 has no form for a tracked row insertion or deletion ("Table Row Revision Tracking"): a document
	// tracking changes refuses them, so its shape is tried with tracking off as well.
	if (untracked) editing.opened.context.review.tracking = false
	let s = editing.session
	let target = nestedCell(s.state.doc)

	for (const story of editing.opened.context.model.stories) {
		if (target >= 0) break
		if (story.kind !== "header" && story.kind !== "footer") continue

		expect(editing.editor.enterStory(story.index, -1)).toBe(true)
		s = editing.editor.storySession(story.index) ?? s
		target = nestedCell(s.state.doc)
	}

	if (target < 0) return "locked"

	s.dispatch(s.state.tr.setSelection(TextSelection.create(s.state.doc, target)))

	if (!(await editing.editor.execute({ type } as never))) return "not applied"

	try {
		await saveRtf(editing)

		return "saved"
	} catch (error) {
		if (!(error instanceof RtfWriteRefusal)) throw error

		return `refused ${error.message.slice(0, 200)}`
	}
}

describe.skipIf(!existsSync(DIRECTORY))("RTF nested tables over the corpus", () => {
	it("inserts and deletes nested rows in every corpus file with nested tables", { timeout: 600_000 }, async () => {
		const files = readdirSync(DIRECTORY)
			.filter(name => name.toLowerCase().endsWith(".rtf"))
			.sort()
			.filter(name => readFileSync(join(DIRECTORY, name), "latin1").includes("\\nesttableprops"))
		const results: string[] = []

		expect(files.length).toBe(4)

		for (const name of files) {
			const text = readFileSync(join(DIRECTORY, name), "latin1")

			for (const type of ["table.row.insertBelow", "table.row.delete"]) {
				const source = await nestedEdit(latin1(text), type)

				// Locked for other words, the file's nested tables save once those are taken out.
				const opened = source === "locked" ? await nestedEdit(latin1(unlocked(text)), type) : source
				const outcome = opened.startsWith("refused rowRevision")
					? `${opened.slice(0, 19)}, untracked ${await nestedEdit(latin1(unlocked(text)), type, true)}`
					: opened

				results.push(`${name} ${type}: ${source === "locked" ? "locked, then " : ""}${outcome}`)
			}
		}

		// The writer's own proof read every save back as edited.
		expect(results.filter(line => !line.endsWith("saved"))).toEqual([])
	})
})
