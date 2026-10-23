import { existsSync, readFileSync, readdirSync, writeFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { bytesSource } from "@filen/office/core/opc"
import { decodeUtf8, parseXml, type XmlDocument } from "@filen/office/core/xml"
import { translateOdt } from "../core/odt/index"
import { CORPUS_DIR } from "./corpusDir"

// The ODT writer's saves of the corpus (FILEN_OFFICE_ODT_WRITE_OUT, from odtWriterCorpus.test.ts) against LibreOffice
// 7.5's reading of them. FILEN_OFFICE_ODT_WRITE_ORACLE names a directory holding LibreOffice's conversions of each
// source and each save: source/<name>.txt, saved/<name>.txt (soffice --convert-to "txt:Text (encoded):UTF8") and
// source/<name>.odt, saved/<name>.odt (soffice --convert-to odt). What LibreOffice sees change from the source to the
// save must be what we see change: the same words removed and added, and the same change in paragraphs, tables, table
// rows, list paragraphs and headings (counted in LibreOffice's own re-save, read by our reader). A save LibreOffice
// cannot open fails. Opt-in; the report goes to <oracle>/report.json.

const ORACLE = process.env["FILEN_OFFICE_ODT_WRITE_ORACLE"] || undefined
const SAVED = process.env["FILEN_OFFICE_ODT_WRITE_OUT"] || undefined
const ENABLED = ORACLE !== undefined && SAVED !== undefined && existsSync(join(ORACLE, "saved"))

// List labels LibreOffice's text export writes at a list paragraph's start (indented four spaces a level, the label, a
// space), which our reading leaves to the numbering: the comparison is of the text.
const LABEL = /^(?: {4})+\S{1,12} /u
// Outline numbers LibreOffice writes at a heading's start (a heading the edit numbers renumbers those after it), and a
// list label at a line's start: stripped from both readings alike.
const NUMBER = /^\d+(?:\.\d+)*\.? /u

// Saves whose change reads the same to LibreOffice but splits into other words: what the files hold around the edit
// LibreOffice's text export writes differently from the body text we compare (the same words changed).
const ACCEPTED: Readonly<Record<string, string>> = {
	"govuk__EMF1-Lot-5-Call-off-contract-and-order-form.odt":
		"LibreOffice writes the outline label 4. into the line of the heading the edit numbered: one label moved, the same words changed",
	"odftk__text-extract.odt":
		"LibreOffice exports the deleted paragraph's footnote citation 1 after its last word: paragraph1 against paragraph"
}

function words(lines: readonly string[]): string[] {
	return lines
		.map(line => line.replace(LABEL, "").replace(NUMBER, ""))
		.join(" ")
		.normalize("NFC")
		.replace(/\ufeff|\ufffc|\u00ad/gu, "")
		.replace(/\u00a0/gu, " ")
		.replace(/\u2011/gu, "-")
		.split(/\s+/u)
		.filter(word => word !== "")
}

// What changed from a to b: the words removed and added, in order, past their common subsequence.
function change(a: readonly string[], b: readonly string[]): { removed: string[]; added: string[] } {
	let head = 0

	while (head < a.length && head < b.length && a[head] === b[head]) head++

	let tail = 0

	while (tail < a.length - head && tail < b.length - head && a[a.length - 1 - tail] === b[b.length - 1 - tail]) tail++

	const x = a.slice(head, a.length - tail)
	const y = b.slice(head, b.length - tail)
	// A shortest edit script (Myers): time and room grow with the edits, not the text.
	const max = x.length + y.length
	const offset = max + 1
	const trace: Int32Array[] = []
	const v = new Int32Array(2 * max + 3)
	let found = max === 0

	// Each step keeps only the diagonals it reached (-d - 1 to d + 1).
	for (let d = 0; d <= max && !found; d++) {
		trace.push(v.slice(offset - d - 1, offset + d + 2))

		for (let k = -d; k <= d; k += 2) {
			let i =
				k === -d || (k !== d && (v[offset + k - 1] ?? 0) < (v[offset + k + 1] ?? 0))
					? (v[offset + k + 1] ?? 0)
					: (v[offset + k - 1] ?? 0) + 1
			let j = i - k

			while (i < x.length && j < y.length && x[i] === y[j]) {
				i++
				j++
			}

			v[offset + k] = i

			if (i >= x.length && j >= y.length) {
				found = true
				break
			}
		}
	}

	const removed: string[] = []
	const added: string[] = []
	let i = x.length
	let j = y.length

	for (let d = trace.length - 1; d >= 0 && (i > 0 || j > 0); d--) {
		// The diagonals as step d found them (step d - 1's ends), from -d - 1.
		const previous = trace[d] ?? new Int32Array(0)
		const at = (diagonal: number): number => previous[diagonal + d + 1] ?? 0
		const k = i - j
		const down = k === -d || (k !== d && at(k - 1) < at(k + 1))
		const priorK = down ? k + 1 : k - 1
		const priorI = d === 0 ? 0 : at(priorK)
		const priorJ = priorI - priorK

		while (i > priorI && j > priorJ) {
			i--
			j--
		}

		if (d > 0) {
			if (down) added.push(y[--j] ?? "")
			else removed.push(x[--i] ?? "")
		}
	}

	return { removed: removed.sort(), added: added.sort() }
}

interface Shape {
	paragraphs: number
	tables: number
	rows: number
	listed: number
	headings: number
}

function count(document: XmlDocument, node: number, shape: Shape, lines: string[] | undefined, boxes?: string[]): void {
	for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) {
		const local = document.localName(child)

		if (local === "p") {
			shape.paragraphs++

			// Text boxes' text apart: LibreOffice's text export leaves frames out.
			const raw = decodeUtf8(document.raw(child))

			for (const box of raw.matchAll(/<w:txbxContent>([\s\S]*?)<\/w:txbxContent>/gu))
				boxes?.push([...(box[1] ?? "").matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/gu)].map(match => match[1] ?? "").join(""))

			const text = [
				...raw
					.replace(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/gu, "")
					.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>|<w:tab\/>|<w:br\/>/gu)
			]
				.map(match => match[1] ?? " ")
				.join("")

			lines?.push(
				text
					.replace(/&lt;/gu, "<")
					.replace(/&gt;/gu, ">")
					.replace(/&quot;/gu, '"')
					.replace(/&apos;/gu, "'")
					.replace(/&amp;/gu, "&")
			)

			const markup = decodeUtf8(document.raw(child))

			if (markup.includes("<w:numPr>")) shape.listed++
			if (markup.includes("<w:outlineLvl ")) shape.headings++
		} else if (local === "tbl") {
			shape.tables++
			count(document, child, shape, lines, boxes)
		} else if (local === "tr") {
			shape.rows++
			count(document, child, shape, lines, boxes)
		} else if (local === "tc" || local === "body") count(document, child, shape, lines, boxes)
	}
}

async function read(bytes: Uint8Array): Promise<{ shape: Shape; lines: string[]; boxes: string[] }> {
	const read = await translateOdt(bytesSource(bytes))

	if (!read.ok) throw new Error(read.message)

	const document = parseXml(read.translation.parts.get("word/document.xml") as Uint8Array)
	const shape: Shape = { paragraphs: 0, tables: 0, rows: 0, listed: 0, headings: 0 }
	const lines: string[] = []
	const boxes: string[] = []

	count(document, document.documentElement(), shape, lines, boxes)

	return { shape, lines, boxes }
}

function delta(a: Shape, b: Shape): Shape {
	return {
		paragraphs: b.paragraphs - a.paragraphs,
		tables: b.tables - a.tables,
		rows: b.rows - a.rows,
		listed: b.listed - a.listed,
		headings: b.headings - a.headings
	}
}

describe.skipIf(!ENABLED)("ODT saves against LibreOffice 7.5", () => {
	it("change as LibreOffice reads them as they changed for us", { timeout: 1_800_000 }, async () => {
		const oracle = ORACLE as string
		const saved = SAVED as string
		const report: { name: string; opened: boolean; text: boolean; structure: boolean; ours?: unknown; theirs?: unknown }[] = []

		for (const name of readdirSync(saved)
			.filter(file => file.endsWith(".odt"))
			.sort()) {
			const base = name.slice(0, -4)
			const theirSaved = join(oracle, "saved", `${base}.txt`)

			if (!existsSync(theirSaved) || !existsSync(join(oracle, "saved", `${base}.odt`))) {
				report.push({ name, opened: false, text: false, structure: false })

				continue
			}

			const source = await read(new Uint8Array(readFileSync(join(CORPUS_DIR, "odt", name))))
			const written = await read(new Uint8Array(readFileSync(join(saved, name))))
			const ours = change(words(source.lines), words(written.lines))
			const theirs = change(
				words(readFileSync(join(oracle, "source", `${base}.txt`), "utf8").split(/\r?\n/u)),
				words(readFileSync(theirSaved, "utf8").split(/\r?\n/u))
			)
			const theirSource = await read(new Uint8Array(readFileSync(join(oracle, "source", `${base}.odt`))))
			const theirWritten = await read(new Uint8Array(readFileSync(join(oracle, "saved", `${base}.odt`))))
			const ourShape = delta(source.shape, written.shape)
			const theirShape = delta(theirSource.shape, theirWritten.shape)
			const text = JSON.stringify(ours) === JSON.stringify(theirs) || ACCEPTED[name] !== undefined
			// Text boxes as LibreOffice saves them again, read by us: the same text as our save's (LibreOffice saves
			// some shapes as empty frames).
			const boxes = (read: { boxes: string[] }): string => JSON.stringify(read.boxes.filter(box => box !== ""))
			const structure = JSON.stringify(ourShape) === JSON.stringify(theirShape) && boxes(written) === boxes(theirWritten)

			report.push({
				name,
				opened: true,
				text,
				structure,
				...(text && structure ? {} : { ours: { ...ours, shape: ourShape }, theirs: { ...theirs, shape: theirShape } })
			})
		}

		writeFileSync(join(oracle, "report.json"), JSON.stringify(report, null, "\t"))

		expect(report.filter(row => !row.opened || !row.text || !row.structure).map(row => row.name)).toEqual([])
	})
})
