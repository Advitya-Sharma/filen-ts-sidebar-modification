import { existsSync, readFileSync, readdirSync, writeFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { bytesSource } from "@filen/office/core/opc"
import { decodeUtf8, parseXml } from "@filen/office/core/xml"
import { translateOdt } from "../core/odt/index"
import { OdtWriteRefusal, proveOdt, writeOdt, type OdtEdits } from "../core/odt/write/writer"
import { CORPUS_DIR } from "./corpusDir"
import { ODT_EDITS, editOdt, odtEditsOf, saveOdt, scriptedOdtEdits, type OdtEditName } from "./odtWriteFixtures"

// The ODT writer over the corpus's odt/ files: an untouched save gives the file byte for byte, Gate 1 proves (or
// locks) every block, and the scripted edits save to a file that reads back as edited (the writer's own proof) or
// are refused with a reason. FILEN_OFFICE_ODT_WRITE_OUT names a directory for the report and the written files
// (for the LibreOffice and validator oracles); FILEN_OFFICE_ODT_WRITE_ONLY a file name filter. Opt-in where the
// corpus exists.

const DIRECTORY = join(CORPUS_DIR, "odt")
const OUT = process.env["FILEN_OFFICE_ODT_WRITE_OUT"] || undefined
const ONLY = process.env["FILEN_OFFICE_ODT_WRITE_ONLY"] || undefined

interface Row {
	name: string
	open: boolean
	untouched?: boolean
	blocks?: number
	locked?: number
	lockReasons?: Record<string, number>
	proveMs?: number
	edits?: (OdtEditName | "header" | "textbox")[]
	saved?: boolean
	refusal?: string
	refusedEdits?: Record<string, string>
	rounds?: number
	written?: number
	saveMs?: number
}

// A header edited as the docx writer writes one: its first text prefixed (the editor edits headers in sessions of
// their own; the part it writes is this).
function headerEdited(parts: ReadonlyMap<string, unknown>, edits: OdtEdits): OdtEdits {
	const name = [...parts.keys()].sort().find(part => /^word\/(header|footer)\d+\.xml$/u.test(part))
	const part = name === undefined ? undefined : parts.get(name)

	if (name === undefined || !(part instanceof Uint8Array)) return edits

	const xml = decodeUtf8(part)
	const at = firstText(xml)

	if (at === null) return edits

	const end = at.index + at[0].length

	return { ...edits, parts: new Map([...edits.parts, [name, new TextEncoder().encode(`${xml.slice(0, end)}Filen ${xml.slice(end)}`)]]) }
}

// A text box typed into as the docx writer writes one: the first text of the first one in a block the proof left
// editable, prefixed.
function textBoxEdited(parts: ReadonlyMap<string, unknown>, edits: OdtEdits, locked: ReadonlySet<number>): OdtEdits {
	const source = parts.get("word/document.xml")
	const part = edits.parts.get("word/document.xml") ?? source

	if (!(source instanceof Uint8Array) || !(part instanceof Uint8Array)) return edits

	const document = parseXml(source)
	const body = document.firstChildElement(document.documentElement())
	const xml = decodeUtf8(part)

	for (let block = document.firstChildElement(body); block >= 0; block = document.nextSiblingElement(block)) {
		if (locked.has(document.start(block))) continue

		const raw = decodeUtf8(document.raw(block))
		const box = raw.indexOf("<w:txbxContent>")

		if (box < 0) continue

		const inner = box + "<w:txbxContent>".length
		const at = firstText(raw.slice(inner, raw.indexOf("</w:txbxContent>", inner)))
		const where = xml.indexOf(raw)

		if (at === null || where < 0) continue

		const end = where + inner + at.index + at[0].length

		return {
			...edits,
			parts: new Map([...edits.parts, ["word/document.xml", new TextEncoder().encode(`${xml.slice(0, end)}Filen ${xml.slice(end)}`)]])
		}
	}

	return edits
}

// Every paragraph's text in order, entities decoded.
function paragraphText(part: Uint8Array): string {
	return [...decodeUtf8(part).matchAll(/<w:p(?:\s[^>]*)?\/>|<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/gu)]
		.map(paragraph => [...paragraph[0].matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/gu)].map(match => match[1]).join(""))
		.join("\n")
		.replace(
			/&(lt|gt|amp|quot|apos);/gu,
			(_entity, entity: string) => ({ lt: "<", gt: ">", amp: "&", quot: '"', apos: "'" })[entity] ?? ""
		)
}

describe.skipIf(!existsSync(DIRECTORY))("ODT writer over the corpus", () => {
	it("saves untouched files byte for byte and scripted edits as edited", { timeout: 3_600_000 }, async () => {
		const rows: Row[] = []

		for (const name of readdirSync(DIRECTORY)
			.filter(file => file.toLowerCase().endsWith(".odt") && (ONLY === undefined || file.includes(ONLY)))
			.sort()) {
			const bytes = new Uint8Array(readFileSync(join(DIRECTORY, name)))
			let editing: Awaited<ReturnType<typeof editOdt>>

			try {
				editing = await editOdt(bytes, { prove: false })
			} catch {
				rows.push({ name, open: false })

				continue
			}

			let started = performance.now()
			const proof = await proveOdt(editing.source)
			const proveMs = performance.now() - started
			const untouched = await writeOdt(editing.source, await odtEditsOf(editing))
			const reasons: Record<string, number> = {}

			for (const reason of proof.reasons) reasons[reason] = (reasons[reason] ?? 0) + 1
			const row: Row = {
				name,
				open: true,
				untouched: untouched.bytes.length === bytes.length && untouched.bytes.every((byte, k) => byte === bytes[k]),
				blocks: editing.source.translation.origins
					.filter(origin => origin.part === "word/document.xml")
					.reduce((sum, origin) => sum + origin.count, 0),
				locked: proof.locked.length,
				lockReasons: reasons,
				proveMs: Math.round(proveMs)
			}
			const gated = await editOdt(bytes)

			row.edits = await scriptedOdtEdits(gated)
			started = performance.now()

			try {
				const headed = headerEdited(gated.source.translation.parts, await odtEditsOf(gated))
				const edits = textBoxEdited(
					gated.source.translation.parts,
					headed,
					new Set(proof.locked.filter(block => block.part === "word/document.xml").map(block => block.offset))
				)

				if (edits.parts.size > 0 && [...edits.parts.keys()].some(part => /^word\/(header|footer)\d+\.xml$/u.test(part)))
					row.edits.push("header")
				if (edits !== headed) row.edits.push("textbox")

				const written = await writeOdt(gated.source, edits)
				const reopened = await translateOdt(bytesSource(written.bytes))

				// Our own reading of the save, independent of the writer's proof: every paragraph's text in order.
				expect(reopened.ok, name).toBe(true)

				if (reopened.ok) {
					const expected =
						edits.parts.get("word/document.xml") ?? (gated.source.translation.parts.get("word/document.xml") as Uint8Array)

					expect(paragraphText(reopened.translation.parts.get("word/document.xml") as Uint8Array), name).toBe(
						paragraphText(expected)
					)
				}

				row.saved = true
				row.rounds = written.rounds
				row.written = written.written
				if (OUT !== undefined) writeFileSync(join(OUT, name), written.bytes)
			} catch (error) {
				if (!(error instanceof OdtWriteRefusal)) throw error

				row.saved = false
				row.refusal = `${error.reason}: ${error.message.slice(0, 600)}`
				row.refusedEdits = {}

				// Each edit on its own, to say which one the writer refuses.
				for (const edit of ODT_EDITS) {
					if (!row.edits.includes(edit)) continue

					const single = await editOdt(bytes)

					await scriptedOdtEdits(single, [edit])

					try {
						await saveOdt(single)
					} catch (inner) {
						if (!(inner instanceof OdtWriteRefusal)) throw inner

						row.refusedEdits[edit] = `${inner.reason}: ${inner.message.slice(0, 300)}`
					}
				}
			}

			row.saveMs = Math.round(performance.now() - started)
			rows.push(row)
		}

		if (OUT !== undefined) writeFileSync(join(OUT, "report.json"), JSON.stringify(rows, null, "\t"))

		const opened = rows.filter(row => row.open)

		expect(opened.filter(row => row.untouched !== true).map(row => row.name)).toEqual([])
	})
})

// The first text outside drawings and fields (a text box's is the frame's, a field's its result: both copied whole).
function firstText(xml: string): { index: number; 0: string } | null {
	const pattern = /<w:drawing\b[\s\S]*?<\/w:drawing>|<w:fldSimple\b[\s\S]*?<\/w:fldSimple>|<w:t(?: [^>]*)?>/gu

	for (let match = pattern.exec(xml); match !== null; match = pattern.exec(xml)) if (match[0].startsWith("<w:t")) return match

	return null
}
