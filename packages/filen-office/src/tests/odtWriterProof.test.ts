import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { describe, expect, it } from "vitest"
import type { OdtSource } from "../core/odt/write/source"
import { OdtWriteRefusal, proveOdt, proveOdtWritten, writeOdt, type OdtEdits, type OdtWritten } from "../core/odt/write/writer"
import { decodeUtf8, encodeUtf8 } from "../core/xml/index"
import { rng, type Rand } from "./fuzz/config"
import { odtBytes, type OdtParts } from "./odtFixtures"
import { odtSource } from "./odtWriteFixtures"
import { cpuMs } from "./timing"

const STYLES =
	'<style:default-style style:family="paragraph"><style:text-properties fo:font-size="12pt" style:font-name="Liberation Serif"/></style:default-style>' +
	'<style:style style:name="Standard" style:family="paragraph" style:class="text"/>'

// Paragraphs each with an automatic style of its own and a span style of its own.
function styled(count: number): Uint8Array {
	let automatic = ""
	let body = ""

	for (let k = 0; k < count; k++) {
		automatic +=
			`<style:style style:name="P${k}" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:margin-left="${k / 10}mm"/></style:style>` +
			`<style:style style:name="T${k}" style:family="text"><style:text-properties fo:color="#${(k * 9973).toString(16).padStart(6, "0").slice(-6)}"/></style:style>`
		body += `<text:p text:style-name="P${k}">Text <text:span text:style-name="T${k}">span ${k}</text:span> end</text:p>`
	}

	return odtBytes({ automatic, styles: STYLES, body })
}

describe("ODT writer proof", () => {
	it("learns each span style once for paragraph styles that give text the same properties", async () => {
		const small = await odtSource(styled(150))
		const large = await odtSource(styled(1200))
		let started = cpuMs()

		expect((await proveOdt(small)).locked).toEqual([])

		const smallMs = cpuMs() - started

		started = cpuMs()
		expect((await proveOdt(large)).locked).toEqual([])

		const largeMs = cpuMs() - started

		// Eight times the paragraphs: about eight times the work, never sixty-four.
		expect(largeMs).toBeLessThan(smallMs * 16)
	})
})

// A document of plain paragraphs and of what a span of blocks cannot be read without: a page break a style asks for
// after its paragraph, a list, numbered headings, bookmarks, a second page style, a section and a table sized to
// the page.
const MIXED: OdtParts = {
	automatic:
		'<style:style style:name="P1" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:text-align="center"/></style:style>' +
		'<style:style style:name="P2" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:break-after="page"/></style:style>' +
		'<style:style style:name="P3" style:family="paragraph" style:parent-style-name="Standard" style:master-page-name="Wide"/>' +
		'<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>' +
		'<style:style style:name="T2" style:family="text"><style:text-properties fo:font-style="italic"/></style:style>' +
		'<style:style style:name="Tab1" style:family="table"><style:table-properties style:rel-width="50%" table:align="left"/></style:style>' +
		'<text:list-style style:name="L1"><text:list-level-style-number text:level="1" style:num-suffix="." style:num-format="1"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" fo:text-indent="-0.25in" fo:margin-left="0.5in"/></style:list-level-properties></text:list-level-style-number></text:list-style>',
	styles:
		'<style:default-style style:family="paragraph"><style:text-properties fo:font-size="12pt" style:font-name="Liberation Serif"/></style:default-style>' +
		'<style:style style:name="Standard" style:family="paragraph" style:class="text"/>' +
		'<style:style style:name="Heading_20_1" style:display-name="Heading 1" style:family="paragraph" style:parent-style-name="Standard" style:default-outline-level="1"><style:text-properties fo:font-size="18pt" fo:font-weight="bold"/></style:style>' +
		'<text:outline-style style:name="Outline"><text:outline-level-style text:level="1" style:num-format="1" style:num-suffix="."/></text:outline-style>',
	stylesAutomatic:
		'<style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2cm" fo:margin-right="2cm"/></style:page-layout>' +
		'<style:page-layout style:name="pm2"><style:page-layout-properties fo:page-width="29.7cm" fo:page-height="21cm" fo:margin-top="1cm" fo:margin-bottom="1cm" fo:margin-left="1cm" fo:margin-right="1cm"/></style:page-layout>',
	masters:
		'<style:master-page style:name="Standard" style:page-layout-name="pm1"/><style:master-page style:name="Wide" style:page-layout-name="pm2"/>',
	body:
		'<text:p text:style-name="P1">First plain paragraph of the document</text:p>' +
		'<text:p>Second plain paragraph with <text:span text:style-name="T1">bold words</text:span> in it</text:p>' +
		'<text:p text:style-name="P2">A paragraph with a page break after it</text:p>' +
		"<text:p>The paragraph the break comes before</text:p>" +
		'<text:list text:style-name="L1"><text:list-item><text:p>One listed item</text:p></text:list-item><text:list-item><text:p>Two listed items</text:p></text:list-item></text:list>' +
		"<text:p>After the list, plain again</text:p>" +
		'<text:h text:style-name="Heading_20_1" text:outline-level="1">A numbered heading</text:h>' +
		'<text:p>Under the heading, <text:span text:style-name="T2">italic</text:span> text</text:p>' +
		'<text:p>A paragraph holding <text:bookmark-start text:name="mark"/>a bookmark<text:bookmark-end text:name="mark"/> here</text:p>' +
		"<text:p>The last paragraph on the first page style</text:p>" +
		'<text:p text:style-name="P3">The first paragraph on the wide page style</text:p>' +
		"<text:p>A plain paragraph on the wide page</text:p>" +
		'<table:table table:name="T" table:style-name="Tab1"><table:table-column table:number-columns-repeated="2"/><table:table-row><table:table-cell office:value-type="string"><text:p>A1 cell</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>B1 cell</text:p></table:table-cell></table:table-row></table:table>' +
		"<text:p>Between the table and the section</text:p>" +
		'<text:section text:name="S1"><text:p>Inside a section</text:p></text:section>' +
		"<text:p>The document's last paragraph</text:p>"
}

// The translation's body as the docx writer writes an edit of it.
function body(source: OdtSource): string {
	return decodeUtf8(source.translation.parts.get("word/document.xml") as Uint8Array)
}

function edited(xml: string): OdtEdits {
	return { parts: new Map([["word/document.xml", encodeUtf8(xml)]]) }
}

// The top-level blocks of the body: their spans in `xml`.
function blocks(xml: string): { start: number; end: number }[] {
	const out: { start: number; end: number }[] = []
	const open = xml.indexOf("<w:body>") + "<w:body>".length
	let depth = 0
	let start = -1

	for (const match of xml.slice(open).matchAll(/<(\/?)w:(p|tbl)\b[^>]*?(\/?)>/gu)) {
		const at = open + (match.index ?? 0)

		if (match[3] === "/") {
			if (depth === 0) out.push({ start: at, end: at + match[0].length })

			continue
		}

		if (match[1] === "/") {
			depth--

			if (depth === 0) out.push({ start, end: at + match[0].length })
		} else {
			if (depth === 0) start = at

			depth++
		}
	}

	return out
}

// `text` typed at the start of block `index`'s first text.
function typed(xml: string, index: number, text: string): string | undefined {
	const block = blocks(xml)[index]
	const found = block === undefined ? null : /<w:t[ >]/u.exec(xml.slice(block.start, block.end))
	const at = block === undefined || found === null ? -1 : block.start + found.index

	if (at < 0) return undefined

	const end = xml.indexOf(">", at) + 1

	return `${xml.slice(0, end)}${text}${xml.slice(end)}`
}

async function saved(source: OdtSource, edits: OdtEdits, spanProof: boolean): Promise<OdtWritten | string> {
	try {
		return await writeOdt(source, edits, { spanProof })
	} catch (error) {
		if (error instanceof OdtWriteRefusal) return error.message

		throw error
	}
}

// Both ways: the same file or the same refusal; a save the span proof proved, one the whole file's proof accepts.
async function bothWays(source: OdtSource, edits: OdtEdits): Promise<OdtWritten | string> {
	const span = await saved(source, edits, true)
	const whole = await saved(source, edits, false)

	if (typeof span === "string" || typeof whole === "string") expect(span).toEqual(whole)
	else {
		expect(span.bytes).toEqual(whole.bytes)

		if (span.proof === "span") expect((await proveOdtWritten(source, edits, span.bytes)).whole).toBe("done")
	}

	return span
}

describe("ODT writer span proof", () => {
	it("proves a typed paragraph from the blocks the save changed alone", async () => {
		const source = await odtSource(odtBytes(MIXED))
		const written = await bothWays(source, edited(typed(body(source), 1, "Typed ") ?? ""))

		expect(typeof written === "string" ? written : written.proof).toBe("span")
	})

	it("proves a typed cell of a table sized to a later page style alone", async () => {
		const source = await odtSource(odtBytes(MIXED))
		const xml = body(source)
		const table = blocks(xml).findIndex(block => xml.startsWith("<w:tbl", block.start))
		const written = await bothWays(source, edited(typed(xml, table, "Typed ") ?? ""))

		expect(typeof written === "string" ? written : written.proof).toBe("span")
	})

	it("leaves to the whole file what a block's state reaches past it or reads from before it", async () => {
		const source = await odtSource(odtBytes(MIXED))
		const xml = body(source)
		const texts = [
			"The paragraph the break comes before",
			"One listed item",
			"A numbered heading",
			"a bookmark",
			"The last paragraph on the first page style",
			"The first paragraph on the wide page style",
			"Inside a section"
		]

		for (const text of texts) {
			const index = blocks(xml).findIndex(block => xml.slice(block.start, block.end).includes(text))
			const written = await bothWays(source, edited(typed(xml, index, "Typed ") ?? ""))

			expect(typeof written === "string" ? "refused" : written.proof, text).not.toBe("span")
		}
	})

	it("agrees with the whole file's proof on random edits", { timeout: 120_000 }, async () => {
		const source = await odtSource(odtBytes(MIXED))
		const xml = body(source)
		const rand = rng(20261023)
		const pick = (list: number): number => Math.floor(rand() * list)
		let spans = 0

		for (let run = 0; run < 60; run++) {
			let next = xml

			for (let step = 0, steps = 1 + pick(3); step < steps; step++) next = randomEdit(next, rand) ?? next

			if (next === xml) continue

			const written = await bothWays(source, edited(next))

			if (typeof written !== "string" && written.proof === "span") spans++
		}

		expect(spans).toBeGreaterThan(5)
	})

	it("never proves a written file the whole file's proof refuses", { timeout: 120_000 }, async () => {
		const source = await odtSource(odtBytes(MIXED))
		const xml = body(source)
		const rand = rng(42)
		let refused = 0

		for (let run = 0; run < 40; run++) {
			const index = Math.floor(rand() * blocks(xml).length)
			const edits = edited(typed(xml, index, "Typed ") ?? xml)
			const written = await saved(source, edits, true)

			if (typeof written === "string") continue

			for (const mutate of MUTATIONS) {
				const mutated = rewritten(written.bytes, mutate, rand)

				if (mutated === undefined) continue

				const verdict = await proveOdtWritten(source, edits, mutated)

				if (verdict.whole !== "done") {
					refused++
					expect(verdict.span, `${mutate.name} of block ${index}`).toBe(false)
				}
			}
		}

		expect(refused).toBeGreaterThan(10)
	})
})

// The written file with content.xml changed by `mutate` (the same length: the span proof reads what it changed).
function rewritten(bytes: Uint8Array, mutate: (content: string, rand: Rand) => string | undefined, rand: Rand): Uint8Array | undefined {
	const entries = unzipSync(bytes)
	const content = strFromU8(entries["content.xml"] ?? new Uint8Array())
	const changed = mutate(content, rand)

	if (changed === undefined || changed === content) return undefined

	return zipSync({ ...entries, "content.xml": strToU8(changed), mimetype: [entries["mimetype"] ?? new Uint8Array(), { level: 0 }] })
}

const MUTATIONS: readonly ((content: string, rand: Rand) => string | undefined)[] = [
	// A letter of the typed text changed.
	function letter(content) {
		const at = content.indexOf("Typed ")

		return at < 0 ? undefined : `${content.slice(0, at)}Tyqed ${content.slice(at + 6)}`
	},
	// A span style swapped for one of the same name's length.
	function spanStyle(content) {
		const at = content.lastIndexOf('text:style-name="T1"')

		return at < 0 ? undefined : `${content.slice(0, at)}text:style-name="T2"${content.slice(at + 20)}`
	},
	// A paragraph style swapped where the typed text is.
	function paragraphStyle(content) {
		const typedAt = content.indexOf("Typed ")
		const at = typedAt < 0 ? -1 : content.lastIndexOf("<text:p", typedAt)

		if (at < 0) return undefined

		const tag = content.slice(at, content.indexOf(">", at) + 1)
		const restyled = tag.includes('"P1"') ? tag.replace('"P1"', '"P2"') : tag.replace("<text:p", '<text:p text:style-name="P3"')

		return `${content.slice(0, at)}${restyled}${content.slice(at + tag.length)}`
	},
	// A page break asked for after a paragraph before the typed one.
	function breakBefore(content, rand) {
		const typedAt = content.indexOf("Typed ")
		const ats = [...content.slice(0, Math.max(0, typedAt)).matchAll(/<text:p(?=[ >])/gu)].map(match => match.index ?? 0)
		const at = ats[Math.floor(rand() * ats.length)]

		if (at === undefined) return undefined

		const tag = content.slice(at, content.indexOf(">", at) + 1)

		return `${content.slice(0, at)}${tag.replace(/ text:style-name="[^"]*"/u, "").replace("<text:p", '<text:p text:style-name="P2"')}${content.slice(at + tag.length)}`
	},
	// A kept paragraph's text changed.
	function keptText(content) {
		const at = content.indexOf("The document's last paragraph")

		return at < 0 ? undefined : `${content.slice(0, at)}The document's lost paragraph${content.slice(at + 29)}`
	}
]

// One edit of the body as an editor makes them: text typed, a paragraph deleted or repeated, a run made bold.
function randomEdit(xml: string, rand: Rand): string | undefined {
	const list = blocks(xml)
	const block = list[Math.floor(rand() * list.length)]

	if (block === undefined) return undefined

	const kind = Math.floor(rand() * 4)
	const raw = xml.slice(block.start, block.end)

	if (kind === 0) return typed(xml, list.indexOf(block), "Typed ")
	if (kind === 1 && list.length > 2 && !raw.includes("<w:sectPr")) return `${xml.slice(0, block.start)}${xml.slice(block.end)}`
	if (kind === 2 && !raw.includes("<w:sectPr"))
		return `${xml.slice(0, block.end)}${raw.replace(/<w:bookmark(Start|End)[^>]*\/>/gu, "")}${xml.slice(block.end)}`

	const run = xml.indexOf("<w:r>", block.start)

	return run < 0 || run > block.end ? undefined : `${xml.slice(0, run)}<w:r><w:rPr><w:b/></w:rPr>${xml.slice(run + "<w:r>".length)}`
}
