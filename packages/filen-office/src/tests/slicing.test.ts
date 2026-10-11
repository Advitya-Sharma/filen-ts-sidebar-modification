import { existsSync, readFileSync, readdirSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { buildDocxModel, openDocxPackage, type DocxModel, type DocxPackage } from "@filen/office/core/docx"
import { SliceClock, ZipArchive, ZipError, bytesSource, type Slicing } from "@filen/office/core/opc"
import { Namespaces, XmlError, parseXml, parseXmlSliced, type XmlDocument } from "@filen/office/core/xml"
import { DocxEngine, type DocxEvent } from "@filen/office/dom/docx"
import { CORPUS_DIR } from "./corpusDir"
import { buildZip, sameBytes } from "./zipBuilder"

const FIXTURES = join(import.meta.dirname, "fixtures")
const DOCX_FIXTURES = [
	...readdirSync(FIXTURES)
		.filter(name => name.endsWith(".docx"))
		.map(name => join(FIXTURES, name)),
	...readdirSync(join(FIXTURES, "docx-corpus"))
		.filter(name => /\.(docx|docm|dotx|dotm)$/.test(name))
		.map(name => join(FIXTURES, "docx-corpus", name))
].filter(file => readFileSync(file).readUInt32BE(0) !== 0xd0cf11e0)
const encoder = new TextEncoder()

// Paths at which two values differ; typed arrays compared as bytes (vitest's deep equality takes seconds on
// megabyte arrays).
function differences(a: unknown, b: unknown, path: string = ""): string[] {
	if (ArrayBuffer.isView(a) && ArrayBuffer.isView(b)) {
		const same =
			a.constructor === b.constructor &&
			Buffer.from(a.buffer, a.byteOffset, a.byteLength).equals(Buffer.from(b.buffer, b.byteOffset, b.byteLength))

		return same ? [] : [path]
	}

	if (a instanceof Map && b instanceof Map) return differences([...a], [...b], path)

	if (Array.isArray(a) && Array.isArray(b)) {
		if (a.length !== b.length) return [`${path}.length`]

		return a.flatMap((value: unknown, k) => differences(value, b[k], `${path}[${k}]`))
	}

	if (typeof a === "object" && typeof b === "object" && a !== null && b !== null && !Array.isArray(a) && !Array.isArray(b)) {
		const keys = new Set([...Object.keys(a), ...Object.keys(b)])

		return [...keys].flatMap(key => differences(Reflect.get(a, key), Reflect.get(b, key), `${path}.${key}`))
	}

	return Object.is(a, b) ? [] : [path]
}

// A host turn at every look at the clock: the finest slicing there is.
function everyCheck(): Slicing & { pauses: number } {
	const slicing = {
		pauses: 0,
		pause: (): Promise<void> => {
			slicing.pauses++

			return Promise.resolve()
		},
		sliceMs: 0,
		now: (): number => 0
	}

	return slicing
}

function arrays(document: XmlDocument): unknown {
	return {
		bytes: document.bytes,
		count: document.count,
		kinds: document.kinds,
		starts: document.starts,
		openEnds: document.openEnds,
		ends: document.ends,
		parents: document.parents,
		firstChildren: document.firstChildren,
		nextSiblings: document.nextSiblings,
		nameIds: document.nameIds,
		names: document.names,
		encoding: document.encoding,
		malformed: document.malformed
	}
}

// What parseXml throws, or its node arrays.
async function outcome(parse: () => XmlDocument | Promise<XmlDocument>): Promise<unknown> {
	try {
		return arrays(await parse())
	} catch (error) {
		if (!(error instanceof XmlError)) throw error

		return { code: error.code, offset: error.offset, message: error.message }
	}
}

async function xmlParts(file: string): Promise<{ name: string; bytes: Uint8Array }[]> {
	const zip = await ZipArchive.open(bytesSource(new Uint8Array(readFileSync(file))))
	const out: { name: string; bytes: Uint8Array }[] = []

	for (const entry of zip.entries) {
		if (/\.(xml|rels)$/i.test(entry.name)) out.push({ name: entry.name, bytes: await zip.read(entry) })
	}

	return out
}

async function open(file: string): Promise<DocxPackage> {
	const opened = await openDocxPackage(bytesSource(new Uint8Array(readFileSync(file))))

	if (!opened.ok) throw new Error(`${file}: ${opened.reason}`)

	return opened.package
}

// Everything a model holds but its parsed documents.
function snapshot(model: DocxModel): unknown {
	return {
		parts: model.parts.map(part => ({
			index: part.index,
			name: part.name,
			count: part.count,
			kinds: part.kinds,
			nodes: part.nodes,
			ends: part.ends,
			parents: part.parents,
			aux: part.aux,
			paragraphs: part.paragraphs,
			runs: part.runs,
			tables: part.tables,
			rows: part.rows,
			cells: part.cells,
			sdts: part.sdts,
			fields: part.fields,
			bookmarks: part.bookmarks,
			revisions: part.revisions,
			alternates: part.alternates,
			relationships: part.relationships.all
		})),
		stories: model.stories,
		sections: model.sections,
		comments: model.comments,
		people: model.people,
		drawings: model.drawings,
		vml: model.vml,
		commentAnchors: model.commentAnchors,
		strings: model.strings,
		maxIds: model.maxIds,
		counts: model.counts,
		background: model.background,
		strict: model.strict,
		missingParts: model.missingParts,
		malformedParts: model.malformedParts,
		rejectedParts: model.rejectedParts
	}
}

async function sameModel(file: string): Promise<{ pauses: number; model: DocxModel }> {
	const sync = await buildDocxModel(await open(file))
	const slicing = everyCheck()
	const sliced = await buildDocxModel(await open(file), { slicing })

	expect(differences(snapshot(sliced), snapshot(sync))).toEqual([])

	return { pauses: slicing.pauses, model: sliced }
}

describe("sliced XML parsing", () => {
	it("gives the same document as parseXml for every part of the fixtures", async () => {
		for (const file of DOCX_FIXTURES) {
			for (const part of await xmlParts(file)) {
				const slicing = everyCheck()

				expect(
					differences(await outcome(() => parseXmlSliced(part.bytes, slicing)), await outcome(() => parseXml(part.bytes)))
				).toEqual([])
			}
		}
	})

	it("gives the same document when the arrays grow and are trimmed in steps", async () => {
		// Four bytes per node, far below the first guess: the node arrays grow several times.
		const dense = encoder.encode(`<r>${"<a/>".repeat(200_000)}x${'<b c="1" c="2">t</b>'.repeat(3000)}</r>`)
		const slicing = everyCheck()

		expect(differences(await outcome(() => parseXmlSliced(dense, slicing)), await outcome(() => parseXml(dense)))).toEqual([])
		// A turn at least every 64 KiB, and between the arrays copied while growing.
		expect(slicing.pauses).toBeGreaterThan(dense.length / 65536 + 8)
	})

	it("fails exactly as parseXml on malformed input", async () => {
		const inputs = [
			"",
			"   ",
			"<r>",
			"<r></s>",
			"text<r/>",
			"<r/><s/>",
			"<r/>tail",
			'<r a="1></r>',
			"<r a='1'",
			"<!DOCTYPE r [<!ENTITY x 'y'>]><r/>",
			"<r><!-- open</r>",
			"<r><![CDATA[x</r>",
			"<r><?pi</r>",
			"<r/><?xml version='1.0'?>",
			"<r><a b<c/></a></r>",
			`<r>${"<a>".repeat(70_000)}${"</a>".repeat(69_999)}</r>`,
			`<r>${"<a/>".repeat(40_000)}</r><extra/>`
		]

		for (const input of inputs) {
			const bytes = encoder.encode(input)

			expect(differences(await outcome(() => parseXmlSliced(bytes, everyCheck())), await outcome(() => parseXml(bytes)))).toEqual([])
		}

		const utf16 = new Uint8Array([0xff, 0xfe, ...[..."<r>é</r>"].flatMap(c => [c.charCodeAt(0), 0])])

		for (const input of [utf16, utf16.subarray(0, 7)]) {
			expect(differences(await outcome(() => parseXmlSliced(input, everyCheck())), await outcome(() => parseXml(input)))).toEqual([])
		}
	})
})

describe("sliced namespace scan", () => {
	it("resolves every element as the one-pass scan does, across scan ranges", async () => {
		// Declarations below the root and every ~200 KB, so the 1 MiB scan ranges cut between them.
		const filler = `<w:p>${'<w:r><w:t xml:space="preserve">text</w:t></w:r>'.repeat(4000)}</w:p>`
		const section = (k: number): string => `<x:s xmlns:x="urn:x${k}" xmlns:w="urn:w${k % 3}">${filler}<x:e/></x:s>`
		const bytes = encoder.encode(
			`<w:document xmlns:w="urn:w" xmlns:x="urn:x"><w:body>${Array.from({ length: 12 }, (_, k) => section(k)).join("")}</w:body></w:document>`
		)
		const document = parseXml(bytes)
		const slicing = everyCheck()
		const sync = new Namespaces(document, { lenient: true })
		const sliced = await Namespaces.sliced(document, { lenient: true }, slicing)

		expect(bytes.length).toBeGreaterThan(2 * 1024 * 1024)
		expect(slicing.pauses).toBeGreaterThanOrEqual(2)
		expect([...sliced.declarations()]).toEqual([...sync.declarations()])

		const differing: number[] = []

		for (let node = 0; node < document.count; node++) if (sliced.elementUri(node) !== sync.elementUri(node)) differing.push(node)

		expect(differing).toEqual([])
		expect(sync.elementUri(document.documentElement())).toBe("urn:w")
	})
})

describe("sliced zip reads", () => {
	it("inflates and checks the CRC in steps, with host turns between them", async () => {
		const data = new Uint8Array(3 * 1024 * 1024)

		for (let k = 0; k < data.length; k++) data[k] = (k * 7919) % 251

		const zip = await ZipArchive.open(
			bytesSource(
				buildZip([
					{ name: "deflated", data, method: 8 },
					{ name: "stored", data, method: 0 },
					{ name: "badCrc", data, method: 8, crc: 1 },
					{ name: "badStoredCrc", data, method: 0, crc: 1 }
				])
			)
		)

		for (const name of ["deflated", "stored"]) {
			const slicing = everyCheck()
			const entry = zip.get(name)

			if (entry === undefined) throw new Error(name)

			expect(sameBytes(await zip.readSliced(entry, slicing), data)).toBe(true)
			expect(slicing.pauses).toBeGreaterThan(1)
		}

		for (const name of ["badCrc", "badStoredCrc"]) {
			const entry = zip.get(name)

			if (entry === undefined) throw new Error(name)

			await expect(zip.readSliced(entry, everyCheck())).rejects.toThrow(ZipError)
			await expect(zip.read(entry)).rejects.toThrow(/CRC mismatch/)
		}
	})

	it("passes on what the host's turn throws instead of calling it damage", async () => {
		const data = new Uint8Array(1024 * 1024).fill(7)
		const zip = await ZipArchive.open(bytesSource(buildZip([{ name: "a", data, method: 8 }])))
		const entry = zip.get("a")
		const closed = new Error("Closed")

		if (entry === undefined) throw new Error("a")

		await expect(zip.readSliced(entry, { pause: () => Promise.reject(closed), sliceMs: 0, now: () => 0 })).rejects.toBe(closed)
	})
})

describe("sliced model build", () => {
	it("builds the same model as buildDocxModel for every fixture", { timeout: 120_000 }, async () => {
		for (const file of DOCX_FIXTURES) await sameModel(file)
	})

	it("gives the host a turn between top-level blocks", async () => {
		const { pauses, model } = await sameModel(join(FIXTURES, "08-large-300p.docx"))
		const body = model.body
		const part = body === undefined ? undefined : model.partOf(body)
		let blocks = 0

		if (body !== undefined && part !== undefined) {
			for (let item = body.item + 1; item < part.end(body.item); item = part.end(item)) blocks++
		}

		expect(blocks).toBeGreaterThan(1000)
		expect(pauses).toBeGreaterThan(blocks)
	})

	it("pauses once the slice is used up, on a clock shared by every step", async () => {
		// Each look at the clock takes a millisecond; a slice ends at the first look 5 ms after the last turn.
		let time = 0
		let resumed = 0
		let longest = 0
		let pauses = 0
		const clock = new SliceClock({
			pause: async () => {
				pauses++
				longest = Math.max(longest, time - resumed)
				await Promise.resolve()
				resumed = time
			},
			sliceMs: 5,
			now: () => time++
		})
		const opened = await openDocxPackage(bytesSource(new Uint8Array(readFileSync(join(FIXTURES, "08-large-300p.docx")))), {
			slicing: clock
		})

		if (!opened.ok) throw new Error(opened.reason)

		await buildDocxModel(opened.package, { slicing: clock })

		expect(pauses).toBeGreaterThan(100)
		// The slice's own five looks, and the one restarting it after the turn.
		expect(longest).toBeLessThanOrEqual(6)
	})
})

const CORPUS_SAMPLES = [
	"docx/govuk__Tariff_of_the_United_Kingdom_v1.34.docx",
	"docx/docx4j__Docx4j_GettingStarted.docx",
	"docx/docx4j__Creating_SmartArt_with_docx4j.docx"
].map(name => join(CORPUS_DIR, name))

describe.skipIf(!CORPUS_SAMPLES.some(file => existsSync(file)))("sliced model build over corpus documents", () => {
	it("builds the same model as buildDocxModel", { timeout: 300_000 }, async () => {
		for (const file of CORPUS_SAMPLES.filter(existsSync)) await sameModel(file)
	})
})

describe("engine open in slices", () => {
	it("yields while opening and stops at the next turn once closed", async () => {
		const events: DocxEvent[] = []
		let turns = 0
		const engine: DocxEngine = new DocxEngine(
			bytesSource(new Uint8Array(readFileSync(join(FIXTURES, "08-large-300p.docx")))),
			{
				fontHost: { addBytes: () => Promise.resolve(true), addLocal: () => Promise.resolve(false), remove: () => undefined },
				replayFonts: false,
				fontAssets: { read: () => Promise.resolve(new Uint8Array(4)) },
				measurer: {
					kerningControl: true,
					measure: (_font, _kerning, texts, count, widths) => {
						for (let i = 0; i < count; i++) widths[i] = (texts[i] ?? "").length * 8
					},
					fontExtents: () => ({ ascent: 12, descent: 4 })
				},
				yield: () => {
					if (++turns === 20) engine.close()

					return Promise.resolve()
				},
				sliceMs: 0
			},
			event => events.push(event)
		)

		await engine.run()

		expect(turns).toBe(20)
		expect(events).toEqual([])
	})
})
