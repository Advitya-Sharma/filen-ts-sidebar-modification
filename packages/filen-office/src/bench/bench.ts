import { unzipSync, zipSync, type Zippable } from "fflate"
import { OpcPackage, ZipArchive, bytesSource, memorySink, planZip, writeZip } from "@filen/office/core/opc"
import { Namespaces, NodeKind, XmlEditor, assertWellFormed, parseXml, type XmlDocument } from "@filen/office/core/xml"
import { wordLikeDocument } from "../tests/wordLike"

// Tokenizer and raw-copy save measurements on generated inputs: a Word-like document.xml of ~4.4 MB (and
// four times that), an ~18.6 MB package holding it plus 50 incompressible images, a 3 MB worksheet and a
// deep chain of nested namespace declarations. Not part of `pnpm test`; run with `pnpm bench` (node
// --expose-gc, so retained memory can be measured).

const gc = (globalThis as { gc?: () => void }).gc ?? (() => undefined)
const PARAGRAPHS = 4100
const RUNS = 25

function retained(): number {
	gc()
	gc()

	const usage = process.memoryUsage()

	return usage.heapUsed + usage.arrayBuffers
}

async function measure(name: string, run: () => unknown): Promise<void> {
	const times: number[] = []

	for (let k = 0; k < RUNS; k++) {
		const started = performance.now()

		await run()
		times.push(performance.now() - started)
	}

	times.sort((a, b) => a - b)
	console.log(
		`${name.padEnd(64)} min ${(times[0] ?? 0).toFixed(2).padStart(7)} ms   median ${(times[RUNS >> 1] ?? 0).toFixed(2).padStart(7)} ms`
	)
}

function middleParagraph(document: XmlDocument): number {
	const id = document.nameIdOf("w:p")

	for (let node = 0, seen = 0; node < document.count; node++) {
		if (document.kinds[node] === NodeKind.Element && document.nameIds[node] === id && ++seen === PARAGRAPHS / 2) return node
	}

	return -1
}

// Pseudo-random bytes, which deflate cannot shrink, like real PNG/JPEG data.
let seed = 0x2545f491

function image(size: number): Uint8Array {
	const out = new Uint8Array(size)

	for (let i = 0; i < size; i++) {
		seed ^= seed << 13
		seed ^= seed >>> 17
		seed ^= seed << 5
		out[i] = seed & 0xff
	}

	return out
}

const documentXml = wordLikeDocument(PARAGRAPHS)
const largeXml = wordLikeDocument(PARAGRAPHS * 4)
const before = retained()
const parsed = parseXml(documentXml)
const kept = retained() - before
const arrayBytes = [
	parsed.kinds,
	parsed.starts,
	parsed.openEnds,
	parsed.ends,
	parsed.parents,
	parsed.firstChildren,
	parsed.nextSiblings,
	parsed.nameIds
].reduce((sum, array) => sum + array.byteLength, 0)
const middle = middleParagraph(parsed)
const styles = new TextEncoder().encode(`<w:styles xmlns:w="urn:w">${'<w:style w:type="paragraph"/>'.repeat(200)}</w:styles>`)
const files: Zippable = {
	"[Content_Types].xml": new TextEncoder().encode(
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/></Types>'
	),
	"word/document.xml": documentXml,
	"word/styles.xml": styles
}

for (let k = 1; k <= 50; k++) files[`word/media/image${k}.png`] = image(370 * 1024)

const docx = zipSync(files, { level: 6 })
const edited = (() => {
	const editor = new XmlEditor(parsed)

	editor.replace(middle, "<w:p><w:r><w:t>edited</w:t></w:r></w:p>")

	return editor.serialize()
})()

console.log(
	`document.xml ${(documentXml.length / 1048576).toFixed(1)} MB: ${parsed.count} nodes, node arrays ${(arrayBytes / parsed.count).toFixed(1)} B/node, ` +
		`retained after parse ${(kept / 1048576).toFixed(1)} MB; large part ${(largeXml.length / 1048576).toFixed(1)} MB; package ${(docx.length / 1048576).toFixed(1)} MB\n`
)

const sheetRows: string[] = []

for (let row = 1; row <= 16_000; row++) {
	let cells = ""

	for (let column = 0; column < 5; column++) cells += `<c r="${"ABCDE"[column]}${row}" t="s"><v>${row * 5 + column}</v></c>`

	sheetRows.push(`<row r="${row}">${cells}</row>`)
}

const sheetXml = new TextEncoder().encode(
	`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${sheetRows.join("")}</sheetData></worksheet>`
)
const deepXml = new TextEncoder().encode(`${'<a xmlns:p="u">'.repeat(100_000)}<p:b/>${"</a>".repeat(100_000)}`)
const parsedSheet = parseXml(sheetXml)
const parsedDeep = parseXml(deepXml)
const inflate = (input: AsyncIterable<Uint8Array>): AsyncIterable<Uint8Array> =>
	ReadableStream.from(input).pipeThrough(new DecompressionStream("deflate-raw"))

await measure("parse 4.4 MB document.xml", () => parseXml(documentXml))
await measure("parse 17.6 MB document.xml", () => parseXml(largeXml))
await measure("namespaces of 4.4 MB document.xml", () => new Namespaces(parsed))
await measure("assertWellFormed 4.4 MB document.xml", () => assertWellFormed(parsed))
await measure("3.1 MB worksheet: parse", () => parseXml(sheetXml))
await measure("3.1 MB worksheet: namespaces", () => new Namespaces(parsedSheet))
await measure("3.1 MB worksheet: assertWellFormed", () => assertWellFormed(parsedSheet))
await measure("100k-deep nested declarations: parse", () => parseXml(deepXml))
await measure("100k-deep nested declarations: namespaces + every element URI", () => {
	const namespaces = new Namespaces(parsedDeep)

	for (let node = 1; node < parsedDeep.count; node++) namespaces.elementUri(node)
})
await measure("serialise with no edits", () => new XmlEditor(parsed).serialize())
await measure("replace one paragraph and serialise", () => {
	const editor = new XmlEditor(parsed)

	editor.replace(middle, "<w:p><w:r><w:t>edited</w:t></w:r></w:p>")

	return editor.serialize()
})
await measure("open package + read document.xml (lazy, CRC-checked)", async () => {
	const opc = await OpcPackage.open(bytesSource(docx))

	await opc.read("word/document.xml")
})
await measure("same, inflated by a native DecompressionStream", async () => {
	const opc = await OpcPackage.open(bytesSource(docx), { inflate })

	await opc.read("word/document.xml")
})
await measure("save, 6 KB part replaced, rest raw-copied", async () => {
	const archive = await ZipArchive.open(bytesSource(docx))
	const sink = memorySink()

	await writeZip(planZip(archive, { replace: [{ name: "word/styles.xml", bytes: styles }] }), archive.source, sink)
	sink.bytes()
})
await measure("save, 4.4 MB document.xml replaced (deflated), rest raw-copied", async () => {
	const archive = await ZipArchive.open(bytesSource(docx))
	const sink = memorySink()

	await writeZip(planZip(archive, { replace: [{ name: "word/document.xml", bytes: edited }] }), archive.source, sink)
	sink.bytes()
})
await measure("same, deflated by a native CompressionStream", async () => {
	const archive = await ZipArchive.open(bytesSource(docx))
	const sink = memorySink()

	await writeZip(planZip(archive, { replace: [{ name: "word/document.xml", bytes: edited }] }), archive.source, sink, {
		deflate: async data => {
			const stream = new Blob([data as Uint8Array<ArrayBuffer>]).stream().pipeThrough(new CompressionStream("deflate-raw"))

			return new Uint8Array(await new Response(stream).arrayBuffer())
		}
	})
	sink.bytes()
})
await measure("baseline: fflate unzipSync + zipSync (re-deflates everything)", () => {
	return zipSync({ ...unzipSync(docx), "word/document.xml": edited }, { level: 6 })
})
