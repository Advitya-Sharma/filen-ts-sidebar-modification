import { describe, expect, it } from "vitest"
import { BodyScanner, OdtScanError, type BodyBatch } from "../core/odt/scan"
import { decodeUtf8, encodeUtf8 } from "../core/xml/index"
import { elementName } from "../core/odf/index"

const ROOT =
	'<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3">'
const BODY = [
	'<t:p t:style-name="a&gt;b">One <t:span>two</t:span></t:p>',
	"<t:h t:outline-level=\"1\" t:x='/>'><!-- <t:p> --><![CDATA[ </t:h> ]]>Head</t:h>",
	'<t:p t:style-name="E"/>',
	"<t:list><t:list-item><t:list><t:list-item><t:p>deep</t:p></t:list-item></t:list></t:list-item></t:list>",
	"<?pi </t:p> ?><t:p>last</t:p>"
]
const PROLOGUE = `<?xml version="1.0" encoding="UTF-8"?>${ROOT}<office:automatic-styles/>`
const STREAM = `${PROLOGUE}<office:body><office:text t:use-soft-page-breaks="true">\n${BODY.join("\n ")}\n</office:text></office:body></office:document-content>`

function run(stream: Uint8Array, step: number, batchBytes?: number): BodyBatch[] {
	const scanner = new BodyScanner(encodeUtf8(ROOT), encodeUtf8(PROLOGUE).length, batchBytes === undefined ? {} : { batchBytes })
	const out: BodyBatch[] = []

	for (let at = 0; at < stream.length; at += step) out.push(...scanner.push(stream.subarray(at, at + step), at))

	out.push(...scanner.finish())

	return out
}

describe("cutting an ODT body into top-level blocks", () => {
	const bytes = encodeUtf8(STREAM)

	for (const step of [1, 7, 4096]) {
		it(`finds each block's byte span whatever the chunking (${step} bytes)`, () => {
			const batches = run(bytes, step)
			const spans = batches.flatMap(batch => batch.blocks).map(block => decodeUtf8(bytes, block.start, block.end))

			expect(spans).toEqual(BODY.map(block => block.replace(/^<\?pi .*?\?>/, "")))
		})
	}

	it("parses each block under the stream's own prefixes", () => {
		const names = run(bytes, 13, 64).flatMap(batch =>
			batch.blocks.map(block => elementName(batch.document, batch.namespaces, block.node))
		)

		expect(names).toEqual(["text:p", "text:h", "text:p", "text:list", "text:p"])
	})

	it("parses blocks in batches no larger than asked", () => {
		expect(run(bytes, 4096, 64).length).toBeGreaterThan(2)
		expect(run(bytes, 4096).length).toBe(1)
	})

	it("refuses a body that never closes", () => {
		expect(() => run(bytes.subarray(0, bytes.length - 60), 50)).toThrow(OdtScanError)
	})

	it("refuses a document type declaration inside the body", () => {
		const hostile = encodeUtf8(`${PROLOGUE}<office:body><office:text><!DOCTYPE x [<!ENTITY a "b">]><t:p/></office:text></office:body>`)

		expect(() => run(hostile, 4096)).toThrow(OdtScanError)
	})
})
