import { describe, expect, it } from "vitest"
import { XmlEditor, XmlError, assertWellFormed, decodeUtf8, parseXml, type XmlDocument } from "@filen/office/core/xml"

const encoder = new TextEncoder()

function parse(source: string): XmlDocument {
	return parseXml(encoder.encode(source))
}

function edit(source: string, apply: (editor: XmlEditor, document: XmlDocument) => void): string {
	const document = parse(source)
	const editor = new XmlEditor(document)

	apply(editor, document)

	const out = editor.serialize()

	assertWellFormed(parseXml(out))

	return decodeUtf8(out)
}

function child(document: XmlDocument, name: string): number {
	for (let node = 0; node < document.count; node++) {
		if (document.name(node) === name) return node
	}

	throw new Error(`no ${name}`)
}

describe("XmlEditor", () => {
	const source = '<?xml version="1.0"?>\n<r xmlns:w="urn:w">  <w:p w:id=\'1\'><w:t>a</w:t></w:p>\n  <w:p w:id="2"><w:pPr/></w:p></r>'

	it("returns the input itself when nothing changed", () => {
		const document = parse(source)
		const editor = new XmlEditor(document)

		expect(editor.changed).toBe(false)
		expect(editor.serialize()).toBe(document.bytes)
	})

	it("replaces, removes and inserts around nodes, leaving every other byte as it was", () => {
		expect(edit(source, (editor, document) => editor.replace(child(document, "w:t"), "<w:t>b</w:t>"))).toBe(
			source.replace("<w:t>a</w:t>", "<w:t>b</w:t>")
		)
		expect(edit(source, (editor, document) => editor.remove(child(document, "w:t")))).toBe(source.replace("<w:t>a</w:t>", ""))
		expect(
			edit(source, (editor, document) => {
				const p = child(document, "w:p")

				editor.insertBefore(p, "<w:x/>")
				editor.insertAfter(p, encoder.encode("<w:y/>"))
			})
		).toBe(source.replace("<w:p w:id='1'><w:t>a</w:t></w:p>", "<w:x/><w:p w:id='1'><w:t>a</w:t></w:p><w:y/>"))
	})

	it("appends and prepends content, expanding self-closing elements", () => {
		expect(
			edit(source, (editor, document) => {
				const pPr = child(document, "w:pPr")

				editor.append(pPr, "<w:b/>")
				editor.append(pPr, "<w:i/>")
				editor.prepend(pPr, "<w:jc/>")
			})
		).toBe(source.replace("<w:pPr/>", "<w:pPr><w:jc/><w:b/><w:i/></w:pPr>"))
		expect(
			edit(source, (editor, document) => {
				const p = child(document, "w:p")

				editor.prepend(p, "<w:pPr/>")
				editor.append(p, "<w:r/>")
			})
		).toBe(source.replace("<w:p w:id='1'><w:t>a</w:t></w:p>", "<w:p w:id='1'><w:pPr/><w:t>a</w:t><w:r/></w:p>"))
	})

	it("expands a self-closing element with whitespace before />", () => {
		expect(edit("<r><a x='1' /></r>", (editor, document) => editor.append(child(document, "a"), "<b/>"))).toBe(
			"<r><a x='1' ><b/></a></r>"
		)
	})

	it("sets text with escaping", () => {
		expect(edit(source, (editor, document) => editor.setText(child(document, "w:t"), "<1 & 2>"))).toBe(
			source.replace(">a<", ">&lt;1 &amp; 2&gt;<")
		)
		expect(edit("<r><t/></r>", (editor, document) => editor.setText(child(document, "t"), "x"))).toBe("<r><t>x</t></r>")
	})

	it("changes, adds and removes attributes in place", () => {
		expect(
			edit(source, (editor, document) => {
				editor.setAttribute(child(document, "w:p"), "w:id", 'say "x"')
			})
		).toBe(source.replace("w:id='1'", "w:id='say &quot;x&quot;'"))
		expect(
			edit(source, (editor, document) => {
				const pPr = child(document, "w:pPr")

				editor.setAttribute(pPr, "w:a", "1")
				editor.setAttribute(pPr, "w:b", "2")
			})
		).toBe(source.replace("<w:pPr/>", '<w:pPr w:a="1" w:b="2"/>'))
		expect(edit(source, (editor, document) => editor.setAttribute(child(document, "w:t"), "x", "&"))).toBe(
			source.replace("<w:t>", '<w:t x="&amp;">')
		)
		expect(edit(source, (editor, document) => editor.removeAttribute(child(document, "w:p"), "w:id"))).toBe(
			source.replace(" w:id='1'", "")
		)
		expect(edit(source, (editor, document) => editor.removeAttribute(child(document, "w:p"), "absent"))).toBe(source)
	})

	it("combines a new attribute with content on a self-closing element", () => {
		expect(
			edit(source, (editor, document) => {
				const pPr = child(document, "w:pPr")

				editor.append(pPr, "<w:b/>")
				editor.setAttribute(pPr, "w:x", "1")
			})
		).toBe(source.replace("<w:pPr/>", '<w:pPr w:x="1"><w:b/></w:pPr>'))
	})

	it("lets the last write to an attribute win", () => {
		expect(
			edit(source, (editor, document) => {
				const p = child(document, "w:p")

				editor.setAttribute(p, "w:id", "x")
				editor.setAttribute(p, "w:id", "y")
			})
		).toBe(source.replace("w:id='1'", "w:id='y'"))
	})

	it("refuses overlapping edits", () => {
		const document = parse(source)
		const editor = new XmlEditor(document)
		const p = child(document, "w:p")

		editor.replace(p, "<w:p/>")
		editor.setAttribute(child(document, "w:t"), "x", "1")

		expect(() => editor.serialize()).toThrow(XmlError)

		const second = new XmlEditor(document)

		second.replace(p, "<w:p/>")
		second.insertAfter(child(document, "w:t"), "<x/>")

		expect(() => second.serialize()).toThrow(XmlError)
	})

	it("refuses element edits on other nodes", () => {
		const document = parse("<r>text</r>")

		expect(() => new XmlEditor(document).append(1 + document.documentElement(), "<x/>")).toThrow(XmlError)
	})

	it("replaces one paragraph of a large part and keeps the rest byte-identical", () => {
		const paragraphs = Array.from({ length: 5000 }, (_, k) => `<w:p w:rsidR="00${k}"><w:r><w:t>Paragraph ${k}</w:t></w:r></w:p>`)
		const big = `<w:document xmlns:w="urn:w"><w:body>${paragraphs.join("")}</w:body></w:document>`
		const document = parse(big)
		const pId = document.nameIdOf("w:p")
		let middle = -1
		let seen = 0

		for (let node = 0; node < document.count; node++) {
			if (document.nameIds[node] === pId && ++seen === 2500) {
				middle = node

				break
			}
		}

		const editor = new XmlEditor(document)

		editor.replace(middle, "<w:p><w:r><w:t>edited</w:t></w:r></w:p>")

		expect(decodeUtf8(editor.serialize())).toBe(big.replace(paragraphs[2499] ?? "", "<w:p><w:r><w:t>edited</w:t></w:r></w:p>"))
	})
})

describe("XmlEditor ordering and repeated edits", () => {
	function run(source: string, apply: (editor: XmlEditor, document: XmlDocument) => void): string {
		const document = parse(source)
		const editor = new XmlEditor(document)

		apply(editor, document)

		return decodeUtf8(editor.serialize())
	}

	it("orders inserts at one byte as a live tree would, whatever the call order", () => {
		const expected = "<p><first/><beforeC/><c/><afterC/><last/></p>"

		expect(
			run("<p><c/></p>", (editor, document) => {
				editor.prepend(1, "<first/>")
				editor.insertBefore(child(document, "c"), "<beforeC/>")
				editor.append(1, "<last/>")
				editor.insertAfter(child(document, "c"), "<afterC/>")
			})
		).toBe(expected)
		expect(
			run("<p><c/></p>", (editor, document) => {
				editor.insertAfter(child(document, "c"), "<afterC/>")
				editor.append(1, "<last/>")
				editor.insertBefore(child(document, "c"), "<beforeC/>")
				editor.prepend(1, "<first/>")
			})
		).toBe(expected)
		expect(
			run("<p><c/></p>", editor => {
				editor.prepend(1, "<two/>")
				editor.prepend(1, "<one/>")
				editor.append(1, "<three/>")
				editor.append(1, "<four/>")
			})
		).toBe("<p><one/><two/><c/><three/><four/></p>")
		expect(
			run("<p/>", editor => {
				editor.append(1, "<b/>")
				editor.prepend(1, "<a/>")
				editor.append(1, "<c/>")
			})
		).toBe("<p><a/><b/><c/></p>")
	})

	it("keeps the last of repeated whole-node edits", () => {
		expect(
			run("<p><t>a</t></p>", (editor, document) => {
				editor.setText(child(document, "t"), "b")
				editor.setText(child(document, "t"), "c")
			})
		).toBe("<p><t>c</t></p>")
		expect(
			run("<p><t/></p>", (editor, document) => {
				editor.setText(child(document, "t"), "b")
				editor.setContent(child(document, "t"), "<x/>")
			})
		).toBe("<p><t><x/></t></p>")
		expect(
			run("<p><t>a</t></p>", (editor, document) => {
				editor.replace(child(document, "t"), "<u/>")
				editor.remove(child(document, "t"))
			})
		).toBe("<p></p>")
	})

	it("drops edits inside content that setContent replaces", () => {
		expect(
			run("<p><t>a</t><u v='1'/></p>", (editor, document) => {
				editor.setText(child(document, "t"), "x")
				editor.setAttribute(child(document, "u"), "v", "2")
				editor.prepend(1, "<s/>")
				editor.setContent(1, "<new/>")
				editor.append(1, "<after/>")
			})
		).toBe("<p><new/><after/></p>")
	})

	it("writes an edited UTF-16 part as UTF-8 and says so in its declaration", () => {
		const source = '<?xml version="1.0" encoding="UTF-16" standalone="yes"?><r>a</r>'
		const units = new Uint8Array(source.length * 2 + 2)

		units.set([0xff, 0xfe])

		for (let k = 0; k < source.length; k++) units[2 + k * 2] = source.charCodeAt(k)

		const document = parseXml(units)
		const editor = new XmlEditor(document)

		editor.setText(child(document, "r"), "é")

		expect(decodeUtf8(editor.serialize())).toBe('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><r>é</r>')
	})
})
