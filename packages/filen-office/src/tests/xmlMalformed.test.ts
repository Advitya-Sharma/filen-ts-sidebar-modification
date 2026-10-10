import { describe, expect, it } from "vitest"
import { NodeKind, parseXml, type XmlDocument } from "@filen/office/core/xml"

function parse(text: string): XmlDocument {
	return parseXml(new TextEncoder().encode(text))
}

function elements(document: XmlDocument): number[] {
	const out: number[] = []

	for (let node = 0; node < document.count; node++) if (document.kinds[node] === NodeKind.Element) out.push(node)

	return out
}

describe("start tags with malformed attributes", () => {
	it("records them while parsing and keeps the tree", () => {
		const document = parse(
			'<r><ok a="1" b=\'2\'/><ok2 a = "x>y/z" /><dup a="1" a="2"/><unquoted a=b/><bare a/><joined a="1"b="2"/><ok3/><slash a="1"/ ></slash></r>'
		)
		const names = elements(document)
			.filter(node => document.isMalformed(node))
			.map(node => document.name(node))

		expect(names).toEqual(["dup", "unquoted", "bare", "joined", "slash"])
		expect(Array.from(document.malformed)).toEqual(elements(document).filter(node => document.isMalformed(node)))
		expect(document.isSelfClosing(elements(document)[2] ?? -1)).toBe(true)
		expect(document.attribute(elements(document)[2] ?? -1, "a")).toBe("x>y/z")
	})

	it("agrees with the validating attribute reader", () => {
		const tags = [
			"",
			' a="1"',
			' a="1" b="2"',
			' a="1"  b = "2" ',
			' a="1" a="1"',
			" a=1",
			" a",
			' a="1"b="2"',
			' x:a="1" y:a="1"',
			' a="&amp;"',
			" a='\"'",
			' a="1" bb="2" a="3"'
		]

		for (const tag of tags) {
			for (const end of [">", "/>", " />"]) {
				const document = parse(`<r><e${tag}${end === ">" ? "></e>" : end}</r>`)
				const element = elements(document)[1] ?? -1
				let throws = false

				try {
					document.attributeSpans(element)
				} catch {
					throws = true
				}

				expect([tag, end, document.isMalformed(element)]).toEqual([tag, end, throws])
			}
		}
	})

	it("still refuses tags that never end", () => {
		expect(() => parse('<r><e a="1></r>')).toThrow()
		expect(() => parse("<r><e a=<b/></r>")).toThrow()
	})
})
