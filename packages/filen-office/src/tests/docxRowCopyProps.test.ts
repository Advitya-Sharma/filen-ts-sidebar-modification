import { TextSelection } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { bytesSource, OpcPackage } from "@filen/office/core/opc"
import type { DocxEditorSession } from "@filen/office/core/docx/edit"
import { buildDocx, p } from "./docxSynthetic"
import { editedLayout, pageDigest, savedDigest } from "./pagedFixtures"

// A row a command makes from a source row (and its cells) takes the properties the schema allows, written as
// every reader accepts them: a switch as a bare element when on and left out when off (the Open XML SDK takes only
// "on"/"off" for w:cantSplit, w:tblHeader, w:noWrap, w:tcFitText and w:hideMark, Word writes neither), and no
// child CT_TrPr or CT_TcPr does not have.

const decoder = new TextDecoder()

function at(session: DocxEditorSession, text: string): number {
	let found = -1

	session.state.doc.descendants((node, pos) => {
		if (found < 0 && node.isText && (node.text ?? "").includes(text)) found = pos + (node.text ?? "").indexOf(text)

		return found < 0
	})

	return found
}

async function insertedBelow(trPr: string, tcPr: string): Promise<string[]> {
	const cell = (text: string): string => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${tcPr}</w:tcPr>${p(text)}</w:tc>`
	const body =
		`${p("Before")}<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>` +
		`<w:tr><w:trPr>${trPr}</w:trPr>${cell("a1")}${cell("b1")}</w:tr><w:tr>${cell("a2")}${cell("b2")}</w:tr></w:tbl>${p("After")}`
	const layout = await editedLayout(buildDocx({ body }))

	layout.session.dispatch(layout.session.state.tr.setSelection(TextSelection.create(layout.session.state.doc, at(layout.session, "b1"))))
	expect(await layout.edit(session => session.execute({ type: "table.row.insertBelow" }))).toBe(true)

	const saved = await layout.save()

	expect(pageDigest(layout.layout)).toEqual(await savedDigest(saved))

	const xml = decoder.decode(await (await OpcPackage.open(bytesSource(saved))).read("word/document.xml"))

	return [...xml.matchAll(/<w:tr>.*?<\/w:tr>|<w:tr [^>]*>.*?<\/w:tr>/g)].map(match => match[0])
}

describe("docx writer: a row made from a source row", () => {
	it("writes the source row's switches in a form every reader accepts", async () => {
		const rows = await insertedBelow('<w:cantSplit w:val="false"/><w:tblHeader w:val="true"/><w:hidden w:val="0"/>', "")

		expect(rows).toHaveLength(3)
		expect(rows[0]).toContain('<w:cantSplit w:val="false"/><w:tblHeader w:val="true"/>')
		expect(rows[1]).toContain("<w:trPr><w:tblHeader/></w:trPr>")
	})

	it("writes the source cells' switches in a form every reader accepts", async () => {
		const rows = await insertedBelow("", '<w:noWrap w:val="true"/><w:tcFitText w:val="1"/><w:hideMark w:val="false"/>')

		expect(rows[1]).toContain('<w:tcW w:w="2000" w:type="dxa"/><w:noWrap/><w:tcFitText/></w:tcPr>')
		expect(rows[1]).not.toMatch(/w:val="(?:true|false|1|0|on|off)"/)
	})

	it("writes switches written as on/off as a bare element or none, a bare one as it is", async () => {
		const rows = await insertedBelow('<w:cantSplit/><w:tblHeader w:val="off"/>', '<w:noWrap w:val="on"/>')

		expect(rows[0]).toContain('<w:tblHeader w:val="off"/>')
		expect(rows[1]).toContain("<w:trPr><w:cantSplit/></w:trPr>")
		expect(rows[1]).toContain('<w:tcW w:w="2000" w:type="dxa"/><w:noWrap/></w:tcPr>')
	})

	it("leaves out a child the row properties do not have", async () => {
		const rows = await insertedBelow('<w:cantSplit/><w:pStyle w:val="x"/>', "")

		expect(rows[0]).toContain('<w:pStyle w:val="x"/>')
		expect(rows[1]).toContain("<w:trPr><w:cantSplit/></w:trPr>")
	})
})
