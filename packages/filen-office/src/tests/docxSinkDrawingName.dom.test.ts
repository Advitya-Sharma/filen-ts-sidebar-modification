// @vitest-environment happy-dom
import { NodeSelection } from "prosemirror-state"
import { afterEach, beforeEach, describe, expect, it } from "vitest"
import { EditSink } from "@filen/office/react/docx/editor"
import { buildDocx, p } from "./docxSynthetic"
import { pagedEditor } from "./pagedDomFixtures"

// The paged editor's input view is what assistive technology reads (the pages are hidden from it): a picture
// there is an image named by its alternative text, else its title (ECMA-376 Part 1 §20.4.2.5).

const PICTURE = "http://schemas.openxmlformats.org/drawingml/2006/picture"

function picture(docPr: string): string {
	return (
		`<w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="476250"/><wp:docPr id="1" name="d"${docPr}/><a:graphic><a:graphicData uri="${PICTURE}">` +
		'<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="p"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
		'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
	)
}

describe("EditSink drawings", () => {
	let host: HTMLDivElement

	beforeEach(() => {
		host = document.createElement("div")
		document.body.appendChild(host)
	})

	afterEach(() => host.remove())

	it("are images named by their alt text, else their title, and follow an alt text edit", async () => {
		const { editor } = await pagedEditor(
			buildDocx({
				body: p("Before") + `<w:p>${picture(' descr="A cat"')}${picture(' title="A title"')}${picture("")}</w:p>`,
				relationships: [{ id: "rIdImg", type: "image", target: "media/image1.png" }],
				parts: [{ name: "word/media/image1.png", data: new Uint8Array([137, 80, 78, 71, 1, 2, 3]) }]
			})
		)
		const sink = new EditSink(editor, host, { label: "Document", caretClientRect: () => undefined })
		const names = () => [...sink.element.querySelectorAll("[role='img']")].map(element => element.getAttribute("aria-label"))

		expect(names()).toEqual(["A cat", "A title"])

		let at = -1

		editor.state.doc.descendants((node, pos) => {
			if (at < 0 && node.type.name === "drawing") at = pos
		})

		editor.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, at)))
		expect(await editor.execute({ type: "drawing.altText", text: "A dog" })).toBe(true)
		await new Promise(resolve => setTimeout(resolve, 0))
		expect(names()).toEqual(["A dog", "A title"])
		sink.destroy()
		editor.dispose()
	})
})
