import { describe, expect, it } from "vitest"
import { sessionOf } from "./docxEditFixtures"
import { p } from "./docxSynthetic"
import { textPositions } from "./fuzz/ops/base"

// The fuzz harness's own choices: an op must find somewhere to act in every editable document.
describe("docx fuzz ops", () => {
	it("types into paragraphs that hold no text (a picture or a break alone) at their start and end", async () => {
		const session = await sessionOf({ body: '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' + p("") })
		const doc = session.state.doc
		const first = doc.child(0)

		expect(first.textContent).toBe("")
		expect(textPositions(doc)).toEqual([1, first.nodeSize - 1, first.nodeSize + 1])
	})
})
