import { describe, expect, it } from "vitest"
import { strFromU8, unzipSync } from "fflate"
import { editRtf, editsOf, latin1 } from "./rtfWriteFixtures"
import { workerHarness } from "./docxEncryptedFixtures"

// An RTF document open in the editor, converted to .docx as it is edited (Download as): the docx writer's parts applied
// to the package the worker holds, never written as RTF first, so an edit the RTF writer refuses still converts.

const RTF = latin1("{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Times;}}\\pard\\plain First paragraph\\par\\pard\\plain Second paragraph\\par}")

describe("an edited RTF document converted to .docx in the docx worker", () => {
	it("converts the edited package, edits included, without the RTF writer", async () => {
		const harness = workerHarness()

		try {
			const editing = await editRtf(RTF)

			editing.session.dispatch(editing.session.state.tr.insertText("Unsaved ", 1))

			const edits = await editsOf(editing)
			const opened = await harness.pool.rtf.open(RTF)

			if (opened.kind !== "opened") throw new Error(opened.kind)

			const result = await opened.document.convertEdited({ parts: [...edits.parts], removed: [...(edits.removed ?? [])] }, {})

			if (result.kind !== "converted") throw new Error(result.kind === "refused" ? result.message : result.kind)

			const main = strFromU8(unzipSync(result.bytes)["word/document.xml"] ?? new Uint8Array())

			expect(main).toContain("Unsaved First paragraph")
			expect(main).toContain("Second paragraph")
			expect(main).not.toMatch(/w:rsidR="[0-9A-F]{8}">/)
			opened.document.release()
		} finally {
			harness.dispose()
		}
	})
})
