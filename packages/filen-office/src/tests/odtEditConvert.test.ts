import { describe, expect, it } from "vitest"
import { strFromU8, unzipSync } from "fflate"
import { odtBytes } from "./odtFixtures"
import { editOdt, odtEditsOf } from "./odtWriteFixtures"
import { workerHarness } from "./docxEncryptedFixtures"

// An OpenDocument text document open in the editor, converted to .docx as it is edited (Download as): the docx writer's
// parts applied to the translation the worker holds and converted there (convertOdtEditToDocx), never written as ODT
// and read again, so an edit the ODT writer refuses still converts. Its signal ends the job.

const SOURCE = odtBytes({ body: "<text:p>First paragraph</text:p><text:p>Second paragraph</text:p>" })

async function edited(text: string) {
	const editing = await editOdt(SOURCE)

	editing.session.dispatch(editing.session.state.tr.insertText(text, 1))

	const edits = await odtEditsOf(editing)

	return { parts: [...edits.parts], removed: [...(edits.removed ?? [])] }
}

describe("an edited OpenDocument text document converted to .docx in the docx worker", () => {
	it("converts the edited translation, edits included and no internal markers", async () => {
		const harness = workerHarness()

		try {
			const opened = await harness.pool.odt.open(SOURCE)

			if (opened.kind !== "opened") throw new Error(opened.kind)

			const result = await opened.document.convertEdited(await edited("Unsaved "), {})

			if (result.kind !== "converted") throw new Error(result.kind)

			const main = strFromU8(unzipSync(result.bytes)["word/document.xml"] ?? new Uint8Array())

			expect(main).toContain("Unsaved First paragraph")
			expect(main).toContain("Second paragraph")
			expect(main).not.toMatch(/w:rsid/)
			opened.document.release()
		} finally {
			harness.dispose()
		}
	})

	it("ends the job when the signal aborts, answering aborted", async () => {
		const harness = workerHarness()

		try {
			const opened = await harness.pool.odt.open(SOURCE)

			if (opened.kind !== "opened") throw new Error(opened.kind)

			const edits = await edited("Unsaved ")
			const before = new AbortController()

			before.abort()
			expect(await opened.document.convertEdited(edits, { signal: before.signal })).toEqual({ kind: "aborted" })

			const during = new AbortController()
			const running = opened.document.convertEdited(edits, { signal: during.signal })

			during.abort()
			expect(await running).toEqual({ kind: "aborted" })

			// The document is still open for the next request.
			expect((await opened.document.convertEdited(edits, {})).kind).toBe("converted")
			opened.document.release()
		} finally {
			harness.dispose()
		}
	})
})
