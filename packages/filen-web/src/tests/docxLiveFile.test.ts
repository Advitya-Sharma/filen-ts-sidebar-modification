import { describe, expect, it, vi } from "vitest"
import type { DocxEditorSession } from "@filen/office/core/docx/edit"
import type { DocxEditingHost } from "@filen/office/dom/docx/editing"

const { prepareDocxSave } = vi.hoisted(() => ({ prepareDocxSave: vi.fn() }))

vi.mock("@filen/office/core/docx/write", () => ({ prepareDocxSave }))

import { editedFile } from "@/features/preview/lib/docxLiveFile"

// Download as from the document editor: the document as the editor holds it now, nothing stored or marked saved,
// from stand-ins so typing meanwhile changes nothing written; the run's signal goes with the write. An OpenDocument
// text document to .docx: the docx writer's parts converted where the document is held (never written as ODT first).

function session(name: string, dirty = true) {
	const standIn = { name: `${name}-frozen` }

	return {
		live: { state: { tr: name }, withTransaction: vi.fn(() => standIn), markSaved: vi.fn(), snapshot: () => ({ dirty }) },
		standIn
	}
}

function hostOf(sessions: ReturnType<typeof session>[], gate: readonly string[] = []) {
	const save = vi.fn((_sessions: unknown, _options: unknown) => Promise.resolve({ bytes: Uint8Array.of(1, 2, 3), encrypted: false }))
	const context = { gate }
	const input = { pkg: "pkg", context, sessions: sessions.map(entry => entry.live) as unknown as DocxEditorSession[] }
	const host = { saveInput: () => input, save } as unknown as DocxEditingHost

	return { host, save, input }
}

describe("editedFile", () => {
	it("writes the editor's sessions as stand-ins in the file's own format, with the run's signal, marking nothing saved", async () => {
		const body = session("body")
		const header = session("header")
		const { host, save } = hostOf([body, header])
		const signal = new AbortController().signal

		expect(await editedFile(host, "odt", signal)).toEqual({ kind: "file", bytes: Uint8Array.of(1, 2, 3) })
		expect(save).toHaveBeenCalledWith([body.standIn, header.standIn], { signal })
		expect(body.live.markSaved).not.toHaveBeenCalled()
	})

	it("converts an OpenDocument text document's edits to .docx where it is held, its ODT writer never run", async () => {
		const body = session("body")
		const { host, save } = hostOf([body])
		const prepared = { replace: [], add: [] }
		const conversion = { kind: "converted" as const, bytes: Uint8Array.of(7), lossy: [] }
		const convertEdited = vi.fn(() => Promise.resolve(conversion))
		const signal = new AbortController().signal

		prepareDocxSave.mockResolvedValue(prepared)

		expect(await editedFile(host, "docx", signal, convertEdited)).toEqual({ kind: "converted", conversion })
		expect(prepareDocxSave.mock.calls[0]?.[0]).toMatchObject({ pkg: "pkg", sessions: [body.standIn] })
		expect(convertEdited).toHaveBeenCalledWith(prepared, signal)
		expect(save).not.toHaveBeenCalled()
	})

	it("hands a document converted where it is held on as .docx for an .odt target, its own writer never run", async () => {
		const { host, save } = hostOf([session("body")])
		const lossy = [{ kind: "rtfShape" as const, count: 3 }]
		const conversion = { kind: "converted" as const, bytes: Uint8Array.of(8), lossy }

		prepareDocxSave.mockResolvedValue({ replace: [], add: [] })

		// What the .docx lacks goes on with it, to be joined with the .odt step's report.
		expect(
			await editedFile(
				host,
				"odt",
				new AbortController().signal,
				vi.fn(() => Promise.resolve(conversion))
			)
		).toEqual({
			kind: "file",
			bytes: Uint8Array.of(8),
			format: "docx",
			lossy
		})
		expect(save).not.toHaveBeenCalled()
	})

	it("gives up the parts' write at its next slice once the signal aborts", async () => {
		const { host } = hostOf([session("body")])
		const controller = new AbortController()

		prepareDocxSave.mockImplementation(async (_input: unknown, options: { slicing: { pause: () => Promise<void> } }) => {
			controller.abort()
			await options.slicing.pause()

			return { replace: [], add: [] }
		})

		await expect(editedFile(host, "docx", controller.signal, vi.fn())).rejects.toMatchObject({ name: "AbortError" })
	})

	it("hands the saved file back when nothing is edited, writing nothing", async () => {
		const { host, save } = hostOf([session("body", false), session("header", false)])

		expect(await editedFile(host, "odt", new AbortController().signal)).toEqual({ kind: "saved" })
		expect(save).not.toHaveBeenCalled()
	})

	it("hands the saved file back for a document that cannot hold edits (view-only: signed, Strict, protected)", async () => {
		const { host, save } = hostOf([session("body")], ["readOnly"])
		const convertEdited = vi.fn()

		expect(await editedFile(host, "docx", new AbortController().signal, convertEdited)).toEqual({ kind: "saved" })
		expect(save).not.toHaveBeenCalled()
		expect(convertEdited).not.toHaveBeenCalled()
	})

	it("refuses a document not open for editing", async () => {
		const host = { saveInput: () => undefined, save: vi.fn() } as unknown as DocxEditingHost

		await expect(editedFile(host, "odt", new AbortController().signal)).rejects.toThrow("not open for editing")
	})
})
