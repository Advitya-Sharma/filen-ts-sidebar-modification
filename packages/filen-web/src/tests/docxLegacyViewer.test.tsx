// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { act, cleanup, fireEvent, render as renderUi, screen, waitFor, within } from "@testing-library/react"
import type { ReactElement } from "react"
import { createTestQueryClient, queryClientWrapper } from "@/tests/testQueryClient"
import type { ByteSource } from "@filen/office/core/opc"
import {
	heldSource,
	type DocxConvertedDocument,
	type DocxLegacyOpen,
	type DocxListener,
	type DocxProtectedOpen,
	type DocxProtectionOptions,
	type DocxSessionApi
} from "@filen/office/dom/docx"
import "@/lib/i18n"
import { registerAllActions } from "@/lib/keymap/actions"
import type { DriveItem } from "@/features/drive/lib/item"
import { linkedFileItem } from "@/tests/fixtures/sdk"

// A Word 97–2003 document in the viewer: a compound file that is no encrypted package is converted in the docx
// worker, shown read-only from the package held there (never from the page's bytes), with the badge listing what
// is shown differently; Save as .docx uploads a copy of the package beside the file and shows it; files the reader
// refuses say why; the package is released once the preview lets go of it.

function render(ui: ReactElement) {
	return renderUi(ui, { wrapper: queryClientWrapper(createTestQueryClient()) })
}

const { host, protectOpen, legacyOpen, uploads, toasts, files, verdicts } = vi.hoisted(() => ({
	files: { bytes: new Uint8Array(0) },
	verdicts: [] as { name: string; verdict: string }[],
	host: { source: null as ByteSource | null },
	protectOpen: vi.fn<(source: ByteSource, options?: DocxProtectionOptions) => Promise<DocxProtectedOpen>>(),
	legacyOpen: vi.fn<(source: ByteSource, options?: DocxProtectionOptions) => Promise<DocxLegacyOpen>>(),
	uploads: [] as { bytes: Uint8Array; extension: string }[],
	toasts: { success: [] as string[], error: [] as string[] }
}))

const CFB = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0, 0, 0])
const PACKAGE_SIZE = 4242

function session(): DocxSessionApi {
	return {
		objects: () => Promise.resolve([]),
		media: () => Promise.resolve(null),
		locate: () => Promise.resolve(null),
		find: () => Promise.resolve(new Int32Array(0)),
		close: () => undefined
	}
}

vi.mock("@/lib/storage/adapter", () => ({ kvGetJson: () => Promise.resolve(null), kvSetJson: () => Promise.resolve() }))
vi.mock("@/lib/log", () => ({ log: { warn: vi.fn(), info: vi.fn(), error: vi.fn(), debug: vi.fn() } }))
vi.mock("sonner", () => ({
	toast: {
		success: (message: string) => toasts.success.push(message),
		error: (message: string) => toasts.error.push(message)
	}
}))
vi.mock("@/features/preview/hooks/usePreviewBytes", () => ({
	usePreviewBytes: () => ({ status: "success", bytes: files.bytes, refetch: () => undefined })
}))
vi.mock("@/features/drive/lib/contentVerdicts", () => ({
	recordContentVerdict: (file: { name: string }, verdict: string) => verdicts.push({ name: file.name, verdict })
}))
vi.mock("@/features/preview/components/textViewer", () => ({ TextViewer: ({ alt }: { alt: string }) => <p>text viewer: {alt}</p> }))
vi.mock("@/features/drive/lib/convertedCopy", () => ({
	uploadConvertedCopy: (item: DriveItem, bytes: Uint8Array, extension: string) => {
		uploads.push({ bytes, extension })

		return Promise.resolve({ status: "saved", item: { ...item, data: { ...item.data, uuid: "copy" } }, name: "report.docx" })
	}
}))
vi.mock("@/features/preview/lib/docxEngine", () => ({
	prewarmDocxWorker: () => undefined,
	DOCX_FONT_URLS: {},
	docxProtection: { open: protectOpen },
	docxLegacy: { open: legacyOpen },
	docxHost: (source: ByteSource, _listener: DocxListener) => {
		host.source = source

		return session()
	}
}))

const { DocxViewer } = await import("@/features/preview/components/docxViewer")

registerAllActions()

const item = linkedFileItem("report.doc", {
	uuid: "aaaaaaaa-0000-0000-0000-000000000009",
	mime: { Decrypted: "application/msword" },
	size: 12n
})

function convertedDocument() {
	const release = vi.fn()
	const opened: ByteSource[] = []
	const document: DocxConvertedDocument = {
		host: source => {
			opened.push(source)

			return session()
		},
		bytes: heldSource(PACKAGE_SIZE),
		notConverted: [
			{ kind: "shape", count: 2 },
			{ kind: "embeddedObject", count: 1 }
		],
		decrypted: false,
		take: () => Promise.resolve(new Uint8Array([0x50, 0x4b, 3, 4])),
		release
	}

	return { document, release, opened }
}

function viewer(props: { canSaveCopy?: boolean; onOpenFile?: (item: DriveItem) => void } = {}): ReactElement {
	return (
		<DocxViewer
			item={item}
			alt="report.doc"
			{...props}
		/>
	)
}

beforeEach(() => {
	host.source = null
	files.bytes = CFB
	verdicts.length = 0
	uploads.length = 0
	toasts.success.length = 0
	toasts.error.length = 0
	protectOpen.mockReset()
	legacyOpen.mockReset()
	protectOpen.mockResolvedValue({ kind: "plain" })
})

afterEach(() => {
	cleanup()
})

describe("a Word 97–2003 document in the docx viewer", () => {
	it("opens the package converted in the worker, read-only, with what is shown differently", async () => {
		const { document, opened } = convertedDocument()

		legacyOpen.mockResolvedValue({ kind: "converted", document })
		render(viewer())

		await waitFor(() => {
			expect(opened).toHaveLength(1)
		})
		expect(opened[0]?.size).toBe(PACKAGE_SIZE)
		// The page's bytes were read by the conversion only, never by an engine.
		expect(host.source).toBeNull()
		expect(legacyOpen.mock.calls[0]?.[0].size).toBe(CFB.length)

		fireEvent.click(await screen.findByRole("button", { name: "Word 97–2003 document" }))

		expect(await screen.findByText("2 drawing shapes, shown as placeholders")).toBeTruthy()
		expect(screen.getByText("1 embedded object, shown as its picture")).toBeTruthy()
		// Without a place to save a copy, no Save as, and the popover does not offer one.
		expect(screen.queryByRole("button", { name: "Save as .docx" })).toBeNull()
		expect(screen.getByText("Converted so it can be shown here, read-only.")).toBeTruthy()
		expect(screen.queryByText(/Save it as \.docx/u)).toBeNull()
	})

	it("offers Save as from the badge's popover too, where a phone's toolbar has no room for its button", async () => {
		const { document } = convertedDocument()

		legacyOpen.mockResolvedValue({ kind: "converted", document })
		render(viewer({ canSaveCopy: true }))

		fireEvent.click(await screen.findByRole("button", { name: "Word 97–2003 document" }))
		expect(await screen.findByText(/Save it as \.docx to get a copy you can edit/u)).toBeTruthy()

		const popover = screen.getByRole("dialog", { name: "Word 97–2003 document" })

		fireEvent.click(within(popover).getByRole("button", { name: "Save as .docx" }))
		await screen.findByRole("alertdialog", { name: "Save as .docx?" })
		expect(screen.queryByRole("dialog", { name: "Word 97–2003 document" })).toBeNull()
	})

	it("says the copy will not keep the document's macros or AutoText", async () => {
		const { document } = convertedDocument()

		legacyOpen.mockResolvedValue({
			kind: "converted",
			document: {
				...document,
				notConverted: [
					{ kind: "macros", count: 1 },
					{ kind: "autoText", count: 1 }
				]
			}
		})
		render(viewer({ canSaveCopy: true }))

		fireEvent.click(await screen.findByRole("button", { name: "Word 97–2003 document" }))
		expect(await screen.findByText("Its macros (a VBA project), which a .docx can't hold")).toBeTruthy()
		expect(screen.getByText("Its AutoText entries, which a .docx copy doesn't keep")).toBeTruthy()

		fireEvent.click(
			within(screen.getByRole("dialog", { name: "Word 97–2003 document" })).getByRole("button", { name: "Save as .docx" })
		)
		expect(await screen.findByText(/The copy has no macros\./u)).toBeTruthy()
		expect(screen.getByText(/It doesn't keep the template's AutoText entries\./u)).toBeTruthy()
	})

	it("saves a copy of the package as .docx beside the file and shows it", async () => {
		const { document } = convertedDocument()
		const onOpenFile = vi.fn()

		legacyOpen.mockResolvedValue({ kind: "converted", document })
		render(viewer({ canSaveCopy: true, onOpenFile }))

		fireEvent.click(await screen.findByRole("button", { name: "Save as .docx" }))
		await screen.findByRole("alertdialog", { name: "Save as .docx?" })

		const confirm = screen.getAllByRole("button", { name: "Save as .docx" }).at(-1)

		if (confirm === undefined) throw new Error("no confirm button")

		fireEvent.click(confirm)

		await waitFor(() => {
			expect(onOpenFile).toHaveBeenCalledTimes(1)
		})
		expect(uploads).toEqual([{ bytes: new Uint8Array([0x50, 0x4b, 3, 4]), extension: "docx" }])
		expect(toasts.success).toEqual(["Saved as report.docx"])
		expect((onOpenFile.mock.calls[0]?.[0] as DriveItem).data.uuid).toBe("copy")
	})

	it("says when the copy will not keep the document's editing password", async () => {
		const { document } = convertedDocument()

		legacyOpen.mockResolvedValue({
			kind: "converted",
			document: { ...document, notConverted: [{ kind: "passwordProtection", count: 1 }] }
		})
		render(viewer({ canSaveCopy: true }))

		fireEvent.click(await screen.findByRole("button", { name: "Word 97–2003 document" }))
		expect(await screen.findByText("The password on its editing restrictions, which could not be kept")).toBeTruthy()

		fireEvent.click(
			within(screen.getByRole("dialog", { name: "Word 97–2003 document" })).getByRole("button", { name: "Save as .docx" })
		)
		expect(await screen.findByText(/anyone can lift them in the copy/u)).toBeTruthy()
	})

	it.each([
		[{ reason: "legacyVersion", nFib: 0x65 }, "This document uses an older Word format that can't be opened here."],
		[{ reason: "encrypted", encryption: "unknown" }, "This document uses an encryption method Filen can't open."],
		[{ reason: "notWordBinary" }, "This file isn't a Word document."],
		[{ reason: "corrupt" }, "This document couldn't be rendered."]
	] as const)("says why it refuses %o", async (refusal, message) => {
		legacyOpen.mockResolvedValue({ kind: "refused", refusal })
		render(viewer())

		expect(await screen.findByText(message)).toBeTruthy()
		expect(host.source).toBeNull()
	})

	it("shows a .dot that is neither a compound file nor a zip as text, and records what it is", async () => {
		const graph = linkedFileItem("graph.dot", {
			uuid: "aaaaaaaa-0000-0000-0000-00000000000d",
			mime: { Decrypted: "application/msword" },
			size: 10n
		})

		files.bytes = new TextEncoder().encode("digraph{}\n")
		render(
			<DocxViewer
				item={graph}
				alt="graph.dot"
			/>
		)

		expect(await screen.findByText("text viewer: graph.dot")).toBeTruthy()
		expect(verdicts).toEqual([{ name: "graph.dot", verdict: "notWordBinary" }])
		expect(protectOpen).not.toHaveBeenCalled()
		expect(legacyOpen).not.toHaveBeenCalled()
	})

	it("opens a .dot compound file as a Word template", async () => {
		const template = linkedFileItem("letter.dot", {
			uuid: "aaaaaaaa-0000-0000-0000-00000000000e",
			mime: { Decrypted: "application/msword" },
			size: 12n
		})
		const { document, opened } = convertedDocument()

		legacyOpen.mockResolvedValue({ kind: "converted", document })
		render(
			<DocxViewer
				item={template}
				alt="letter.dot"
			/>
		)

		await waitFor(() => {
			expect(opened).toHaveLength(1)
		})
		expect(verdicts).toEqual([])
	})

	it("releases the converted package once the preview lets go of it", async () => {
		const { document, release, opened } = convertedDocument()

		legacyOpen.mockResolvedValue({ kind: "converted", document })

		const view = render(viewer())

		await waitFor(() => {
			expect(opened).toHaveLength(1)
		})
		view.unmount()
		await act(() => new Promise<void>(resolve => setTimeout(resolve, 5)))

		expect(release).toHaveBeenCalledTimes(1)
	})
})
