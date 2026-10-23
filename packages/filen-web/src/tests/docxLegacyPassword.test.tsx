// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { act, cleanup, fireEvent, render as renderUi, screen, waitFor } from "@testing-library/react"
import type { ReactElement } from "react"
import { createTestQueryClient, queryClientWrapper } from "@/tests/testQueryClient"
import type { ByteSource } from "@filen/office/core/opc"
import {
	heldSource,
	type DocxConvertedDocument,
	type DocxLegacy,
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

// A Word 97–2003 document with a password to open: refused as encrypted by the worker, the password is asked in the
// document's place and the file converted again with it; a wrong password asks again; the copy Save as .docx makes
// is said not to be password-protected.

function render(ui: ReactElement) {
	return renderUi(ui, { wrapper: queryClientWrapper(createTestQueryClient()) })
}

const { host, protectOpen, legacyOpen, uploads, toasts, files, verdicts } = vi.hoisted(() => ({
	files: { bytes: new Uint8Array(0) },
	verdicts: [] as { name: string; verdict: string }[],
	host: { source: null as ByteSource | null },
	protectOpen: vi.fn<(source: ByteSource, options?: DocxProtectionOptions) => Promise<DocxProtectedOpen>>(),
	legacyOpen: vi.fn<(source: ByteSource, options?: Parameters<DocxLegacy["open"]>[1]) => Promise<DocxLegacyOpen>>(),
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
		decrypted: true,
		take: () => Promise.resolve(new Uint8Array([0x50, 0x4b, 3, 4])),
		release
	}

	return { document, release, opened }
}

function viewer(): ReactElement {
	return (
		<DocxViewer
			item={item}
			alt="report.doc"
			canSaveCopy
		/>
	)
}

function submit(password: string): void {
	fireEvent.change(screen.getByLabelText("Password", { selector: "input" }), { target: { value: password } })
	fireEvent.click(screen.getByRole("button", { name: "Open" }))
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

describe("a Word 97–2003 document with a password to open", () => {
	it("asks for the password, asks again after a wrong one, then shows the decrypted conversion", async () => {
		const { document, opened } = convertedDocument()

		legacyOpen.mockImplementation((_source, options) => {
			if (options?.password === undefined)
				return Promise.resolve({ kind: "refused", refusal: { reason: "encrypted", encryption: "cryptoApiRc4" } })

			return Promise.resolve(
				options.password === "right" ? { kind: "converted", document } : { kind: "refused", refusal: { reason: "wrongPassword" } }
			)
		})
		render(viewer())

		await screen.findByRole("dialog", { name: "Password-protected document" })
		submit("wrong")

		expect((await screen.findByRole("alert")).textContent).toBe("That password isn't right. Try again.")

		submit("right")

		await waitFor(() => {
			expect(opened).toHaveLength(1)
		})
		expect(legacyOpen.mock.calls.map(call => call[1]?.password)).toEqual([undefined, "wrong", "right"])
		// Every try converts the page's bytes in the worker; no engine read them.
		expect(host.source).toBeNull()

		fireEvent.click(screen.getByRole("button", { name: "Save as .docx" }))
		expect(await screen.findByText(/The copy isn't password-protected: anyone who can open it can read it\./u)).toBeTruthy()
	})

	it.each(["xor", "rc4", "cryptoApiRc4"] as const)("asks for the password of a %s-encrypted file", async encryption => {
		legacyOpen.mockResolvedValue({ kind: "refused", refusal: { reason: "encrypted", encryption } })
		render(viewer())

		expect(await screen.findByRole("dialog", { name: "Password-protected document" })).toBeTruthy()
	})

	it("says an encryption it cannot decrypt is unsupported, without asking", async () => {
		legacyOpen.mockResolvedValue({ kind: "refused", refusal: { reason: "encrypted", encryption: "unknown" } })
		render(viewer())

		expect(await screen.findByText("This document uses an encryption method Filen can't open.")).toBeTruthy()
		expect(screen.queryByRole("dialog")).toBeNull()
	})

	it("says why when the password verified but the file could not be read", async () => {
		legacyOpen.mockImplementation((_source, options) =>
			Promise.resolve(
				options?.password === undefined
					? { kind: "refused", refusal: { reason: "encrypted", encryption: "rc4" } }
					: { kind: "refused", refusal: { reason: "corrupt" } }
			)
		)
		render(viewer())

		await screen.findByRole("dialog", { name: "Password-protected document" })
		submit("right")

		expect(await screen.findByText("This document couldn't be rendered.")).toBeTruthy()
	})

	it("drops a conversion that finished after the prompt went away", async () => {
		const { document, release } = convertedDocument()
		let finish: (opened: DocxLegacyOpen) => void = () => undefined

		legacyOpen.mockImplementation((_source, options) =>
			options?.password === undefined
				? Promise.resolve({ kind: "refused", refusal: { reason: "encrypted", encryption: "xor" } })
				: new Promise(resolve => {
						finish = resolve
					})
		)

		const { unmount } = render(viewer())

		await screen.findByRole("dialog", { name: "Password-protected document" })
		submit("right")
		await waitFor(() => {
			expect(legacyOpen).toHaveBeenCalledTimes(2)
		})
		unmount()
		await act(async () => {
			finish({ kind: "converted", document })
			await Promise.resolve()
		})

		expect(release).toHaveBeenCalled()
	})
})
