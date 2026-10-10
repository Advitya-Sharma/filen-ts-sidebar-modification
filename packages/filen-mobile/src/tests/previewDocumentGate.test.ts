// @vitest-environment happy-dom

import { vi, describe, it, expect, beforeEach } from "vitest"
import { createElement } from "react"
import { render } from "@testing-library/react"

const { mockUseFileUriQuery, mockUseRangeSource } = vi.hoisted(() => ({
	mockUseFileUriQuery: vi.fn(),
	mockUseRangeSource: vi.fn()
}))

vi.mock("react-native", async () => {
	const actual = await import("@/tests/mocks/reactNative")
	const { createElement: h } = await import("react")

	return {
		...actual,
		ActivityIndicator: () => h("div", { "data-testid": "activity" })
	}
})

vi.mock("react-i18next", async () => await import("@/tests/mocks/reactI18next"))
vi.mock("zustand/shallow", async () => await import("@/tests/mocks/zustandShallow"))

vi.mock("@shopify/flash-list", async () => {
	const { useState } = await import("react")

	return {
		useRecyclingState: (init: unknown) => useState(typeof init === "function" ? (init as () => unknown)() : init)
	}
})

vi.mock("@/components/drivePreview/gallery", () => ({
	galleryItemKey: (item: { data: { data: { uuid: string } } }) => item.data.data.uuid
}))

vi.mock("@/queries/useFileUri.query", () => ({
	default: mockUseFileUriQuery
}))

vi.mock("@/hooks/useRangeSource", () => ({
	default: mockUseRangeSource
}))

vi.mock("@/hooks/useIsOnline", () => ({
	default: () => true
}))

vi.mock("@/stores/useDrivePreview.store", () => ({
	default: (selector: (state: Record<string, unknown>) => unknown) => selector({ headerHeight: 0, setContentScrolled: vi.fn() })
}))

vi.mock("react-native-safe-area-context", () => ({
	useSafeAreaInsets: () => ({ top: 0, bottom: 0, left: 0, right: 0 })
}))

vi.mock("uniwind", () => ({
	useResolveClassNames: () => ({ backgroundColor: "#000" }),
	useUniwind: () => ({ theme: "dark" })
}))

vi.mock("@/components/drivePreview/useEditableSave", () => ({
	default: () => ({
		hasEdits: false,
		setHasEdits: vi.fn(),
		saveHandleRef: { current: null },
		readOnly: false,
		save: vi.fn(),
		isOnline: true
	})
}))

vi.mock("@/components/drivePreview/previewStatus", async () => {
	const { createElement: h } = await import("react")

	return {
		PreviewSpinner: () => h("div", { "data-testid": "spinner" }),
		PreviewMessage: (props: { text: string }) => h("div", { "data-testid": "message" }, props.text),
		PreviewStatusMessage: (props: { text: string }) => h("div", null, props.text)
	}
})

vi.mock("@/components/ui/view", async () => {
	const { createElement: h } = await import("react")

	return {
		default: (props: { children?: unknown }) => h("div", null, props.children as never)
	}
})

function marker(testId: string) {
	return async () => {
		const { createElement: h } = await import("react")

		return { default: () => h("div", { "data-testid": testId }), backgroundColors: { normal: { dark: "#000", light: "#fff" } } }
	}
}

vi.mock("@/components/textEditor", marker("text-viewer"))
vi.mock("@/components/pdfPreview", marker("pdf-viewer"))
vi.mock("@/components/docxPreview", marker("docx-viewer"))
vi.mock("@/components/drivePreview/previewSaveButton", marker("save-button"))
vi.mock("@/components/drivePreview/unavailableOfflineNotice", marker("offline"))
vi.mock("@/components/drivePreview/previewLoadFailedNotice", marker("failed"))

import PreviewText from "@/components/drivePreview/previewText"
import PreviewPdf from "@/components/drivePreview/previewPdf"
import PreviewDocx from "@/components/drivePreview/previewDocx"
import { MAX_TEXT_BYTES } from "@/components/textEditor/constants"
import { MAX_PDF_BYTES } from "@/components/pdfPreview/constants"
import { MAX_DOCX_BYTES } from "@/components/docxPreview/constants"

function page(name: string, size: number) {
	return {
		type: "drive" as const,
		data: {
			type: "file",
			data: { uuid: `uuid-${name}`, size: BigInt(size), decryptedMeta: { name, size: BigInt(size) }, undecryptable: false }
		}
	} as never
}

const VIEWERS = [
	["text", PreviewText, "notes.txt", MAX_TEXT_BYTES, "text_file_too_large", "text-viewer"],
	["pdf", PreviewPdf, "scan.pdf", MAX_PDF_BYTES, "pdf_too_large", "pdf-viewer"],
	["docx", PreviewDocx, "report.docx", MAX_DOCX_BYTES, "document_too_large", "docx-viewer"]
] as const

describe("PreviewDocumentGate — a drive file over its viewer's cap", () => {
	beforeEach(() => {
		mockUseFileUriQuery.mockReset().mockReturnValue({ status: "success", fetchStatus: "idle", data: { uri: "file:///cache/f" } })
		mockUseRangeSource.mockReset().mockImplementation((_uri: string | null, { maxBytes }: { maxBytes: number }) => ({
			status: "ready",
			size: maxBytes,
			readRange: vi.fn()
		}))
	})

	it.each(VIEWERS)("%s: opens a file at the cap, fetching it", (_kind, Viewer, name, cap, _message, viewerId) => {
		const { container } = render(createElement(Viewer, { item: page(name, cap) }))

		expect(mockUseFileUriQuery).toHaveBeenCalledWith(expect.objectContaining({ type: "drive" }), { enabled: true })
		expect(mockUseRangeSource).toHaveBeenCalledWith("file:///cache/f", expect.objectContaining({ maxBytes: cap }))
		expect(container.querySelector(`[data-testid='${viewerId}']`)).not.toBeNull()
		expect(container.querySelector("[data-testid='message']")).toBeNull()
	})

	it.each(VIEWERS)(
		"%s: refuses a file one byte over the cap with the viewer's own message, never fetching it",
		(_kind, Viewer, name, cap, message, viewerId) => {
			const { container } = render(createElement(Viewer, { item: page(name, cap + 1) }))

			expect(mockUseFileUriQuery).toHaveBeenCalledWith(expect.objectContaining({ type: "drive" }), { enabled: false })
			expect(container.querySelector("[data-testid='message']")?.textContent).toBe(message)
			expect(container.querySelector(`[data-testid='${viewerId}']`)).toBeNull()
		}
	)

	it("still refuses by the downloaded copy's own size, with the same message", () => {
		mockUseRangeSource.mockReturnValue({ status: "refused", reason: "tooLarge", size: MAX_PDF_BYTES + 1 })

		const { container } = render(createElement(PreviewPdf, { item: page("scan.pdf", 10) }))

		expect(container.querySelector("[data-testid='message']")?.textContent).toBe("pdf_too_large")
	})
})
