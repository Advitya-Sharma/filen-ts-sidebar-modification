// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { act, cleanup, fireEvent, render, screen } from "@testing-library/react"
import { createElement, type ReactNode } from "react"
import type { File } from "@filen/sdk-rs"
import "@/lib/i18n"

// Download as in the editors' File menus: beside the PDF, the formats the open file converts to (never its own), each
// a local download made as the item menu makes it; nothing where the file may not be downloaded. A workbook converts
// from its open session (unsaved edits included).

const { start } = vi.hoisted(() => ({ start: vi.fn() }))

vi.mock("@/features/drive/lib/downloadAs", () => ({ startDownloadAs: start }))
vi.mock("@/lib/useIsOnline", () => ({ useIsOnline: () => true }))
vi.mock("@/features/office/lib/officeSettings", () => ({
	useOfficeSettings: () => ({ loaded: true, autosave: false }),
	saveOfficeAutosave: vi.fn()
}))

import { narrowItem, type DriveItem } from "@/features/drive/lib/item"
import { DownloadAsItemProvider, type LiveConversion } from "@/features/drive/components/editorDownloadAs"
import { DocxFileMenu } from "@/features/preview/components/docxFileMenu"
import { SheetFileMenu } from "@/features/spreadsheet/save/sheetFileMenu"
import { PreviewDownloadableProvider } from "@/features/preview/lib/accessMode"

function file(name: string): DriveItem {
	return narrowItem({
		uuid: "33333333-3333-3333-3333-333333333333",
		stableUUID: undefined,
		parent: "22222222-2222-2222-2222-222222222222",
		size: 1n,
		favorited: false,
		region: "de-1",
		bucket: "filen-1",
		timestamp: 0n,
		chunks: 1n,
		canMakeThumbnail: false,
		meta: { type: "decoded", data: { name, mime: "application/octet-stream", modified: 0n, size: 1n, key: "key", version: 2 } }
	} satisfies File)
}

function within(item: DriveItem, menu: ReactNode): void {
	render(
		<PreviewDownloadableProvider
			downloadable
			onDownload={vi.fn()}
		>
			<DownloadAsItemProvider item={item}>{menu}</DownloadAsItemProvider>
		</PreviewDownloadableProvider>
	)
}

async function settle(): Promise<void> {
	for (let i = 0; i < 5; i++) await new Promise(resolve => setTimeout(resolve, 0))
}

async function openSub(trigger: string): Promise<void> {
	await act(async () => {
		fireEvent.click(screen.getByRole("button", { name: "File" }))
		await settle()
	})

	const sub = screen.getByRole("menuitem", { name: trigger })

	await act(async () => {
		sub.focus()
		fireEvent.keyDown(sub, { key: "ArrowRight" })
		await settle()
	})
}

beforeEach(() => {
	start.mockReset()
})

afterEach(() => {
	cleanup()
})

describe("the document File menu's Download as", () => {
	it("offers .odt beside the PDF for a Word document and converts the saved file", async () => {
		const item = file("report.docx")

		within(item, createElement(DocxFileMenu, { onDownloadPdf: vi.fn() }))
		await openSub("Download as")

		expect(screen.getByRole("menuitem", { name: "PDF document (.pdf)" })).toBeTruthy()
		expect(screen.queryByRole("menuitem", { name: "Word document (.docx)" })).toBeNull()

		act(() => {
			fireEvent.click(screen.getByRole("menuitem", { name: "OpenDocument text (.odt)" }))
		})

		expect(start).toHaveBeenCalledWith([item], "odt", { accessMode: "authed", cacheScope: "authed" })
	})

	it("converts the editor's document as it holds it, unsaved edits included, where the editor hands it", async () => {
		const item = file("report.docx")
		const liveFile = vi.fn((_target: "docx" | "odt", _signal: AbortSignal) =>
			Promise.resolve({ kind: "file" as const, bytes: new Uint8Array() })
		)

		within(
			item,
			<DocxFileMenu
				onDownloadPdf={vi.fn()}
				liveFile={liveFile}
			/>
		)
		await openSub("Download as")

		act(() => {
			fireEvent.click(screen.getByRole("menuitem", { name: "OpenDocument text (.odt)" }))
		})

		expect(start).toHaveBeenCalledWith([item], "odt", { accessMode: "authed", cacheScope: "authed", liveFile })
	})

	it("offers .docx and .odt for an RTF document", async () => {
		within(file("letter.rtf"), createElement(DocxFileMenu, { onDownloadPdf: vi.fn() }))
		await openSub("Download as")

		expect(screen.getByRole("menuitem", { name: "Word document (.docx)" })).toBeTruthy()
		expect(screen.getByRole("menuitem", { name: "OpenDocument text (.odt)" })).toBeTruthy()
	})

	it("offers nothing where the file may not be downloaded", () => {
		render(
			<PreviewDownloadableProvider downloadable={false}>
				<DownloadAsItemProvider item={file("report.docx")}>
					<DocxFileMenu onDownloadPdf={undefined} />
				</DownloadAsItemProvider>
			</PreviewDownloadableProvider>
		)

		expect(screen.queryByRole("button", { name: "File" })).toBeNull()
	})
})

describe("the spreadsheet File menu's Download as", () => {
	it("offers .xlsx for an OpenDocument spreadsheet and converts the open session, unsaved edits included", async () => {
		const item = file("book.ods")
		const live: LiveConversion = vi.fn()

		within(
			item,
			createElement(SheetFileMenu, {
				text: false,
				onDownloadCsv: vi.fn(),
				onDownloadPdf: vi.fn(),
				onSaveAsXlsx: undefined,
				autosave: false,
				liveConversion: live
			})
		)
		await openSub("Download as")

		expect(screen.queryByRole("menuitem", { name: "OpenDocument spreadsheet (.ods)" })).toBeNull()

		act(() => {
			fireEvent.click(screen.getByRole("menuitem", { name: "Excel workbook (.xlsx)" }))
		})

		expect(start).toHaveBeenCalledWith([item], "xlsx", { accessMode: "authed", cacheScope: "authed", live })
	})
})
