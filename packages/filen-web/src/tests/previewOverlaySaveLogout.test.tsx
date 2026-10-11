// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { createElement, useEffect } from "react"
import { act, cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react"
import { QueryClient } from "@tanstack/react-query"

// A sign-out asked while a Word document or a spreadsheet is being saved: nothing is asked until the save
// settles, so a discard can never be answered while the editor's write or the upload is under way (it would
// still land). Saved, the sign-out proceeds with nothing to lose; failed, the prompt asks, and a discard then
// uploads nothing.

const harness = vi.hoisted(() => ({
	// The editor's write, settled by the test.
	write: null as null | { resolve: (bytes: Uint8Array) => void; reject: (error: unknown) => void },
	uploads: 0,
	fail: false,
	commits: 0
}))

vi.mock("@/queries/client", () => ({ queryClient: new QueryClient() }))
vi.mock("sonner", () => ({ toast: { success: vi.fn(), error: vi.fn(), warning: vi.fn(), info: vi.fn() } }))
vi.mock("@tanstack/react-router", () => ({
	useBlocker: () => ({ status: "idle" }),
	useNavigate: () => vi.fn(),
	useRouterState: () => "/drive"
}))
vi.mock("@/lib/keymap/useAction", () => ({ useAction: vi.fn(), IN_EDITORS: {}, IN_EDITORS_AND_FIELDS: {} }))
vi.mock("@/lib/useIsOnline", () => ({ useIsOnline: () => true }))
vi.mock("@/features/preview/lib/docxEditingFlag", () => ({ DOCX_EDITING: true }))
vi.mock("@/features/preview/lib/previewCache", async () => ({
	...(await vi.importActual<typeof import("@/features/preview/lib/previewCache")>("@/features/preview/lib/previewCache")),
	loadPreviewBytes: () => Promise.resolve()
}))
vi.mock("@/features/drive/lib/actions", async () => ({
	...(await vi.importActual<typeof import("@/features/drive/lib/actions")>("@/features/drive/lib/actions")),
	currentRootUuid: () => "00000000-0000-0000-0000-0000000000aa"
}))
vi.mock("@/features/drive/lib/previewSave.logic", async () => {
	const actual = await vi.importActual<typeof import("@/features/drive/lib/previewSave.logic")>("@/features/drive/lib/previewSave.logic")

	return {
		...actual,
		runPreviewSave: (_deps: unknown, args: { item: import("@/features/drive/lib/item").DriveItem }) => {
			harness.uploads++

			return Promise.resolve(
				harness.fail ? { status: "error", dto: { species: "plain", message: "offline" } } : { status: "success", item: args.item }
			)
		}
	}
})

// The editors: each hands the overlay a save source writing when the test says, and reports unsaved edits.
function source(): Promise<{ bytes: Uint8Array; commit: () => void }> {
	return new Promise((resolve, reject) => {
		harness.write = {
			resolve: bytes => {
				resolve({
					bytes,
					commit: () => {
						harness.commits++
						dirtyChange?.(false)
					}
				})
			},
			reject
		}
	})
}

let dirtyChange: ((dirty: boolean) => void) | undefined

vi.mock("@/features/preview/components/docxViewer", () => ({
	DocxViewer: ({ editing }: { editing?: { saveRef: { current: unknown }; onDirtyChange: (dirty: boolean) => void } }) => {
		useEffect(() => {
			if (editing === undefined) return

			editing.saveRef.current = source
			dirtyChange = editing.onDirtyChange
			editing.onDirtyChange(true)
		}, [editing])

		return null
	}
}))
vi.mock("@/features/spreadsheet/components/spreadsheetViewer", () => ({
	SpreadsheetViewer: ({ saveRef, onDirtyChange }: { saveRef: { current: unknown }; onDirtyChange: (dirty: boolean) => void }) => {
		useEffect(() => {
			saveRef.current = source
			dirtyChange = onDirtyChange
			onDirtyChange(true)
		}, [saveRef, onDirtyChange])

		return null
	}
}))

import "@/lib/i18n"
import { registerAllActions } from "@/lib/keymap/actions"
import { usePreviewUnsavedGuardStore } from "@/features/preview/store/usePreviewUnsavedGuard"
import { narrowItem } from "@/features/drive/lib/item"
import { PreviewOverlay } from "@/features/preview/components/previewOverlay"

registerAllActions()

function itemNamed(name: string, mime: string) {
	return narrowItem({
		uuid: "11111111-1111-1111-1111-111111111111",
		stableUUID: undefined,
		parent: "22222222-2222-2222-2222-222222222222",
		size: 1_024n,
		favorited: false,
		region: "de-1",
		bucket: "filen-1",
		timestamp: 1_700_000_000_000n,
		chunks: 1n,
		canMakeThumbnail: false,
		meta: { type: "decoded", data: { name, mime, modified: 1_700_000_000_000n, size: 1_024n, key: "key", version: 2 } }
	})
}

const FILES = [
	["report.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"],
	["budget.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"]
] as const

async function openAndStartSave(name: string, mime: string) {
	const onClose = vi.fn()

	render(
		createElement(PreviewOverlay, {
			variant: "drive" as const,
			items: [itemNamed(name, mime)],
			index: 0,
			onStep: vi.fn(),
			onClose,
			onItemRemoved: vi.fn()
		})
	)

	const save = await screen.findByRole("button", { name: "Save" })

	fireEvent.click(save)
	await waitFor(() => {
		expect(harness.write).not.toBeNull()
	})

	// A sign-out asks while the editor is still writing.
	const resolve = vi.fn()

	act(() => {
		usePreviewUnsavedGuardStore.setState({ logoutRequest: { resolve } })
	})

	return { onClose, resolve }
}

beforeEach(() => {
	harness.write = null
	harness.uploads = 0
	harness.fail = false
	harness.commits = 0
	usePreviewUnsavedGuardStore.setState({ dirty: false, logoutRequest: null })
})

afterEach(() => {
	cleanup()
	usePreviewUnsavedGuardStore.setState({ dirty: false, logoutRequest: null })
})

describe.each(FILES)("PreviewOverlay — signing out while %s saves", (name, mime) => {
	it("asks nothing until the save settles, then proceeds with nothing lost", async () => {
		const { onClose, resolve } = await openAndStartSave(name, mime)

		expect(screen.queryByRole("alertdialog")).toBeNull()
		expect(resolve).not.toHaveBeenCalled()

		await act(async () => {
			harness.write?.resolve(new Uint8Array([1, 2, 3]))
			await Promise.resolve()
		})
		await waitFor(() => {
			expect(resolve).toHaveBeenCalledExactlyOnceWith(true)
		})
		expect(harness.uploads).toBe(1)
		expect(harness.commits).toBe(1)
		expect(onClose).toHaveBeenCalled()
		expect(usePreviewUnsavedGuardStore.getState().logoutRequest).toBeNull()
	})

	it("asks once a failed save settled, and a discard then uploads nothing", async () => {
		harness.fail = true

		const { onClose, resolve } = await openAndStartSave(name, mime)

		await act(async () => {
			harness.write?.resolve(new Uint8Array([1, 2, 3]))
			await Promise.resolve()
		})

		const prompt = await screen.findByRole("alertdialog")

		expect(prompt).toBeDefined()
		expect(harness.uploads).toBe(1)

		fireEvent.click(screen.getByRole("button", { name: "Discard" }))
		expect(resolve).toHaveBeenCalledExactlyOnceWith(true)
		expect(onClose).toHaveBeenCalled()
		// Nothing more went up after the discard.
		await new Promise(resolve => setTimeout(resolve, 50))
		expect(harness.uploads).toBe(1)
		expect(harness.commits).toBe(0)
	})
})
