// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { createElement, useEffect } from "react"
import { act, cleanup, render, screen, waitFor } from "@testing-library/react"
import { QueryClient } from "@tanstack/react-query"

// The spreadsheet's idle autosave goes through the overlay's one save path: the viewer asks (onAutosave), the
// overlay writes and uploads as for Save, and refuses while a save is in flight; a sign-out asked during such a
// save waits for it to settle, then proceeds with nothing lost.

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

// The spreadsheet viewer's double: hands the overlay a save source writing when the test says, reports unsaved
// edits, and keeps the overlay's autosave request for the test to make.
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
let autosave: (() => boolean) | undefined

vi.mock("@/features/spreadsheet/next/sheetNextFlag", () => ({ SHEET_NEXT: false }))
vi.mock("@/features/spreadsheet/components/spreadsheetViewer", () => ({
	SpreadsheetViewer: ({
		saveRef,
		onDirtyChange,
		onAutosave
	}: {
		saveRef: { current: unknown }
		onDirtyChange: (dirty: boolean) => void
		onAutosave?: () => boolean
	}) => {
		useEffect(() => {
			saveRef.current = source
			dirtyChange = onDirtyChange
			autosave = onAutosave
			onDirtyChange(true)
		}, [saveRef, onDirtyChange, onAutosave])

		return null
	}
}))

import "@/lib/i18n"
import { registerAllActions } from "@/lib/keymap/actions"
import { usePreviewUnsavedGuardStore } from "@/features/preview/store/usePreviewUnsavedGuard"
import { narrowItem } from "@/features/drive/lib/item"
import { PreviewOverlay } from "@/features/preview/components/previewOverlay"

registerAllActions()

function budget() {
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
		meta: {
			type: "decoded",
			data: {
				name: "budget.xlsx",
				mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
				modified: 1_700_000_000_000n,
				size: 1_024n,
				key: "key",
				version: 2
			}
		}
	})
}

beforeEach(() => {
	harness.write = null
	harness.uploads = 0
	harness.fail = false
	harness.commits = 0
	autosave = undefined
	usePreviewUnsavedGuardStore.setState({ dirty: false, logoutRequest: null })
})

afterEach(() => {
	cleanup()
	usePreviewUnsavedGuardStore.setState({ dirty: false, logoutRequest: null })
})

describe("PreviewOverlay — a spreadsheet's idle autosave", () => {
	it("saves through the overlay, refuses a second while one is in flight, and a sign-out waits for it", async () => {
		const onClose = vi.fn()

		render(
			createElement(PreviewOverlay, {
				variant: "drive" as const,
				items: [budget()],
				index: 0,
				onStep: vi.fn(),
				onClose,
				onItemRemoved: vi.fn()
			})
		)

		await screen.findByRole("button", { name: "Save" })
		await waitFor(() => {
			expect(autosave).toBeDefined()
		})

		let started = false

		act(() => {
			started = autosave?.() === true
		})
		expect(started).toBe(true)
		await waitFor(() => {
			expect(harness.write).not.toBeNull()
		})
		// One save at a time: the next pause tries again.
		act(() => {
			started = autosave?.() === true
		})
		expect(started).toBe(false)

		const resolve = vi.fn()

		act(() => {
			usePreviewUnsavedGuardStore.setState({ logoutRequest: { resolve } })
		})
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
	})
})
