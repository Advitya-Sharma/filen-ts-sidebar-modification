// @vitest-environment jsdom

import { afterEach, describe, expect, it, vi } from "vitest"
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react"
import type { File as SdkFile } from "@filen/sdk-rs"

const { uploadFileBytes, nameTakenInDirectory } = vi.hoisted(() => ({
	uploadFileBytes: vi.fn<(parent: string | null, data: Uint8Array, name: string, mime: string) => Promise<SdkFile>>(),
	nameTakenInDirectory: vi.fn<(parent: string | null, name: string) => Promise<boolean>>()
}))

vi.mock("@/lib/sdk/client", () => ({ sdkApi: { uploadFileBytes } }))
vi.mock("@/features/drive/lib/nameTaken", () => ({ nameTakenInDirectory }))
vi.mock("@/features/drive/queries/drive", async importOriginal => ({
	...(await importOriginal<typeof import("@/features/drive/queries/drive")>()),
	driveListingQueryUpdate: vi.fn()
}))

import "@/lib/i18n"
import { useUploadMenuActions } from "@/features/drive/hooks/useUploadMenuActions"

afterEach(() => {
	cleanup()
})

function created(name: string): SdkFile {
	return {
		uuid: "11111111-1111-1111-1111-111111111111",
		stableUUID: undefined,
		parent: "22222222-2222-2222-2222-222222222222",
		chunks: 0n,
		size: 0n,
		favorited: false,
		region: "",
		bucket: "",
		timestamp: 1_700_000_000_000n,
		canMakeThumbnail: false,
		meta: { type: "decoded", data: { name, mime: "text/plain", size: 0n, modified: 0n, key: "k", version: 2 } }
	}
}

function Harness({ openPreview }: { openPreview: (items: unknown[], index: number) => void }) {
	const actions = useUploadMenuActions({ parentUuid: "parent", disabled: false, openPreview, hiddenNotice: false })

	return (
		<>
			{actions.host}
			<button
				type="button"
				onClick={actions.newTextFile}
			>
				open
			</button>
		</>
	)
}

function openDialog() {
	const openPreview = vi.fn()

	render(<Harness openPreview={openPreview} />)
	fireEvent.click(screen.getByRole("button", { name: "open" }))

	return { openPreview }
}

function submit(name: string): void {
	fireEvent.change(screen.getByLabelText("Name"), { target: { value: name } })
	fireEvent.click(screen.getByRole("button", { name: "Create" }))
}

describe("New text file dialog", () => {
	it("shows an inline error on a taken name and never uploads onto it", async () => {
		nameTakenInDirectory.mockResolvedValue(true)
		const { openPreview } = openDialog()

		submit("notes")

		expect(await screen.findByText("A file or directory with this name already exists here.")).toBeTruthy()
		expect(nameTakenInDirectory).toHaveBeenCalledWith("parent", "notes.txt")
		expect(uploadFileBytes).not.toHaveBeenCalled()
		expect(openPreview).not.toHaveBeenCalled()
		expect(screen.getByRole("button", { name: "Create" }).hasAttribute("disabled")).toBe(true)

		// A different name clears the error and can be created.
		nameTakenInDirectory.mockResolvedValue(false)
		uploadFileBytes.mockResolvedValue(created("notes 2.txt"))
		submit("notes 2")

		await waitFor(() => {
			expect(openPreview).toHaveBeenCalledOnce()
		})
		expect(screen.queryByText("A file or directory with this name already exists here.")).toBeNull()
		expect(uploadFileBytes).toHaveBeenCalledWith("parent", new Uint8Array(0), "notes 2.txt", "text/plain")
	})

	it("forgets the taken name when the dialog opens again", async () => {
		nameTakenInDirectory.mockResolvedValue(true)
		openDialog()
		submit("notes")
		await screen.findByText("A file or directory with this name already exists here.")

		fireEvent.keyDown(screen.getByLabelText("Name"), { key: "Escape" })
		await waitFor(() => {
			expect(screen.queryByLabelText("Name")).toBeNull()
		})
		fireEvent.click(screen.getByRole("button", { name: "open" }))
		fireEvent.change(screen.getByLabelText("Name"), { target: { value: "notes" } })

		expect(screen.queryByText("A file or directory with this name already exists here.")).toBeNull()
	})
})
