// @vitest-environment jsdom

// A public link's download without File System Access (Firefox, Safari): the page's SDK streams into the
// service worker's download, so the browser's own download shows its progress, at any size.
import { beforeEach, describe, expect, it, vi } from "vitest"
import type { AnyLinkedDirWithContext } from "@filen/sdk-rs"
import type { SwSaveTarget } from "@/features/drive/lib/saveDownload"
import { ErrorWithDTO, plainErrorDTO } from "@/lib/sdk/errors"
import { PUBLIC_BUFFERED_DOWNLOAD_MAX_BYTES } from "@/features/publicLinks/lib/download.logic"
import { linkedFileItem } from "@/tests/fixtures/sdk"

const { downloadLinkedDirToZipAnon, downloadLinkedFileToWriterAnon, cancelTransfer, triggerSwStreamDownload } = vi.hoisted(() => ({
	downloadLinkedDirToZipAnon:
		vi.fn<
			(
				dir: AnyLinkedDirWithContext,
				id: string,
				writable: WritableStream<Uint8Array>,
				onProgress: (written: bigint, total: bigint) => void
			) => Promise<void>
		>(),
	downloadLinkedFileToWriterAnon:
		vi.fn<(file: unknown, id: string, writable: WritableStream<Uint8Array>, onProgress: (bytes: bigint) => void) => Promise<void>>(),
	cancelTransfer: vi.fn<(id: string) => void>(),
	triggerSwStreamDownload:
		vi.fn<(save: SwSaveTarget, id: string, readable: ReadableStream<Uint8Array>, size: number | null) => Promise<void>>()
}))

vi.mock("@/lib/sdk/client", () => ({ sdkApi: { downloadLinkedDirToZipAnon, downloadLinkedFileToWriterAnon, cancelTransfer } }))
vi.mock("@/features/drive/lib/saveDownload", async importOriginal => ({
	...(await importOriginal<typeof import("@/features/drive/lib/saveDownload")>()),
	isFsaAvailable: () => false,
	canStreamToSw: () => true,
	assertSwControlled: vi.fn(),
	triggerSwStreamDownload
}))

const { startAnonDirZipDownload, startAnonFileDownload } = await import("@/features/publicLinks/lib/download")
const { narrowToAnyFile } = await import("@/features/drive/lib/download")
const { clearPreviewCache } = await import("@/features/preview/lib/previewCache")

const dir = {} as AnyLinkedDirWithContext

// What the browser's download received.
async function drain(readable: ReadableStream<Uint8Array>): Promise<number[]> {
	const bytes: number[] = []

	for await (const chunk of readable) {
		bytes.push(...chunk)
	}

	return bytes
}

beforeEach(() => {
	vi.clearAllMocks()
	clearPreviewCache()
})

describe("public link downloads through the service worker", () => {
	it("streams a file over the in-memory cap into the browser's download with its size", async () => {
		const size = PUBLIC_BUFFERED_DOWNLOAD_MAX_BYTES + 1n
		const received: Promise<number[]>[] = []

		triggerSwStreamDownload.mockImplementation(async (_save, _id, readable) => {
			received.push(drain(readable))

			await received[0]
		})
		downloadLinkedFileToWriterAnon.mockImplementation(async (_file, _id, writable, onProgress) => {
			const writer = writable.getWriter()

			await writer.write(new Uint8Array([1, 2, 3]))
			onProgress(3n)
			await writer.close()
		})

		const onProgress = vi.fn()
		const outcome = await startAnonFileDownload({
			file: narrowToAnyFile(linkedFileItem("big.bin")),
			name: "big.bin",
			size,
			linkScope: "link",
			onProgress
		})

		expect(outcome).toEqual({ status: "success" })
		expect(triggerSwStreamDownload).toHaveBeenCalledWith(
			expect.objectContaining({ kind: "sw", name: "big.bin" }),
			expect.any(String),
			expect.any(ReadableStream),
			Number(size),
			expect.any(Function)
		)
		expect(await received[0]).toEqual([1, 2, 3])
		expect(onProgress).toHaveBeenCalledWith(3, Number(size))
	})

	it("starts no browser download for a file that fails before its first byte", async () => {
		downloadLinkedFileToWriterAnon.mockRejectedValue({ species: "sdk", kind: "Network", message: "offline", label: "offline" })

		const outcome = await startAnonFileDownload({
			file: narrowToAnyFile(linkedFileItem("a.txt")),
			name: "a.txt",
			size: 4n,
			linkScope: "link",
			onProgress: vi.fn()
		})

		expect(outcome.status).toBe("error")
		expect(triggerSwStreamDownload).not.toHaveBeenCalled()
	})

	it("reports a cancel from the browser's download and stops the SDK", async () => {
		triggerSwStreamDownload.mockRejectedValue(new ErrorWithDTO(plainErrorDTO("cancelled", "Cancelled")))
		downloadLinkedDirToZipAnon.mockImplementation(async (_dir, id, writable) => {
			const writer = writable.getWriter()

			await writer.write(new Uint8Array([1]))
			await vi.waitFor(() => {
				expect(cancelTransfer).toHaveBeenCalledWith(id)
			})
			await writer.abort()

			throw new ErrorWithDTO(plainErrorDTO("cancelled", "Cancelled"))
		})

		const outcome = await startAnonDirZipDownload({ dir, name: "Photos", onProgress: vi.fn() })

		expect(outcome).toEqual({ status: "cancelled" })
		expect(triggerSwStreamDownload).toHaveBeenCalledWith(
			expect.objectContaining({ name: "Photos.zip" }),
			expect.any(String),
			expect.any(ReadableStream),
			null,
			expect.any(Function)
		)
	})
})
