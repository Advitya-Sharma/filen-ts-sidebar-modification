import { beforeEach, describe, expect, it, vi } from "vitest"
import type { Dir, File, NormalDirsAndFiles, UuidStr } from "@filen/sdk-rs"

const { listDirectory } = vi.hoisted(() => ({
	listDirectory: vi.fn<(target: unknown) => Promise<NormalDirsAndFiles>>()
}))

vi.mock("@/lib/sdk/client", () => ({ sdkApi: { listDirectory } }))

vi.mock("@/queries/client", async () => ({ queryClient: (await import("@/tests/testQueryClient")).createTestQueryClient() }))

import { queryClient } from "@/queries/client"
import { discardListingPatches, driveListingQueryKey, driveListingQueryOptions, queueListingCreate } from "@/features/drive/queries/drive"
import { nameTakenInDirectory } from "@/features/drive/lib/nameTaken"
import { narrowItem } from "@/features/drive/lib/item"
import { socketAuthenticated, socketDropped } from "@/lib/sdk/socketSession"
import { testUuid } from "@/tests/support/uuid"

function mockFile(name: string, parent: UuidStr): File {
	return {
		uuid: testUuid(`file-${name}`),
		stableUUID: undefined,
		parent,
		size: 0n,
		favorited: false,
		region: "",
		bucket: "",
		timestamp: 1_700_000_000_000n,
		chunks: 0n,
		canMakeThumbnail: false,
		meta: { type: "decoded", data: { name, mime: "text/plain", modified: 1n, size: 0n, key: "k", version: 2 } }
	}
}

function mockDir(name: string, parent: UuidStr): Dir {
	return {
		uuid: testUuid(`dir-${name}`),
		parent,
		color: "default",
		timestamp: 1_700_000_000_000n,
		favorited: false,
		meta: { type: "decoded", data: { name } }
	}
}

// Read-this-session is page-session state, so every test lists a directory of its own.
let dirCounter = 0

function nextDir(): UuidStr {
	dirCounter++

	return testUuid(`taken-parent${String(dirCounter)}`)
}

beforeEach(() => {
	queryClient.clear()
	discardListingPatches()
	socketAuthenticated()
	listDirectory.mockReset()
})

describe("nameTakenInDirectory", () => {
	it("answers from a listing read this session under the live socket, without a request", async () => {
		const dir = nextDir()

		listDirectory.mockResolvedValue({ dirs: [], files: [mockFile("notes.txt", dir)] })
		await queryClient.query(driveListingQueryOptions("drive", dir))
		expect(listDirectory).toHaveBeenCalledTimes(1)

		await expect(nameTakenInDirectory(dir, "notes.txt")).resolves.toBe(true)
		await expect(nameTakenInDirectory(dir, "other.txt")).resolves.toBe(false)
		expect(listDirectory).toHaveBeenCalledTimes(1)
	})

	it("is case-insensitive, as the server's own name uniqueness is", async () => {
		const dir = nextDir()

		listDirectory.mockResolvedValue({ dirs: [], files: [mockFile("Report.DOCX", dir)] })

		await expect(nameTakenInDirectory(dir, "report.docx")).resolves.toBe(true)
	})

	it("counts a directory of that name as taken", async () => {
		const dir = nextDir()

		listDirectory.mockResolvedValue({ dirs: [mockDir("Untitled.txt", dir)], files: [] })

		await expect(nameTakenInDirectory(dir, "untitled.txt")).resolves.toBe(true)
	})

	it("reads the directory when nothing is cached, and keeps the read for the listing", async () => {
		const dir = nextDir()

		listDirectory.mockResolvedValue({ dirs: [], files: [] })

		await expect(nameTakenInDirectory(dir, "notes.txt")).resolves.toBe(false)
		expect(listDirectory).toHaveBeenCalledTimes(1)
		expect(queryClient.getQueryData(driveListingQueryKey({ variant: "drive", uuid: dir }))).toEqual([])
	})

	it("rereads a cached listing it cannot trust, so a file made elsewhere is never missed", async () => {
		const dir = nextDir()
		const existing = mockFile("notes.txt", dir)

		// Restored from disk: rows from an earlier session, not read under this socket.
		queryClient.setQueryData(driveListingQueryKey({ variant: "drive", uuid: dir }), [], { updatedAt: Date.now() })
		listDirectory.mockResolvedValue({ dirs: [], files: [existing] })

		await expect(nameTakenInDirectory(dir, "notes.txt")).resolves.toBe(true)
		expect(listDirectory).toHaveBeenCalledTimes(1)
		expect(queryClient.getQueryData(driveListingQueryKey({ variant: "drive", uuid: dir }))).toEqual([narrowItem(existing)])
	})

	it("rereads a listing read while the socket was down", async () => {
		const dir = nextDir()

		socketDropped()
		listDirectory.mockResolvedValue({ dirs: [], files: [] })
		await queryClient.query(driveListingQueryOptions("drive", dir))
		socketAuthenticated()
		listDirectory.mockResolvedValue({ dirs: [], files: [mockFile("notes.txt", dir)] })

		await expect(nameTakenInDirectory(dir, "notes.txt")).resolves.toBe(true)
		expect(listDirectory).toHaveBeenCalledTimes(2)
	})

	it("sees a create still queued for the listing", async () => {
		const dir = nextDir()

		listDirectory.mockResolvedValue({ dirs: [], files: [] })
		await queryClient.query(driveListingQueryOptions("drive", dir))
		queueListingCreate(dir, narrowItem(mockFile("fresh.docx", dir)))

		await expect(nameTakenInDirectory(dir, "fresh.docx")).resolves.toBe(true)
		expect(listDirectory).toHaveBeenCalledTimes(1)
	})
})
