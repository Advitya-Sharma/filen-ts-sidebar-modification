import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import type { Query } from "@tanstack/react-query"
import type { File } from "@filen/sdk-rs"
import { testUuid } from "@/tests/support/uuid"

// A save's listing patch reaches disk: the persister writes a row only after a read, so without this a reload
// restores the listing as last read, naming the version the save replaced, and a click before its first read
// lands opens that version. Patched listings are written once per window, never once a sign-out started.

const hooks = vi.hoisted(() => ({ persisted: [] as Query[], signOut: false }))

vi.mock("@/lib/sdk/client", () => ({ sdkApi: {} }))
vi.mock("@/queries/client", async () => ({ queryClient: (await import("@/tests/testQueryClient")).createTestQueryClient() }))
vi.mock("@/queries/persist", () => ({
	persister: {
		persistQuery: (query: Query) => {
			hooks.persisted.push(query)

			return Promise.resolve()
		}
	}
}))
vi.mock("@/lib/signOutPending", () => ({ readSignOutPending: () => hooks.signOut }))

const { queryClient } = await import("@/queries/client")
const { driveListingQueryKey, driveListingQueryUpdate } = await import("@/features/drive/queries/drive")
const { runPreviewSave } = await import("@/features/drive/lib/previewSave.logic")
const { narrowItem } = await import("@/features/drive/lib/item")

type DriveItem = ReturnType<typeof narrowItem>

function file(label: string, name = "budget.xlsx"): DriveItem {
	return narrowItem({
		uuid: testUuid(label),
		stableUUID: "lineage" as File["stableUUID"],
		parent: testUuid("parent"),
		size: 1n,
		favorited: false,
		region: "de-1",
		bucket: "filen-1",
		timestamp: 0n,
		chunks: 1n,
		canMakeThumbnail: false,
		meta: { type: "decoded", data: { name, mime: "", modified: 0n, size: 1n, key: "k", version: 2 } }
	})
}

const KEY = driveListingQueryKey({ variant: "drive", uuid: testUuid("parent") })

async function save(from: DriveItem, to: DriveItem): Promise<void> {
	const outcome = await runPreviewSave(
		{
			uploadFileBytes: () => Promise.resolve(to.data as File),
			patchListing: driveListingQueryUpdate,
			rootUuid: testUuid("root")
		},
		{ item: from, content: new Uint8Array([1]) }
	)

	expect(outcome.status).toBe("success")
}

function uuidsOf(query: Query | undefined): string[] {
	return ((query?.state.data ?? []) as DriveItem[]).map(item => item.data.uuid)
}

beforeEach(() => {
	vi.useFakeTimers()
	hooks.persisted = []
	hooks.signOut = false
	queryClient.clear()
	queryClient.setQueryData(KEY, [file("v0"), file("other", "notes.txt")])
})

afterEach(() => {
	vi.useRealTimers()
})

describe("a drive listing patched after a save", () => {
	it("is written to disk naming the saved version, once for every patch in the window", async () => {
		await save(file("v0"), file("v1"))
		await save(file("v1"), file("v2"))
		expect(hooks.persisted).toEqual([])

		await vi.advanceTimersByTimeAsync(1_000)

		expect(hooks.persisted).toHaveLength(1)
		expect(uuidsOf(hooks.persisted[0])).toEqual(expect.arrayContaining([testUuid("v2"), testUuid("other")]))
		expect(uuidsOf(hooks.persisted[0])).not.toContain(testUuid("v0"))
		expect(uuidsOf(hooks.persisted[0])).not.toContain(testUuid("v1"))
	})

	it("is not written once a sign-out started, nor once the listing was dropped", async () => {
		await save(file("v0"), file("v1"))
		hooks.signOut = true
		await vi.advanceTimersByTimeAsync(1_000)
		expect(hooks.persisted).toEqual([])

		hooks.signOut = false
		await save(file("v1"), file("v2"))
		queryClient.removeQueries({ queryKey: KEY })
		await vi.advanceTimersByTimeAsync(1_000)
		expect(hooks.persisted).toEqual([])
	})
})
