import { beforeEach, describe, expect, it, vi } from "vitest"
import type { Dir } from "@filen/sdk-rs"
import { noteFoundDir, resolveParentDir } from "@/workers/parentDir"
import { cacheDirs, classifyUnderRoot, clearDirectoryCache, forgetFoundDirs } from "@/features/drive/lib/cache"

const DRIVE_ROOT = crypto.randomUUID()

function dir(uuid: string, parent: string = DRIVE_ROOT): Dir {
	return { uuid, parent, meta: { type: "decoded", data: { name: uuid } }, timestamp: 0n, color: "default" } as Dir
}

const lookup = vi.fn((uuid: string) => Promise.resolve<Dir | undefined>(dir(uuid)))

beforeEach(() => {
	clearDirectoryCache()
})

describe("resolveParentDir", () => {
	it("passes a found directory as a uuid-only parent, with no lookup", async () => {
		const uuid = crypto.randomUUID()

		noteFoundDir(uuid, DRIVE_ROOT)

		await expect(resolveParentDir(uuid, lookup)).resolves.toEqual({ uuid })
		expect(lookup).not.toHaveBeenCalled()
	})

	it("prefers a cached Dir over a found uuid", async () => {
		const cached = dir(crypto.randomUUID())

		noteFoundDir(cached.uuid, DRIVE_ROOT)
		cacheDirs([cached])

		await expect(resolveParentDir(cached.uuid, lookup)).resolves.toBe(cached)
		expect(lookup).not.toHaveBeenCalled()
	})

	it("looks up any other directory, and one forgotten after a socket gap", async () => {
		const uuid = crypto.randomUUID()

		noteFoundDir(uuid, DRIVE_ROOT)
		forgetFoundDirs()

		await expect(resolveParentDir(uuid, lookup)).resolves.toEqual(dir(uuid))
		expect(lookup).toHaveBeenCalledExactlyOnceWith(uuid)
	})

	it("never notes a malformed uuid", async () => {
		noteFoundDir("not-a-uuid", DRIVE_ROOT)

		await resolveParentDir("not-a-uuid", lookup)

		expect(lookup).toHaveBeenCalledOnce()
	})
})

describe("found directories place a uuid under the drive root", () => {
	it("a directory found under one found at the root is provably outside a Photos root", () => {
		const dotFilen = crypto.randomUUID()
		const chatUploads = crypto.randomUUID()
		const photosRoot = crypto.randomUUID()

		expect(classifyUnderRoot(chatUploads, photosRoot, DRIVE_ROOT)).toBe("unknown")

		noteFoundDir(dotFilen, DRIVE_ROOT)
		noteFoundDir(chatUploads, dotFilen)

		expect(classifyUnderRoot(chatUploads, photosRoot, DRIVE_ROOT)).toBe("nonMember")
		expect(classifyUnderRoot(chatUploads, dotFilen, DRIVE_ROOT)).toBe("member")
	})
})
