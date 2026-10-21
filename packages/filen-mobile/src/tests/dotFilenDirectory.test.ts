import { vi, describe, it, expect, beforeEach } from "vitest"

const { mockSdkClient } = vi.hoisted(() => ({
	mockSdkClient: {
		findDirUuid: vi.fn(),
		getDir: vi.fn(),
		createDir: vi.fn(),
		root: vi.fn(() => ({ uuid: "root-uuid" }))
	}
}))

vi.mock("@/lib/auth", () => ({
	default: {
		getSdkClients: vi.fn(async () => ({ authedSdkClient: mockSdkClient }))
	}
}))

vi.mock("@/lib/signals", () => ({
	toSignalOpts: (signal?: AbortSignal) => (signal ? { signal } : undefined)
}))

vi.mock("@filen/sdk-rs", () => {
	class WrapClass {
		public inner: unknown

		public constructor(inner: unknown) {
			this.inner = inner
		}
	}

	return {
		AnyNormalDir: { Root: WrapClass, Dir: WrapClass }
	}
})

import { ensureDotFilenSubdirectory, clearDotFilenDirectoryMemo } from "@/lib/dotFilenDirectory"
import { useSocketStore } from "@/stores/useSocket.store"
import events from "@/lib/events"

function dir(name: string) {
	return { uuid: `${name}-uuid`, meta: { tag: "Decoded", inner: [{ name }] } }
}

const dotFilen = dir(".filen")
const playlists = dir("Playlists")
const chats = dir("Chat Uploads")
const existing = [dotFilen, playlists, chats]

// The drive as the name-hash lookup sees it: `.filen` under the root, both subdirectories under `.filen`.
function serveExisting(): void {
	mockSdkClient.findDirUuid.mockImplementation(async (parent: string, name: string) => {
		const lower = name.toLowerCase()

		if (parent === "root-uuid") {
			return lower === ".filen" ? dotFilen.uuid : undefined
		}

		return parent === dotFilen.uuid ? existing.find(d => d.meta.inner[0]?.name.toLowerCase() === lower)?.uuid : undefined
	})
	mockSdkClient.getDir.mockImplementation(async (uuid: string) => existing.find(d => d.uuid === uuid))
}

function resetClient(): void {
	mockSdkClient.findDirUuid.mockReset()
	mockSdkClient.getDir.mockReset()
	mockSdkClient.createDir.mockReset()
}

describe("ensureDotFilenSubdirectory", () => {
	beforeEach(() => {
		clearDotFilenDirectoryMemo()
		useSocketStore.setState({ state: "disconnected", connectedAt: 0 })
		resetClient()
	})

	it("finds existing directories by name and fetches only the subdirectory, listing nothing", async () => {
		serveExisting()

		const signal = new AbortController().signal

		await expect(ensureDotFilenSubdirectory("Playlists", signal)).resolves.toBe(playlists)
		expect(mockSdkClient.findDirUuid).toHaveBeenNthCalledWith(1, "root-uuid", ".filen", { signal })
		expect(mockSdkClient.findDirUuid).toHaveBeenNthCalledWith(2, dotFilen.uuid, "Playlists", { signal })
		expect(mockSdkClient.getDir).toHaveBeenCalledTimes(1)
		expect(mockSdkClient.getDir).toHaveBeenCalledWith(playlists.uuid, { signal })
		expect(mockSdkClient.createDir).not.toHaveBeenCalled()
	})

	it("creates both levels when .filen is missing, without looking under the new directory", async () => {
		const createdChild = dir("Chat Uploads")

		mockSdkClient.findDirUuid.mockResolvedValueOnce(undefined)
		mockSdkClient.createDir.mockResolvedValueOnce(dotFilen)
		mockSdkClient.createDir.mockResolvedValueOnce(createdChild)

		await expect(ensureDotFilenSubdirectory("Chat Uploads")).resolves.toBe(createdChild)
		expect(mockSdkClient.findDirUuid).toHaveBeenCalledTimes(1)
		expect(mockSdkClient.getDir).not.toHaveBeenCalled()
		expect(mockSdkClient.createDir).toHaveBeenNthCalledWith(1, { inner: { uuid: "root-uuid" } }, ".filen", undefined)
		expect(mockSdkClient.createDir).toHaveBeenNthCalledWith(2, { inner: dotFilen }, "Chat Uploads", undefined)
	})

	it("fetches .filen to create a missing subdirectory in it", async () => {
		const createdChild = dir("Playlists")

		mockSdkClient.findDirUuid.mockResolvedValueOnce(dotFilen.uuid)
		mockSdkClient.findDirUuid.mockResolvedValueOnce(undefined)
		mockSdkClient.getDir.mockResolvedValueOnce(dotFilen)
		mockSdkClient.createDir.mockResolvedValueOnce(createdChild)

		await expect(ensureDotFilenSubdirectory("Playlists")).resolves.toBe(createdChild)
		expect(mockSdkClient.getDir).toHaveBeenCalledWith(dotFilen.uuid, undefined)
		expect(mockSdkClient.createDir).toHaveBeenCalledTimes(1)
		expect(mockSdkClient.createDir).toHaveBeenCalledWith({ inner: dotFilen }, "Playlists", undefined)
	})
})

describe("ensureDotFilenSubdirectory memo", () => {
	function connect(connectedAt: number): void {
		useSocketStore.setState({ state: "connected", connectedAt })
	}

	beforeEach(() => {
		clearDotFilenDirectoryMemo()
		useSocketStore.setState({ state: "disconnected", connectedAt: 0 })
		resetClient()
		serveExisting()
	})

	it("reuses each resolved directory within the same connected socket session", async () => {
		connect(100)

		await expect(ensureDotFilenSubdirectory("Playlists")).resolves.toBe(playlists)
		await expect(ensureDotFilenSubdirectory("Playlists")).resolves.toBe(playlists)
		expect(mockSdkClient.findDirUuid).toHaveBeenCalledTimes(2)

		await expect(ensureDotFilenSubdirectory("Chat Uploads")).resolves.toBe(chats)
		await expect(ensureDotFilenSubdirectory("Chat Uploads")).resolves.toBe(chats)
		expect(mockSdkClient.findDirUuid).toHaveBeenCalledTimes(4)
	})

	it("looks up every time while the socket is not connected", async () => {
		await ensureDotFilenSubdirectory("Playlists")
		await ensureDotFilenSubdirectory("Playlists")

		useSocketStore.setState({ state: "reconnecting", connectedAt: 100 })

		await ensureDotFilenSubdirectory("Playlists")

		expect(mockSdkClient.findDirUuid).toHaveBeenCalledTimes(6)
	})

	it("looks up again once the socket reconnected", async () => {
		connect(100)

		await ensureDotFilenSubdirectory("Playlists")

		connect(200)

		await ensureDotFilenSubdirectory("Playlists")
		await ensureDotFilenSubdirectory("Playlists")

		expect(mockSdkClient.findDirUuid).toHaveBeenCalledTimes(4)
	})

	it("keeps nothing a clear overlapped", async () => {
		connect(100)

		let release: () => void = () => {}
		const gate = new Promise<void>(resolve => {
			release = resolve
		})

		mockSdkClient.findDirUuid.mockImplementationOnce(async () => {
			await gate

			return dotFilen.uuid
		})

		const pending = ensureDotFilenSubdirectory("Playlists")

		clearDotFilenDirectoryMemo()
		release()

		await expect(pending).resolves.toBe(playlists)

		await ensureDotFilenSubdirectory("Playlists")

		expect(mockSdkClient.findDirUuid).toHaveBeenCalledTimes(4)
	})

	it("keeps nothing resolved across a connect: what changed while down was missed", async () => {
		useSocketStore.setState({ state: "disconnected", connectedAt: 100 })
		mockSdkClient.findDirUuid.mockImplementationOnce(async () => {
			connect(300)

			return dotFilen.uuid
		})

		await ensureDotFilenSubdirectory("Playlists")
		await ensureDotFilenSubdirectory("Playlists")

		expect(mockSdkClient.findDirUuid).toHaveBeenCalledTimes(4)
	})

	it("local drive changes and logout clear it", async () => {
		connect(100)

		await ensureDotFilenSubdirectory("Playlists")

		events.emit("driveItemRemoved", { uuid: "x" })
		await ensureDotFilenSubdirectory("Playlists")

		events.emit("driveItemUpdated", { previousUuid: "x", item: playlists as never })
		await ensureDotFilenSubdirectory("Playlists")

		events.emit("logout")
		await ensureDotFilenSubdirectory("Playlists")

		expect(mockSdkClient.findDirUuid).toHaveBeenCalledTimes(8)
	})
})
