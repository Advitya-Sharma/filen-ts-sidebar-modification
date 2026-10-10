import { beforeEach, describe, expect, it, vi } from "vitest"
import type { Dir, SocketEvent, UuidStr } from "@filen/sdk-rs"
import { testUuid } from "@/tests/support/uuid"
import { driveFile } from "@/tests/support/videoDrive"

const DOT_FILEN = testUuid("dotfilen")
const SUBDIRS: Record<string, UuidStr> = {
	Playlists: testUuid("playlists"),
	Videos: testUuid("videos"),
	"Chat Uploads": testUuid("chatuploads")
}

const { createDirectory, directoryExists, listDirectory, subscribeToSocket } = vi.hoisted(() => ({
	createDirectory: vi.fn<(parent: string | null, name: string) => Promise<Dir>>(),
	directoryExists: vi.fn<(parent: string | null, name: string) => Promise<boolean>>(),
	listDirectory: vi.fn<(target: { kind: "uuid"; uuid: string }) => Promise<{ dirs: Dir[]; files: [] }>>(),
	subscribeToSocket: vi.fn<(onEvent: (event: SocketEvent) => void) => Promise<void>>()
}))

vi.mock("@/lib/sdk/client", () => ({ sdkApi: { createDirectory, directoryExists, listDirectory, subscribeToSocket } }))
vi.mock("@/queries/client", async () => ({ queryClient: (await import("@/tests/testQueryClient")).createTestQueryClient() }))
vi.mock("@/queries/persist", () => ({ persister: { persistQuery: () => Promise.resolve() } }))

function dir(uuid: UuidStr, name: string = uuid): Dir {
	return { uuid, meta: { type: "decoded", data: { name } }, timestamp: 0n, color: "default" } as Dir
}

// `.filen` as a listing shows it: every subdirectory, under its own name.
function dotFilenListing(target: { uuid: string }): Promise<{ dirs: Dir[]; files: [] }> {
	return Promise.resolve({
		dirs: target.uuid === DOT_FILEN ? Object.entries(SUBDIRS).map(([name, uuid]) => dir(uuid, name)) : [],
		files: []
	})
}

async function fresh() {
	vi.resetModules()

	return import("@/lib/dotFilen")
}

function driveEvent(inner: Extract<SocketEvent, { type: "drive" }>["inner"]): Extract<SocketEvent, { type: "drive" }> {
	return { type: "drive", inner, driveMessageId: 0n }
}

beforeEach(() => {
	vi.clearAllMocks()
	directoryExists.mockImplementation(() => Promise.resolve(true))
	listDirectory.mockImplementation(dotFilenListing)
	createDirectory.mockImplementation((parent, name) => {
		if (parent === null) {
			return Promise.resolve(dir(DOT_FILEN))
		}

		const uuid = SUBDIRS[name]

		return uuid === undefined ? Promise.reject(new Error(`unexpected ${name}`)) : Promise.resolve(dir(uuid))
	})
})

describe(".filen subdirectories", () => {
	it("resolves .filen once for every subdirectory", async () => {
		const { getDotFilenSubdir } = await fresh()

		await expect(Promise.all([getDotFilenSubdir("Playlists"), getDotFilenSubdir("Videos")])).resolves.toEqual([
			SUBDIRS["Playlists"],
			SUBDIRS["Videos"]
		])
		await getDotFilenSubdir("Chat Uploads")

		// .filen once, then one call per subdirectory.
		expect(createDirectory).toHaveBeenCalledTimes(4)
	})

	it("knows the resolved directories synchronously", async () => {
		const { getDotFilenSubdir, isDotFilenSystemDir, dotFilenSubdirUuid } = await fresh()

		expect(dotFilenSubdirUuid("Videos")).toBeNull()

		await getDotFilenSubdir("Videos")

		expect(dotFilenSubdirUuid("Videos")).toBe(SUBDIRS["Videos"])
		expect(isDotFilenSystemDir(DOT_FILEN)).toBe(true)
		expect(isDotFilenSystemDir(SUBDIRS["Videos"] ?? "")).toBe(true)
		expect(isDotFilenSystemDir(testUuid("other"))).toBe(false)
	})

	it("forgets only the subdirectory an event concerns, and tells its feature", async () => {
		const { getDotFilenSubdir, handleDotFilenDriveEvent, onDotFilenSubdirForgotten, dotFilenSubdirUuid } = await fresh()
		const forgotten = vi.fn()

		onDotFilenSubdirForgotten("Videos", forgotten)
		await getDotFilenSubdir("Videos")
		await getDotFilenSubdir("Playlists")

		handleDotFilenDriveEvent(driveEvent({ type: "folderTrash", parent: DOT_FILEN, uuid: SUBDIRS["Videos"] ?? DOT_FILEN }))

		expect(forgotten).toHaveBeenCalledTimes(1)
		expect(dotFilenSubdirUuid("Videos")).toBeNull()
		expect(dotFilenSubdirUuid("Playlists")).toBe(SUBDIRS["Playlists"])
	})

	it("forgets everything when .filen itself goes", async () => {
		const { getDotFilenSubdir, handleDotFilenDriveEvent, dotFilenSubdirUuid } = await fresh()

		await getDotFilenSubdir("Videos")
		await getDotFilenSubdir("Playlists")

		handleDotFilenDriveEvent(
			driveEvent({ type: "folderMetadataChanged", uuid: DOT_FILEN, meta: { type: "decoded", data: { name: "old" } } })
		)

		expect(dotFilenSubdirUuid("Videos")).toBeNull()
		expect(dotFilenSubdirUuid("Playlists")).toBeNull()
	})

	it("remembers files saved into a system directory, so a later uuid-only event is known as app data", async () => {
		const { getDotFilenSubdir, handleDotFilenDriveEvent, isDotFilenSystemFile } = await fresh()
		const file = testUuid("state")

		await getDotFilenSubdir("Videos")

		expect(isDotFilenSystemFile(file)).toBe(false)

		handleDotFilenDriveEvent(
			driveEvent({
				type: "fileNew",
				file: {
					uuid: file,
					stableUUID: file,
					parent: SUBDIRS["Videos"] ?? DOT_FILEN,
					size: 1n,
					favorited: false,
					region: "r",
					bucket: "b",
					timestamp: 0n,
					chunks: 1n,
					canMakeThumbnail: false,
					meta: {
						type: "decoded",
						data: { name: "library.v1.json", mime: "application/json", modified: 0n, size: 1n, key: "k", version: 2 }
					}
				}
			})
		)

		expect(isDotFilenSystemFile(file)).toBe(true)
	})

	it("drops everything at logout, a lookup in flight included", async () => {
		const { getDotFilenSubdir, resetDotFilen, dotFilenSubdirUuid } = await fresh()
		const pending = getDotFilenSubdir("Videos")

		resetDotFilen()

		await expect(pending).resolves.toBe(SUBDIRS["Videos"])
		expect(dotFilenSubdirUuid("Videos")).toBeNull()

		await getDotFilenSubdir("Videos")

		expect(createDirectory).toHaveBeenCalledTimes(4)
	})

	it("a `.filen` event reaches every consumer whatever order the socket host registered them in", async () => {
		vi.resetModules()

		const socket = await import("@/lib/sdk/socket")
		const dotFilen = await import("@/lib/dotFilen")
		const playlists = await import("@/features/audio/lib/playlists")
		const audioHandlers = await import("@/features/audio/lib/socketHandlers")
		const library = await import("@/features/videos/lib/library")
		let deliver: (event: SocketEvent) => void = () => undefined

		subscribeToSocket.mockImplementation(onEvent => {
			deliver = onEvent

			return Promise.resolve()
		})

		// The socket host's order: playlists' handler ahead of the `.filen` one and the library's.
		const unregister = [
			audioHandlers.registerPlaylistSocketHandlers(),
			dotFilen.registerDotFilenSocketHandlers(),
			library.registerLibraryConfigSocketHandlers()
		]

		await socket.socketBridge.start()
		await playlists.getPlaylistsDirectoryUuid()
		await dotFilen.getDotFilenSubdir("Videos")
		await dotFilen.getDotFilenSubdir("Chat Uploads")

		const playlistsForgotten = vi.fn()

		dotFilen.onDotFilenSubdirForgotten("Playlists", playlistsForgotten)

		deliver(driveEvent({ type: "folderTrash", parent: testUuid("root"), uuid: DOT_FILEN }))

		expect(["Playlists", "Videos", "Chat Uploads"].map(name => dotFilen.dotFilenSubdirUuid(name))).toEqual([null, null, null])
		expect(dotFilen.isDotFilenSystemDir(DOT_FILEN)).toBe(false)
		expect(playlistsForgotten).toHaveBeenCalledTimes(1)

		// Nothing is saved into the trashed directories: every lookup starts over from `.filen`.
		createDirectory.mockClear()
		await dotFilen.getDotFilenSubdir("Chat Uploads")
		expect(createDirectory).toHaveBeenCalledTimes(2)

		for (const off of unregister) {
			off()
		}

		await socket.socketBridge.stop().catch(() => undefined)
	})

	it("forgetting one subdirectory keeps `.filen` while a sibling is resolved under it", async () => {
		const { getDotFilenSubdir, forgetDotFilenSubdir, isDotFilenSystemDir } = await fresh()

		await getDotFilenSubdir("Videos")
		await getDotFilenSubdir("Playlists")
		forgetDotFilenSubdir("Playlists")

		expect(isDotFilenSystemDir(DOT_FILEN)).toBe(true)

		createDirectory.mockClear()
		await getDotFilenSubdir("Playlists")
		expect(createDirectory).toHaveBeenCalledTimes(1)
	})

	it("a directory trashed elsewhere while a subdirectory lookup is out forgets nothing", async () => {
		const { getDotFilenSubdir, handleDotFilenDriveEvent, dotFilenSubdirUuid } = await fresh()
		const forgotten = vi.fn()

		await getDotFilenSubdir("Playlists")

		let release: () => void = () => undefined

		createDirectory.mockImplementationOnce(
			(_parent, name) =>
				new Promise(resolve => {
					release = () => {
						resolve(dir(SUBDIRS[name] ?? DOT_FILEN))
					}
				})
		)

		const { onDotFilenSubdirForgotten } = await import("@/lib/dotFilen")

		onDotFilenSubdirForgotten("Playlists", forgotten)

		const pending = getDotFilenSubdir("Videos")

		await Promise.resolve()
		// Its parent isn't `.filen`: it can't be the Videos directory being looked up.
		handleDotFilenDriveEvent(driveEvent({ type: "folderTrash", parent: testUuid("elsewhere"), uuid: testUuid("x") }))
		release()
		await pending

		expect(dotFilenSubdirUuid("Videos")).toBe(SUBDIRS["Videos"])
		expect(dotFilenSubdirUuid("Playlists")).toBe(SUBDIRS["Playlists"])
		expect(forgotten).not.toHaveBeenCalled()
	})

	it("only files that can't be photos count as app data, and one moved out stops counting", async () => {
		const { getDotFilenSubdir, handleDotFilenDriveEvent, isDotFilenSystemFile } = await fresh()
		const uploads = SUBDIRS["Chat Uploads"] ?? DOT_FILEN

		await getDotFilenSubdir("Chat Uploads")

		const json = driveFile(uploads, "state.json")
		const photo = driveFile(uploads, "holiday.jpg")

		handleDotFilenDriveEvent(driveEvent({ type: "fileNew", file: json }))
		handleDotFilenDriveEvent(driveEvent({ type: "fileNew", file: photo }))

		expect(isDotFilenSystemFile(json.uuid)).toBe(true)
		expect(isDotFilenSystemFile(photo.uuid)).toBe(false)

		handleDotFilenDriveEvent(driveEvent({ type: "fileMove", file: { ...json, parent: testUuid("elsewhere") } }))

		expect(isDotFilenSystemFile(json.uuid)).toBe(false)
	})

	it("a find creates nothing where nothing exists, and remembers no absence", async () => {
		const { findDotFilenSubdir, dotFilenSubdirUuid } = await fresh()

		directoryExists.mockImplementation(() => Promise.resolve(false))

		await expect(findDotFilenSubdir("Videos")).resolves.toBeNull()
		await expect(findDotFilenSubdir("Videos")).resolves.toBeNull()

		expect(createDirectory).not.toHaveBeenCalled()
		expect(listDirectory).not.toHaveBeenCalled()
		expect(directoryExists).toHaveBeenCalledTimes(2)
		expect(dotFilenSubdirUuid("Videos")).toBeNull()
	})

	it("a find stops at the first missing level", async () => {
		const { findDotFilenSubdir } = await fresh()

		listDirectory.mockImplementation(() => Promise.resolve({ dirs: [dir(testUuid("other"), "Other")], files: [] }))

		await expect(findDotFilenSubdir("Videos")).resolves.toBeNull()

		// `.filen` exists: its find-or-create returns it (creating nothing), and its listing has no Videos.
		expect(createDirectory).toHaveBeenCalledTimes(1)
		expect(createDirectory).toHaveBeenCalledWith(null, ".filen")
		expect(directoryExists).toHaveBeenCalledTimes(1)
		expect(listDirectory).toHaveBeenCalledExactlyOnceWith({ kind: "uuid", uuid: DOT_FILEN })
	})

	it("a find matches a subdirectory's name case-insensitively and never creates it", async () => {
		const { findDotFilenSubdir } = await fresh()

		listDirectory.mockImplementation(() => Promise.resolve({ dirs: [dir(SUBDIRS["Videos"] ?? testUuid("x"), "videos")], files: [] }))

		await expect(findDotFilenSubdir("Videos")).resolves.toBe(SUBDIRS["Videos"])
		expect(createDirectory).toHaveBeenCalledExactlyOnceWith(null, ".filen")
	})

	it("a found directory is remembered like a created one", async () => {
		const { findDotFilenSubdir, getDotFilenSubdir, dotFilenSubdirUuid } = await fresh()

		await expect(findDotFilenSubdir("Videos")).resolves.toBe(SUBDIRS["Videos"])
		expect(dotFilenSubdirUuid("Videos")).toBe(SUBDIRS["Videos"])

		const calls = createDirectory.mock.calls.length + directoryExists.mock.calls.length

		await expect(getDotFilenSubdir("Videos")).resolves.toBe(SUBDIRS["Videos"])
		await expect(findDotFilenSubdir("Videos")).resolves.toBe(SUBDIRS["Videos"])
		expect(createDirectory.mock.calls.length + directoryExists.mock.calls.length).toBe(calls)
	})

	it("a create joining a find that comes back empty creates", async () => {
		const { findDotFilenSubdir, getDotFilenSubdir } = await fresh()

		directoryExists.mockImplementation(() => Promise.resolve(false))

		const finding = findDotFilenSubdir("Videos")
		const creating = getDotFilenSubdir("Videos")

		await expect(finding).resolves.toBeNull()
		await expect(creating).resolves.toBe(SUBDIRS["Videos"])
		expect(createDirectory).toHaveBeenCalledTimes(2)
	})

	it("a find under way when the directories are forgotten isn't remembered", async () => {
		const { findDotFilenSubdir, dotFilenSubdirUuid, forgetDotFilenSubdir } = await fresh()
		const finding = findDotFilenSubdir("Videos")

		forgetDotFilenSubdir("Videos")

		await expect(finding).resolves.toBe(SUBDIRS["Videos"])
		expect(dotFilenSubdirUuid("Videos")).toBeNull()
	})
})
