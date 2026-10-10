import { vi, describe, it, expect } from "vitest"

vi.mock("@/constants", () => ({
	EXPO_IMAGE_SUPPORTED_EXTENSIONS: new Set<string>([".jpg", ".jpeg", ".png", ".gif", ".webp", ".avif", ".heic", ".heif", ".svg"]),
	EXPO_VIDEO_SUPPORTED_EXTENSIONS: new Set<string>([".mp4", ".mov", ".m4v", ".3gp", ".webm", ".mkv"]),
	EXPO_AUDIO_SUPPORTED_EXTENSIONS: new Set<string>([".mp3", ".m4a", ".aac", ".wav", ".ogg", ".flac"])
}))

// driveSelectors pulls in the serializer, which would otherwise load uniffi-bindgen-react-native.
vi.mock("@/lib/serializer", () => ({
	serialize: (value: unknown) => JSON.stringify(value)
}))

import { holdsAnyOf, isPhotoGridItem } from "@/features/photos/utils"
import { type DriveItem } from "@/types"
import fileKindVerdicts from "@/lib/fileKindVerdicts"

function makeItem({ type, name, size = 0 }: { type: string; name: string | null; size?: number }): DriveItem {
	return {
		type,
		data: {
			uuid: `uuid-${name ?? ""}`,
			size: BigInt(size),
			decryptedMeta: name === null ? null : { name }
		}
	} as unknown as DriveItem
}

describe("isPhotoGridItem", () => {
	it("accepts a supported image file", () => {
		expect(isPhotoGridItem(makeItem({ type: "file", name: "photo.jpg" }))).toBe(true)
	})

	it("accepts a video file", () => {
		expect(isPhotoGridItem(makeItem({ type: "file", name: "clip.mp4" }))).toBe(true)
	})

	it("accepts a RAW camera file (previewed via the SDK-extracted JPEG)", () => {
		expect(isPhotoGridItem(makeItem({ type: "file", name: "shot.cr2" }))).toBe(true)
	})

	it("accepts an svg (image-equivalent, renders via react-native-svg)", () => {
		expect(isPhotoGridItem(makeItem({ type: "file", name: "logo.svg" }))).toBe(true)
	})

	it("accepts shared file types", () => {
		expect(isPhotoGridItem(makeItem({ type: "sharedFile", name: "photo.png" }))).toBe(true)
		expect(isPhotoGridItem(makeItem({ type: "sharedRootFile", name: "photo.png" }))).toBe(true)
	})

	it("matches extensions case-insensitively", () => {
		expect(isPhotoGridItem(makeItem({ type: "file", name: "IMG.HEIC" }))).toBe(true)
		expect(isPhotoGridItem(makeItem({ type: "file", name: "photo.JPG" }))).toBe(true)
		expect(isPhotoGridItem(makeItem({ type: "file", name: "photo.AVIF" }))).toBe(true)
	})

	it("rejects an image whose extension expo-image cannot render", () => {
		expect(isPhotoGridItem(makeItem({ type: "file", name: "scan.tiff" }))).toBe(false)
	})

	it("rejects directories", () => {
		expect(isPhotoGridItem(makeItem({ type: "directory", name: "vacation.jpg" }))).toBe(false)
	})

	it("rejects items without decrypted metadata", () => {
		expect(isPhotoGridItem(makeItem({ type: "file", name: null }))).toBe(false)
	})

	it("rejects non-image / non-video files", () => {
		expect(isPhotoGridItem(makeItem({ type: "file", name: "document.pdf" }))).toBe(false)
		expect(isPhotoGridItem(makeItem({ type: "file", name: "song.mp3" }))).toBe(false)
	})

	it("filters a mixed listing down to the grid items", () => {
		const items = [
			makeItem({ type: "file", name: "photo.jpg" }),
			makeItem({ type: "file", name: "clip.mp4" }),
			makeItem({ type: "file", name: "shot.dng" }),
			makeItem({ type: "directory", name: "album.jpg" }),
			makeItem({ type: "file", name: "scan.tiff" }),
			makeItem({ type: "file", name: "document.pdf" }),
			makeItem({ type: "sharedFile", name: "shared.png" }),
			makeItem({ type: "file", name: null })
		]

		expect(items.filter(item => isPhotoGridItem(item)).map(item => item.data.decryptedMeta?.name)).toEqual([
			"photo.jpg",
			"clip.mp4",
			"shot.dng",
			"shared.png"
		])
	})

	describe("content verdicts", () => {
		const STREAM_SIZED = 2 * 1024 * 1024

		it("drops a video a player found to hold audio alone", () => {
			expect(isPhotoGridItem(makeItem({ type: "file", name: "voice.mp4" }), "audio")).toBe(false)
			expect(isPhotoGridItem(makeItem({ type: "file", name: "clip.mp4" }), null)).toBe(true)
		})

		it("takes a .ts only once a read found it to be a stream the platform plays", () => {
			const capture = makeItem({ type: "file", name: "capture.ts", size: STREAM_SIZED })

			expect(isPhotoGridItem(capture, null)).toBe(false)
			expect(isPhotoGridItem(capture, "notTransportStream")).toBe(false)
			// iOS (the mocked platform): AVPlayer plays no progressive transport stream.
			expect(isPhotoGridItem(capture, "transportStream")).toBe(false)
		})

		it("reads the store's verdict when none is passed", () => {
			fileKindVerdicts.record("uuid-voice.m4v.mp4", "audio")

			expect(isPhotoGridItem(makeItem({ type: "file", name: "voice.m4v.mp4" }))).toBe(false)
		})
	})
})

describe("isPhotoGridItem on Android", () => {
	it("takes a .ts a read found to be a transport stream, which Media3 plays", async () => {
		vi.resetModules()

		const { Platform } = await import("react-native")

		Platform.OS = "android"

		try {
			const { isPhotoGridItem: isAndroidPhoto } = await import("@/features/photos/utils")
			const capture = {
				type: "file",
				data: { uuid: "c", size: BigInt(2 * 1024 * 1024), decryptedMeta: { name: "capture.ts" } }
			} as unknown as DriveItem

			expect(isAndroidPhoto(capture, "transportStream")).toBe(true)
			expect(isAndroidPhoto(capture, null)).toBe(false)
		} finally {
			Platform.OS = "ios"
			vi.resetModules()
		}
	})
})

describe("holdsAnyOf", () => {
	const items = [makeItem({ type: "file", name: "a.mp4" }), makeItem({ type: "file", name: "b.jpg" })]

	it("finds a changed file in the listing, alone or among several", () => {
		expect(holdsAnyOf(items, ["uuid-a.mp4"])).toBe(true)
		expect(holdsAnyOf(items, ["elsewhere", "uuid-b.jpg"])).toBe(true)
	})

	it("is false for files outside the listing, no change, or no listing yet", () => {
		expect(holdsAnyOf(items, ["elsewhere"])).toBe(false)
		expect(holdsAnyOf(items, ["elsewhere", "other"])).toBe(false)
		expect(holdsAnyOf(items, [])).toBe(false)
		expect(holdsAnyOf(undefined, ["uuid-a.mp4"])).toBe(false)
	})
})
