import { describe, expect, it, vi } from "vitest"

vi.mock("@/constants", async () => await import("@/tests/mocks/constants"))

import {
	galleryItemFollowing,
	galleryItemNeedsSniff,
	galleryItemPreviewType,
	galleryItemRenderName,
	galleryItemTypeExtension
} from "@/components/drivePreview/galleryRenderName"
import type { GalleryItemTagged } from "@/components/drivePreview/gallery"
import type { DriveItemFileExtracted } from "@/types"
import fileKindVerdicts from "@/lib/fileKindVerdicts"

function file(uuid: string, name: string, mime?: string, size = 0): DriveItemFileExtracted {
	return { type: "file", data: { uuid, size: BigInt(size), decryptedMeta: { name, mime } } } as unknown as DriveItemFileExtracted
}

const STREAM_SIZED = 2 * 1024 * 1024

describe("a gallery page's render name", () => {
	it("is the file's own name until the file is replaced", () => {
		expect(galleryItemRenderName({ type: "drive", data: file("v1", "todo.txt") })).toBe("todo.txt")
	})

	it("keeps the name the page opened with across a rename, so an open editor is never swapped out", () => {
		const opened: GalleryItemTagged = { type: "drive", data: file("v1", "todo.txt") }
		const renamed = galleryItemFollowing(opened, file("v1", "todo"))

		expect(renamed.type === "drive" ? renamed.data.data.decryptedMeta?.name : undefined).toBe("todo")
		expect(galleryItemRenderName(renamed)).toBe("todo.txt")
		// A save or newer version after it keeps it too.
		expect(galleryItemRenderName(galleryItemFollowing(renamed, file("v2", "todo.bak")))).toBe("todo.txt")
	})
})

describe("a gallery page's renderer", () => {
	it("reads the render name, then the file's stored mime", () => {
		expect(galleryItemPreviewType({ type: "drive", data: file("v1", "clip", "video/mp4") })).toBe("video")
		expect(galleryItemPreviewType({ type: "drive", data: file("v1", "Makefile") })).toBe("code")
		expect(galleryItemPreviewType({ type: "drive", data: file("v1", "data.bin", "application/octet-stream") })).toBe("unknown")
	})

	it("is the text viewer for a file opened as text, whatever its type", () => {
		expect(galleryItemPreviewType({ type: "drive", data: file("v1", "data.bin"), asText: true })).toBe("text")
		expect(galleryItemPreviewType({ type: "drive", data: file("v1", "photo.jpg"), asText: true })).toBe("text")
	})

	it("stays the text viewer across a rename or a newer version", () => {
		const opened: GalleryItemTagged = { type: "drive", data: file("v1", "data.bin"), asText: true }
		const followed = galleryItemFollowing(opened, file("v2", "data.dat"))

		expect(followed).toMatchObject({ asText: true })
		expect(galleryItemPreviewType(followed)).toBe("text")
		expect(galleryItemFollowing({ type: "drive", data: file("v1", "a.txt") }, file("v2", "a.txt"))).not.toHaveProperty("asText")
	})

	it("picks the editor mode by the type extension, plain for a file opened as text", () => {
		expect(galleryItemTypeExtension({ type: "drive", data: file("v1", ".bashrc") })).toBe("sh")
		expect(galleryItemTypeExtension({ type: "drive", data: file("v1", "README.MD") })).toBe("md")
		expect(galleryItemTypeExtension({ type: "drive", data: file("v1", "data.bin"), asText: true })).toBe("")
	})
})

describe("a gallery page's content verdict", () => {
	it("is the one the page holds, never the store's", () => {
		fileKindVerdicts.record("held", "audio")

		expect(galleryItemPreviewType({ type: "drive", data: file("held", "voice.mp4") })).toBe("video")
		expect(galleryItemPreviewType({ type: "drive", data: file("held", "voice.mp4"), verdict: "audio" })).toBe("audio")
		expect(galleryItemTypeExtension({ type: "drive", data: file("held", "voice.mp4"), verdict: "audio" })).toBe("m4a")
	})

	it("holds an unchecked .ts or .mts page for its read, and nothing else", () => {
		expect(galleryItemNeedsSniff({ type: "drive", data: file("v1", "capture.ts", undefined, STREAM_SIZED) })).toBe(true)
		// iOS (the mocked platform) shows neither a stream nor an .mts: nothing to read for.
		expect(galleryItemNeedsSniff({ type: "drive", data: file("v1", "camera.mts", undefined, STREAM_SIZED) })).toBe(false)
		expect(galleryItemNeedsSniff({ type: "drive", data: file("v1", "main.ts", undefined, 1024) })).toBe(false)
		expect(galleryItemNeedsSniff({ type: "drive", data: file("v1", "lib.d.ts", undefined, STREAM_SIZED) })).toBe(false)
		expect(
			galleryItemNeedsSniff({ type: "drive", data: file("v1", "capture.ts", undefined, STREAM_SIZED), verdict: "notTransportStream" })
		).toBe(false)
		expect(galleryItemNeedsSniff({ type: "drive", data: file("v1", "capture.ts", undefined, STREAM_SIZED), asText: true })).toBe(false)
		expect(galleryItemNeedsSniff({ type: "drive", data: file("v1", "clip.mp4", undefined, STREAM_SIZED) })).toBe(false)
	})

	it("follows the file across a rename or a save", () => {
		const opened: GalleryItemTagged = { type: "drive", data: file("v1", "voice.mp4"), verdict: "audio" }

		expect(galleryItemFollowing(opened, file("v2", "voice.mp4"))).toMatchObject({ verdict: "audio" })
		expect(galleryItemFollowing({ type: "drive", data: file("v1", "a.txt") }, file("v2", "a.txt"))).not.toHaveProperty("verdict")
	})

	it("keeps a .ts page in the text viewer when a save grows it past the size a stream starts at", () => {
		const opened: GalleryItemTagged = { type: "drive", data: file("v1", "main.ts", undefined, 1024) }
		const saved = galleryItemFollowing(opened, file("v2", "main.ts", undefined, STREAM_SIZED))

		expect(saved).toMatchObject({ verdict: "notTransportStream" })
		expect(galleryItemNeedsSniff(saved)).toBe(false)
		expect(galleryItemPreviewType(saved)).toBe("code")

		// A page still behind its gate stays there.
		const gated: GalleryItemTagged = { type: "drive", data: file("v1", "capture.ts", undefined, STREAM_SIZED) }

		expect(galleryItemNeedsSniff(galleryItemFollowing(gated, file("v1", "capture.ts", undefined, STREAM_SIZED)))).toBe(true)
	})

	it("shows no preview for a stream the gate found unplayable, and keeps that across a save", () => {
		const unplayable: GalleryItemTagged = { type: "drive", data: file("v1", "clip.m2ts", undefined, STREAM_SIZED), unplayable: true }

		expect(galleryItemPreviewType(unplayable)).toBe("unknown")
		expect(galleryItemNeedsSniff(unplayable)).toBe(false)
		expect(galleryItemFollowing(unplayable, file("v1", "renamed.m2ts", undefined, STREAM_SIZED))).toMatchObject({ unplayable: true })
	})
})
