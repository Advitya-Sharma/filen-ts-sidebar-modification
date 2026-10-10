import type { GalleryItemTagged } from "@/components/drivePreview/gallery"
import type { DriveItemFileExtracted } from "@/types"
import type { FileSource } from "@/queries/fileSource"
import { getPreviewType, fileTypeExtension, isUncheckedTransportStream, type PreviewType } from "@/lib/previewType"

// The name a page picks its renderer and editor mode by: the one it opened with, for as long as it shows that
// file. A rename elsewhere must never swap an open editor, and its unsaved edits, for another renderer; saves
// still write under the file's current name.
export function galleryItemRenderName(item: GalleryItemTagged): string {
	return item.type === "drive" ? (item.openedName ?? item.data.data.decryptedMeta?.name ?? "") : item.data.name
}

function galleryItemMime(item: GalleryItemTagged): string | undefined {
	return item.type === "drive" ? item.data.data.decryptedMeta?.mime : undefined
}

function galleryItemVerdict(item: GalleryItemTagged) {
	return item.type === "drive" ? (item.verdict ?? null) : null
}

// The renderer a page uses: by its render name, the file's stored mime and the verdict the page holds, or the
// text viewer for a file opened as text.
export function galleryItemPreviewType(item: GalleryItemTagged): PreviewType {
	if (item.type === "drive" && item.asText) {
		return "text"
	}

	if (item.type === "drive" && item.unplayable) {
		return "unknown"
	}

	return getPreviewType(galleryItemRenderName(item), galleryItemMime(item), galleryItemVerdict(item))
}

// A `.ts`/`.mts`/`.m2ts` page whose renderer waits for a read of the file (transportStreamGate.tsx).
export function galleryItemNeedsSniff(item: GalleryItemTagged): boolean {
	return (
		item.type === "drive" &&
		!item.asText &&
		!item.unplayable &&
		item.verdict === undefined &&
		isUncheckedTransportStream(galleryItemRenderName(item), galleryItemMime(item), item.data.data.size, null)
	)
}

// The extension the text viewer picks its editor mode and highlighting by; "" for a file opened as text, which
// shows as plain text.
export function galleryItemTypeExtension(item: GalleryItemTagged): string {
	if (item.type === "drive" && item.asText) {
		return ""
	}

	return fileTypeExtension(galleryItemRenderName(item), galleryItemMime(item), galleryItemVerdict(item))
}

// `next` shown in place of `existing`: a rename, move, newer version or save of the same file. A file opened as
// text stays in the text viewer, and a page keeps the verdict it renders by. A page already showing a renderer
// keeps it: a `.ts` saved past the size a transport stream starts at stays in the text viewer.
export function galleryItemFollowing(existing: GalleryItemTagged, next: DriveItemFileExtracted): GalleryItemTagged {
	const following: Extract<GalleryItemTagged, { type: "drive" }> = {
		type: "drive",
		data: next,
		openedName: galleryItemRenderName(existing),
		...(existing.type === "drive" && existing.asText ? { asText: true } : {}),
		...(existing.type === "drive" && existing.verdict !== undefined ? { verdict: existing.verdict } : {}),
		...(existing.type === "drive" && existing.unplayable ? { unplayable: true } : {})
	}

	return following.verdict === undefined && !galleryItemNeedsSniff(existing) && galleryItemNeedsSniff(following)
		? { ...following, verdict: "notTransportStream" }
		: following
}

export function galleryItemFileSource(item: GalleryItemTagged): FileSource {
	// The held item goes by value so a cross-directory search hit (not in the global uuid cache) still resolves.
	return item.type === "drive"
		? { type: "drive", data: { uuid: item.data.data.uuid, item: item.data } }
		: { type: "external", data: item.data }
}
