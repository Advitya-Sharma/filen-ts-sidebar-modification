import { type DriveItem, type DriveItemFileExtracted } from "@/types"
import { getDriveItemPreviewType, isImagePreviewType } from "@/lib/previewType"
import { isFileItem } from "@/features/drive/driveSelectors"
import { type ContentVerdict } from "@filen/shared"

/**
 * Whether a drive item belongs in the photos grid: a decrypted image (incl. svg / RAW) or video file, by its
 * content verdict too: an `.mp4` a player found to hold audio alone is no photo, and a `.ts`/`.mts` is one only
 * once a read found it to be a transport stream the platform plays. `verdict` as getDriveItemPreviewType takes
 * it. The photos gallery is handed the grid list as-is, so this is the only filter either applies.
 */
export function isPhotoGridItem(item: DriveItem, verdict?: ContentVerdict | null): item is DriveItemFileExtracted {
	if (!isFileItem(item) || !item.data.decryptedMeta) {
		return false
	}

	const previewType = getDriveItemPreviewType(item, verdict)

	return isImagePreviewType(previewType) || previewType === "video"
}
