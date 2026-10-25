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

// Whether `items` holds a file among `changed`: a scan with no kind resolved, one Set for a bulk change.
export function holdsAnyOf(items: readonly DriveItem[] | undefined, changed: readonly string[]): boolean {
	if (items === undefined || changed.length === 0) {
		return false
	}

	if (changed.length === 1) {
		const [uuid] = changed

		return items.some(item => item.data.uuid === uuid)
	}

	const uuids = new Set(changed)

	return items.some(item => uuids.has(item.data.uuid))
}

export type PhotoTimelineMonth = {
	year: number
	/** 0-based month */
	month: number
	/** Index of the newest photo of that month in the (capture-desc sorted) grid list. */
	index: number
}

export type PhotoTimeline = {
	/** Capture time (ms) per grid index, read by the scrubber bubble. */
	timestamps: Float64Array
	/** One scrubber marker per year, at its first (newest) photo. */
	yearMarkers: { label: string; index: number }[]
	/** Every month that holds photos, newest first — the targets of the "Jump to date" menu. */
	months: PhotoTimelineMonth[]
}

/**
 * One pass over the sorted grid list: capture timestamps plus the year/month boundaries.
 * `getTimestamp` is the same capture key the grid is SORTED by, so boundaries match what is on screen.
 */
export function buildPhotoTimeline<T>(items: readonly T[], getTimestamp: (item: T) => number): PhotoTimeline {
	const timestamps = new Float64Array(items.length)
	const yearMarkers: PhotoTimeline["yearMarkers"] = []
	const months: PhotoTimelineMonth[] = []
	let lastYear = Number.NaN
	let lastMonthKey = Number.NaN

	for (let i = 0; i < items.length; i++) {
		const item = items[i]

		if (item === undefined) {
			continue
		}

		const timestamp = getTimestamp(item)

		timestamps[i] = timestamp

		const date = new Date(timestamp)
		const year = date.getFullYear()
		const month = date.getMonth()
		const monthKey = year * 12 + month

		if (year !== lastYear) {
			yearMarkers.push({ label: String(year), index: i })

			lastYear = year
		}

		if (monthKey !== lastMonthKey) {
			months.push({ year, month, index: i })

			lastMonthKey = monthKey
		}
	}

	return { timestamps, yearMarkers, months }
}
