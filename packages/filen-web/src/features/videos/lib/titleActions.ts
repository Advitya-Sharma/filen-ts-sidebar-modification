import { videoLibrary } from "@/features/videos/lib/index"
import { narrowItem, type DriveItem } from "@/features/drive/lib/item"
import { defaultRevealDeps, runOpenContainingDirectory } from "@/features/drive/lib/reveal"
import { startDownloads } from "@/features/drive/lib/download"
import type { DriveNavigationTarget } from "@/features/drive/lib/navigate"

// The detail pages' More menu: the drive file behind a video (rebuilt by the index's worker, no request), shown in
// Cloud Drive or downloaded through the drive's own paths.

async function driveItem(stableUuid: string): Promise<DriveItem | null> {
	const [file] = await videoLibrary.playableFiles([stableUuid])

	return file === undefined ? null : narrowItem(file)
}

// Opens the directory holding the file with the file revealed; false when the index no longer has it.
export async function revealVideoFile(stableUuid: string, navigate: (target: DriveNavigationTarget) => void): Promise<boolean> {
	const item = await driveItem(stableUuid)

	if (item === null) {
		return false
	}

	await runOpenContainingDirectory(defaultRevealDeps, item, navigate)

	return true
}

export async function downloadVideoFile(stableUuid: string): Promise<boolean> {
	const item = await driveItem(stableUuid)

	if (item === null) {
		return false
	}

	await startDownloads([item])

	return true
}
