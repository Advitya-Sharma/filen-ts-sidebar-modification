import { wipeThumbnailStore } from "@/features/drive/lib/thumbCache"
import { wipeArtStore } from "@/features/videos/lib/artStore"
import { wipeVideoIndexStore } from "@/features/videos/lib/indexStore"

// Every OPFS cache derived from the account's files: drive thumbnails, Videos artwork and the Videos library
// index. All run to the end; the first failure is reported.
export async function wipeImageCaches(): Promise<void> {
	const failed = (await Promise.allSettled([wipeThumbnailStore(), wipeArtStore(), wipeVideoIndexStore()])).find(
		outcome => outcome.status === "rejected"
	)

	if (failed !== undefined) {
		throw failed.reason
	}
}
