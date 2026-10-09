import { wipeThumbnailStore } from "@/features/drive/lib/thumbCache"
import { wipeArtStore } from "@/features/videos/lib/artStore"

// Every OPFS cache of images derived from the account's files: drive thumbnails and Videos artwork. Both
// run to the end; the first failure is reported.
export async function wipeImageCaches(): Promise<void> {
	const failed = (await Promise.allSettled([wipeThumbnailStore(), wipeArtStore()])).find(outcome => outcome.status === "rejected")

	if (failed !== undefined) {
		throw failed.reason
	}
}
