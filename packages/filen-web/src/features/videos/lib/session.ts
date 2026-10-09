import { probeService } from "@/features/videos/lib/probeService"
import { probeFacts } from "@/features/videos/lib/probeFacts"
import { artLoader } from "@/features/videos/lib/artwork"
import { frameGenerator, releaseFrameWorker } from "@/features/videos/lib/frames"
import { videoLibrary } from "@/features/videos/lib/index"
import { resetLibraryConfig } from "@/features/videos/lib/library"
import { resetNfoReader } from "@/features/videos/lib/nfoReader"

// Logout: nothing of the account's library, probes, artwork or frames stays in memory, and the frame worker
// goes once idle. The worker's index goes with its client.
export function resetVideoSession(): void {
	videoLibrary.reset()
	resetLibraryConfig()
	videoLibrary.hold()
	probeService.reset()
	probeFacts.reset()
	artLoader.reset()
	frameGenerator.reset()
	releaseFrameWorker()
	resetNfoReader()
}
