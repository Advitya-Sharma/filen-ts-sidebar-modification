import { probeService } from "@/features/videos/lib/probeService"
import { artLoader } from "@/features/videos/lib/artwork"
import { frameGenerator, releaseFrameWorker } from "@/features/videos/lib/frames"

// Logout: nothing of the account's probes, artwork or frames stays in memory, and the frame worker goes
// once idle.
export function resetVideoSession(): void {
	probeService.reset()
	artLoader.reset()
	frameGenerator.reset()
	releaseFrameWorker()
}
