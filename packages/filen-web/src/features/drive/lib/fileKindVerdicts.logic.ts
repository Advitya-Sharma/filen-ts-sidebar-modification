import type { ContentVerdict } from "@filen/shared"
import type { VideoProbe } from "@filen/shared/videos"

// What a container probe of a file says about its kind: a transport stream or not (for a `.ts`), else
// whether it holds a picture. Null when it says nothing (a probe that found no tracks).
export function probeVerdict(
	probe: Pick<VideoProbe, "container" | "video" | "audio" | "partial"> | null,
	transportStreamName: boolean
): ContentVerdict | null {
	if (transportStreamName) {
		return probe?.container === "ts" ? "transportStream" : "notTransportStream"
	}

	if (probe === null) {
		return null
	}

	if (probe.video !== null) {
		return "video"
	}

	// Its picture track may lie past what a partial walk read.
	return !probe.partial && probe.audio.length > 0 ? "audio" : null
}
