import { type ContentVerdict } from "@filen/shared"
import type { SourceLoadEventPayload } from "expo-video"
import { recordContentVerdict, type VerdictFile } from "@/lib/previewType"

// What a native video player that loaded a file says about its kind, at no cost: the tracks it listed. Only
// positive evidence counts: a listed picture track makes it a video, listed audio tracks without one make it
// audio. Anything else (no track listed, as iOS does for a lone audio track outside a media selection group, or
// a failed load, which both platforms report with empty lists) says nothing: a video taken for audio would open
// in the audio player from then on.
export function playerTrackVerdict(
	loaded: Pick<SourceLoadEventPayload, "availableVideoTracks" | "availableAudioTracks">,
	failed: boolean
): ContentVerdict | null {
	if (failed) {
		return null
	}

	if (loaded.availableVideoTracks.length > 0) {
		return "video"
	}

	return loaded.availableAudioTracks.length > 0 ? "audio" : null
}

// Keeps what a player's load says about `file` (recordContentVerdict keeps only a verdict that changes its kind).
export function recordPlayerVerdict(
	file: VerdictFile,
	loaded: Pick<SourceLoadEventPayload, "availableVideoTracks" | "availableAudioTracks">,
	failed: boolean
): void {
	const verdict = playerTrackVerdict(loaded, failed)

	if (verdict !== null) {
		recordContentVerdict(file, verdict)
	}
}
