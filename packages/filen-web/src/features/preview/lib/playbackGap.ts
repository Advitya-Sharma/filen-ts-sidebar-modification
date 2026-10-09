import type { TrackKind } from "@/features/preview/lib/containerTracks"

// Whether a playing element decoded any audio, from whichever non-standard counter this engine has;
// null where it has none.
function decodedAudio(video: HTMLVideoElement): boolean | null {
	if ("webkitAudioDecodedByteCount" in video && typeof video.webkitAudioDecodedByteCount === "number") {
		return video.webkitAudioDecodedByteCount > 0
	}

	if ("mozHasAudio" in video && typeof video.mozHasAudio === "boolean") {
		return video.mozHasAudio
	}

	if ("audioTracks" in video && typeof video.audioTracks === "object" && video.audioTracks !== null && "length" in video.audioTracks) {
		return video.audioTracks.length !== 0
	}

	return null
}

// The track kinds a playing element shows no sign of decoding. Only a suspicion: a file without an
// audio track decodes no audio either, which the probe settles (previewVideo.ts).
export function undecodedKinds(video: HTMLVideoElement): TrackKind[] {
	const kinds: TrackKind[] = []

	if (video.videoWidth === 0 || video.getVideoPlaybackQuality().totalVideoFrames === 0) {
		kinds.push("video")
	}

	if (decodedAudio(video) === false) {
		kinds.push("audio")
	}

	return kinds
}
