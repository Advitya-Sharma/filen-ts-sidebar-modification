import { selectAudioTrack } from "@filen/shared/videos"
import type { VideoProbe } from "@/lib/media/probe/types"

// The audio track an item starts on, the same for the player and the detail page's select and Direct/Remux badge:
// the series' remembered language, then the user's languages and the file's default, preferring a track this browser
// plays without converting.

const NATIVE_AUDIO: ReadonlySet<string> = new Set(["aac", "mp3", "opus", "vorbis", "flac"])

export function defaultAudioIndex({
	probe,
	rememberedLang,
	userLanguages
}: {
	probe: Pick<VideoProbe, "audio">
	rememberedLang: string | null
	userLanguages: readonly string[]
}): number {
	return selectAudioTrack(
		probe.audio.map(track => ({ lang: track.lang, default: track.def })),
		{
			userLanguages,
			rememberedLang,
			playable: (_track, index) => {
				const codec = probe.audio[index]?.codec ?? null

				return codec !== null && NATIVE_AUDIO.has(codec)
			}
		}
	)
}
