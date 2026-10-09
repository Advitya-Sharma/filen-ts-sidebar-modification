import { type } from "arktype"
import { create } from "zustand"
import { kvPreference } from "@/lib/storage/preference"
import { log } from "@/lib/log"
import {
	DEFAULT_SUBTITLE_STYLE,
	SUBTITLE_BACKGROUNDS,
	SUBTITLE_COLORS,
	SUBTITLE_EDGES,
	SUBTITLE_FONTS,
	normalizeSubtitleStyle,
	type SubtitleStyle
} from "@/lib/media/subtitles/subtitleStyle.logic"

// The subtitle style on this device (kv media.subtitles.v1), shared by every video player.

export const SUBTITLE_STYLE_KEY = "media.subtitles.v1"

const subtitleStylePreference = kvPreference<SubtitleStyle>({
	key: SUBTITLE_STYLE_KEY,
	schema: type({
		size: "number",
		font: type.enumerated(...SUBTITLE_FONTS),
		color: type.enumerated(...SUBTITLE_COLORS),
		background: type.enumerated(...SUBTITLE_BACKGROUNDS),
		edge: type.enumerated(...SUBTITLE_EDGES),
		position: "number"
	}),
	fallback: DEFAULT_SUBTITLE_STYLE,
	normalize: normalizeSubtitleStyle
})

export const useSubtitleStyleStore = create<SubtitleStyle>(() => DEFAULT_SUBTITLE_STYLE)

let warm: Promise<void> | null = null
let userChanged = false

// Loads the stored style once per tab. Never rejects.
export function warmSubtitleStyle(): Promise<void> {
	warm ??= subtitleStylePreference.get().then(
		stored => {
			if (!userChanged) {
				useSubtitleStyleStore.setState(stored, true)
			}
		},
		(error: unknown) => {
			log.warn("media", "failed to load the subtitle style", error)
		}
	)

	return warm
}

export function setSubtitleStyle(patch: Partial<SubtitleStyle>): void {
	const next = normalizeSubtitleStyle({ ...useSubtitleStyleStore.getState(), ...patch })

	userChanged = true
	useSubtitleStyleStore.setState(next, true)
	subtitleStylePreference.set(next).catch((error: unknown) => {
		log.warn("media", "failed to persist the subtitle style", error)
	})
}

export function resetSubtitleStyle(): void {
	setSubtitleStyle(DEFAULT_SUBTITLE_STYLE)
}
