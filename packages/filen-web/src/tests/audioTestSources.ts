import type { TrackSource } from "@/features/audio/lib/engine"

// Track sources for the engine suites: a URL the element plays, whose release is recorded in `released`.
export function urlSource(url: string, released: string[] = []): TrackSource {
	return {
		media: { kind: "url", url },
		classify: () => Promise.resolve("other"),
		release: () => {
			released.push(url)
		}
	}
}

// What a fake element was loaded with: the URL, or "session" for a MediaSource source.
export function loadedUrl(source: TrackSource): string {
	return source.media.kind === "url" ? source.media.url : "session"
}
