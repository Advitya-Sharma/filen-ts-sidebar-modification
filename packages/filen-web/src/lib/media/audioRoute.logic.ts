// How an audio file reaches an <audio> element: as it is, through the service worker's range route
// ("direct"), or decoded in the remux worker into a WAV stream the same route serves ("transcode"). Either
// way the element plays a URL, so background playback, the lock screen, prefetch and gapless switching work
// the same. Where no service worker serves the page, both go through MediaSource instead.

export type AudioRoute = "direct" | "transcode"

// Files no browser plays as they are: Matroska audio (MKV is never played direct), and bare AC-3, E-AC-3
// and DTS streams.
const TRANSCODED_EXTENSIONS: ReadonlySet<string> = new Set(["mka", "ac3", "eac3", "ec3", "dts"])

// `extension` is the file's type extension (an `.mkv` holding only audio reads as `mka`); `contentType` its
// allowlisted inline type, null when it has none the route may serve it as; `demoted` once the browser
// failed to decode the file as it is (an ALAC `.m4a` outside Safari).
export function audioRoute(params: { extension: string; contentType: string | null; demoted: boolean }): AudioRoute {
	return params.demoted || params.contentType === null || TRANSCODED_EXTENSIONS.has(params.extension) ? "transcode" : "direct"
}

// Files this browser failed to decode as they are, for the rest of the session: their next play goes
// straight to the decoded stream. Bounded, oldest out first.
export class AudioDemotions {
	private readonly uuids = new Set<string>()
	private readonly capacity: number

	public constructor(capacity: number) {
		this.capacity = capacity
	}

	public add(uuid: string): void {
		this.uuids.delete(uuid)
		this.uuids.add(uuid)

		for (const oldest of this.uuids) {
			if (this.uuids.size <= this.capacity) {
				break
			}

			this.uuids.delete(oldest)
		}
	}

	public has(uuid: string): boolean {
		return this.uuids.has(uuid)
	}

	public clear(): void {
		this.uuids.clear()
	}
}

export const audioDemotions = new AudioDemotions(1024)
