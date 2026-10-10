import type { ReadRange } from "@filen/shared/videos"

// A Range request against a same-origin URL that answers 206 (the service worker's inline-preview
// route). Anything else is a failure: a 200 would be the whole file, so its body is never read.
export function urlReadRange(url: string): ReadRange {
	return async (start, end) => {
		const response = await fetch(url, { headers: { Range: `bytes=${String(start)}-${String(end - 1)}` } })

		if (response.status !== 206) {
			await response.body?.cancel().catch(() => undefined)

			throw new Error(`range read answered ${String(response.status)}`)
		}

		return new Uint8Array(await response.arrayBuffer())
	}
}
