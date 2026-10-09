// Key packet timestamps a pump has read, in order. A file without an index (Matroska without Cues) can
// only be seeked within what has been read, and the demuxer's own lookup there misses a key packet that
// lies in an earlier cluster than the one it starts searching from; the exact timestamp of a known key
// packet always finds it.
export class KeyIndex {
	private readonly times: number[] = []

	public add(timestamp: number): void {
		const last = this.times.at(-1)

		if (last === undefined || timestamp > last) {
			this.times.push(timestamp)

			return
		}

		const at = this.insertionPoint(timestamp)

		if (this.times[at] !== timestamp) {
			this.times.splice(at, 0, timestamp)
		}
	}

	// The last known key packet at or before `time`, null when none is known.
	public atOrBefore(time: number): number | null {
		const at = this.insertionPoint(time)
		const exact = this.times[at]

		if (exact === time) {
			return exact
		}

		return this.times[at - 1] ?? null
	}

	// First index whose timestamp is >= `time`.
	private insertionPoint(time: number): number {
		let low = 0
		let high = this.times.length

		while (low < high) {
			const mid = (low + high) >>> 1

			if ((this.times[mid] ?? Number.POSITIVE_INFINITY) < time) {
				low = mid + 1
			} else {
				high = mid
			}
		}

		return low
	}
}

// How much of a file without an index a seek may scan past what has been read: the demuxer finds a far
// position only by reading every cluster before it. About a second and a half of download at broadband
// speed; further seeks stop at the end of what has been scanned.
export const MAX_SEEK_SCAN_BYTES = 16 * 1_048_576

// Where a seek into a file without an index starts: `time` itself when the clusters still to scan up to
// it fit the budget (estimated from the file's average bitrate, or the whole file when its duration is
// unknown), else the end of what has been scanned.
export function seekReach(params: { time: number; scannedUntil: number; size: number; duration: number }): {
	time: number
	clamped: boolean
} {
	const { time, scannedUntil, size, duration } = params

	if (time <= scannedUntil) {
		return { time, clamped: false }
	}

	const toScan = duration > 0 ? ((time - scannedUntil) * size) / duration : size

	return toScan <= MAX_SEEK_SCAN_BYTES ? { time, clamped: false } : { time: scannedUntil, clamped: true }
}
