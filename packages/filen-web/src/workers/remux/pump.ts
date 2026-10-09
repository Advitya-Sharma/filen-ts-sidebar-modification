// The packet pump between a demuxed track and a fragmented-MP4 muxer. It reads one track in decode
// order from a starting point, decides which packets the output gets, and stops reading once the
// muxer has emitted media up to the buffer horizon the player asked for, so a paused or well-buffered
// player costs no reads.
//
// Timestamps stay on the source timeline: playback starts at a key packet K, nothing presented before
// K is written, and the player seeks its element to K. Shifting the output to start at zero would
// need an edit list, and Chromium applies a fragmented file's edit list before MSE trims at
// appendWindowStart, dropping the whole first GOP.

export interface PumpItem {
	// Presentation timestamp, seconds.
	timestamp: number
	duration: number
	type: "key" | "delta"
}

export interface PumpReader<T extends PumpItem> {
	// The item playback from `time` starts with: for video the key packet at or before `time`, for
	// audio the packet covering it. Null for a track with nothing there.
	first(time: number): Promise<T | null>
	// The item after `item` in decode order, null at the end of the track.
	next(item: T): Promise<T | null>
}

export interface PumpWriter<T extends PumpItem> {
	write(item: T): Promise<void>
	// The track ended: flush whatever the muxer still holds.
	end(): Promise<void>
}

export interface PumpOptions<T extends PumpItem = PumpItem> {
	// The earliest timestamp written. -Infinity keeps a video key packet whatever its timestamp; 0 drops
	// an audio track's priming packets, which carry negative timestamps.
	floor: number
	// Called once with the start timestamp, before anything is written.
	onStart?: (startTimestamp: number) => void
	// Called for every item read but not written (dropped before the start, or read as the pump was
	// cancelled), so one holding decoded media can free it.
	discard?: (item: T) => void
}

export type PumpResult = { type: "ended"; startTimestamp: number | null } | { type: "cancelled" }

// Whether an item read after the start is written. Pictures presented before the key packet (an open
// GOP's leading pictures) reference frames before it and cannot be decoded.
export function keepItem(item: PumpItem, startTimestamp: number): boolean {
	return item.timestamp >= startTimestamp
}

// Shared between one pump and whatever tells it how far to read: the player raises the horizon, the
// muxer reports how far it has emitted, and a seek or track switch cancels.
export class PumpGate {
	private horizonValue: number
	private flushedUntilValue = Number.NEGATIVE_INFINITY
	private cancelledValue = false
	private wake: (() => void) | null = null

	public constructor(horizon: number) {
		this.horizonValue = horizon
	}

	public get horizon(): number {
		return this.horizonValue
	}

	public get flushedUntil(): number {
		return this.flushedUntilValue
	}

	// A method rather than a getter: it changes across awaits.
	public isCancelled(): boolean {
		return this.cancelledValue
	}

	public setHorizon(horizon: number): void {
		this.horizonValue = horizon
		this.release()
	}

	// The muxer has emitted every packet presented before `timestamp`.
	public noteFlushed(timestamp: number): void {
		this.flushedUntilValue = Math.max(this.flushedUntilValue, timestamp)
	}

	public cancel(): void {
		this.cancelledValue = true
		this.release()
	}

	// Resolves true once there is room to read on (emitted media ends before the horizon), false when
	// the pump was cancelled meanwhile.
	public async waitForRoom(): Promise<boolean> {
		while (!this.cancelledValue && this.flushedUntilValue >= this.horizonValue) {
			await new Promise<void>(resolve => {
				this.wake = resolve
			})
		}

		return !this.cancelledValue
	}

	private release(): void {
		const wake = this.wake

		this.wake = null
		wake?.()
	}
}

export async function runPump<T extends PumpItem>(
	reader: PumpReader<T>,
	writer: PumpWriter<T>,
	gate: PumpGate,
	time: number,
	options: PumpOptions<T>
): Promise<PumpResult> {
	let item = await reader.first(time)

	if (gate.isCancelled()) {
		if (item !== null) {
			options.discard?.(item)
		}

		return { type: "cancelled" }
	}

	if (item === null) {
		await writer.end()

		return { type: "ended", startTimestamp: null }
	}

	const startTimestamp = Math.max(options.floor, item.timestamp)

	options.onStart?.(startTimestamp)

	while (item !== null) {
		if (keepItem(item, startTimestamp)) {
			await writer.write(item)
		} else {
			options.discard?.(item)
		}

		// Checked before reading on, so a pump at its horizon has read nothing past what it wrote.
		if (!(await gate.waitForRoom())) {
			return { type: "cancelled" }
		}

		item = await reader.next(item)

		if (gate.isCancelled()) {
			if (item !== null) {
				options.discard?.(item)
			}

			return { type: "cancelled" }
		}
	}

	await writer.end()

	return { type: "ended", startTimestamp }
}
