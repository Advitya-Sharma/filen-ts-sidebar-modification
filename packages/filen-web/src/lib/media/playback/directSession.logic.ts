// The decisions behind a direct (service worker) stream's keepalive and recovery, as pure functions.

// A worker kept alive for a player paused longer than this is let go; a resume after it recovers instead.
export const KEEPALIVE_IDLE_PAUSE_MS = 30 * 60_000
// After a recovery, another one is tried only once playback has gone on this long.
export const RECOVERY_COOLDOWN_MS = 60_000
// A recovery that has not brought the stream back by then has failed (measured: 1.1-1.7 s, most of it a
// cold worker compiling its wasm).
export const RECOVERY_TIMEOUT_MS = 10_000

const MEDIA_ERR_ABORTED = 1
const MEDIA_ERR_NETWORK = 2
const MEDIA_ERR_DECODE = 3

// What a MediaError on a direct stream calls for.
// - "recover": the stream broke (the service worker was terminated and lost its registration, a download
//   failed), so register the file again and reload at the same position.
// - "verify": MEDIA_ERR_SRC_NOT_SUPPORTED before any metadata, which is a format verdict or a fetch that
//   failed at once (Chrome reports both so): asking the worker whether it still serves the stream tells them
//   apart.
// - "demote": this browser cannot play it.
// - "fail": the stream broke again during or right after a recovery; never a verdict on the browser.
export function directErrorAction(params: {
	code: number
	hadMetadata: boolean
	// A recovery's reload has not played yet.
	recovering: boolean
	now: number
	lastRecoveryAt: number | null
}): "ignore" | "recover" | "verify" | "demote" | "fail" {
	const { code, hadMetadata, recovering, now, lastRecoveryAt } = params

	if (code === MEDIA_ERR_ABORTED) {
		return "ignore"
	}

	if (recovering) {
		return "fail"
	}

	if (code === MEDIA_ERR_DECODE) {
		return "demote"
	}

	if (code === MEDIA_ERR_NETWORK || hadMetadata) {
		return mayRecover(now, lastRecoveryAt) ? "recover" : "fail"
	}

	return "verify"
}

// Whether another recovery may start: none yet, or the last one is RECOVERY_COOLDOWN_MS old.
export function mayRecover(now: number, lastRecoveryAt: number | null): boolean {
	return lastRecoveryAt === null || now - lastRecoveryAt >= RECOVERY_COOLDOWN_MS
}

// A playhead that has not moved for this long while it should (playing or seeking) has its stream asked
// whether it is still there. WebKit retries a range request the service worker answers 404 for ever, with no
// MediaError and readyState left at HAVE_FUTURE_DATA, so a frozen playhead is the only sign of it.
export const STREAM_STALL_MS = 8_000
export const STREAM_STALL_POLL_MS = 1_000

// The playhead should be moving: told to play or seeking, and not at the end.
export function expectsProgress(element: { paused: boolean; seeking: boolean; ended: boolean }): boolean {
	return !element.ended && (!element.paused || element.seeking)
}

// What a stall that lasted STREAM_STALL_MS (or the worker's report that the stream is gone) calls for, once
// the worker said whether it still serves the stream: one it serves is only slow (wait on); one it lost or
// never had is registered again, unless a recovery ran within RECOVERY_COOLDOWN_MS, when the route is given
// up on for the next rung. A worker that could not be asked ("unknown") is waited on, as before it could be.
export function stallAction(params: {
	status: "present" | "missing" | "lost" | "unknown"
	now: number
	lastRecoveryAt: number | null
}): "wait" | "recover" | "demote" {
	if (params.status === "present" || params.status === "unknown") {
		return "wait"
	}

	return mayRecover(params.now, params.lastRecoveryAt) ? "recover" : "demote"
}

// A pause this close before an error came with it (Chrome pauses a playing element as its stream breaks):
// the user's intent is what it was before that pause.
export const ERROR_PAUSE_GRACE_MS = 1_500
