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
//   failed at once (Chrome reports both so): one tiny range request through the stream tells them apart.
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

	const mayRecover = lastRecoveryAt === null || now - lastRecoveryAt >= RECOVERY_COOLDOWN_MS

	if (code === MEDIA_ERR_NETWORK || hadMetadata) {
		return mayRecover ? "recover" : "fail"
	}

	return "verify"
}
