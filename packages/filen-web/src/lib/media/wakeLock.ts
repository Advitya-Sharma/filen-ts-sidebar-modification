// Keeps the screen on while a video plays. Chromium already holds a wake lock for a playing, visible
// <video>, so this asks only elsewhere. The browser drops the lock whenever the page is hidden; it is asked
// for again when the page comes back while still wanted.

interface WakeLockSentinelLike {
	readonly released: boolean
	release: () => Promise<void>
	addEventListener: (type: "release", listener: () => void) => void
}

interface WakeLockLike {
	request: (type: "screen") => Promise<WakeLockSentinelLike>
}

export interface ScreenWakeLock {
	// Whether the screen should stay on; repeated calls with the same value cost nothing.
	set: (wanted: boolean) => void
	dispose: () => void
}

export function createScreenWakeLock(options: { needed: boolean; wakeLock?: WakeLockLike | null; doc?: Document }): ScreenWakeLock {
	const api =
		options.wakeLock !== undefined
			? options.wakeLock
			: typeof navigator !== "undefined" && "wakeLock" in navigator
				? navigator.wakeLock
				: null
	const doc = options.doc ?? (typeof document === "undefined" ? null : document)

	if (!options.needed || api === null || doc === null) {
		return {
			set: () => undefined,
			dispose: () => undefined
		}
	}

	let wanted = false
	let sentinel: WakeLockSentinelLike | null = null
	let requesting = false
	let disposed = false

	function request(): void {
		if (requesting || sentinel !== null || !wanted || disposed || doc?.visibilityState !== "visible" || api === null) {
			return
		}

		requesting = true
		api.request("screen").then(
			lock => {
				requesting = false

				if (!wanted || disposed) {
					void lock.release().catch(() => undefined)

					return
				}

				sentinel = lock
				lock.addEventListener("release", () => {
					if (sentinel === lock) {
						sentinel = null
					}
				})
			},
			() => {
				// Refused (battery saver, no user activation yet): the screen may dim, nothing else changes.
				requesting = false
			}
		)
	}

	function drop(): void {
		const held = sentinel

		sentinel = null

		if (held !== null && !held.released) {
			void held.release().catch(() => undefined)
		}
	}

	function onVisibility(): void {
		if (doc?.visibilityState === "visible") {
			request()
		}
	}

	doc.addEventListener("visibilitychange", onVisibility)

	return {
		set: next => {
			if (wanted === next) {
				return
			}

			wanted = next

			if (next) {
				request()
			} else {
				drop()
			}
		},
		dispose: () => {
			disposed = true
			wanted = false
			doc.removeEventListener("visibilitychange", onVisibility)
			drop()
		}
	}
}
