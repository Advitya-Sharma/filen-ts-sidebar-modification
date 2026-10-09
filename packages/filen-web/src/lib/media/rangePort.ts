import * as Comlink from "comlink"
import type { AbortableReadRange } from "@/lib/media/blockCache"

// Byte ranges of one file served over a MessagePort, so a reader in one worker (the remuxer) downloads
// straight from another (the SDK worker) without the page relaying every block. Each read carries an id
// the reader can cancel; the bytes come back transferred, never copied.
export interface RangePortApi {
	read(start: number, end: number, id: number): Promise<Uint8Array>
	cancel(id: number): void
	// The reader is done with the file: in-flight reads are aborted and the port closed.
	close(): void
	// Takes the Web Lock `name` and holds it until the port shuts down; true once it is held, false when
	// it cannot be (already shut down, no Web Locks). A port whose serving worker dies posts nothing, but
	// the lock goes with the worker, so a reader waiting for it learns at once. The other way round,
	// `readerLock` is one the reader holds for as long as it reads: the port then lives exactly as long as
	// its reader (however long a player stays paused), and shuts down when the reader's worker dies.
	holdLife(name: string, readerLock?: string): Promise<boolean>
}

// Posted by the server as it shuts down (a cancel from its own side, the idle timeout), so the reader
// fails its reads at once instead of waiting on a port nobody answers any more.
const CLOSED_NOTICE = "filen.rangePort.closed"

// A served port nobody has read from for this long is shut down: a reader that went away without
// closing (a crashed worker) must not keep the file's entry alive for the rest of the session. Only for a
// reader that holds no lock of its own (holdLife's readerLock), whose death cannot be seen otherwise.
export const RANGE_PORT_IDLE_MS = 30 * 60_000

// Serves `read` on `port` until `lifetime` aborts, which the reader's close() and the idle timeout do too.
export function exposeRangePort(port: MessagePort, read: AbortableReadRange, lifetime: AbortController): void {
	const closed = lifetime.signal
	const reads = new Map<number, AbortController>()
	let idle: ReturnType<typeof setTimeout> | undefined
	let readerWatched = false
	const armIdle = () => {
		clearTimeout(idle)

		if (reads.size === 0 && !readerWatched) {
			idle = setTimeout(() => {
				lifetime.abort()
			}, RANGE_PORT_IDLE_MS)
		}
	}
	const shutdown = () => {
		clearTimeout(idle)

		for (const controller of reads.values()) {
			controller.abort()
		}

		reads.clear()
		port.postMessage({ [CLOSED_NOTICE]: true })
		// A task later, so the rejections of the reads just aborted are posted before the port goes.
		setTimeout(() => {
			port.close()
		}, 0)
	}

	if (closed.aborted) {
		port.postMessage({ [CLOSED_NOTICE]: true })
		port.close()

		return
	}

	closed.addEventListener("abort", shutdown, { once: true })
	armIdle()

	const shutDown = new Promise<void>(resolve => {
		closed.addEventListener(
			"abort",
			() => {
				resolve()
			},
			{ once: true }
		)
	})

	const api: RangePortApi = {
		async read(start, end, id) {
			const controller = new AbortController()

			reads.set(id, controller)
			clearTimeout(idle)

			try {
				const bytes = await read(start, end, AbortSignal.any([controller.signal, closed]))

				return Comlink.transfer(bytes, [bytes.buffer])
			} finally {
				reads.delete(id)
				armIdle()
			}
		},
		cancel(id) {
			reads.get(id)?.abort()
		},
		close() {
			// After this call's own reply has been posted.
			setTimeout(() => {
				lifetime.abort()
			}, 0)
		},
		async holdLife(name, readerLock) {
			if (closed.aborted || !locksAvailable()) {
				return false
			}

			const held = Promise.withResolvers<boolean>()

			void navigator.locks
				.request(name, () => {
					held.resolve(true)

					return shutDown
				})
				.catch(() => {
					held.resolve(false)
				})

			const holding = await held.promise

			if (holding && readerLock !== undefined) {
				readerWatched = true
				clearTimeout(idle)
				// Granted once the reader lets go: it closed, or its worker died.
				void navigator.locks
					.request(readerLock, () => {
						lifetime.abort()
					})
					.catch(() => undefined)
			}

			return holding
		}
	}

	Comlink.expose(api, port)
}

function locksAvailable(): boolean {
	return "locks" in navigator
}

export class RangePortClosedError extends Error {
	public constructor() {
		super("range port closed")
		this.name = "RangePortClosedError"
	}
}

const CLOSE_REPLY_TIMEOUT_MS = 1000

// A read with no reply for this long has lost its server (a port closed without notice, which only some
// engines report); the SDK retries transient failures well within it.
export const RANGE_READ_TIMEOUT_MS = 60_000

export interface RangePortReader {
	readRange: AbortableReadRange
	// Called once if the serving side goes away (not after this side's own close()).
	onLost(callback: () => void): void
	close(): void
}

function isClosedNotice(data: unknown): boolean {
	return typeof data === "object" && data !== null && CLOSED_NOTICE in data
}

// The reading end of a port served by exposeRangePort. Every read settles: with the bytes, when its
// signal aborts, when the server closes, or after RANGE_READ_TIMEOUT_MS.
export function rangePortReader(port: MessagePort): RangePortReader {
	const remote = Comlink.wrap<RangePortApi>(port)
	const pending = new Set<(error: Error) => void>()
	const lostCallbacks: (() => void)[] = []
	let nextId = 0
	let gone = false
	let closing = false
	const lose = () => {
		if (gone) {
			return
		}

		gone = true

		if (!closing) {
			for (const callback of lostCallbacks.splice(0)) {
				callback()
			}
		}

		for (const fail of pending) {
			fail(new RangePortClosedError())
		}

		pending.clear()
	}

	port.addEventListener("message", event => {
		if (isClosedNotice(event.data)) {
			lose()
		}
	})
	// Where the engine reports a port closed on the other side.
	port.addEventListener("close", lose)

	// No engine reports a port whose serving worker died (crashed, terminated); the server's life lock is
	// released then, and this request for it is granted. Also granted after an orderly shutdown, which
	// has already been noticed by then.
	// This side's own lock, held until close(), tells the server the reader is still there.
	const readerDone = Promise.withResolvers<undefined>()

	if (locksAvailable()) {
		const lifeLock = `filen-range-port-${crypto.randomUUID()}`
		const readerLock = `filen-range-reader-${crypto.randomUUID()}`
		const readerHeld = Promise.withResolvers<undefined>()

		void navigator.locks
			.request(readerLock, () => {
				readerHeld.resolve(undefined)

				return readerDone.promise
			})
			.catch(() => {
				readerHeld.resolve(undefined)
			})
		void readerHeld.promise
			.then(() => remote.holdLife(lifeLock, readerLock))
			.then(held =>
				held
					? navigator.locks.request(lifeLock, () => {
							lose()
						})
					: undefined
			)
			.catch(() => undefined)
	}

	return {
		readRange: (start, end, signal) => {
			if (gone) {
				return Promise.reject(new RangePortClosedError())
			}

			signal.throwIfAborted()

			const id = nextId++

			return new Promise<Uint8Array>((resolve, reject) => {
				const timeout = setTimeout(() => {
					settle(new Error("range read timed out"))
				}, RANGE_READ_TIMEOUT_MS)
				const onAbort = () => {
					void remote.cancel(id).catch(() => undefined)
					settle(signal.reason instanceof Error ? signal.reason : new DOMException("aborted", "AbortError"))
				}
				const settle = (error: Error | null, bytes?: Uint8Array) => {
					clearTimeout(timeout)
					signal.removeEventListener("abort", onAbort)
					pending.delete(fail)

					if (error === null && bytes !== undefined) {
						resolve(bytes)
					} else {
						reject(error ?? new Error("range read failed"))
					}
				}
				const fail = (error: Error) => {
					settle(error)
				}

				pending.add(fail)
				signal.addEventListener("abort", onAbort, { once: true })
				remote.read(start, end, id).then(
					bytes => {
						settle(null, bytes)
					},
					(error: unknown) => {
						settle(error instanceof Error ? error : new Error(String(error)))
					}
				)
			})
		},
		onLost: callback => {
			lostCallbacks.push(callback)
		},
		close: () => {
			closing = true
			// A server already shut down from its side never answers.
			void Promise.race([remote.close(), new Promise(resolve => setTimeout(resolve, CLOSE_REPLY_TIMEOUT_MS))])
				.catch(() => undefined)
				.finally(() => {
					lose()
					readerDone.resolve(undefined)
					remote[Comlink.releaseProxy]()
					port.close()
				})
		}
	}
}
