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
}

// Posted by the server as it shuts down (a cancel from its own side, the idle timeout), so the reader
// fails its reads at once instead of waiting on a port nobody answers any more.
const CLOSED_NOTICE = "filen.rangePort.closed"

// A served port nobody has read from for this long is shut down: a reader that went away without
// closing (a crashed worker) must not keep the file's entry alive for the rest of the session.
export const RANGE_PORT_IDLE_MS = 30 * 60_000

// Serves `read` on `port` until `lifetime` aborts, which the reader's close() and the idle timeout do too.
export function exposeRangePort(port: MessagePort, read: AbortableReadRange, lifetime: AbortController): void {
	const closed = lifetime.signal
	const reads = new Map<number, AbortController>()
	let idle: ReturnType<typeof setTimeout> | undefined
	const armIdle = () => {
		clearTimeout(idle)

		if (reads.size === 0) {
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
		}
	}

	Comlink.expose(api, port)
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
					remote[Comlink.releaseProxy]()
					port.close()
				})
		}
	}
}
