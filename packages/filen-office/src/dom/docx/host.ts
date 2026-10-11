import type { ByteSource } from "../../core/opc/index"
import { DocxEngine, type DocxEngineServices } from "./engine"
import type { DocxAnnotations } from "./annotations"
import type { DocxEvent, DocxFindOptions, DocxSessionApi, DocxLocation, DocxMedia, DocxObjectData, DocxOpenOptions } from "./protocol"

// Two ways to run the engine behind the same DocxSessionApi: in a worker (web), talking over a small
// message protocol, or on the calling thread, time-sliced (an Expo DOM component, where workers are not
// assumed). The document's bytes stay with the caller either way: the engine reads the ranges it needs.

export type DocxListener = (event: DocxEvent) => void

// Opens a document; the listener hears every event in order until close().
export type DocxHost = (source: ByteSource, listener: DocxListener, options?: DocxOpenOptions) => DocxSessionApi

interface MessageChannelLike {
	readonly port1: { onmessage: ((event: unknown) => void) | null; close?(): void }
	readonly port2: { postMessage(message: unknown): void; close?(): void }
}

// A macrotask turn, so messages queued for the thread (requests, reads) are handled between layout slices.
// MessageChannel avoids the 4 ms clamp nested timers get; setTimeout is the fallback.
export function macrotaskYield(): () => Promise<void> {
	const Channel: unknown = Reflect.get(globalThis, "MessageChannel")

	if (typeof Channel !== "function") {
		const timer: unknown = Reflect.get(globalThis, "setTimeout")

		return () =>
			new Promise(resolve => {
				if (typeof timer === "function") timer(resolve, 0)
				else resolve()
			})
	}

	const channel = new (Channel as new () => MessageChannelLike)()
	const waiting: (() => void)[] = []

	channel.port1.onmessage = () => {
		waiting.shift()?.()
	}

	return () =>
		new Promise<void>(resolve => {
			waiting.push(resolve)
			channel.port2.postMessage(0)
		})
}

// --- Messages ---

type Method = "objects" | "media" | "locate" | "find" | "annotations"

// `s`: the session (one open document) a message belongs to.
type ToWorker =
	| { readonly t: "open"; readonly s: number; readonly size: number; readonly options: DocxOpenOptions }
	| { readonly t: "read"; readonly s: number; readonly id: number; readonly bytes?: Uint8Array; readonly error?: string }
	| {
			readonly t: "call"
			readonly s: number
			readonly id: number
			readonly method: Method
			readonly arg: unknown
			readonly options?: unknown
	  }
	| { readonly t: "close"; readonly s: number }

type FromWorker =
	| { readonly t: "event"; readonly s: number; readonly event: DocxEvent }
	| { readonly t: "read"; readonly s: number; readonly id: number; readonly offset: number; readonly length: number }
	| { readonly t: "result"; readonly s: number; readonly id: number; readonly value?: unknown; readonly error?: string }

function bufferOf(bytes: ArrayBufferView): ArrayBuffer | undefined {
	return bytes.buffer instanceof ArrayBuffer ? bytes.buffer : undefined
}

function isRecord(value: unknown): value is Record<string, unknown> {
	return typeof value === "object" && value !== null
}

function findOptions(value: unknown): DocxFindOptions {
	if (!isRecord(value)) return {}

	const caseSensitive = value["caseSensitive"]
	const limit = value["limit"]

	return {
		...(typeof caseSensitive === "boolean" ? { caseSensitive } : {}),
		...(typeof limit === "number" && Number.isInteger(limit) && limit > 0 ? { limit } : {})
	}
}

// --- Worker side ---

// The worker's end of its channel to the page (adapt `self`).
export interface DocxWorkerScope {
	post(message: unknown, transfer: readonly ArrayBuffer[]): void
	listen(onMessage: (data: unknown) => void): void
}

interface WorkerSession {
	readonly engine: DocxEngine
	readonly reads: Map<number, { resolve: (bytes: Uint8Array) => void; reject: (error: Error) => void }>
	nextRead: number
}

// Serves documents for the page, any number at a time, each its own engine. A closed document's engine
// drops everything it held; the worker itself lives on for the next one (and keeps fetched font files).
export function serveDocxWorker(scope: DocxWorkerScope, services: (options: DocxOpenOptions) => DocxEngineServices): void {
	const sessions = new Map<number, WorkerSession>()
	const post = (message: FromWorker, transfer: readonly ArrayBuffer[] = []): void => {
		scope.post(message, transfer)
	}

	scope.listen(data => {
		if (!isRecord(data) || typeof data["s"] !== "number") return

		const message = data as ToWorker
		const id = message.s

		switch (message.t) {
			case "open": {
				if (sessions.has(id) || typeof message.size !== "number") return

				const session: WorkerSession = {
					engine: new DocxEngine(
						{
							size: message.size,
							read(offset, length) {
								if (session.engine.closed) return Promise.reject(new Error("closed"))

								const read = session.nextRead++

								return new Promise((resolve, reject) => {
									session.reads.set(read, { resolve, reject })
									post({ t: "read", s: id, id: read, offset, length })
								})
							}
						},
						services(message.options),
						(event, transfer) => post({ t: "event", s: id, event }, transfer),
						message.options
					),
					reads: new Map(),
					nextRead: 1
				}

				sessions.set(id, session)
				void session.engine.run()

				break
			}
			case "read": {
				const session = sessions.get(id)
				const pending = session?.reads.get(message.id)

				session?.reads.delete(message.id)

				if (message.bytes instanceof Uint8Array) pending?.resolve(message.bytes)
				else pending?.reject(new RangeError(message.error ?? "Read failed"))

				break
			}
			case "call": {
				const engine = sessions.get(id)?.engine
				const call = message.id

				if (engine === undefined) {
					post({ t: "result", s: id, id: call, error: "closed" })

					return
				}

				const transfer: ArrayBuffer[] = []
				const run = async (): Promise<unknown> => {
					switch (message.method) {
						case "objects":
							return engine.objects(
								Array.isArray(message.arg) ? message.arg.filter(item => typeof item === "number") : [],
								transfer
							)
						case "media": {
							const media = typeof message.arg === "string" ? await engine.media(message.arg) : null

							if (media === null) return null

							// A view into a larger inflate buffer is copied so only its own bytes move.
							const bytes =
								media.bytes.byteOffset === 0 && media.bytes.byteLength === media.bytes.buffer.byteLength
									? media.bytes
									: media.bytes.slice()
							const buffer = bufferOf(bytes)

							if (buffer !== undefined) transfer.push(buffer)

							return { bytes, type: media.type }
						}
						case "locate":
							return typeof message.arg === "string" ? engine.locate(message.arg) : null
						case "find": {
							const matches =
								typeof message.arg === "string"
									? await engine.find(message.arg, findOptions(message.options))
									: new Int32Array(0)
							const buffer = bufferOf(matches)

							if (buffer !== undefined) transfer.push(buffer)

							return matches
						}
						case "annotations":
							return engine.annotations()
					}
				}

				run().then(
					value => post({ t: "result", s: id, id: call, value }, transfer),
					(error: unknown) =>
						post({ t: "result", s: id, id: call, error: error instanceof Error ? error.message : String(error) })
				)

				break
			}
			case "close": {
				const session = sessions.get(id)

				sessions.delete(id)
				session?.engine.close()

				for (const pending of session?.reads.values() ?? []) pending.reject(new Error("closed"))

				break
			}
		}
	})
}

// --- Caller side ---

// The page's end of a worker running serveDocxWorker (adapt a Worker).
export interface DocxWorkerLike {
	post(message: unknown, transfer: readonly ArrayBuffer[]): void
	listen(onMessage: (data: unknown) => void, onError: () => void): void
	terminate(): void
}

export interface DocxWorkerPool {
	readonly host: DocxHost
	// Starts the worker ahead of the first document (its script is parsed while the file downloads).
	prewarm(): void
	dispose(): void
}

// Timers through globalThis: the module type-checks without DOM or Node typings.
function setTimer(run: () => void, ms: number): unknown {
	const set: unknown = Reflect.get(globalThis, "setTimeout")

	return typeof set === "function" ? set(run, ms) : undefined
}

function clearTimer(handle: unknown): void {
	const clear: unknown = Reflect.get(globalThis, "clearTimeout")

	if (typeof clear === "function") clear(handle)
}

interface Connection {
	readonly worker: DocxWorkerLike
	readonly sessions: Map<number, (data: FromWorker) => void>
	readonly failures: Map<number, () => void>
	idle: unknown
}

// One worker shared by every open document and kept `idleMs` after the last one closes, so a second
// document skips the worker's start-up and font downloads. Reads are answered from each document's source
// in the caller's thread; each answer is a copy of just that range.
export function docxWorkerPool(create: () => DocxWorkerLike, idleMs = 60_000): DocxWorkerPool {
	let connection: Connection | undefined
	let nextSession = 1

	const scheduleIdle = (current: Connection): void => {
		if (current.sessions.size > 0 || current.idle !== undefined) return

		current.idle = setTimer(() => {
			current.idle = undefined

			if (current.sessions.size === 0 && connection === current) {
				connection = undefined
				current.worker.terminate()
			}
		}, idleMs)
	}

	const connect = (): Connection => {
		if (connection !== undefined) {
			if (connection.idle !== undefined) {
				clearTimer(connection.idle)
				connection.idle = undefined
			}

			return connection
		}

		const worker = create()
		const current: Connection = { worker, sessions: new Map(), failures: new Map(), idle: undefined }

		worker.listen(
			data => {
				if (isRecord(data) && typeof data["s"] === "number") current.sessions.get(data["s"])?.(data as FromWorker)
			},
			() => {
				// The worker died (a defect or out of memory): every document on it fails, the next starts a new one.
				if (connection === current) connection = undefined

				for (const fail of [...current.failures.values()]) fail()

				worker.terminate()
			}
		)
		connection = current

		return current
	}

	const host: DocxHost = (source, listener, options = {}) => {
		const current = connect()
		const id = nextSession++
		const worker = current.worker
		let closed = false
		let nextCall = 1
		let mainPartBytes: number | undefined
		const calls = new Map<number, { resolve: (value: unknown) => void; reject: (error: Error) => void }>()
		const end = (): void => {
			closed = true
			current.sessions.delete(id)
			current.failures.delete(id)

			for (const call of calls.values()) call.reject(new Error("closed"))

			calls.clear()
			scheduleIdle(current)
		}
		const close = (): void => {
			if (closed) return

			worker.post({ t: "close", s: id } satisfies ToWorker, [])
			end()
		}

		current.failures.set(id, () => {
			if (closed) return

			end()
			listener({
				type: "error",
				reason: "failed",
				message: "The document worker failed",
				...(mainPartBytes === undefined ? {} : { mainPartBytes })
			})
		})
		current.sessions.set(id, message => {
			if (closed) return

			switch (message.t) {
				case "event":
					if (message.event.type === "opened") mainPartBytes = message.event.info.mainPartBytes

					listener(message.event)

					if (message.event.type === "error") close()

					break
				case "read":
					source.read(message.offset, message.length).then(
						view => {
							if (closed) return

							const copy = view.slice()

							worker.post({ t: "read", s: id, id: message.id, bytes: copy } satisfies ToWorker, [copy.buffer])
						},
						(error: unknown) => {
							if (!closed) {
								worker.post(
									{
										t: "read",
										s: id,
										id: message.id,
										error: error instanceof Error ? error.message : String(error)
									} satisfies ToWorker,
									[]
								)
							}
						}
					)

					break
				case "result": {
					const call = calls.get(message.id)

					calls.delete(message.id)

					if (message.error === undefined) call?.resolve(message.value)
					else call?.reject(new Error(message.error))

					break
				}
			}
		})
		worker.post({ t: "open", s: id, size: source.size, options } satisfies ToWorker, [])

		const call = <T>(method: Method, arg: unknown, callOptions?: unknown): Promise<T> => {
			if (closed) return Promise.reject(new Error("closed"))

			const callId = nextCall++

			return new Promise<T>((resolve, reject) => {
				calls.set(callId, { resolve: value => resolve(value as T), reject })
				worker.post({ t: "call", s: id, id: callId, method, arg, options: callOptions } satisfies ToWorker, [])
			})
		}

		return {
			objects: ids => call<readonly DocxObjectData[]>("objects", ids),
			media: partName => call<DocxMedia | null>("media", partName),
			locate: anchor => call<DocxLocation | null>("locate", anchor),
			find: (query, findOptions) => call<Int32Array>("find", query, findOptions),
			annotations: () => call<DocxAnnotations | null>("annotations", null),
			close
		}
	}

	return {
		host,
		prewarm() {
			scheduleIdle(connect())
		},
		dispose() {
			const current = connection

			connection = undefined

			if (current === undefined) return

			if (current.idle !== undefined) clearTimer(current.idle)

			for (const fail of [...current.failures.values()]) fail()

			current.worker.terminate()
		}
	}
}

// A host that runs the engine on the calling thread, yielding between slices. `services` is called per
// document (a fresh font manager scope and measurer state per document is the engine's own business).
export function inThreadDocxHost(services: (options: DocxOpenOptions) => DocxEngineServices): DocxHost {
	return (source, listener, options = {}) => {
		let closed = false
		const engine = new DocxEngine(
			source,
			services(options),
			event => {
				if (!closed) listener(event)
			},
			options
		)

		void engine.run()

		return {
			objects: ids => engine.objects(ids),
			media: partName => engine.media(partName),
			locate: anchor => engine.locate(anchor),
			find: (query, findOptions) => engine.find(query, findOptions),
			annotations: () => engine.annotations(),
			close() {
				closed = true
				engine.close()
			}
		}
	}
}
