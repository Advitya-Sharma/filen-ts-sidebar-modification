import * as Comlink from "comlink"
import type {
	RemuxAudioStart,
	RemuxAudioStartParams,
	RemuxEvent,
	RemuxMediaInfo,
	RemuxOpenParams,
	RemuxSeekParams,
	RemuxSeekResult,
	RemuxStats,
	TrackKind
} from "@/lib/media/remuxTypes"
import { RemuxEngine } from "@/workers/remux/engine"

// Remuxes files the browser cannot play as they are (Matroska, or audio it cannot decode) into
// fragmented MP4 for the page's MediaSource. Its own worker so the demuxer's heap is freed with it once
// playback ends, and so a parse never runs on the page or in the SDK worker, whose heap never shrinks.
// More than one session at a time: the next episode is opened while the current one plays.
// The page's callback, a Comlink proxy here: releasing it frees the page's end too.
type EventSink = ((event: RemuxEvent) => unknown) & Partial<Comlink.ProxyMethods>

const sessions = new Map<number, { engine: RemuxEngine; onEvent: EventSink }>()
let nextSession = 0

// A file opened ahead of playback (the next episode, once Up Next shows): its head and first keyframe interval
// are in its block cache when the player's own open adopts it. One at a time; one nobody took closes itself.
interface Warmed {
	key: string
	engine: RemuxEngine
	info: Promise<RemuxMediaInfo>
	adopt: (sink: EventSink) => void
	expiry: ReturnType<typeof setTimeout>
}

const WARM_EXPIRY_MS = 5 * 60_000
// Subtitle cues found while warming wait for the session that adopts the file.
const MAX_WARM_EVENTS = 256
let warmed: Warmed | null = null

function dropWarmed(key?: string): Promise<void> {
	const entry = warmed

	if (entry === null || (key !== undefined && entry.key !== key)) {
		return Promise.resolve()
	}

	warmed = null
	clearTimeout(entry.expiry)

	return entry.engine.close().catch(() => undefined)
}

function session(id: number): RemuxEngine {
	const entry = sessions.get(id)

	if (entry === undefined) {
		throw new Error("remux session closed")
	}

	return entry.engine
}

async function closeSession(id: number): Promise<void> {
	const entry = sessions.get(id)

	sessions.delete(id)

	if (entry !== undefined) {
		await entry.engine.close()
		// Frees the page's end of the event callback too.
		entry.onEvent[Comlink.releaseProxy]?.()
	}
}

const api = {
	// Opens `params` under `key` ahead of playback: the head, the index and the first keyframe interval are read
	// into its cache. The port belongs to the warmed file from now on.
	async warm(params: RemuxOpenParams, key: string): Promise<void> {
		await dropWarmed()

		let sink: EventSink | null = null
		const held: RemuxEvent[] = []
		const engine = new RemuxEngine(params, event => {
			if (sink !== null) {
				void sink(event.type === "segment" ? Comlink.transfer(event, [event.bytes.buffer]) : event)
			} else if (event.type === "textCues" && held.length < MAX_WARM_EVENTS) {
				held.push(event)
			}
		})
		const entry: Warmed = {
			key,
			engine,
			info: engine.open(params),
			adopt: next => {
				sink = next

				for (const event of held.splice(0)) {
					void next(event)
				}
			},
			expiry: setTimeout(() => {
				void dropWarmed(key)
			}, WARM_EXPIRY_MS)
		}

		warmed = entry

		try {
			await entry.info
			await engine.canStartAt(0)
		} catch {
			await dropWarmed(key)
		}
	},
	unwarm(key: string): Promise<void> {
		return dropWarmed(key)
	},
	// `onEvent` is the page's Comlink proxy; segments cross transferred. With `warmKey` naming the file warmed
	// last, that file is taken over and the port passed along closes unused.
	async open(params: RemuxOpenParams & { warmKey?: string }, onEvent: EventSink): Promise<{ id: number; info: RemuxMediaInfo }> {
		const ready = warmed

		if (params.warmKey !== undefined && ready?.key === params.warmKey) {
			warmed = null
			clearTimeout(ready.expiry)
			params.port.close()

			const id = nextSession++

			ready.adopt(onEvent)
			sessions.set(id, { engine: ready.engine, onEvent })

			try {
				return { id, info: await ready.info }
			} catch (error) {
				await closeSession(id)

				throw error
			}
		}

		const id = nextSession++
		const engine = new RemuxEngine(params, event => {
			void onEvent(event.type === "segment" ? Comlink.transfer(event, [event.bytes.buffer]) : event)
		})

		sessions.set(id, { engine, onEvent })

		try {
			return { id, info: await engine.open(params) }
		} catch (error) {
			await closeSession(id)

			throw error
		}
	},
	seek(id: number, params: RemuxSeekParams): Promise<RemuxSeekResult> {
		return session(id).seek(params)
	},
	switchAudio(id: number, params: RemuxAudioStartParams): Promise<RemuxAudioStart> {
		return session(id).switchAudio(params)
	},
	canStartAt(id: number, time: number): Promise<boolean> {
		return session(id).canStartAt(time)
	},
	stopAudio(id: number): Promise<void> {
		return session(id).stopAudio()
	},
	setHorizon(id: number, kind: TrackKind, generation: number, horizon: number): void {
		sessions.get(id)?.engine.setHorizon(kind, generation, horizon)
	},
	stats(id: number): RemuxStats | null {
		return sessions.get(id)?.engine.snapshotStats() ?? null
	},
	close(id: number): Promise<void> {
		return closeSession(id)
	},
	// Held for the worker's whole life, so the page learns the moment it is gone (crashed, terminated):
	// the page waits for the same lock, which it gets only then.
	async holdLifeLock(name: string): Promise<void> {
		const acquired = Promise.withResolvers<undefined>()

		void navigator.locks.request(name, () => {
			acquired.resolve(undefined)

			// Never settles: the lock goes with the worker.
			return new Promise<never>(() => undefined)
		})
		await acquired.promise
	}
}

export type RemuxWorkerApi = typeof api

Comlink.expose(api)
