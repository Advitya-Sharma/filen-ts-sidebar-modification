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
	// `onEvent` is the page's Comlink proxy; segments cross transferred.
	async open(params: RemuxOpenParams, onEvent: EventSink): Promise<{ id: number; info: RemuxMediaInfo }> {
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
