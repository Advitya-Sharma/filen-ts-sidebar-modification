import type { CDPSession, Page } from "@playwright/test"

// Memory of a page and its workers through the Chrome DevTools Protocol (Chromium only), each sample taken
// after a full collection: the JS heap, the DOM's node, document and listener counts, layout counts, and each
// dedicated worker's heap (attached one message at a time: Playwright's CDP sessions do not reach child
// sessions).

export interface WorkerHeap {
	readonly url: string
	readonly usedMb: number
}

export interface MemorySample {
	readonly jsHeapUsedMb: number
	readonly jsHeapTotalMb: number
	readonly domNodes: number
	readonly documents: number
	readonly listeners: number
	readonly layoutCount: number
	readonly recalcStyleCount: number
	readonly workers: readonly WorkerHeap[]
}

export interface CdpMemory {
	sample(): Promise<MemorySample>
	dispose(): Promise<void>
}

const MB = 1048576

interface TargetInfo {
	readonly targetId: string
	readonly type: string
	readonly url: string
}

// How long a worker may take to answer: the SDK's wasm threads block in Atomics.wait and never do.
const WORKER_REPLY_MS = 2_000

// One CDP command to a target attached without a session of its own (Target.sendMessageToTarget).
async function sendToTarget(cdp: CDPSession, sessionId: string, method: string): Promise<unknown> {
	const id = Math.floor(Math.random() * 1e9)
	let listener: ((event: { sessionId?: string; message: string }) => void) | undefined
	let timer: ReturnType<typeof setTimeout> | undefined
	const reply = new Promise<unknown>((resolve, reject) => {
		timer = setTimeout(() => {
			reject(new Error(`${method}: no reply`))
		}, WORKER_REPLY_MS)
		listener = event => {
			if (event.sessionId !== sessionId) {
				return
			}

			const message = JSON.parse(event.message) as { id?: number; result?: unknown; error?: { message: string } }

			if (message.id !== id) {
				return
			}

			if (message.error !== undefined) {
				reject(new Error(message.error.message))
			} else {
				resolve(message.result)
			}
		}

		cdp.on("Target.receivedMessageFromTarget", listener)
	})

	try {
		await cdp.send("Target.sendMessageToTarget", { sessionId, message: JSON.stringify({ id, method }) })

		return await reply
	} finally {
		clearTimeout(timer)

		if (listener !== undefined) {
			cdp.off("Target.receivedMessageFromTarget", listener)
		}
	}
}

async function workerHeaps(cdp: CDPSession): Promise<WorkerHeap[]> {
	const { targetInfos } = (await cdp.send("Target.getTargets")) as { targetInfos: TargetInfo[] }
	const heaps: WorkerHeap[] = []

	for (const target of targetInfos.filter(info => info.type === "worker")) {
		let sessionId: string | undefined

		try {
			sessionId = (await cdp.send("Target.attachToTarget", { targetId: target.targetId, flatten: false })).sessionId

			await sendToTarget(cdp, sessionId, "HeapProfiler.collectGarbage")

			const usage = (await sendToTarget(cdp, sessionId, "Runtime.getHeapUsage")) as { usedSize: number }

			heaps.push({ url: target.url, usedMb: usage.usedSize / MB })
		} catch {
			// The worker ended meanwhile (the zip worker of a finished save), or is blocked (a wasm thread).
		} finally {
			if (sessionId !== undefined) {
				await cdp.send("Target.detachFromTarget", { sessionId }).catch(() => undefined)
			}
		}
	}

	return heaps
}

export async function cdpMemory(page: Page): Promise<CdpMemory> {
	const cdp = await page.context().newCDPSession(page)

	await cdp.send("Performance.enable")

	return {
		async sample() {
			await cdp.send("HeapProfiler.collectGarbage")

			const heap = (await cdp.send("Runtime.getHeapUsage")) as { usedSize: number; totalSize: number }
			const counters = await cdp.send("Memory.getDOMCounters")
			const { metrics } = await cdp.send("Performance.getMetrics")
			const metric = (name: string): number => metrics.find(entry => entry.name === name)?.value ?? 0

			return {
				jsHeapUsedMb: heap.usedSize / MB,
				jsHeapTotalMb: heap.totalSize / MB,
				domNodes: counters.nodes,
				documents: counters.documents,
				listeners: counters.jsEventListeners,
				layoutCount: metric("LayoutCount"),
				recalcStyleCount: metric("RecalcStyleCount"),
				workers: await workerHeaps(cdp)
			}
		},
		async dispose() {
			await cdp.detach().catch(() => undefined)
		}
	}
}

// Installed before the page loads (page.addInitScript): each key's time to the frame after it, long tasks,
// and blob URLs made and revoked. Read with readPageTimings.
export function pageTimingsScript(): void {
	const state = {
		keys: [] as number[],
		longTasks: [] as { start: number; duration: number }[],
		created: 0,
		revoked: 0
	}

	Object.defineProperty(window, "__docxSoak", { value: state })

	const channel = new MessageChannel()
	const pending: number[] = []

	// After the frame the key was painted in: a task queued from that frame's animation callback.
	channel.port1.onmessage = () => {
		const start = pending.shift()

		if (start !== undefined) {
			state.keys.push(performance.now() - start)
		}
	}

	window.addEventListener(
		"keydown",
		event => {
			// Typed keys only: not shortcuts, nor the modifiers on their own.
			if (event.ctrlKey || event.metaKey || event.altKey || ["Control", "Meta", "Alt", "Shift"].includes(event.key)) {
				return
			}

			pending.push(event.timeStamp)
			requestAnimationFrame(() => {
				channel.port2.postMessage(0)
			})
		},
		{ capture: true }
	)

	try {
		new PerformanceObserver(list => {
			for (const entry of list.getEntries()) {
				state.longTasks.push({ start: entry.startTime, duration: entry.duration })
			}
		}).observe({ type: "longtask", buffered: true })
	} catch {
		// No Long Tasks API (Firefox, WebKit): timings only.
	}

	const create = URL.createObjectURL.bind(URL)
	const revoke = URL.revokeObjectURL.bind(URL)

	URL.createObjectURL = object => {
		state.created++

		return create(object)
	}
	URL.revokeObjectURL = url => {
		state.revoked++
		revoke(url)
	}
}

export interface PageTimings {
	readonly keys: readonly number[]
	readonly longTasks: readonly { readonly start: number; readonly duration: number }[]
	readonly created: number
	readonly revoked: number
	readonly now: number
}

export function readPageTimings(page: Page): Promise<PageTimings> {
	return page.evaluate(() => {
		const state = (window as unknown as { __docxSoak: Omit<PageTimings, "now"> }).__docxSoak

		return {
			keys: [...state.keys],
			longTasks: [...state.longTasks],
			created: state.created,
			revoked: state.revoked,
			now: performance.now()
		}
	})
}
