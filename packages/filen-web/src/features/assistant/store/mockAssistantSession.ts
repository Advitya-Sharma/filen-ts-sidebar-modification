import type {
	AssistantEvent,
	AssistantExchange,
	AssistantModelInfo,
	AssistantModelState,
	AssistantProposal,
	AssistantSession,
	AssistantStep
} from "@/features/assistant/types"

// A scripted AssistantSession for tests and for development builds before a real runtime registers. It
// walks the whole UI surface: a chunked download, a load that starts only on request (the Load action or a
// sent message, which waits for it), the turn's steps (loading, routing, a search with its prompt-reading
// progress and result count), a streamed markdown reply, and a confirmation card
// when the message asks to draft, move, copy, trash or download something. Timer-driven, so tests step it
// with fake timers.

export const MOCK_ASSISTANT_MODEL: AssistantModelInfo = {
	id: "mock-model",
	name: "Qwen3.5 4B",
	sizeBytes: 2_740_000_000
}

const PROPOSAL_PATTERN = /\b(draft|move|copy|trash|download)\b/i

const FILE_TARGETS = [
	{ handle: "f1", name: "Taxes 2025.pdf" },
	{ handle: "f2", name: "Receipts.xlsx" },
	{ handle: "d1", name: "Scans" }
]

function mockProposal(id: string, verb: string): AssistantProposal {
	switch (verb) {
		case "draft":
			return {
				id,
				action: "draft-message",
				targets: [{ handle: "c1", name: "Family" }],
				destination: null,
				text: { text: "Running late, see you at 8!" }
			}
		case "move":
		case "copy":
			return { id, action: verb, targets: FILE_TARGETS, destination: "Documents/Archive", text: {} }
		case "trash":
			return { id, action: "trash", targets: FILE_TARGETS, destination: null, text: {} }
		default:
			return { id, action: "download", targets: FILE_TARGETS, destination: null, text: {} }
	}
}

export interface MockAssistantSessionOptions {
	initialState?: AssistantModelState
	// Delay between scripted steps (download chunks, reply words).
	stepMs?: number
	downloadSteps?: number
}

export function createMockAssistantSession({
	initialState = { kind: "not-downloaded", model: MOCK_ASSISTANT_MODEL },
	stepMs = 120,
	downloadSteps = 10
}: MockAssistantSessionOptions = {}): AssistantSession {
	const listeners = new Set<(event: AssistantEvent) => void>()
	const timers = new Set<ReturnType<typeof setTimeout>>()
	let modelState = initialState
	let turnCount = 0
	let activeTurnId: string | null = null
	let history: readonly AssistantExchange[] = []

	function emit(event: AssistantEvent): void {
		if (event.type === "model") {
			modelState = event.state
		}

		for (const listener of listeners) {
			listener(event)
		}
	}

	function later(run: () => void, steps: number): void {
		const timer = setTimeout(() => {
			timers.delete(timer)
			run()
		}, stepMs * steps)

		timers.add(timer)
	}

	function clearTimers(): void {
		for (const timer of timers) {
			clearTimeout(timer)
		}

		timers.clear()
	}

	const LOAD_STEPS = 2
	// Apart from the turn's timers: like the runtime's, a load carries on when its turn is stopped.
	let loadTimer: ReturnType<typeof setTimeout> | undefined

	// Steps until the model is ready: 0 when it is, the rest of a load already under way, or a new load.
	function loadModel(): number {
		if (modelState.kind === "ready") {
			return 0
		}

		if (modelState.kind === "loading") {
			return LOAD_STEPS
		}

		if (modelState.kind !== "downloaded") {
			return 0
		}

		const model = modelState.model

		emit({ type: "model", state: { kind: "loading", model, verifying: false } })
		loadTimer = setTimeout(() => {
			loadTimer = undefined
			emit({ type: "model", state: { kind: "ready", model } })
		}, stepMs * LOAD_STEPS)

		return LOAD_STEPS
	}

	function downloadModel(): void {
		if (modelState.kind !== "not-downloaded" && modelState.kind !== "error") {
			return
		}

		const model = modelState.model ?? MOCK_ASSISTANT_MODEL

		emit({ type: "model", state: { kind: "downloading", model, receivedBytes: 0, waiting: false } })

		for (let step = 1; step <= downloadSteps; step++) {
			later(() => {
				emit({
					type: "model",
					state: {
						kind: "downloading",
						model,
						receivedBytes: Math.round((model.sizeBytes * step) / downloadSteps),
						waiting: false
					}
				})
			}, step)
		}

		later(() => {
			emit({ type: "model", state: { kind: "downloaded", model, idleUnloaded: false } })
		}, downloadSteps + 1)
	}

	function cancelDownload(): void {
		if (modelState.kind !== "downloading") {
			return
		}

		clearTimers()
		emit({ type: "model", state: { kind: "not-downloaded", model: modelState.model } })
	}

	function deleteModel(): Promise<void> {
		clearTimers()
		clearTimeout(loadTimer)
		activeTurnId = null
		emit({ type: "model", state: { kind: "not-downloaded", model: MOCK_ASSISTANT_MODEL } })

		return Promise.resolve()
	}

	function send(message: string): string {
		turnCount++

		const turnId = `turn-${String(turnCount)}`
		const step = (next: AssistantStep) => {
			emit({ type: "step", turnId, step: next, at: Date.now() })
		}
		const search = (patch: Partial<Extract<AssistantStep, { kind: "tool" }>>) => {
			step({
				id: `${turnId}-search`,
				kind: "tool",
				status: "active",
				progress: null,
				tool: "search_items",
				subject: "taxes",
				resultCount: null,
				resultMore: false,
				failure: null,
				...patch
			})
		}
		const verb = PROPOSAL_PATTERN.exec(message)?.[1]?.toLowerCase()
		const words = (
			verb === undefined
				? "I found **3 files** matching that in your drive:\n\n- `Taxes 2025.pdf`\n- `Receipts.xlsx`\n- `Notes.md`"
				: "Here is what I would do. Nothing changes until you confirm."
		).split(/(?<= )/)

		activeTurnId = turnId

		// A model on disk loads first, inside the turn.
		const loadSteps = loadModel()
		const loading = (status: AssistantStep["status"]) => {
			step({ id: `${turnId}-loading`, kind: "loading", status, progress: null })
		}
		const startSearch = () => {
			if (loadSteps > 0) {
				loading("done")
			}

			step({ id: `${turnId}-routing`, kind: "routing", status: "done", progress: null, route: "files" })
			search({ progress: 0.4 })
		}

		if (loadSteps === 0) {
			startSearch()
		} else {
			loading("active")
			later(startSearch, loadSteps)
		}

		later(() => {
			search({ progress: 1 })
		}, loadSteps + 1)
		later(() => {
			search({ status: "done", resultCount: 3 })
		}, loadSteps + 2)

		words.forEach((_, index) => {
			later(
				() => {
					emit({
						type: "reply",
						turnId,
						replyId: `${turnId}-reply`,
						text: words.slice(0, index + 1).join(""),
						final: index === words.length - 1
					})
				},
				loadSteps + 3 + index
			)
		})

		later(
			() => {
				if (verb !== undefined) {
					emit({ type: "proposal", turnId, proposal: mockProposal(`${turnId}-proposal`, verb) })
				}

				activeTurnId = null
				history = [...history, { user: message, reply: words.join("") }]
				emit({ type: "turn-end", turnId, at: Date.now(), outcome: "complete" })
			},
			loadSteps + 4 + words.length
		)

		return turnId
	}

	function cancel(): void {
		if (activeTurnId === null) {
			return
		}

		const turnId = activeTurnId

		activeTurnId = null
		clearTimers()
		emit({ type: "turn-end", turnId, at: Date.now(), outcome: "cancelled" })
	}

	function confirm(proposalId: string): void {
		later(() => {
			emit({ type: "proposal-settled", proposalId, outcome: "done" })
		}, 2)
	}

	return {
		getModelState: () => modelState,
		subscribe: listener => {
			listeners.add(listener)

			return () => {
				listeners.delete(listener)
			}
		},
		setPanelOpen: () => undefined,
		downloadModel,
		loadModel: () => {
			loadModel()
		},
		noteActivity: () => undefined,
		cancelDownload,
		deleteModel,
		send,
		cancel,
		confirm,
		reject: () => undefined,
		reset: (resumed = []) => {
			cancel()
			history = resumed
		},
		getHistory: () => history
	}
}
