import {
	AssistantAbortError,
	AssistantContextError,
	runAssistantTurn,
	type AssistantConfirmationRequest,
	type AssistantHistoryTurn,
	type AssistantModel,
	type AssistantToolCallRecord,
	type AssistantToolCallStatus,
	type AssistantTurnEvent
} from "@filen/shared/assistant"
import type { ToolExecutor } from "@/features/assistant/runtime/executor"
import {
	endedStep,
	type AssistantEvent,
	type AssistantModelState,
	type AssistantSession,
	type AssistantStep,
	type AssistantStepFailure,
	type AssistantStepStatus
} from "@/features/assistant/types"
import { buildProposal, turnExecutor, type TurnExecutor } from "@/features/assistant/loop/executorAdapter"
import { log } from "@/lib/log"

// What the session needs of the model runtime (runtime/client.ts), so tests drive it with a fake.
export interface AssistantRuntimeLike {
	readonly model: AssistantModel
	getState(): AssistantModelState
	subscribe(listener: (state: AssistantModelState) => void): () => void
	init(): Promise<void>
	setPanelOpen(open: boolean): void
	// Settles once the model is ready or has failed to load; the state says which.
	load(): Promise<void>
	// Postpones the idle unload.
	touch(): void
	retry(): void
	cancelDownload(): void
	deleteModel(): Promise<void>
	// Keeps the model loaded until the returned release runs.
	hold(): () => void
}

// One executed (or refused) non-read call, newest last: what the assistant did or tried to do this session.
export interface AssistantActionLogEntry {
	at: number
	turnId: string
	tool: string
	args: Record<string, unknown>
	status: AssistantToolCallStatus
	error?: string
}

const ACTION_LOG_MAX = 200

// Prior exchanges kept for context; fitMessages drops the oldest of them when the 8K context runs short.
const HISTORY_MAX = 12

const FINAL_STATUSES = new Set<AssistantToolCallStatus>(["ok", "error", "invalid", "blocked", "declined", "repeated"])

const STEP_STATUS: Record<AssistantToolCallStatus, AssistantStepStatus> = {
	preparing: "active",
	queued: "queued",
	running: "active",
	confirming: "waiting",
	ok: "done",
	error: "failed",
	invalid: "failed",
	blocked: "failed",
	declined: "failed",
	discarded: "discarded",
	repeated: "failed"
}

const STEP_FAILURES = new Map<AssistantToolCallStatus, AssistantStepFailure>([
	["error", "error"],
	["invalid", "invalid"],
	["blocked", "blocked"],
	["declined", "declined"],
	["repeated", "repeated"]
])

function sameStep(a: AssistantStep, b: AssistantStep): boolean {
	const keys = new Set([...Object.keys(a), ...Object.keys(b)])

	return [...keys].every(key => (a as Record<string, unknown>)[key] === (b as Record<string, unknown>)[key])
}

export interface TurnSteps {
	// Brackets the model load a turn waits for.
	loading: (status: AssistantStepStatus) => void
	event: (event: AssistantTurnEvent) => void
	// The turn ended: rows still running end with it, as stopped by the user or failed with the turn (a turn
	// that completed leaves none). Whatever arrives afterwards is stale and dropped.
	close: (outcome: "done" | "stopped" | "failed") => void
}

// The loop's events as the turn's progress list. Routing, each step's thinking and the closing reply are
// phase rows, at most one active at a time; a call is a row of its own, re-emitted only when what it
// shows changed. Prompt-reading progress goes to the active phase, or to the running call while the
// loop extracts its content.
export function turnSteps({ turnId, emit, now }: { turnId: string; emit: (event: AssistantEvent) => void; now: () => number }): TurnSteps {
	const steps = new Map<string, AssistantStep>()
	let phase: AssistantStep | null = null
	let inTool = false
	let runningTool: string | null = null
	let count = 0
	let closed = false

	const put = (step: AssistantStep) => {
		steps.set(step.id, step)
		emit({ type: "step", turnId, step, at: now() })
	}

	const endPhase = (status: AssistantStepStatus) => {
		if (phase !== null) {
			put({ ...phase, status, progress: null })
			phase = null
		}
	}

	const startPhase = (kind: "routing" | "thinking" | "replying") => {
		endPhase("done")
		count++
		phase =
			kind === "routing"
				? { id: `${turnId}:routing`, kind, status: "active", progress: null, route: null }
				: { id: `${turnId}:${kind}${String(count)}`, kind, status: "active", progress: null }
		put(phase)
	}

	const progress = (share: number) => {
		const target = inTool && runningTool !== null ? steps.get(runningTool) : phase

		if (target?.status !== "active" || target.progress === share) {
			return
		}

		const next = { ...target, progress: share }

		if (target === phase) {
			phase = next
		}

		put(next)
	}

	return {
		loading: status => {
			if (!closed) {
				put({ id: `${turnId}:loading`, kind: "loading", status, progress: null })
			}
		},

		event: event => {
			if (closed) {
				return
			}

			switch (event.type) {
				case "phase":
					inTool = event.phase === "tool" || event.phase === "confirming"

					if (event.phase === "routing" || event.phase === "thinking" || event.phase === "replying") {
						startPhase(event.phase)
					}

					break
				case "route":
					if (phase?.kind === "routing") {
						put({ ...phase, status: "done", progress: null, route: event.route })
						phase = null
					}

					break
				case "progress": {
					const { processed, total } = event.progress

					progress(total > 0 ? Math.min(1, Math.max(0, processed / total)) : 1)

					break
				}
				case "call": {
					endPhase("done")

					const id = `${turnId}:${event.call.id}`
					const step: AssistantStep = {
						id,
						kind: "tool",
						status: STEP_STATUS[event.call.status],
						progress: null,
						tool: event.call.tool,
						subject: event.call.subject ?? null,
						resultCount: event.call.resultCount ?? null,
						resultMore: event.call.resultMore ?? false,
						failure: STEP_FAILURES.get(event.call.status) ?? null
					}
					const previous = steps.get(id)

					runningTool = event.call.status === "running" ? id : runningTool === id ? null : runningTool

					// The loop re-emits a call whenever its record changes; the list hears only what it shows.
					if (previous === undefined || !sameStep({ ...previous, progress: null }, step)) {
						put(step)
					}

					break
				}
				case "reply":
					// The reply itself shows from its first word on.
					if (event.final || event.text.length > 0) {
						endPhase("done")
					}

					break
			}
		},

		close: outcome => {
			if (closed) {
				return
			}

			const status = outcome === "failed" ? "failed" : "stopped"

			endPhase(outcome === "done" ? "done" : status)

			for (const step of steps.values()) {
				if (step.status === "active" || step.status === "queued" || step.status === "waiting") {
					put(endedStep(step, status))
				}
			}

			closed = true
		}
	}
}

// Settles with the promise, or rejects as soon as the signal aborts: Stop ends a turn that is still waiting
// for the model to load (the load itself carries on, the user asked for it).
function unlessAborted<T>(promise: Promise<T>, signal: AbortSignal): Promise<T> {
	return new Promise((resolve, reject) => {
		const onAbort = () => {
			reject(new AssistantAbortError())
		}

		if (signal.aborted) {
			onAbort()

			return
		}

		signal.addEventListener("abort", onAbort, { once: true })
		void promise.then(resolve, reject).finally(() => {
			signal.removeEventListener("abort", onAbort)
		})
	})
}

function nonce(): string {
	const bytes = crypto.getRandomValues(new Uint8Array(6))

	return Array.from(bytes, byte => byte.toString(16).padStart(2, "0")).join("")
}

interface PendingProposal {
	request: AssistantConfirmationRequest
	turn: TurnExecutor
	decide: (approved: boolean) => void
}

export type AssistantRuntimeSession = AssistantSession & {
	getActionLog: () => readonly AssistantActionLogEntry[]
}

// The AssistantSession the UI talks to: the shared agent loop (@filen/shared/assistant runAssistantTurn)
// over the on-device model and the app's tool executor, one turn at a time.
export function createAssistantSession({
	runtime,
	executor,
	now = Date.now
}: {
	runtime: AssistantRuntimeLike
	executor: ToolExecutor
	// The clock of the steps' `at`, which the UI compares with Date.now() for live durations.
	now?: () => number
}): AssistantRuntimeSession {
	const listeners = new Set<(event: AssistantEvent) => void>()
	const proposals = new Map<string, PendingProposal>()
	const actionLog: AssistantActionLogEntry[] = []
	let history: AssistantHistoryTurn[] = []
	// Bumped by reset: a turn of the previous conversation that still completes must not add its exchange to
	// the new one's history.
	let conversation = 0
	let turnCount = 0
	let active: AbortController | null = null

	function emit(event: AssistantEvent): void {
		for (const listener of listeners) {
			listener(event)
		}
	}

	runtime.subscribe(state => {
		emit({ type: "model", state })
	})

	function logCall(turnId: string, call: AssistantToolCallRecord): void {
		if (call.effect === "read" || !FINAL_STATUSES.has(call.status)) {
			return
		}

		actionLog.push({
			at: Date.now(),
			turnId,
			tool: call.tool,
			args: call.args,
			status: call.status,
			...(call.error === undefined ? {} : { error: call.error })
		})

		if (actionLog.length > ACTION_LOG_MAX) {
			actionLog.shift()
		}
	}

	async function runTurn({
		turnId,
		message,
		controller
	}: {
		turnId: string
		message: string
		controller: AbortController
	}): Promise<void> {
		const release = runtime.hold()
		const turnConversation = conversation
		const turn = turnExecutor({ executor, signal: controller.signal })
		let segment = 0
		let segmentHasText = false
		const steps = turnSteps({ turnId, emit, now })

		const onEvent = (event: AssistantTurnEvent) => {
			steps.event(event)

			if (event.type === "reply") {
				if (event.text.length > 0 || segmentHasText) {
					segmentHasText = true
					emit({ type: "reply", turnId, replyId: `${turnId}:r${String(segment)}`, text: event.text, final: event.final })
				}

				if (event.final && segmentHasText) {
					segment++
					segmentHasText = false
				}
			} else if (event.type === "call") {
				logCall(turnId, event.call)
			}
		}

		// Shows the card and waits for the user; aborting the turn declines it.
		const confirm = async (request: AssistantConfirmationRequest): Promise<boolean> => {
			const id = `${turnId}:${request.callId}`
			const preview = await executor.preview({ tool: request.tool, args: request.args })
			const proposal = preview.ok ? buildProposal({ id, tool: request.tool, args: request.args, preview }) : null
			const signal = controller.signal

			if (proposal === null || signal.aborted) {
				return false
			}

			return new Promise<boolean>(resolve => {
				const onAbort = () => {
					proposals.delete(id)
					resolve(false)
				}

				signal.addEventListener("abort", onAbort, { once: true })
				proposals.set(id, {
					request,
					turn,
					decide: approved => {
						signal.removeEventListener("abort", onAbort)
						proposals.delete(id)
						resolve(approved)
					}
				})
				emit({ type: "proposal", turnId, proposal })
			})
		}

		try {
			if (runtime.getState().kind !== "ready") {
				steps.loading("active")
				await unlessAborted(runtime.load(), controller.signal)
				steps.loading(runtime.getState().kind === "ready" ? "done" : "failed")
			}

			// The load failed or another tab holds the model; the panel shows that state.
			if (runtime.getState().kind !== "ready") {
				emit({ type: "turn-end", turnId, at: now(), outcome: "error", error: "model-unavailable" })

				return
			}

			const result = await runAssistantTurn({
				model: runtime.model,
				executor: turn.shared,
				history,
				message,
				signal: controller.signal,
				onEvent,
				confirm,
				nonce
			})

			if (turnConversation === conversation) {
				history = [...history, { user: message, reply: result.reply }].slice(-HISTORY_MAX)
			}

			steps.close("done")
			emit({ type: "turn-end", turnId, at: now(), outcome: "complete" })
		} catch (e) {
			const cancelled = e instanceof AssistantAbortError || controller.signal.aborted

			steps.close(cancelled ? "stopped" : "failed")

			if (cancelled) {
				emit({ type: "turn-end", turnId, at: now(), outcome: "cancelled" })
			} else if (e instanceof AssistantContextError) {
				emit({ type: "turn-end", turnId, at: now(), outcome: "error", error: "context-full" })
			} else {
				log.error("assistant", "turn failed", e)
				emit({
					type: "turn-end",
					turnId,
					at: now(),
					outcome: "error",
					error: runtime.getState().kind === "ready" ? "failed" : "model-unavailable"
				})
			}
		} finally {
			for (const [id, pending] of proposals) {
				if (pending.turn === turn) {
					pending.decide(false)
					proposals.delete(id)
				}
			}

			if (active === controller) {
				active = null
			}

			release()
		}
	}

	return {
		getModelState: () => runtime.getState(),

		subscribe: listener => {
			listeners.add(listener)

			return () => {
				listeners.delete(listener)
			}
		},

		setPanelOpen: open => {
			runtime.setPanelOpen(open)
		},

		downloadModel: () => {
			runtime.retry()
		},

		loadModel: () => {
			void runtime.load()
		},

		noteActivity: () => {
			runtime.touch()
		},

		cancelDownload: () => {
			runtime.cancelDownload()
		},

		deleteModel: async () => {
			active?.abort()
			await runtime.deleteModel()
		},

		send: message => {
			const turnId = `turn-${String(++turnCount)}`
			const controller = new AbortController()

			active?.abort()
			active = controller
			void runTurn({ turnId, message, controller })

			return turnId
		},

		cancel: () => {
			active?.abort()
		},

		confirm: proposalId => {
			const pending = proposals.get(proposalId)

			if (pending === undefined) {
				emit({ type: "proposal-settled", proposalId, outcome: "failed" })

				return
			}

			// Started here, synchronously, inside the click that confirmed it.
			const running = pending.turn.start(pending.request)

			pending.decide(true)

			const release = runtime.hold()

			void running
				.then(
					result => {
						emit({ type: "proposal-settled", proposalId, outcome: result.ok ? "done" : "failed" })
					},
					(e: unknown) => {
						log.error("assistant", "confirmed action failed", e)
						emit({ type: "proposal-settled", proposalId, outcome: "failed" })
					}
				)
				.finally(release)
		},

		reject: proposalId => {
			proposals.get(proposalId)?.decide(false)
		},

		reset: (resumed = []) => {
			active?.abort()
			conversation++
			history = resumed.slice(-HISTORY_MAX).map(({ user, reply }) => ({ user, reply }))
			executor.reset()
		},

		getHistory: () => history,

		getActionLog: () => actionLog
	}
}
