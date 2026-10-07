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
import type { AssistantActivity, AssistantEvent, AssistantModelState, AssistantSession } from "@/features/assistant/types"
import { buildProposal, turnExecutor, type TurnExecutor } from "@/features/assistant/loop/executorAdapter"
import { log } from "@/lib/log"

// What the session needs of the model runtime (runtime/client.ts), so tests drive it with a fake.
export interface AssistantRuntimeLike {
	readonly model: AssistantModel
	getState(): AssistantModelState
	subscribe(listener: (state: AssistantModelState) => void): () => void
	init(): Promise<void>
	setPanelOpen(open: boolean): void
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

const FINAL_STATUSES = new Set<AssistantToolCallStatus>(["ok", "error", "invalid", "blocked", "declined"])

function activityStatus(status: AssistantToolCallStatus): AssistantActivity["status"] {
	return status === "ok" ? "done" : status === "running" || status === "confirming" ? "running" : "failed"
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
	executor
}: {
	runtime: AssistantRuntimeLike
	executor: ToolExecutor
}): AssistantRuntimeSession {
	const listeners = new Set<(event: AssistantEvent) => void>()
	const proposals = new Map<string, PendingProposal>()
	const actionLog: AssistantActionLogEntry[] = []
	let history: AssistantHistoryTurn[] = []
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
		const turn = turnExecutor({ executor, signal: controller.signal })
		let segment = 0
		let segmentHasText = false
		const activities = new Map<string, AssistantActivity["status"]>()

		const onEvent = (event: AssistantTurnEvent) => {
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
				const id = `${turnId}:${event.call.id}`
				const status = activityStatus(event.call.status)

				// The loop re-emits a call on every change; the UI only hears about its status.
				if (activities.get(id) !== status) {
					activities.set(id, status)
					emit({ type: "activity", turnId, activity: { id, tool: event.call.tool, status } })
				}

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

			history = [...history, { user: message, reply: result.reply }].slice(-HISTORY_MAX)
			emit({ type: "turn-end", turnId, outcome: "complete" })
		} catch (e) {
			if (e instanceof AssistantAbortError || controller.signal.aborted) {
				emit({ type: "turn-end", turnId, outcome: "cancelled" })
			} else if (e instanceof AssistantContextError) {
				emit({ type: "turn-end", turnId, outcome: "error", error: "context-full" })
			} else {
				log.error("assistant", "turn failed", e)
				emit({
					type: "turn-end",
					turnId,
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

		reset: () => {
			active?.abort()
			history = []
			executor.reset()
		},

		getActionLog: () => actionLog
	}
}
