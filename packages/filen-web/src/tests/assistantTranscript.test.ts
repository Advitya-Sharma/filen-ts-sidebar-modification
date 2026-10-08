import { describe, expect, it } from "vitest"
import {
	applyAssistantEvent,
	beginTurn,
	pendingProposalIds,
	setProposalStatus,
	setStepsExpanded,
	stepsExpanded,
	INITIAL_TRANSCRIPT_STATE,
	type TranscriptState
} from "@/features/assistant/store/transcript"
import type { AssistantEvent, AssistantProposal, AssistantStep } from "@/features/assistant/types"

const PROPOSAL: AssistantProposal = {
	id: "p1",
	action: "move",
	targets: [{ handle: "f1", name: "a.txt" }],
	destination: "Documents",
	text: {}
}

function tool(patch: Partial<Extract<AssistantStep, { kind: "tool" }>> = {}): AssistantStep {
	return {
		id: "t1:call_1_0",
		kind: "tool",
		status: "active",
		progress: null,
		tool: "search_items",
		subject: "taxes",
		resultCount: null,
		resultMore: false,
		failure: null,
		...patch
	}
}

function phase(kind: "loading" | "thinking" | "replying", status: AssistantStep["status"], id = `t1:${kind}`): AssistantStep {
	return { id, kind, status, progress: null }
}

function fold(state: TranscriptState, events: readonly AssistantEvent[]): TranscriptState {
	return events.reduce(applyAssistantEvent, state)
}

function started(turnId = "t1"): TranscriptState {
	return beginTurn(INITIAL_TRANSCRIPT_STATE, turnId, "hello")
}

describe("beginTurn", () => {
	it("appends the user's message and marks the turn active", () => {
		const state = started()

		expect(state.entries).toEqual([{ kind: "user", id: "user:t1", turnId: "t1", text: "hello" }])
		expect(state.activeTurnId).toBe("t1")
	})

	it("expires cards still waiting from earlier turns, and only those", () => {
		const withCards = fold(started(), [
			{ type: "proposal", turnId: "t1", proposal: PROPOSAL },
			{ type: "proposal", turnId: "t1", proposal: { ...PROPOSAL, id: "p2" } }
		])
		const answered = { ...withCards, entries: setProposalStatus(withCards.entries, "p2", "cancelled") }

		expect(pendingProposalIds(answered.entries)).toEqual(["p1"])

		const next = beginTurn(answered, "t2", "again")

		expect(next.entries.flatMap(entry => (entry.kind === "proposal" ? [entry.status] : []))).toEqual(["expired", "cancelled"])
		expect(pendingProposalIds(next.entries)).toEqual([])
	})
})

describe("applyAssistantEvent", () => {
	it("expires a card still waiting when its turn ends, stopped or failed, so it can no longer be confirmed", () => {
		for (const end of [
			{ type: "turn-end", turnId: "t1", at: 0, outcome: "cancelled" },
			{ type: "turn-end", turnId: "t1", at: 0, outcome: "error", error: "failed" }
		] as const) {
			const state = fold(started(), [{ type: "proposal", turnId: "t1", proposal: PROPOSAL }, end])

			expect(state.entries.flatMap(entry => (entry.kind === "proposal" ? [entry.status] : []))).toEqual(["expired"])
			expect(pendingProposalIds(state.entries)).toEqual([])
		}
	})

	it("drops a reply that ended empty and shows a withheld one in its place", () => {
		const state = fold(started(), [
			{ type: "reply", turnId: "t1", replyId: "r1", text: "df.sort_", final: false },
			{ type: "reply", turnId: "t1", replyId: "r1", text: "", final: true },
			{ type: "reply", turnId: "t1", replyId: "r2", text: "", final: true, withheld: true }
		])

		expect(state.entries.filter(entry => entry.kind === "reply")).toEqual([
			{ kind: "reply", id: "reply:r2", turnId: "t1", text: "", final: true, withheld: true }
		])
	})

	it("replaces a reply bubble's text by reply id rather than appending a new bubble per emission", () => {
		const state = fold(started(), [
			{ type: "reply", turnId: "t1", replyId: "r1", text: "Hel", final: false },
			{ type: "reply", turnId: "t1", replyId: "r1", text: "Hello there", final: true }
		])

		expect(state.entries.filter(entry => entry.kind === "reply")).toEqual([
			{ kind: "reply", id: "reply:r1", turnId: "t1", text: "Hello there", final: true }
		])
	})

	it("keeps one steps entry right below the turn's message, with replies after it in arrival order", () => {
		const state = fold(started(), [
			{ type: "reply", turnId: "t1", replyId: "r1", text: "Let me look", final: true },
			{ type: "step", turnId: "t1", step: tool({ status: "active" }), at: 10 },
			{ type: "step", turnId: "t1", step: tool({ status: "done", resultCount: 4 }), at: 30 },
			{ type: "reply", turnId: "t1", replyId: "r2", text: "Found", final: false }
		])

		expect(state.entries.map(entry => entry.id)).toEqual(["user:t1", "steps:t1", "reply:r1", "reply:r2"])
		expect(state.entries[1]).toMatchObject({
			kind: "steps",
			startedAt: 10,
			lastAt: 30,
			outcome: null,
			steps: [{ id: "t1:call_1_0", status: "done", resultCount: 4, startedAt: 10, endedAt: 30 }]
		})
	})

	it("adds a pending card for a proposal and settles it from the runtime's outcome", () => {
		const state = fold(started(), [
			{ type: "proposal", turnId: "t1", proposal: PROPOSAL },
			{ type: "proposal-settled", proposalId: "p1", outcome: "done" }
		])

		expect(state.entries.at(-1)).toEqual({ kind: "proposal", id: "proposal:p1", turnId: "t1", proposal: PROPOSAL, status: "done" })
	})

	it("ends a completed turn without a row, freeing the composer and recording the settled turn", () => {
		const state = fold(started(), [{ type: "turn-end", turnId: "t1", at: 0, outcome: "complete" }])

		expect(state.entries).toHaveLength(1)
		expect(state.activeTurnId).toBeNull()
		expect(state.settled).toEqual({ turnId: "t1", seq: 1 })
	})

	it("freezes the streaming reply and notes a cancelled or failed turn", () => {
		const cancelled = fold(started(), [
			{ type: "reply", turnId: "t1", replyId: "r1", text: "Partial", final: false },
			{ type: "turn-end", turnId: "t1", at: 0, outcome: "cancelled" }
		])

		expect(cancelled.entries.at(-2)).toMatchObject({ kind: "reply", final: true })
		expect(cancelled.entries.at(-1)).toEqual({ kind: "turnEnd", id: "turnEnd:t1", turnId: "t1", outcome: "cancelled" })

		const failed = fold(started(), [{ type: "turn-end", turnId: "t1", at: 0, outcome: "error", error: "context-full" }])

		expect(failed.entries.at(-1)).toMatchObject({ kind: "turnEnd", outcome: "context-full" })
	})

	it("counts every settled turn, so the same turn id never reads as unchanged", () => {
		const first = fold(started(), [{ type: "turn-end", turnId: "t1", at: 0, outcome: "complete" }])
		const second = fold(beginTurn(first, "t2", "next"), [{ type: "turn-end", turnId: "t2", at: 0, outcome: "complete" }])

		expect(second.settled).toEqual({ turnId: "t2", seq: 2 })
	})

	it("tracks the model state", () => {
		const state = applyAssistantEvent(INITIAL_TRANSCRIPT_STATE, {
			type: "model",
			state: { kind: "loading", model: { id: "m", name: "M", sizeBytes: 1 }, verifying: false }
		})

		expect(state.modelState.kind).toBe("loading")
	})
})

describe("steps", () => {
	function stepsOf(state: TranscriptState) {
		const entry = state.entries.find(candidate => candidate.kind === "steps")

		if (entry?.kind !== "steps") {
			throw new Error("no steps entry")
		}

		return entry
	}

	it("keeps a step's start time across updates and ends it when it stops running", () => {
		const state = fold(started(), [
			{ type: "step", turnId: "t1", step: phase("loading", "active"), at: 100 },
			{ type: "step", turnId: "t1", step: { ...phase("loading", "active"), progress: 0.5 }, at: 900 },
			{ type: "step", turnId: "t1", step: phase("loading", "done"), at: 2100 }
		])

		expect(stepsOf(state).steps).toEqual([
			{ id: "t1:loading", kind: "loading", status: "done", progress: null, startedAt: 100, endedAt: 2100 }
		])
	})

	it("drops thinking and replying rows once they end, and discarded calls, but counts their time", () => {
		const state = fold(started(), [
			{ type: "step", turnId: "t1", step: phase("thinking", "active"), at: 0 },
			{ type: "step", turnId: "t1", step: tool({ status: "active", subject: null }), at: 500 },
			{ type: "step", turnId: "t1", step: phase("thinking", "done"), at: 500 },
			{ type: "step", turnId: "t1", step: tool({ status: "discarded" }), at: 600 },
			{ type: "step", turnId: "t1", step: phase("replying", "active", "t1:replying2"), at: 700 },
			{ type: "step", turnId: "t1", step: phase("replying", "done", "t1:replying2"), at: 4200 }
		])

		expect(stepsOf(state).steps).toEqual([])
		expect(stepsOf(state)).toMatchObject({ startedAt: 0, lastAt: 4200 })
	})

	it("opens while the turn runs, collapses once it completed, and follows the user's choice after", () => {
		const running = fold(started(), [{ type: "step", turnId: "t1", step: tool({ status: "done" }), at: 1 }])

		expect(stepsExpanded(stepsOf(running))).toBe(true)

		const done = fold(running, [{ type: "turn-end", turnId: "t1", at: 0, outcome: "complete" }])

		expect(stepsOf(done).outcome).toBe("complete")
		expect(stepsExpanded(stepsOf(done))).toBe(false)

		const opened = { ...done, entries: setStepsExpanded(done.entries, "t1", true) }

		expect(stepsExpanded(stepsOf(opened))).toBe(true)
		expect(setStepsExpanded(opened.entries, "t1", true)).toBe(opened.entries)
		expect(setStepsExpanded(opened.entries, "missing", false)).toBe(opened.entries)
	})

	it("stops what still runs when the turn is cancelled or fails, and stays open", () => {
		const state = fold(started(), [
			{ type: "step", turnId: "t1", step: tool({ status: "waiting" }), at: 5 },
			{ type: "step", turnId: "t1", step: phase("thinking", "active"), at: 8 },
			{ type: "turn-end", turnId: "t1", at: 0, outcome: "cancelled" }
		])

		expect(stepsOf(state)).toMatchObject({ outcome: "cancelled", steps: [{ status: "stopped", endedAt: 8, progress: null }] })
		expect(stepsExpanded(stepsOf(state))).toBe(true)

		const failed = fold(started(), [
			{ type: "step", turnId: "t1", step: tool({ status: "failed", failure: "blocked" }), at: 5 },
			{ type: "turn-end", turnId: "t1", at: 0, outcome: "error", error: "failed" }
		])

		expect(stepsOf(failed)).toMatchObject({ outcome: "failed", steps: [{ status: "failed", failure: "blocked" }] })
	})

	it("fails what still runs when the turn fails, rather than reading as stopped by the user", () => {
		const state = fold(started(), [
			{ type: "step", turnId: "t1", step: tool({ status: "active" }), at: 5 },
			{ type: "step", turnId: "t1", step: phase("loading", "active"), at: 6 },
			{ type: "turn-end", turnId: "t1", at: 9, outcome: "error", error: "failed" }
		])

		expect(stepsOf(state).steps).toMatchObject([
			{ kind: "tool", status: "failed", failure: "error", endedAt: 9 },
			{ kind: "loading", status: "failed", endedAt: 9 }
		])
	})

	it("counts the reply that streamed after the last step into the turn's duration", () => {
		const state = fold(started(), [
			{ type: "step", turnId: "t1", step: tool({ status: "done" }), at: 1000 },
			{ type: "step", turnId: "t1", step: phase("replying", "active", "t1:replying2"), at: 2000 },
			{ type: "step", turnId: "t1", step: phase("replying", "done", "t1:replying2"), at: 2500 },
			{ type: "reply", turnId: "t1", replyId: "t1:r0", text: "A long answer", final: true },
			{ type: "turn-end", turnId: "t1", at: 9000, outcome: "complete" }
		])

		expect(stepsOf(state)).toMatchObject({ startedAt: 1000, lastAt: 9000 })
	})

	it("ignores a step that arrives after its turn ended, so nothing restarts", () => {
		const ended = fold(started(), [
			{ type: "step", turnId: "t1", step: tool({ status: "done" }), at: 1 },
			{ type: "turn-end", turnId: "t1", at: 2, outcome: "complete" }
		])

		expect(fold(ended, [{ type: "step", turnId: "t1", step: tool({ status: "active" }), at: 3 }])).toBe(ended)
	})

	it("starts a queued call's time when it runs", () => {
		const state = fold(started(), [
			{ type: "step", turnId: "t1", step: tool({ status: "queued" }), at: 100 },
			{ type: "step", turnId: "t1", step: tool({ status: "active" }), at: 3000 },
			{ type: "step", turnId: "t1", step: tool({ status: "done" }), at: 3500 }
		])

		expect(stepsOf(state).steps).toMatchObject([{ status: "done", startedAt: 3000, endedAt: 3500 }])
	})

	it("ignores events of a turn that is not in this conversation", () => {
		const other = beginTurn(INITIAL_TRANSCRIPT_STATE, "t2", "x")
		const state = fold(other, [
			{ type: "step", turnId: "t1", step: tool(), at: 1 },
			{ type: "reply", turnId: "t1", replyId: "r", text: "late", final: true },
			{ type: "turn-end", turnId: "t1", at: 0, outcome: "cancelled" }
		])

		expect(state).toBe(other)

		const stale = { ...INITIAL_TRANSCRIPT_STATE, activeTurnId: "t1" }

		expect(applyAssistantEvent(stale, { type: "turn-end", turnId: "t1", at: 0, outcome: "cancelled" })).toEqual(
			INITIAL_TRANSCRIPT_STATE
		)
	})
})

describe("setProposalStatus", () => {
	it("returns the same array when the card is missing or already in that status", () => {
		const state = fold(started(), [{ type: "proposal", turnId: "t1", proposal: PROPOSAL }])

		expect(setProposalStatus(state.entries, "missing", "done")).toBe(state.entries)
		expect(setProposalStatus(state.entries, "p1", "pending")).toBe(state.entries)
	})
})
