import { describe, expect, it } from "vitest"
import {
	applyAssistantEvent,
	beginTurn,
	pendingProposalIds,
	setProposalStatus,
	INITIAL_TRANSCRIPT_STATE,
	type TranscriptState
} from "@/features/assistant/store/transcript"
import type { AssistantEvent, AssistantProposal } from "@/features/assistant/types"

const PROPOSAL: AssistantProposal = {
	id: "p1",
	action: "move",
	targets: [{ handle: "f1", name: "a.txt" }],
	destination: "Documents",
	text: {}
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
			{ type: "proposal", turnId: "t1", proposal: { ...PROPOSAL, id: "p2" } },
			{ type: "turn-end", turnId: "t1", outcome: "complete" }
		])
		const answered = { ...withCards, entries: setProposalStatus(withCards.entries, "p2", "cancelled") }

		expect(pendingProposalIds(answered.entries)).toEqual(["p1"])

		const next = beginTurn(answered, "t2", "again")

		expect(next.entries.flatMap(entry => (entry.kind === "proposal" ? [entry.status] : []))).toEqual(["expired", "cancelled"])
		expect(pendingProposalIds(next.entries)).toEqual([])
	})
})

describe("applyAssistantEvent", () => {
	it("replaces a reply bubble's text by reply id rather than appending a new bubble per emission", () => {
		const state = fold(started(), [
			{ type: "reply", turnId: "t1", replyId: "r1", text: "Hel", final: false },
			{ type: "reply", turnId: "t1", replyId: "r1", text: "Hello there", final: true }
		])

		expect(state.entries.filter(entry => entry.kind === "reply")).toEqual([
			{ kind: "reply", id: "reply:r1", turnId: "t1", text: "Hello there", final: true }
		])
	})

	it("keeps replies and activities in arrival order, upserting activities by id", () => {
		const state = fold(started(), [
			{ type: "activity", turnId: "t1", activity: { id: "a1", tool: "search_items", status: "running" } },
			{ type: "reply", turnId: "t1", replyId: "r1", text: "Found it", final: true },
			{ type: "activity", turnId: "t1", activity: { id: "a1", tool: "search_items", status: "done" } },
			{ type: "reply", turnId: "t1", replyId: "r2", text: "Opening", final: false }
		])

		expect(state.entries.map(entry => entry.id)).toEqual(["user:t1", "activity:a1", "reply:r1", "reply:r2"])
		expect(state.entries[1]).toMatchObject({ kind: "activity", tool: "search_items", status: "done" })
	})

	it("adds a pending card for a proposal and settles it from the runtime's outcome", () => {
		const state = fold(started(), [
			{ type: "proposal", turnId: "t1", proposal: PROPOSAL },
			{ type: "proposal-settled", proposalId: "p1", outcome: "done" }
		])

		expect(state.entries.at(-1)).toEqual({ kind: "proposal", id: "proposal:p1", turnId: "t1", proposal: PROPOSAL, status: "done" })
	})

	it("ends a completed turn without a row, freeing the composer and recording the settled turn", () => {
		const state = fold(started(), [{ type: "turn-end", turnId: "t1", outcome: "complete" }])

		expect(state.entries).toHaveLength(1)
		expect(state.activeTurnId).toBeNull()
		expect(state.settled).toEqual({ turnId: "t1", seq: 1 })
	})

	it("freezes the streaming reply and notes a cancelled or failed turn", () => {
		const cancelled = fold(started(), [
			{ type: "reply", turnId: "t1", replyId: "r1", text: "Partial", final: false },
			{ type: "turn-end", turnId: "t1", outcome: "cancelled" }
		])

		expect(cancelled.entries.at(-2)).toMatchObject({ kind: "reply", final: true })
		expect(cancelled.entries.at(-1)).toEqual({ kind: "turnEnd", id: "turnEnd:t1", turnId: "t1", outcome: "cancelled" })

		const failed = fold(started(), [{ type: "turn-end", turnId: "t1", outcome: "error", error: "context-full" }])

		expect(failed.entries.at(-1)).toMatchObject({ kind: "turnEnd", outcome: "context-full" })
	})

	it("counts every settled turn, so the same turn id never reads as unchanged", () => {
		const first = fold(started(), [{ type: "turn-end", turnId: "t1", outcome: "complete" }])
		const second = fold(beginTurn(first, "t2", "next"), [{ type: "turn-end", turnId: "t2", outcome: "complete" }])

		expect(second.settled).toEqual({ turnId: "t2", seq: 2 })
	})

	it("tracks the model state", () => {
		const state = applyAssistantEvent(INITIAL_TRANSCRIPT_STATE, {
			type: "model",
			state: { kind: "loading", model: { id: "m", name: "M", sizeBytes: 1 } }
		})

		expect(state.modelState.kind).toBe("loading")
	})
})

describe("setProposalStatus", () => {
	it("returns the same array when the card is missing or already in that status", () => {
		const state = fold(started(), [{ type: "proposal", turnId: "t1", proposal: PROPOSAL }])

		expect(setProposalStatus(state.entries, "missing", "done")).toBe(state.entries)
		expect(setProposalStatus(state.entries, "p1", "pending")).toBe(state.entries)
	})
})
