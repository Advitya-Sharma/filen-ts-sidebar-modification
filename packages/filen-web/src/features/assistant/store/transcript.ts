import type { AssistantEvent, AssistantModelState, AssistantProposal, AssistantTurnError } from "@/features/assistant/types"

// The panel's transcript as a flat, ordered row list, folded from AssistantEvents by one pure reducer so
// the whole conversation flow is testable without React or a session.

// pending: awaiting the user. running: confirmed, the write is in flight. expired: a newer message was
// sent before the card was answered, so it can no longer be confirmed out of context.
export type ProposalStatus = "pending" | "running" | "done" | "failed" | "cancelled" | "expired"

export type TranscriptEntry =
	| { kind: "user"; id: string; turnId: string; text: string }
	| { kind: "reply"; id: string; turnId: string; text: string; final: boolean }
	| { kind: "activity"; id: string; turnId: string; tool: string; status: "running" | "done" | "failed" }
	| { kind: "proposal"; id: string; turnId: string; proposal: AssistantProposal; status: ProposalStatus }
	| { kind: "turnEnd"; id: string; turnId: string; outcome: "cancelled" | AssistantTurnError }

export interface TranscriptState {
	modelState: AssistantModelState
	entries: readonly TranscriptEntry[]
	// The turn still producing output; the composer shows Stop while it is set.
	activeTurnId: string | null
	// The last turn to finish, read by the polite live region. A counter rather than the id alone, so a
	// re-announcement of the same turn is impossible and a new one always registers as a change.
	settled: { turnId: string; seq: number } | null
}

export const INITIAL_MODEL_STATE: AssistantModelState = { kind: "unavailable", reason: "unsupported-browser", model: null }

export const INITIAL_TRANSCRIPT_STATE: TranscriptState = {
	modelState: INITIAL_MODEL_STATE,
	entries: [],
	activeTurnId: null,
	settled: null
}

// Updates land on the newest rows, so the search runs from the end.
function indexOfEntry(entries: readonly TranscriptEntry[], id: string): number {
	for (let index = entries.length - 1; index >= 0; index--) {
		if (entries[index]?.id === id) {
			return index
		}
	}

	return -1
}

function upsert(entries: readonly TranscriptEntry[], entry: TranscriptEntry): readonly TranscriptEntry[] {
	const index = indexOfEntry(entries, entry.id)

	if (index === -1) {
		return [...entries, entry]
	}

	return entries.with(index, entry)
}

// Returns the same array when nothing changed, so an untouched transcript keeps its identity.
function mapEntries(entries: readonly TranscriptEntry[], update: (entry: TranscriptEntry) => TranscriptEntry): readonly TranscriptEntry[] {
	let next: TranscriptEntry[] | null = null

	for (const [index, entry] of entries.entries()) {
		const updated = update(entry)

		if (updated !== entry) {
			next ??= entries.slice()
			next[index] = updated
		}
	}

	return next ?? entries
}

export function replyEntryId(replyId: string): string {
	return `reply:${replyId}`
}

export function proposalEntryId(proposalId: string): string {
	return `proposal:${proposalId}`
}

export function applyAssistantEvent(state: TranscriptState, event: AssistantEvent): TranscriptState {
	switch (event.type) {
		case "model": {
			return { ...state, modelState: event.state }
		}

		case "reply": {
			return {
				...state,
				entries: upsert(state.entries, {
					kind: "reply",
					id: replyEntryId(event.replyId),
					turnId: event.turnId,
					text: event.text,
					final: event.final
				})
			}
		}

		case "activity": {
			return {
				...state,
				entries: upsert(state.entries, {
					kind: "activity",
					id: `activity:${event.activity.id}`,
					turnId: event.turnId,
					tool: event.activity.tool,
					status: event.activity.status
				})
			}
		}

		case "proposal": {
			return {
				...state,
				entries: upsert(state.entries, {
					kind: "proposal",
					id: proposalEntryId(event.proposal.id),
					turnId: event.turnId,
					proposal: event.proposal,
					status: "pending"
				})
			}
		}

		case "proposal-settled": {
			return { ...state, entries: setProposalStatus(state.entries, event.proposalId, event.outcome) }
		}

		case "turn-end": {
			// A cancelled or failed turn leaves its last reply as it stood: nothing more will stream into it.
			const entries = mapEntries(state.entries, entry =>
				entry.kind === "reply" && entry.turnId === event.turnId && !entry.final ? { ...entry, final: true } : entry
			)
			const outcome = event.outcome === "error" ? event.error : event.outcome

			return {
				...state,
				entries:
					outcome === "complete"
						? entries
						: [...entries, { kind: "turnEnd", id: `turnEnd:${event.turnId}`, turnId: event.turnId, outcome }],
				activeTurnId: state.activeTurnId === event.turnId ? null : state.activeTurnId,
				settled: { turnId: event.turnId, seq: (state.settled?.seq ?? 0) + 1 }
			}
		}
	}
}

export function setProposalStatus(
	entries: readonly TranscriptEntry[],
	proposalId: string,
	status: ProposalStatus
): readonly TranscriptEntry[] {
	const index = indexOfEntry(entries, proposalEntryId(proposalId))
	const entry = entries[index]

	if (entry?.kind !== "proposal" || entry.status === status) {
		return entries
	}

	return entries.with(index, { ...entry, status })
}

// The user's message opens its turn. Cards still waiting from earlier turns expire: answering them after
// the conversation moved on would act on a request the user may no longer mean.
export function beginTurn(state: TranscriptState, turnId: string, text: string): TranscriptState {
	const entries = mapEntries(state.entries, entry =>
		entry.kind === "proposal" && entry.status === "pending" ? { ...entry, status: "expired" } : entry
	)

	return {
		...state,
		entries: [...entries, { kind: "user", id: `user:${turnId}`, turnId, text }],
		activeTurnId: turnId
	}
}

export function pendingProposalIds(entries: readonly TranscriptEntry[]): string[] {
	const ids: string[] = []

	for (const entry of entries) {
		if (entry.kind === "proposal" && entry.status === "pending") {
			ids.push(entry.proposal.id)
		}
	}

	return ids
}
