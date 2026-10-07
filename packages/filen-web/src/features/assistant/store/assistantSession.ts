import { useAssistantStore } from "@/features/assistant/store/useAssistantStore"
import { pendingProposalIds } from "@/features/assistant/store/transcript"
import { createMockAssistantSession } from "@/features/assistant/store/mockAssistantSession"
import type { AssistantEvent, AssistantSession } from "@/features/assistant/types"

// Where the runtime plugs in: it calls registerAssistantSession once at import, and the store follows its
// events from then on, panel open or not (a download keeps reporting progress with the panel closed).
// The UI only ever talks to the session through the functions below.
let session: AssistantSession | null = null
let unsubscribe: (() => void) | null = null
// Mounted panel surfaces. A count rather than a flag: a closing drawer's exit animation can still be
// mounted when the inline panel mounts at a breakpoint crossing.
let panelsShown = 0
// Events a session emits synchronously inside send(), held until the user's message is in the transcript
// so the turn's first activity can never land above the message that started it.
let heldDuringSend: AssistantEvent[] | null = null

export function registerAssistantSession(next: AssistantSession): () => void {
	unsubscribe?.()
	session = next
	useAssistantStore.getState().resetConversation(next.getModelState())
	unsubscribe = next.subscribe(event => {
		if (heldDuringSend !== null) {
			heldDuringSend.push(event)

			return
		}

		useAssistantStore.getState().applyEvent(event)
	})
	next.setPanelOpen(panelsShown > 0)

	return () => {
		if (session !== next) {
			return
		}

		unsubscribe?.()
		unsubscribe = null
		session = null
	}
}

export function getAssistantSession(): AssistantSession | null {
	return session
}

// Development builds render against the scripted mock until a real runtime registers over it; production
// stays on the store's initial "unavailable" state. The branch is dropped from production bundles.
if (import.meta.env.DEV) {
	registerAssistantSession(createMockAssistantSession())
}

// The panel surface mounts exactly while the panel is on screen (inline only when open, the drawer's only
// while it is open), so its mount is the open signal the runtime loads and unloads the model by. Returns
// the matching cleanup for a React ref callback.
export function trackAssistantPanelShown(): () => void {
	panelsShown++

	if (panelsShown === 1) {
		session?.setPanelOpen(true)
	}

	return () => {
		panelsShown--

		if (panelsShown === 0) {
			session?.setPanelOpen(false)
		}
	}
}

export function downloadAssistantModel(): void {
	session?.downloadModel()
}

export function cancelAssistantDownload(): void {
	session?.cancelDownload()
}

// Rejects when the runtime could not delete the files; the caller reports it.
export async function removeAssistantModel(): Promise<void> {
	await session?.deleteModel()
}

// Returns whether a turn started. Earlier unanswered cards are rejected first (beginTurn expires them in
// the transcript), so the runtime never holds a proposal the UI can no longer confirm.
export function sendAssistantMessage(text: string): boolean {
	const store = useAssistantStore.getState()
	const message = text.trim()

	if (session === null || message.length === 0 || store.activeTurnId !== null || store.modelState.kind !== "ready") {
		return false
	}

	for (const proposalId of pendingProposalIds(store.entries)) {
		session.reject(proposalId)
	}

	const held: AssistantEvent[] = []
	let turnId: string

	heldDuringSend = held

	try {
		turnId = session.send(message)
	} finally {
		heldDuringSend = null
	}

	store.beginTurn(turnId, message)

	for (const event of held) {
		store.applyEvent(event)
	}

	return true
}

export function cancelAssistantTurn(): void {
	session?.cancel()
}

// Must run inside the Confirm button's own click handler: the session starts the write synchronously
// (see AssistantSession.confirm).
export function confirmAssistantProposal(proposalId: string): void {
	if (session === null) {
		return
	}

	useAssistantStore.getState().setProposalStatus(proposalId, "running")
	session.confirm(proposalId)
}

export function rejectAssistantProposal(proposalId: string): void {
	if (session === null) {
		return
	}

	useAssistantStore.getState().setProposalStatus(proposalId, "cancelled")
	session.reject(proposalId)
}

export function startNewAssistantConversation(): void {
	if (session === null) {
		return
	}

	const store = useAssistantStore.getState()

	if (store.activeTurnId !== null) {
		session.cancel()
	}

	for (const proposalId of pendingProposalIds(store.entries)) {
		session.reject(proposalId)
	}

	session.reset()
	store.resetConversation(session.getModelState())
}
