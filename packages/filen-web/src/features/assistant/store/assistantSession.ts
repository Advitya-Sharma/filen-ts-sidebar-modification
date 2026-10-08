import { useAssistantStore } from "@/features/assistant/store/useAssistantStore"
import { pendingProposalIds } from "@/features/assistant/store/transcript"
import { createMockAssistantSession } from "@/features/assistant/store/mockAssistantSession"
import type { AssistantEvent, AssistantModelState, AssistantSession } from "@/features/assistant/types"

// Where the runtime plugs in: it calls registerAssistantSession once at import, and the store follows its
// events from then on, panel open or not (a download keeps reporting progress with the panel closed).
// The UI only ever talks to the session through the functions below.
let session: AssistantSession | null = null
let unsubscribe: (() => void) | null = null
// Mounted panel surfaces. A count rather than a flag: a closing drawer's exit animation can still be
// mounted when the inline panel mounts at a breakpoint crossing.
let panelsShown = 0
// Events a session emits synchronously inside send(), held until the user's message is in the transcript
// so the turn's first step can never land above the message that started it.
let heldDuringSend: AssistantEvent[] | null = null

// Streamed reply text and step updates reach the store once per frame, as one update: a reply re-renders
// its markdown on every change. Everything else lands at once, after whatever is pending, so the order
// holds; a hidden tab (no frames) catches up at the turn's end.
let pending: AssistantEvent[] = []
let cancelFlush: (() => void) | null = null

function isUrgent(event: AssistantEvent): boolean {
	return event.type !== "step" && !(event.type === "reply" && !event.final)
}

// Only the newest of consecutive texts of one reply matters. Steps all stay: they carry the times.
function supersedes(next: AssistantEvent, last: AssistantEvent | undefined): boolean {
	return next.type === "reply" && last?.type === "reply" && next.replyId === last.replyId
}

function flushPending(): void {
	cancelFlush?.()
	cancelFlush = null

	const events = pending

	pending = []
	useAssistantStore.getState().applyEvents(events)
}

function scheduleFlush(): void {
	if (cancelFlush !== null) {
		return
	}

	if (typeof requestAnimationFrame === "function") {
		const id = requestAnimationFrame(() => {
			cancelFlush = null
			flushPending()
		})

		cancelFlush = () => {
			cancelAnimationFrame(id)
		}
	} else {
		const id = setTimeout(() => {
			cancelFlush = null
			flushPending()
		}, 16)

		cancelFlush = () => {
			clearTimeout(id)
		}
	}
}

function deliver(event: AssistantEvent): void {
	if (isUrgent(event)) {
		pending.push(event)
		flushPending()

		return
	}

	if (supersedes(event, pending.at(-1))) {
		pending[pending.length - 1] = event
	} else {
		pending.push(event)
	}

	scheduleFlush()
}

function dropPending(): void {
	cancelFlush?.()
	cancelFlush = null
	pending = []
}

export function registerAssistantSession(next: AssistantSession): () => void {
	unsubscribe?.()
	dropPending()
	session = next
	useAssistantStore.getState().resetConversation(next.getModelState())
	unsubscribe = next.subscribe(event => {
		if (heldDuringSend !== null) {
			heldDuringSend.push(event)

			return
		}

		deliver(event)
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
// while it is open), so its mount is the open signal: the runtime checks the model early, but never loads
// it for a panel alone. Returns the matching cleanup for a React ref callback.
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

export function loadAssistantModel(): void {
	session?.loadModel()
}

export function noteAssistantActivity(): void {
	session?.noteActivity()
}

export type ConversationModelState = Extract<AssistantModelState, { kind: "ready" | "loading" | "downloaded" }>

// The states a message can be sent from, and so the ones the panel shows the conversation in: a model on
// disk is loaded by the turn itself.
export function modelAcceptsMessages(state: AssistantModelState): state is ConversationModelState {
	return state.kind === "ready" || state.kind === "loading" || state.kind === "downloaded"
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

	if (session === null || message.length === 0 || store.activeTurnId !== null || !modelAcceptsMessages(store.modelState)) {
		return false
	}

	for (const proposalId of pendingProposalIds(store.entries)) {
		session.reject(proposalId)
	}

	const held: AssistantEvent[] = []
	let turnId: string

	flushPending()
	heldDuringSend = held

	try {
		turnId = session.send(message)
	} finally {
		heldDuringSend = null
	}

	store.beginTurn(turnId, message)

	store.applyEvents(held)

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
	dropPending()
	store.resetConversation(session.getModelState())
}
