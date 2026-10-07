import { type } from "arktype"
import { create } from "zustand"
import { kvLoadOnce, kvSetJsonQuiet } from "@/lib/storage/kvBestEffort"
import {
	applyAssistantEvent,
	beginTurn,
	setProposalStatus,
	INITIAL_TRANSCRIPT_STATE,
	type ProposalStatus,
	type TranscriptState
} from "@/features/assistant/store/transcript"
import type { AssistantEvent, AssistantModelState } from "@/features/assistant/types"

// The assistant panel's UI state. The transcript is in-memory only: it holds decrypted names and note
// or chat excerpts, so it never reaches storage. Only whether the panel is open persists (kv is wiped on
// logout, so it is per account without a prefix).
const PANEL_OPEN_KV_KEY = "assistant.v1.panelOpen"

interface AssistantStore extends TranscriptState {
	// The inline panel at desktop widths. Persisted.
	panelOpen: boolean
	// The drawer below the layout breakpoint. Never persisted: the shell closes it on every navigation.
	drawerOpen: boolean
	draft: string
	// Set by a user-initiated open of the inline panel; its first focus target takes focus once and clears
	// it. A panel restored open at boot never steals focus. The drawer has Base UI's own initial focus.
	focusRequested: boolean
	setPanelOpen: (open: boolean) => void
	setDrawerOpen: (open: boolean) => void
	setDraft: (draft: string) => void
	takeFocusRequest: () => boolean
	applyEvent: (event: AssistantEvent) => void
	beginTurn: (turnId: string, text: string) => void
	setProposalStatus: (proposalId: string, status: ProposalStatus) => void
	// A different session took over: its model state, and an empty conversation (the old one's context is gone).
	resetConversation: (modelState: AssistantModelState) => void
}

export const useAssistantStore = create<AssistantStore>((set, get) => ({
	...INITIAL_TRANSCRIPT_STATE,
	panelOpen: false,
	drawerOpen: false,
	draft: "",
	focusRequested: false,
	setPanelOpen: open => {
		panelOpenUserModified = true
		set({ panelOpen: open, focusRequested: open })
		void kvSetJsonQuiet(PANEL_OPEN_KV_KEY, open, "assistant", "panel open state")
	},
	setDrawerOpen: open => {
		set({ drawerOpen: open })
	},
	setDraft: draft => {
		set({ draft })
	},
	takeFocusRequest: () => {
		if (!get().focusRequested) {
			return false
		}

		set({ focusRequested: false })

		return true
	},
	applyEvent: event => {
		set(state => applyAssistantEvent(state, event))
	},
	beginTurn: (turnId, text) => {
		set(state => beginTurn(state, turnId, text))
	},
	setProposalStatus: (proposalId, status) => {
		set(state => ({ entries: setProposalStatus(state.entries, proposalId, status) }))
	},
	resetConversation: modelState => {
		set({ ...INITIAL_TRANSCRIPT_STATE, modelState })
	}
}))

// A toggle that lands before the persisted value loads must win over it.
let panelOpenUserModified = false

export const hydrateAssistantPanelOpen = kvLoadOnce(
	PANEL_OPEN_KV_KEY,
	type("boolean"),
	open => {
		if (!panelOpenUserModified) {
			useAssistantStore.setState({ panelOpen: open })
		}
	},
	"assistant",
	"panel open state"
)

export function toggleAssistant(narrow: boolean): void {
	const { panelOpen, drawerOpen, setPanelOpen, setDrawerOpen } = useAssistantStore.getState()

	if (narrow) {
		setDrawerOpen(!drawerOpen)

		return
	}

	setPanelOpen(!panelOpen)
}

// Whether the panel is showing at the current width; what the rail toggle reports as pressed.
export function useAssistantVisible(narrow: boolean): boolean {
	return useAssistantStore(state => (narrow ? state.drawerOpen : state.panelOpen))
}

export default useAssistantStore
