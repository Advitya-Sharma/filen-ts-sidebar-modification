// The one contract between the assistant's UI (components/, store/) and whatever drives the model (runtime/,
// tools/). The UI never imports the runtime: an implementation hands itself to registerAssistantSession
// (store/assistantSession.ts), and every state change reaches the UI as an AssistantEvent.
//
// Names, paths and free text in here are shown verbatim. Tool activity carries only the tool's name; the
// UI owns its wording (locales/en/assistant.ts).

export interface AssistantModelInfo {
	// Stable build id (cache directory, logs). Never shown.
	id: string
	// Shown on the setup screen, e.g. "Qwen3.5 4B".
	name: string
	// The whole download, every shard, in bytes.
	sizeBytes: number
}

// Why the model cannot run here. Terminal for this device, except "in-use-elsewhere": another tab holds the
// model (one tab loads it), and the runtime re-emits the real state once that tab lets go.
export type AssistantUnavailableReason = "unsupported-browser" | "insufficient-memory" | "insufficient-storage" | "in-use-elsewhere"

export type AssistantModelError = "download-failed" | "load-failed" | "out-of-memory" | "storage-full"

// "downloaded": the files are on disk but not in memory. The runtime loads them on its own once the panel
// opens (setPanelOpen), passing through "loading", and unloads them a while after it closes. `model` on
// unavailable/error is what would have been (or was being) installed, when known.
export type AssistantModelState =
	| { kind: "unavailable"; reason: AssistantUnavailableReason; model: AssistantModelInfo | null }
	| { kind: "not-downloaded"; model: AssistantModelInfo }
	| { kind: "downloading"; model: AssistantModelInfo; receivedBytes: number }
	| { kind: "downloaded"; model: AssistantModelInfo }
	| { kind: "loading"; model: AssistantModelInfo }
	| { kind: "ready"; model: AssistantModelInfo }
	| { kind: "error"; error: AssistantModelError; model: AssistantModelInfo | null }

// One tool call inside a turn, by the tool's registry name ("search_items"). Re-emitted with the same id as
// its status changes; the UI upserts by id and words it itself, with a generic line for a name it does
// not know.
export interface AssistantActivity {
	id: string
	tool: string
	status: "running" | "done" | "failed"
}

export type AssistantProposalAction =
	"draft-message" | "move" | "copy" | "trash" | "download" | "rename" | "favorite" | "unfavorite" | "create-directory" | "create-note"

// An item the call acts on, as the tool preview resolved it. The handle's leading letter picks the icon
// (f file, d directory, n note, c chat, p playlist, t track); the handle itself is never shown.
export interface AssistantProposalTarget {
	handle: string
	name: string
}

// A consequential call waiting for the user. Nothing happens until its card is confirmed. `targets` are
// every item it acts on (the card lists each name and counts them); `destination` is the target
// directory's path for move/copy; `text` holds the free-text arguments verbatim, keyed by argument name
// ("text" is a draft's message, "name" a new name).
export interface AssistantProposal {
	id: string
	action: AssistantProposalAction
	targets: readonly AssistantProposalTarget[]
	destination: string | null
	text: Readonly<Record<string, string>>
}

export type AssistantTurnError = "model-unavailable" | "context-full" | "failed"

export type AssistantEvent =
	| { type: "model"; state: AssistantModelState }
	// The FULL screened text of one reply bubble so far, not a delta: the runtime re-screens the whole reply
	// on every emission (a link the user never typed must not flash up mid-stream). A turn can hold several
	// replies interleaved with activities; `final` marks a reply that will not change again.
	| { type: "reply"; turnId: string; replyId: string; text: string; final: boolean }
	| { type: "activity"; turnId: string; activity: AssistantActivity }
	| { type: "proposal"; turnId: string; proposal: AssistantProposal }
	// A confirmed proposal finished. Rejected and expired proposals are settled by the UI itself.
	| { type: "proposal-settled"; proposalId: string; outcome: "done" | "failed" }
	| { type: "turn-end"; turnId: string; outcome: "complete" | "cancelled" }
	| { type: "turn-end"; turnId: string; outcome: "error"; error: AssistantTurnError }

export interface AssistantSession {
	// Current snapshot; changes are announced through a "model" event.
	getModelState: () => AssistantModelState
	// Returns its own unsubscribe. Events are delivered synchronously, in order.
	subscribe: (listener: (event: AssistantEvent) => void) => () => void
	// Whether the panel (or its narrow-screen drawer) is on screen. Called on every change, and once right
	// after registration. Opening is what loads a downloaded model.
	setPanelOpen: (open: boolean) => void
	// Starts (or, from "error", retries) the opt-in download and load. Called from a user click.
	downloadModel: () => void
	cancelDownload: () => void
	// Frees the disk space; settles with the model back in "not-downloaded".
	deleteModel: () => Promise<void>
	// Starts a turn and returns its id. Only called while the model is ready and no turn is running.
	send: (message: string) => string
	// Stops the running turn; it still ends with a "turn-end" event (outcome "cancelled").
	cancel: () => void
	// Runs the proposal. Called inside the trusted click handler of the Confirm button, so it MUST start the
	// work synchronously: a download started after an await has lost the user activation.
	confirm: (proposalId: string) => void
	// The user declined, or the card expired because a new message was sent first.
	reject: (proposalId: string) => void
	// Forgets the conversation so the next turn starts from an empty context.
	reset: () => void
}
