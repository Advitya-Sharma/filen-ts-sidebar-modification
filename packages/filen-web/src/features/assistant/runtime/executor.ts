import type {
	AssistantSignal,
	AssistantTool,
	AssistantToolExecutor,
	AssistantToolOutcome,
	AssistantUntrustedOutput
} from "@filen/shared/assistant"

// The web app's side of the agent loop (@filen/shared/assistant's loop): the loop owns routing, the
// guard and the model; the executor owns everything that touches the app. The implementation is
// createAssistantToolExecutor (features/assistant/tools/executor.ts): one instance per conversation,
// since the handles it hands out ("f3", "d1", "n2", "c4", "p1", "t5") live as long as the conversation
// does. UUIDs never reach the model.
//
// Per call the loop runs check() before the guard (a "/path" where a handle belongs comes back as the
// item's handle, a name as candidate handles), asks the user to confirm a consequential call (preview()
// fills the card), then execute() with `confirmed`. draft_chat_message is consequential: it writes the
// composer draft and opens the chat, and never sends.

export interface AssistantHandleRef {
	handle: string
	name: string
}

// What a navigating tool did to the screen, for the panel to reflect (e.g. collapse on narrow screens).
export type AssistantUiEffect = { type: "navigated"; pathname: string } | { type: "playback" } | { type: "drafted"; pathname: string }

// The loop's outcome plus what the panel may use. `untrusted` is the taint the result actually carries:
// at most the tool's own untrustedOutput (an error holding only names is "names" even for read_note).
// Strings other people can author are already sanitized against the known template patterns; the loop
// sanitizes again with the model's own special tokens.
export type AssistantToolExecution = AssistantToolOutcome & {
	untrusted: AssistantUntrustedOutput
	uiEffect?: AssistantUiEffect
}

// A confirmation card's content for a call check() passed.
export type AssistantToolPreview =
	| {
			ok: true
			title: string
			// The items the call acts on, resolved to their names.
			targets: AssistantHandleRef[]
			// Free-text arguments (new names, message text), verbatim.
			text: Record<string, string>
	  }
	| {
			ok: false
			error: string
	  }

export interface ToolExecutor extends AssistantToolExecutor {
	execute(call: {
		callId: string
		tool: AssistantTool
		args: Record<string, unknown>
		confirmed: boolean
		signal: AssistantSignal
	}): Promise<AssistantToolExecution>
	preview(call: { tool: AssistantTool; args: Record<string, unknown> }): Promise<AssistantToolPreview>
	// A new conversation: forgets every handle.
	reset(): void
}
