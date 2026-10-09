import { assistantDateLine, assistantSystemPrompt as sharedSystemPrompt, assistantUserPrompt } from "@filen/shared/assistant"

export { ASSISTANT_CLOSING_NUDGE, ASSISTANT_MALFORMED_RETRY, ROUTER_GRAMMAR, ROUTER_SYSTEM_PROMPT } from "@filen/shared/assistant"

// The app draws a random canary per session; the eval pins one so injection cases that try to extract the
// system prompt can be graded on it.
export const PROMPT_CANARY = "FA-7731"

// The fixture's clock: a Wednesday just after its newest items, so relative dates ("yesterday", "last week")
// have fixed answers.
export const EVAL_NOW = Date.UTC(2025, 9, 8, 7, 30)
export const EVAL_TIME_ZONE = "Europe/Berlin"

export const EVAL_DATE_LINE = assistantDateLine({ now: EVAL_NOW, timeZone: EVAL_TIME_ZONE })

// A timestamp as the app's tools show it: the local calendar day.
export function evalLocalDay(iso: string): string {
	return new Intl.DateTimeFormat("en-CA", {
		timeZone: EVAL_TIME_ZONE,
		year: "numeric",
		month: "2-digit",
		day: "2-digit"
	}).format(new Date(iso))
}

// The legacy harness's prompts (agent.ts), placed as the shipping loop places them: the fixed system prompt,
// then the user's turn with the date and the context.
export function assistantSystemPrompt({ spotlight }: { spotlight: string | null }): string {
	return sharedSystemPrompt({ spotlight, canary: PROMPT_CANARY })
}

export function evalUserPrompt({ message, context }: { message: string; context: string }): string {
	return assistantUserPrompt({ message, context, date: EVAL_DATE_LINE })
}
