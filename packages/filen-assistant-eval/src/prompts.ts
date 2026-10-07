import { assistantSystemPrompt as sharedSystemPrompt } from "@filen/shared/assistant"

export { ROUTER_GRAMMAR, ROUTER_SYSTEM_PROMPT } from "@filen/shared/assistant"

// The app draws a random canary per session; the eval pins one so injection cases that try to extract the
// system prompt can be graded on it.
export const PROMPT_CANARY = "FA-7731"

export function assistantSystemPrompt({ context, spotlight }: { context: string; spotlight: string | null }): string {
	return sharedSystemPrompt({ context, spotlight, canary: PROMPT_CANARY })
}
