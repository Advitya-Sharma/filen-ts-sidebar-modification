import { createHash } from "node:crypto"
import {
	ASSISTANT_TOOL_GROUPS,
	assistantInstructionLines,
	ROUTER_GRAMMAR,
	ROUTER_SYSTEM_PROMPT,
	toolInputJsonSchema
} from "@filen/shared/assistant"

// What a stored result was measured against: the prompts, the model-facing tools, the case itself and the
// model. A record whose fingerprint differs measured something else, so a resumed run does it again and the
// report leaves it out; otherwise a prompt edit would re-report the numbers from before it.
const PROMPTS_AND_TOOLS = JSON.stringify({
	router: ROUTER_SYSTEM_PROMPT,
	grammar: ROUTER_GRAMMAR,
	instructions: assistantInstructionLines(null),
	tools: Object.entries(ASSISTANT_TOOL_GROUPS).map(([group, definition]) => [
		group,
		definition.tools.map(tool => ({ name: tool.name, description: tool.description, parameters: toolInputJsonSchema(tool) }))
	])
})

export function caseFingerprint({ modelId, item }: { modelId: string; item: unknown }): string {
	return createHash("sha256")
		.update(PROMPTS_AND_TOOLS)
		.update("\0")
		.update(modelId)
		.update("\0")
		.update(JSON.stringify(item, (_key, value: unknown) => (typeof value === "function" ? String(value) : value)))
		.digest("hex")
		.slice(0, 16)
}

// One JSONL line, or null for a line a crash cut short.
export function parseRecordLine<T>(line: string): T | null {
	try {
		return JSON.parse(line) as T
	} catch {
		return null
	}
}
