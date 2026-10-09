import { createHash } from "node:crypto"
import { readFileSync } from "node:fs"
import {
	ASSISTANT_CLOSING_NUDGE,
	ASSISTANT_MALFORMED_RETRY,
	ASSISTANT_TOOL_GROUPS,
	assistantInstructionLines,
	ROUTER_GRAMMAR,
	ROUTER_SYSTEM_PROMPT,
	toolInputJsonSchema
} from "@filen/shared/assistant"
import { EVAL_DATE_LINE } from "./prompts"

// What a stored result was measured against: the prompts, the model-facing tools, the fake app's data and
// outputs, the case itself and the model. A record whose fingerprint differs measured something else, so a
// resumed run does it again and the report leaves it out; otherwise a prompt edit would re-report the
// numbers from before it.
const PROMPTS_AND_TOOLS = JSON.stringify({
	router: ROUTER_SYSTEM_PROMPT,
	grammar: ROUTER_GRAMMAR,
	instructions: assistantInstructionLines(null),
	date: EVAL_DATE_LINE,
	nudges: [ASSISTANT_MALFORMED_RETRY, ASSISTANT_CLOSING_NUDGE],
	fakeApp: ["./fakeApp.ts", "./fixture.ts"].map(file => readFileSync(new URL(file, import.meta.url), "utf8")),
	// What the shipping loop presets measure, besides the prompts above.
	loop: [
		"./loopHarness.ts",
		...["loop", "guard", "untrusted", "prompts", "budget", "activity"].map(file => `../../filen-shared/src/assistant/${file}.ts`)
	].map(file => readFileSync(new URL(file, import.meta.url), "utf8")),
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
