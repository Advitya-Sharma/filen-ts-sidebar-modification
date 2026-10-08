import type { ChatCompletionChunk, ChatCompletionMessage } from "@wllama/wllama/esm/index.js"
import type { AssistantChatMessage } from "@filen/shared/assistant"

// Between the loop's engine-neutral messages and wllama's OpenAI-shaped chat API.

export interface GenerationState {
	content: string
	calls: { id: string; name: string; arguments: string }[]
}

export function toEngineMessage(message: AssistantChatMessage): ChatCompletionMessage {
	switch (message.role) {
		case "system":
		case "user":
			return { role: message.role, content: message.content }
		case "assistant":
			return {
				role: "assistant",
				content: message.content,
				...(message.toolCalls === undefined || message.toolCalls.length === 0
					? {}
					: {
							tool_calls: message.toolCalls.map(call => ({
								id: call.id,
								type: "function" as const,
								function: { name: call.name, arguments: call.arguments }
							}))
						})
			}
		case "tool":
			return { role: "tool", content: message.content, tool_call_id: message.toolCallId }
	}
}

// What a generation reports to the page while it runs, through one callback.
export type GenerationUpdate =
	{ type: "text"; text: string } | { type: "progress"; processed: number; total: number } | { type: "call"; index: number; name: string }

export interface ChunkChanges {
	text: boolean
	// The prompt being read, the reused cache prefix left out.
	progress: { processed: number; total: number } | null
	// Calls whose name grew in this chunk, by position among the named calls (the order the result lists them).
	named: { index: number; name: string }[]
}

// Folds streamed chunks into the reply text and the tool calls the engine parsed out of the model's own
// tool-call syntax (constrained by the schema-derived grammar llama.cpp applies once a call starts).
export function accumulateChunk({ state, chunk }: { state: GenerationState; chunk: ChatCompletionChunk }): ChunkChanges {
	const changes: ChunkChanges = { text: false, progress: null, named: [] }
	const prompt = chunk.prompt_progress

	if (prompt !== undefined) {
		changes.progress = {
			processed: Math.max(0, prompt.processed - prompt.cache),
			total: Math.max(0, prompt.total - prompt.cache)
		}
	}

	const delta = chunk.choices[0]?.delta

	if (delta === undefined) {
		return changes
	}

	if (typeof delta.content === "string" && delta.content.length > 0) {
		state.content += delta.content
		changes.text = true
	}

	const renamed = new Set<number>()

	for (const part of delta.tool_calls ?? []) {
		const call = (state.calls[part.index] ??= { id: "", name: "", arguments: "" })
		const name = part.function?.name ?? ""

		call.id = part.id ?? call.id
		call.name += name
		call.arguments += part.function?.arguments ?? ""

		if (name.length > 0) {
			renamed.add(part.index)
		}
	}

	if (renamed.size > 0) {
		let position = 0

		// forEach skips the holes a sparse engine index leaves, as the result's filter drops unnamed calls.
		state.calls.forEach((call, index) => {
			if (call.name.length === 0) {
				return
			}

			if (renamed.has(index)) {
				changes.named.push({ index: position, name: call.name })
			}

			position++
		})
	}

	return changes
}
