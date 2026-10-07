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

// Folds streamed chunks into the reply text and the tool calls the engine parsed out of the model's own
// tool-call syntax (constrained by the schema-derived grammar llama.cpp applies once a call starts).
export function accumulateChunk({ state, chunk }: { state: GenerationState; chunk: ChatCompletionChunk }): boolean {
	const delta = chunk.choices[0]?.delta
	let textChanged = false

	if (delta === undefined) {
		return false
	}

	if (typeof delta.content === "string" && delta.content.length > 0) {
		state.content += delta.content
		textChanged = true
	}

	for (const part of delta.tool_calls ?? []) {
		const call = (state.calls[part.index] ??= { id: "", name: "", arguments: "" })

		call.id = part.id ?? call.id
		call.name += part.function?.name ?? ""
		call.arguments += part.function?.arguments ?? ""
	}

	return textChanged
}
