// The engines expose no tokenizer, so prompt size is estimated from characters, on the high side: ASCII
// runs at about four characters per token on Qwen-family vocabularies, everything else (CJK, Cyrillic,
// Arabic, emoji) at one to two. Overestimating only drops history a little early.
export function estimateTokens(text: string): number {
	let ascii = 0
	let other = 0

	for (const char of text) {
		if (char.charCodeAt(0) < 0x80) {
			ascii++
		} else {
			other++
		}
	}

	return Math.ceil(ascii / 3 + other)
}

// Template tokens around every message (role markers, separators).
export const MESSAGE_OVERHEAD_TOKENS = 8

const MAX_STRING_CHARS = 2000
const MAX_ARRAY_ITEMS = 50

// Caps what one tool result can add to the prompt: long strings are cut and long lists shortened, with a
// note the model can relay, before the result is serialized.
export function capToolResult(value: unknown): unknown {
	if (typeof value === "string") {
		return value.length > MAX_STRING_CHARS ? `${value.slice(0, MAX_STRING_CHARS)}… (truncated)` : value
	}

	if (Array.isArray(value)) {
		const items = value.slice(0, MAX_ARRAY_ITEMS).map(capToolResult)

		return value.length > MAX_ARRAY_ITEMS ? [...items, `(${value.length - MAX_ARRAY_ITEMS} more not shown)`] : items
	}

	if (typeof value === "object" && value !== null) {
		return Object.fromEntries(Object.entries(value).map(([key, entry]) => [key, capToolResult(entry)]))
	}

	return value
}
