import { sanitizeUntrustedText } from "@filen/shared/assistant"
import type { AssistantHandleRef } from "@/features/assistant/runtime/executor"

// Model-facing text is English, like the system prompt; names keep the app's own labels.

// A call the app refused, worded for the model. `invalid` marks a malformed call (wrong handle kind, a
// name where a handle belongs) as opposed to a valid one that failed; `candidates` are handles whose
// names match what the model passed, so it can call again with one.
export class AssistantToolError extends Error {
	readonly invalid: boolean
	readonly candidates: AssistantHandleRef[]

	constructor(message: string, options: { invalid?: boolean; candidates?: AssistantHandleRef[] } = {}) {
		super(message)
		this.name = "AssistantToolError"
		this.invalid = options.invalid ?? false
		this.candidates = options.candidates ?? []
	}
}

// Every string in a result, sanitized: values mix the app's own words with names and text other people
// wrote, and sanitizing the app's own words changes nothing.
export function sanitizeStrings(value: unknown, specialTokens: readonly string[]): unknown {
	if (typeof value === "string") {
		return sanitizeUntrustedText(value, specialTokens)
	}

	if (Array.isArray(value)) {
		return value.map(entry => sanitizeStrings(entry, specialTokens))
	}

	if (typeof value === "object" && value !== null) {
		return Object.fromEntries(Object.entries(value).map(([key, entry]) => [key, sanitizeStrings(entry, specialTokens)]))
	}

	return value
}

// Long text cut at `max` characters, never inside a surrogate pair.
export function clipText(text: string, max: number): { text: string; truncated: boolean } {
	if (text.length <= max) {
		return { text, truncated: false }
	}

	const end = /[\uD800-\uDBFF]/.test(text.charAt(max - 1)) ? max - 1 : max

	return { text: text.slice(0, end), truncated: true }
}

export function isoDate(timestamp: bigint | number): string {
	const date = new Date(Number(timestamp))

	return Number.isNaN(date.getTime()) ? "" : date.toISOString().slice(0, 10)
}

export function normalizeText(value: string): string {
	return value.normalize("NFKD").replace(/\p{M}/gu, "").toLowerCase()
}

export function searchTokens(value: string): string[] {
	return normalizeText(value)
		.split(/[^\p{L}\p{N}]+/u)
		.filter(token => token.length > 0)
}

// "invoices" finds "Invoice_2024.pdf": names are matched word by word with crude plural folding.
export function tokenMatches(haystack: string, token: string): boolean {
	if (haystack.includes(token)) {
		return true
	}

	if (token.length > 4 && token.endsWith("es") && haystack.includes(token.slice(0, -2))) {
		return true
	}

	return token.length > 3 && token.endsWith("s") && haystack.includes(token.slice(0, -1))
}

export function matchScore(name: string, tokens: readonly string[]): number {
	const haystack = normalizeText(name)

	return tokens.filter(token => tokenMatches(haystack, token)).length
}

export function sameName(a: string, b: string): boolean {
	return normalizeText(a.trim()) === normalizeText(b.trim())
}
