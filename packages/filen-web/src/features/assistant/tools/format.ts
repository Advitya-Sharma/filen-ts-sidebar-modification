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

// After an await that preceded a side effect (a navigation, playback): a stopped turn does nothing more.
export function throwIfStopped(scope: { signal: { readonly aborted: boolean } | undefined }): void {
	if (scope.signal?.aborted === true) {
		throw new AssistantToolError("Stopped")
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

// The user's own calendar date, as the "Today" line in the system prompt counts days.
export function isoDate(timestamp: bigint | number): string {
	const date = new Date(Number(timestamp))

	if (Number.isNaN(date.getTime())) {
		return ""
	}

	const pad = (value: number) => String(value).padStart(2, "0")

	return `${String(date.getFullYear()).padStart(4, "0")}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`
}

export function normalizeText(value: string): string {
	return value.normalize("NFKD").replace(/\p{M}/gu, "").toLowerCase()
}

const UNSPACED =
	/[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Thai}\p{Script=Lao}\p{Script=Khmer}\p{Script=Myanmar}]/u

const KANA_PARTICLE = /^[\p{Script=Hiragana}]$/u

// Created once: a segmenter loads the locale's word data.
let segmenter: Intl.Segmenter | null = null

// The words of a query, in the case and form the user wrote them: scripts written without spaces
// ("2024年の請求書", "報告書をArchiveに") are split into words, a change of script splits too, and a lone
// kana particle is dropped.
function queryWords(value: string): string[] {
	const text = value.normalize("NFC").toLowerCase()

	if (!UNSPACED.test(text)) {
		return text.split(/[^\p{L}\p{M}\p{N}]+/u).filter(word => word.length > 0)
	}

	segmenter ??= new Intl.Segmenter(undefined, { granularity: "word" })

	return [...segmenter.segment(text)]
		.filter(segment => segment.isWordLike === true)
		.flatMap(segment =>
			segment.segment.split(
				/(?<=\p{Script=Latin}|\p{N})(?=[^\p{Script=Latin}\p{N}])|(?<=[^\p{Script=Latin}\p{N}])(?=\p{Script=Latin}|\p{N})/u
			)
		)
		.filter(word => word.length > 0 && !KANA_PARTICLE.test(word))
}

// For the app's own matching, where both sides are folded the same way (accents and marks removed).
export function searchTokens(value: string): string[] {
	return queryWords(value)
		.map(normalizeText)
		.filter(token => token.length > 0)
}

// For a matcher that compares names as written (the SDK's search engine: NFC, case-folded, nothing else
// folded): accents, Devanagari vowel signs and voiced kana stay.
export function rawSearchTokens(value: string): string[] {
	return queryWords(value)
}

// The one word a substring engine is asked for: the one likeliest to narrow the hits, a word over a number,
// the longest otherwise, as its singular stem so it also finds the other form ("invoices" asks for
// "invoice").
export function engineToken(tokens: readonly string[]): string | null {
	const ranked = tokens
		.filter(token => token.length > 0)
		.toSorted((a, b) => Number(/^\p{N}+$/u.test(a)) - Number(/^\p{N}+$/u.test(b)) || b.length - a.length)
	const token = ranked[0]

	if (token === undefined) {
		return null
	}

	if (token.length > 4 && token.endsWith("es")) {
		return token.slice(0, -2)
	}

	return token.length > 3 && token.endsWith("s") && !token.endsWith("ss") ? token.slice(0, -1) : token
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
