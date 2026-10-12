import { sanitizeUntrustedText, type AssistantSignal } from "@filen/shared/assistant"
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

// Settles with `work`, or as stopped as soon as the turn is: a call waiting on an offline fetch (a listing
// query stays paused while the browser is offline) or a blocked navigation does not hold the turn, and with
// it the model.
export function unlessStopped<T>(work: Promise<T>, signal: AssistantSignal): Promise<T> {
	if (signal.aborted) {
		return Promise.reject(new AssistantToolError("Stopped"))
	}

	return new Promise((resolve, reject) => {
		const onAbort = () => {
			reject(new AssistantToolError("Stopped"))
		}

		signal.addEventListener("abort", onAbort, { once: true })
		void work.then(resolve, reject).finally(() => {
			signal.removeEventListener("abort", onAbort)
		})
	})
}

// After an await that preceded a side effect (a navigation, playback): a stopped turn does nothing more.
export function throwIfStopped(scope: { signal: { readonly aborted: boolean } | undefined }): void {
	if (scope.signal?.aborted === true) {
		throw new AssistantToolError("Stopped")
	}
}

// Every string in a result, sanitized: values mix the app's own words with names and text other people
// wrote, and sanitizing the app's own words changes nothing. The joiners a word is spelled with stay, so
// the model reads and quotes Persian or Indic text as written; the loop adds the model's own tokens.
export function sanitizeStrings(value: unknown): unknown {
	if (typeof value === "string") {
		return sanitizeUntrustedText(value, [], { keepJoiners: true })
	}

	if (Array.isArray(value)) {
		return value.map(entry => sanitizeStrings(entry))
	}

	if (typeof value === "object" && value !== null) {
		return Object.fromEntries(Object.entries(value).map(([key, entry]) => [key, sanitizeStrings(entry)]))
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

// Accents, marks and invisible characters go: a name the model copies back matches with or without a
// Persian name's zero-width non-joiner.
export function normalizeText(value: string): string {
	return value
		.normalize("NFKD")
		.replace(/[\p{M}\p{Default_Ignorable_Code_Point}]/gu, "")
		.toLowerCase()
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

// A word's singular, for a substring engine: "invoices" asks for "invoice", "boxes" for "box".
function singularStem(token: string): string {
	if (token.length > 4 && /(?:s|x|z|ch|sh)es$/u.test(token)) {
		return token.slice(0, -2)
	}

	return token.length > 3 && token.endsWith("s") && !token.endsWith("ss") ? token.slice(0, -1) : token
}

// The words a substring engine may be asked for, the likeliest to narrow the hits first: words before
// numbers, longer before shorter, each as its singular so it also finds the other form, none twice.
export function engineTokens(tokens: readonly string[], max: number): string[] {
	const ranked = tokens
		.filter(token => token.length > 0)
		.toSorted((a, b) => Number(/^\p{N}+$/u.test(a)) - Number(/^\p{N}+$/u.test(b)) || b.length - a.length)
		.map(singularStem)

	return [...new Set(ranked)].slice(0, max)
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
