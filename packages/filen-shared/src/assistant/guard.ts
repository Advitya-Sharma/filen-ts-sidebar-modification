import { isConsequentialTool, type AssistantTool, type AssistantUntrustedOutput } from "./tool"

// Deterministic injection policy, enforced by the dispatcher whatever the model says:
//
// 0. Scope. The router reads only the user's message, before any tool output exists, and decides
//    the tool group and whether the request asks for a change at all. A read request never gets
//    consequential tools, so injected text cannot add capabilities the user did not ask for.
// 1. Grounding. A consequential call runs only if every handle argument traces back to the user:
//    the visible context (current directory, selection, playing track), a lookup whose own
//    arguments came from the user's message, an item the user named, or an item a grounded call
//    created this turn. Free-text arguments (new names, titles, recipients) must appear in the
//    user's own message. Text that exists only inside tool results never grounds anything.
// 2. Taint. Once a tool returns content other people can write (note bodies, file text, chat
//    messages), the turn is tainted: consequential tools are withdrawn for the rest of the turn,
//    navigation needs grounded targets and open_route is withdrawn. Acting on what was read takes
//    a new request from the user.
// 3. Calls that pass still reach the user as a confirmation card naming exactly what happens;
//    there is no "yes to all". Permanent deletion does not exist as a tool at all.

export type AssistantGuardOptions = {
	// From the router: false when the user's message asks for no change.
	consequentialAllowed: boolean
	grounding: boolean
	taintGate: boolean
}

export type AssistantGuardDecision =
	| {
			allowed: true
	  }
	| {
			allowed: false
			reason: string
	  }

const TAINT_ORDER: readonly AssistantUntrustedOutput[] = ["none", "names", "content"]

function normalize(value: string): string {
	return value.normalize("NFKD").replace(/\p{M}/gu, "").toLowerCase()
}

const CJK = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u

const CJK_RUN = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]+/gu

const EXTENSION = /\.[a-z0-9]{1,5}$/i

function words(value: string): string[] {
	// CJK writes Latin words without spaces ("報告書をArchiveに"), so script changes split too.
	return normalize(value.replace(EXTENSION, ""))
		.replace(CJK_RUN, run => ` ${run} `)
		.split(/[^\p{L}\p{N}]+/u)
		.filter(word => word.length > 0)
}

function distinctive(word: string): boolean {
	return CJK.test(word) || /^\d+$/.test(word) ? word.length >= 2 : word.length >= 3
}

// "archiv" grounds "Archive", "invoices" grounds "Invoice"; containment also covers attached
// clitics (Arabic "وفواتير" = "and invoices") and compounds. CJK words match as substrings.
function wordInMessage({ word, message, messageWords }: { word: string; message: string; messageWords: readonly string[] }): boolean {
	if (CJK.test(word)) {
		return message.includes(word)
	}

	return messageWords.some(
		candidate =>
			candidate === word ||
			(!/^\d+$/.test(word) &&
				Math.min(candidate.length, word.length) >= 4 &&
				(candidate.includes(word) || word.includes(candidate)))
	)
}

export class AssistantTurnGuard {
	private readonly message: string
	private readonly messageWords: string[]
	private readonly grounded = new Set<string>()
	private readonly names = new Map<string, string>()
	private taintLevel: AssistantUntrustedOutput = "none"

	private readonly options: AssistantGuardOptions

	constructor({
		options,
		userMessage,
		contextHandles
	}: {
		options: AssistantGuardOptions
		userMessage: string
		contextHandles: readonly string[]
	}) {
		this.options = options
		this.message = normalize(userMessage)
		this.messageWords = words(userMessage)

		for (const handle of contextHandles) {
			this.grounded.add(handle)
		}
	}

	get taint(): AssistantUntrustedOutput {
		return this.taintLevel
	}

	// Every distinctive word of `text` occurs in the user's message (`share` = fraction required).
	textGrounded(text: string, share = 1): boolean {
		const candidates = words(text).filter(distinctive)

		if (candidates.length === 0) {
			return normalize(text.trim()).length > 0 && this.message.includes(normalize(text.trim()))
		}

		const hits = candidates.filter(word =>
			wordInMessage({
				word,
				message: this.message,
				messageWords: this.messageWords
			})
		).length

		return hits >= Math.ceil(candidates.length * share)
	}

	handleGrounded(handle: string): boolean {
		if (this.grounded.has(handle)) {
			return true
		}

		const name = this.names.get(handle)

		return (
			name !== undefined &&
			words(name)
				.filter(distinctive)
				.some(word =>
					wordInMessage({
						word,
						message: this.message,
						messageWords: this.messageWords
					})
				)
		)
	}

	private argumentHandles({ tool, args }: { tool: AssistantTool; args: Record<string, unknown> }): string[] {
		return (tool.grounded.handles ?? []).flatMap(key => {
			const value = args[key]

			return typeof value === "string" ? [value] : Array.isArray(value) ? value.filter(item => typeof item === "string") : []
		})
	}

	private argumentsGrounded({
		tool,
		args,
		share
	}: {
		tool: AssistantTool
		args: Record<string, unknown>
		share: number
	}): AssistantGuardDecision {
		for (const handle of this.argumentHandles({
			tool,
			args
		})) {
			if (!this.handleGrounded(handle)) {
				return {
					allowed: false,
					reason: `"${handle}" was not part of the user's request`
				}
			}
		}

		for (const key of tool.grounded.text ?? []) {
			const value = args[key]

			if (typeof value === "string" && !this.textGrounded(value, share)) {
				return {
					allowed: false,
					reason: `"${value}" does not come from the user's message`
				}
			}
		}

		return {
			allowed: true
		}
	}

	// The tools to offer for the next step.
	offered<T extends AssistantTool>(tools: readonly T[]): T[] {
		const tainted = this.options.taintGate && this.taintLevel === "content"

		return tools.filter(
			tool =>
				(this.options.consequentialAllowed && !tainted) || (!isConsequentialTool(tool) && !(tainted && tool.name === "open_route"))
		)
	}

	check({ tool, args }: { tool: AssistantTool; args: Record<string, unknown> }): AssistantGuardDecision {
		const consequential = isConsequentialTool(tool)
		const tainted = this.options.taintGate && this.taintLevel === "content"

		if (consequential && !this.options.consequentialAllowed) {
			return {
				allowed: false,
				reason: "The user's request does not ask for a change."
			}
		}

		if (tainted && (consequential || tool.name === "open_route")) {
			return {
				allowed: false,
				reason: "Not available after reading content from files, notes or chats in this request. Ask the user to request it directly."
			}
		}

		if ((consequential && this.options.grounding) || (tainted && tool.effect === "navigate")) {
			return this.argumentsGrounded({
				tool,
				args,
				share: 1
			})
		}

		return {
			allowed: true
		}
	}

	// Call after a tool ran. `returned` lists the handles its result exposed, with display names.
	record({
		tool,
		args,
		returned,
		created = []
	}: {
		tool: AssistantTool
		args: Record<string, unknown>
		returned: readonly {
			handle: string
			name: string
		}[]
		created?: readonly string[]
	}): void {
		if (TAINT_ORDER.indexOf(tool.untrustedOutput) > TAINT_ORDER.indexOf(this.taintLevel)) {
			this.taintLevel = tool.untrustedOutput
		}

		for (const { handle, name } of returned) {
			this.names.set(handle, name)
		}

		// A consequential call only got here grounded, so what it created is the user's too.
		for (const handle of created) {
			this.grounded.add(handle)
		}

		const keyed =
			(tool.grounded.handles ?? []).some(key => args[key] !== undefined) ||
			(tool.grounded.text ?? []).some(key => args[key] !== undefined)

		if (
			tool.effect === "read" &&
			keyed &&
			this.argumentsGrounded({
				tool,
				args,
				share: 0.5
			}).allowed
		) {
			for (const { handle } of returned) {
				this.grounded.add(handle)
			}
		}
	}
}
