import { resultItemCount, toolCallSubject } from "./activity"
import { capToolResult, estimateTokens, MESSAGE_OVERHEAD_TOKENS } from "./budget"
import { AssistantTurnGuard } from "./guard"
import { assistantInstructionLines, assistantSystemPrompt, ROUTER_GRAMMAR, ROUTER_SYSTEM_PROMPT } from "./prompts"
import {
	isConsequentialTool,
	toolInputJsonSchema,
	type AssistantRoute,
	type AssistantTool,
	type AssistantToolEffect,
	type AssistantToolGroup,
	type PortableObjectSchema
} from "./tool"
import { ASSISTANT_TOOL_GROUPS } from "./tools"
import {
	mapUntrustedStrings,
	quarantinedExtractionPrompt,
	sanitizeUntrustedText,
	screenAssistantReply,
	spotlightInstruction,
	spotlightText,
	type SpotlightMode,
	type SpotlightOptions
} from "./untrusted"

// The agent loop the eval (packages/filen-assistant-eval) validated, for any engine that renders the
// model's own chat template and constrains tool calls to their JSON schema (wllama, llama.rn):
//
// 1. The router reads only the user's message and picks one tool group plus read or change, under a
//    grammar that admits nothing else.
// 2. Up to MAX_STEPS steps offer that group's tools, minus what the guard withdraws (guard.ts). Every
//    call is schema-validated, checked by the app (unknown handles, names where handles belong) and by
//    the guard; one malformed step is retried with the error, a second ends the loop.
// 3. Content other people wrote (note bodies, file text, chat messages) is summarized by a tool-less
//    quarantine call before the main loop sees it, and every untrusted string is sanitized and
//    spotlighted.
// 4. A consequential call runs only after the user confirmed it. A turn that ends without a text reply
//    gets one more, tool-less, generation. The reply is screened before the user sees any of it.

// The part of AbortSignal the loop needs; the shared package compiles without DOM types, and a DOM or
// React Native AbortSignal satisfies it as is.
export interface AssistantSignal {
	readonly aborted: boolean
	addEventListener(type: "abort", listener: () => void, options?: { once?: boolean }): void
	removeEventListener(type: "abort", listener: () => void): void
}

export type AssistantModelToolCall = {
	id: string
	name: string
	// Raw JSON object text, as the engine parsed it out of the model's tool-call syntax.
	arguments: string
}

export type AssistantChatMessage =
	| {
			role: "system" | "user"
			content: string
	  }
	| {
			role: "assistant"
			content: string
			toolCalls?: readonly AssistantModelToolCall[]
	  }
	| {
			role: "tool"
			content: string
			toolCallId: string
	  }

export type AssistantModelTool = {
	name: string
	description: string
	parameters: PortableObjectSchema
}

export type AssistantGenerateRequest = {
	messages: readonly AssistantChatMessage[]
	tools?: readonly AssistantModelTool[]
	// GBNF; never combined with tools.
	grammar?: string
	maxTokens: number
}

export type AssistantGenerateResult = {
	content: string
	toolCalls: readonly AssistantModelToolCall[]
}

// How far the engine has read the prompt, in tokens, leaving out a prefix it reused from its cache.
export type AssistantPromptProgress = {
	processed: number
	total: number
}

export interface AssistantModel {
	readonly contextTokens: number
	// The loaded model's own special-token strings, for sanitizeUntrustedText.
	readonly specialTokens: readonly string[]
	// `onText` receives the whole text content generated so far, never a delta. `onProgress` reports the
	// prompt being read, before the first token; `onToolCall` reports a tool call as soon as its complete
	// name is known, by its position among the calls this generation returns, before its arguments are
	// decoded. Both are optional for an engine that cannot tell.
	generate(
		request: AssistantGenerateRequest,
		options: {
			signal: AssistantSignal
			onText?: (text: string) => void
			onProgress?: (progress: AssistantPromptProgress) => void
			onToolCall?: (call: { index: number; name: string }) => void
		}
	): Promise<AssistantGenerateResult>
}

export type AssistantToolCandidate = {
	handle: string
	name: string
}

export type AssistantToolOutcome =
	| {
			ok: true
			// What the model gets back; serialized as JSON.
			result: unknown
			// Every handle the result shows, with its name; derived from `handle` fields when absent.
			returned?: readonly AssistantToolCandidate[]
			// Handles of items a consequential call created.
			created?: readonly string[]
	  }
	| AssistantToolFailure

export type AssistantToolFailure = {
	ok: false
	// A malformed call (schema, unknown or wrong-kind handle) as opposed to a valid call the app refused;
	// the turn's single retry is spent on these.
	invalid: boolean
	error: string
	// The tool result for the model; { error, candidates } when absent.
	result?: unknown
	// Items whose names match a name the model passed where a handle belongs.
	candidates?: readonly AssistantToolCandidate[]
}

// The app side of the loop. Implementations hand out short handles ("f3", "n2") and keep the mapping to
// real items for the session; the model never sees a UUID.
export interface AssistantToolExecutor {
	// What the user sees when the turn starts: one line for the system prompt, and the handles it names
	// (current directory, selection, playing track), which ground consequential calls.
	context(): Promise<{
		description: string
		handles: readonly string[]
	}>
	// Runs before the guard, as the eval did: resolves paths given where a handle belongs and rejects
	// handles the conversation never handed out, with candidates when a name was passed instead. A tool
	// the app cannot run yet fails here too (invalid: false). The returned args replace the model's.
	check(call: { callId: string; tool: AssistantTool; args: Record<string, unknown> }): Promise<
		| {
				ok: true
				args: Record<string, unknown>
		  }
		| AssistantToolFailure
	>
	// The name a handle this executor handed out stands for, to label a call in the UI.
	name?(handle: string): string | undefined
	// `confirmed` is true only for a consequential call the user confirmed.
	execute(call: {
		callId: string
		tool: AssistantTool
		args: Record<string, unknown>
		confirmed: boolean
		signal: AssistantSignal
	}): Promise<AssistantToolOutcome>
}

// "preparing": the model is still writing the call's arguments. "queued": written, waiting for the calls
// before it in the same step. "discarded": a call announced early that never reached the dispatcher; it did
// not run. "repeated": the same call already ran this request, so it was not run again.
export type AssistantToolCallStatus =
	"preparing" | "queued" | "running" | "confirming" | "ok" | "error" | "invalid" | "blocked" | "declined" | "discarded" | "repeated"

export type AssistantToolCallRecord = {
	id: string
	step: number
	tool: string
	title: string
	effect: AssistantToolEffect | null
	consequential: boolean
	args: Record<string, unknown>
	status: AssistantToolCallStatus
	error?: string
	// For the UI (activity.ts): what the checked call acts on, sanitized and shortened, when it names one
	// thing; and how many items a listing returned, `resultMore` when the list was cut off.
	subject?: string
	resultCount?: number
	resultMore?: boolean
}

export type AssistantConfirmationRequest = {
	callId: string
	tool: AssistantTool
	args: Record<string, unknown>
}

export type AssistantTurnPhase = "routing" | "thinking" | "tool" | "confirming" | "replying"

export type AssistantTurnEvent =
	| {
			type: "phase"
			phase: AssistantTurnPhase
	  }
	| {
			type: "route"
			route: AssistantRoute | null
			intent: "read" | "change" | null
	  }
	| {
			// The screened text of the reply being written; each event replaces the previous text until one
			// with `final`, after which the next reply text starts a new reply (text written before tool
			// calls stays as its own reply).
			type: "reply"
			text: string
			final: boolean
	  }
	| {
			type: "call"
			call: AssistantToolCallRecord
	  }
	| {
			// The prompt of the generation under way (routing, a step, a content extraction or the closing
			// reply) being read.
			type: "progress"
			progress: AssistantPromptProgress
	  }

export type AssistantHistoryTurn = {
	user: string
	reply: string
}

export type AssistantTurnResult = {
	reply: string
	route: AssistantRoute | null
	intent: "read" | "change" | null
	calls: AssistantToolCallRecord[]
	unrecovered: boolean
	stepCapHit: boolean
}

export class AssistantAbortError extends Error {
	override readonly name = "AbortError"

	constructor() {
		super("The assistant turn was cancelled")
	}
}

export class AssistantContextError extends Error {
	constructor() {
		super("The message does not fit the model's context")
	}
}

export const MAX_STEPS = 5
export const STEP_MAX_TOKENS = 384
export const QUARANTINE_MAX_TOKENS = 200
const ROUTER_MAX_TOKENS = 8

const MALFORMED_CALL =
	/<tool_call>|<\|tool_call_start\|>|<function=|\|\|call:|"name"\s*:\s*"[a-z_]+"\s*,\s*"(?:arguments|parameters)"|\b[a-z]+_[a-z_]+\s*\(\s*(?:[a-z_]+\s*=|\{)/

const MALFORMED_RETRY =
	"(Your last reply contained tool-call syntax the app could not run. Call the tool properly, or reply in plain text.)"

const CLOSING_NUDGE = "(No more tools are available in this request. Reply to the user now, in plain text.)"

const CONTENT_FIELDS = ["content", "messages"]

const SPOTLIGHT_DELIMITER = /⟦\/?[^⟦⟧\s]{1,32}⟧/g

type GenerateOptions = Parameters<AssistantModel["generate"]>[1]

// An engine may deliver a callback after its generation settled (a worker's updates and its result travel
// on different ports): by then the result supersedes it, so it is dropped rather than reopening a row or a
// reply that already ended.
async function generateSettled(
	model: AssistantModel,
	request: AssistantGenerateRequest,
	{ signal, onText, onProgress, onToolCall }: GenerateOptions
): Promise<AssistantGenerateResult> {
	let settled = false
	const live =
		<T>(callback: ((value: T) => void) | undefined) =>
		(value: T) => {
			if (!settled) {
				callback?.(value)
			}
		}

	try {
		return await model.generate(request, {
			signal,
			...(onText === undefined ? {} : { onText: live(onText) }),
			...(onProgress === undefined ? {} : { onProgress: live(onProgress) }),
			...(onToolCall === undefined ? {} : { onToolCall: live(onToolCall) })
		})
	} finally {
		settled = true
	}
}

function checkAborted(signal: AssistantSignal): void {
	if (signal.aborted) {
		throw new AssistantAbortError()
	}
}

export function parseRoute(raw: string): { route: AssistantRoute | null; intent: "read" | "change" | null } {
	const [routeWord, intentWord] = raw.trim().split(" ")
	const route =
		routeWord === "answer"
			? "answer"
			: routeWord !== undefined && Object.hasOwn(ASSISTANT_TOOL_GROUPS, routeWord)
				? (routeWord as AssistantToolGroup)
				: null

	return {
		route,
		intent: intentWord === "read" || intentWord === "change" ? intentWord : null
	}
}

// Every handle a tool result exposes, with the name or title next to it.
function handlesIn(value: unknown, found = new Map<string, string>()): Map<string, string> {
	if (Array.isArray(value)) {
		for (const item of value) {
			handlesIn(item, found)
		}
	} else if (typeof value === "object" && value !== null) {
		const record = value as Record<string, unknown>
		const handle = record["handle"]
		const name = record["name"] ?? record["title"]

		if (typeof handle === "string") {
			found.set(handle, typeof name === "string" ? name : handle)
		}

		for (const entry of Object.values(record)) {
			handlesIn(entry, found)
		}
	}

	return found
}

function failureResult(failure: AssistantToolFailure): unknown {
	return (
		failure.result ?? {
			error: failure.error,
			...(failure.candidates === undefined
				? {}
				: {
						candidates: failure.candidates
					})
		}
	)
}

function normalizeForEcho(text: string): string {
	return text.toLowerCase().replace(/\s+/g, " ")
}

const ECHO_FRAGMENT = 40

// Withholds a reply that repeats the canary or any instruction line (from its first words on), and
// defangs links the user did not type. Runs on partial replies too, so nothing unscreened is shown.
export function screenReply({
	reply,
	userMessage,
	canary,
	instructions
}: {
	reply: string
	userMessage: string
	canary: string
	instructions: readonly string[]
}): string {
	const normalized = normalizeForEcho(reply)
	const echoes = instructions.some(line => {
		const fragment = normalizeForEcho(line.replace(/^- /, "")).slice(0, ECHO_FRAGMENT)

		return fragment.length === ECHO_FRAGMENT && normalized.includes(fragment)
	})

	return screenAssistantReply({
		// An echo is withheld exactly like a canary hit.
		reply: echoes ? canary : reply,
		userMessage,
		canary
	})
}

function messageTokens(message: AssistantChatMessage): number {
	const calls = message.role === "assistant" ? (message.toolCalls ?? []).map(call => call.name + call.arguments).join("") : ""

	return estimateTokens(message.content + calls) + MESSAGE_OVERHEAD_TOKENS
}

// Keeps the prompt inside the context: older conversation turns go first, then the oldest tool results
// of this turn are replaced by a stub. The system prompt and the user's message always stay.
export function fitMessages({
	system,
	history,
	turn,
	tools,
	maxTokens,
	contextTokens
}: {
	system: AssistantChatMessage
	history: readonly AssistantChatMessage[]
	turn: readonly AssistantChatMessage[]
	tools: readonly AssistantModelTool[]
	maxTokens: number
	contextTokens: number
}): AssistantChatMessage[] {
	const budget = contextTokens - maxTokens - estimateTokens(JSON.stringify(tools)) - messageTokens(system)
	let kept = history.slice()
	const current = turn.slice()
	const total = () => [...kept, ...current].reduce((sum, message) => sum + messageTokens(message), 0)

	while (total() > budget && kept.length > 0) {
		// History alternates user / assistant; drop whole exchanges.
		kept = kept.slice(2)
	}

	for (let index = 0; index < current.length && total() > budget; index++) {
		const message = current[index]

		if (message?.role === "tool") {
			current[index] = {
				...message,
				content: JSON.stringify({
					omitted: "Result dropped to fit the context; call the tool again if it is still needed."
				})
			}
		}
	}

	if (total() > budget) {
		throw new AssistantContextError()
	}

	return [system, ...kept, ...current]
}

function toModelTools(tools: readonly AssistantTool[]): AssistantModelTool[] {
	return tools.map(tool => ({
		name: tool.name,
		description: tool.description,
		parameters: toolInputJsonSchema(tool)
	}))
}

function parseArguments(raw: string): Record<string, unknown> | null {
	try {
		const value: unknown = JSON.parse(raw.trim() === "" ? "{}" : raw)

		return typeof value === "object" && value !== null && !Array.isArray(value) ? (value as Record<string, unknown>) : null
	} catch {
		return null
	}
}

export async function runAssistantTurn({
	model,
	executor,
	history,
	message,
	signal,
	onEvent,
	confirm,
	nonce,
	spotlight = "delimit"
}: {
	model: AssistantModel
	executor: AssistantToolExecutor
	history: readonly AssistantHistoryTurn[]
	message: string
	signal: AssistantSignal
	onEvent: (event: AssistantTurnEvent) => void
	// Resolves true only when the user confirmed this exact call.
	confirm: (request: AssistantConfirmationRequest, signal: AssistantSignal) => Promise<boolean>
	// Unpredictable to other people's text: per-turn canary and spotlight marker.
	nonce: () => string
	spotlight?: SpotlightMode | "off"
}): Promise<AssistantTurnResult> {
	checkAborted(signal)
	onEvent({
		type: "phase",
		phase: "routing"
	})

	const onProgress = (progress: AssistantPromptProgress) => {
		onEvent({
			type: "progress",
			progress
		})
	}
	const routed = await generateSettled(
		model,
		{
			messages: [
				{
					role: "system",
					content: ROUTER_SYSTEM_PROMPT
				},
				{
					role: "user",
					content: message
				}
			],
			grammar: ROUTER_GRAMMAR,
			maxTokens: ROUTER_MAX_TOKENS
		},
		{
			signal,
			onProgress
		}
	)
	const { route, intent } = parseRoute(routed.content)

	onEvent({
		type: "route",
		route,
		intent
	})
	checkAborted(signal)

	const context = await executor.context()
	const canary = nonce()
	const spotlightOptions: SpotlightOptions | null =
		spotlight === "off"
			? null
			: {
					mode: spotlight,
					marker: nonce()
				}
	const instructionSpotlight = spotlightOptions === null ? null : spotlightInstruction(spotlightOptions)
	const instructions = assistantInstructionLines(instructionSpotlight)
	const guard = new AssistantTurnGuard({
		options: {
			consequentialAllowed: intent === "change",
			grounding: true,
			taintGate: true
		},
		userMessage: message,
		contextHandles: context.handles
	})
	const groupTools: readonly AssistantTool[] = route === null || route === "answer" ? [] : ASSISTANT_TOOL_GROUPS[route].tools
	const system: AssistantChatMessage = {
		role: "system",
		content: assistantSystemPrompt({
			context: context.description,
			spotlight: instructionSpotlight,
			canary
		})
	}
	const past: AssistantChatMessage[] = history.flatMap(entry => [
		{
			role: "user" as const,
			content: entry.user
		},
		{
			role: "assistant" as const,
			content: entry.reply
		}
	])
	const turn: AssistantChatMessage[] = [
		{
			role: "user",
			content: message
		}
	]
	const calls: AssistantToolCallRecord[] = []
	const specialTokens = model.specialTokens
	const screen = (reply: string) =>
		screenReply({
			// The model copies spotlight delimiters it saw around item names into its reply.
			reply: reply.replace(SPOTLIGHT_DELIMITER, ""),
			userMessage: message,
			canary,
			instructions
		})
	const nameOf = (handle: string) => executor.name?.(handle)
	const sanitized = (value: unknown) => mapUntrustedStrings(value, text => sanitizeUntrustedText(text, specialTokens))
	const untrusted = (value: unknown) =>
		spotlightOptions === null ? sanitized(value) : mapUntrustedStrings(sanitized(value), text => spotlightText(text, spotlightOptions))
	// Streams the reply as it is written, screened, unless it turns out to be tool-call syntax.
	const streamReply = (text: string) => {
		if (!MALFORMED_CALL.test(text)) {
			onEvent({
				type: "reply",
				text: screen(text.trim()),
				final: false
			})
		}
	}
	const update = (record: AssistantToolCallRecord, patch: Partial<AssistantToolCallRecord>): AssistantToolCallRecord => {
		const next = {
			...record,
			...patch
		}

		calls[calls.indexOf(record)] = next
		onEvent({
			type: "call",
			call: next
		})

		return next
	}

	// Calls that ran this turn, by tool and arguments: a small model can loop on one lookup until the step cap.
	const executedCalls = new Set<string>()
	let answer = ""
	let retryUsed = false
	let blockedSteps = 0
	let unrecovered = false
	let stepCapHit = false
	let declined = false

	for (let step = 1; step <= MAX_STEPS; step++) {
		checkAborted(signal)
		onEvent({
			type: "phase",
			phase: "thinking"
		})

		const offered = guard.offered(groupTools)
		const tools = toModelTools(offered)
		// Calls the engine named while the model was still writing their arguments, by position. Only the
		// newest is being written; those before it wait to be dispatched.
		const announced = new Map<number, AssistantToolCallRecord>()
		const queueAnnounced = (except: string | null) => {
			for (const [index, record] of announced) {
				if (record.status === "preparing" && record.id !== except) {
					const queued: AssistantToolCallRecord = {
						...record,
						status: "queued"
					}

					announced.set(index, queued)
					onEvent({
						type: "call",
						call: queued
					})
				}
			}
		}
		const announce = ({ index, name }: { index: number; name: string }) => {
			const tool = offered.find(candidate => candidate.name === name)

			if (tool === undefined || announced.has(index)) {
				return
			}

			queueAnnounced(null)

			const record: AssistantToolCallRecord = {
				id: `call_${step}_${index}`,
				step,
				tool: tool.name,
				title: tool.title,
				effect: tool.effect,
				consequential: isConsequentialTool(tool),
				args: {},
				status: "preparing"
			}

			announced.set(index, record)
			onEvent({
				type: "call",
				call: record
			})
		}
		// An announced call that never reached the dispatcher (malformed, or cut off by a declined call
		// before it) did not run.
		const discardUnfinished = (finished: ReadonlySet<string>) => {
			for (const record of announced.values()) {
				if (!finished.has(record.id)) {
					onEvent({
						type: "call",
						call: {
							...record,
							status: "discarded"
						}
					})
				}
			}
		}
		const generated = await generateSettled(
			model,
			{
				messages: fitMessages({
					system,
					history: past,
					turn,
					tools,
					maxTokens: STEP_MAX_TOKENS,
					contextTokens: model.contextTokens
				}),
				...(tools.length > 0
					? {
							tools
						}
					: {}),
				maxTokens: STEP_MAX_TOKENS
			},
			{
				signal,
				onText: streamReply,
				onProgress,
				onToolCall: announce
			}
		)

		if (generated.toolCalls.length === 0) {
			discardUnfinished(new Set())

			const text = generated.content.trim()

			if (MALFORMED_CALL.test(text)) {
				if (!retryUsed && step < MAX_STEPS) {
					retryUsed = true
					turn.push(
						{
							role: "assistant",
							content: text
						},
						{
							role: "user",
							content: MALFORMED_RETRY
						}
					)

					continue
				}

				unrecovered = true

				break
			}

			answer = text

			break
		}

		// The first call is dispatched next; the rest wait their turn.
		queueAnnounced(`call_${step}_0`)

		// Text written before the calls ("Let me look that up") stays as a reply of its own.
		if (generated.content.trim().length > 0) {
			onEvent({
				type: "reply",
				text: screen(generated.content.trim()),
				final: true
			})
		}

		const stepCalls = generated.toolCalls.map((call, index) => ({
			...call,
			id: `call_${step}_${index}`
		}))
		const results: AssistantChatMessage[] = []
		const processed: AssistantModelToolCall[] = []
		let invalidThisStep = false
		let blockedThisStep = 0

		for (const call of stepCalls) {
			checkAborted(signal)

			// Only the routed group's tools exist for this turn, whatever the engine let through; the guard
			// refuses those it withdrew.
			const tool = groupTools.find(candidate => candidate.name === call.name)
			const raw = parseArguments(call.arguments)
			let record: AssistantToolCallRecord = {
				id: call.id,
				step,
				tool: call.name,
				title: tool?.title ?? call.name,
				effect: tool?.effect ?? null,
				consequential: tool === undefined ? false : isConsequentialTool(tool),
				args: raw ?? {},
				status: "running"
			}
			let result: unknown

			calls.push(record)
			onEvent({
				type: "call",
				call: record
			})

			const parsed = tool === undefined || raw === null ? null : tool.input.safeParse(raw)

			if (tool === undefined || parsed === null || !parsed.success) {
				const error =
					tool === undefined
						? `Tool "${call.name}" is not available for this request`
						: parsed === null || parsed.success
							? "Arguments are not a JSON object"
							: parsed.error.issues.map(issue => `${issue.path.join(".") || "input"}: ${issue.message}`).join("; ")

				record = update(record, {
					status: "invalid",
					error
				})
				result = {
					error: `Invalid call: ${error}`
				}
			} else {
				const checked = await executor.check({
					callId: call.id,
					tool,
					args: parsed.data as Record<string, unknown>
				})

				if (!checked.ok) {
					record = update(record, {
						status: checked.invalid ? "invalid" : "error",
						error: checked.error
					})
					result = failureResult(checked)

					if (checked.candidates !== undefined) {
						guard.record({
							tool,
							args: parsed.data as Record<string, unknown>,
							returned: checked.candidates
						})
					}
				} else {
					const args = checked.args
					const decision = guard.check({
						tool,
						args
					})

					// A refused call shows no subject: what the guard did not let through is not repeated to the user.
					const subject = decision.allowed
						? toolCallSubject({
								tool,
								args,
								name: nameOf,
								grounded: text => guard.textGrounded(text),
								userMessage: message,
								canary,
								specialTokens
							})
						: undefined

					record = update(record, {
						args,
						...(subject === undefined
							? {}
							: {
									subject
								})
					})

					const signature = `${tool.name}:${JSON.stringify(args)}`
					const repeated = executedCalls.has(signature)

					if (!decision.allowed || repeated) {
						const reason = repeated
							? "Already called with these arguments in this request; use that result."
							: decision.allowed
								? ""
								: decision.reason

						blockedThisStep++
						record = update(record, {
							status: repeated ? "repeated" : "blocked",
							error: reason
						})
						result = {
							error: `Blocked: ${reason}`
						}
					} else {
						executedCalls.add(signature)

						let approved = true

						if (isConsequentialTool(tool)) {
							record = update(record, {
								status: "confirming"
							})
							onEvent({
								type: "phase",
								phase: "confirming"
							})
							approved = await confirm(
								{
									callId: record.id,
									tool,
									args
								},
								signal
							)
							checkAborted(signal)
						}

						if (!approved) {
							declined = true
							record = update(record, {
								status: "declined"
							})
							result = {
								error: "The user declined this action."
							}
						} else {
							if (record.status === "confirming") {
								record = update(record, {
									status: "running"
								})
							}

							onEvent({
								type: "phase",
								phase: "tool"
							})

							const outcome = await executor.execute({
								callId: record.id,
								tool,
								args,
								confirmed: isConsequentialTool(tool),
								signal
							})

							if (outcome.ok) {
								guard.record({
									tool,
									args,
									returned:
										outcome.returned ??
										[...handlesIn(outcome.result)].map(([handle, name]) => ({
											handle,
											name
										})),
									created: outcome.created ?? []
								})
								// The call reads as running until its content is extracted, which is a generation of its own.
								result =
									tool.untrustedOutput === "content"
										? await quarantine({
												model,
												request: message,
												// Sanitized first: the quarantine prompt is rendered and tokenized like any other.
												result: sanitized(outcome.result),
												signal,
												onProgress
											})
										: outcome.result

								const listed = resultItemCount({
									tool,
									result: outcome.result
								})

								record = update(record, {
									status: "ok",
									...(listed === undefined
										? {}
										: {
												resultCount: listed.count,
												resultMore: listed.more
											})
								})
							} else {
								record = update(record, {
									status: outcome.invalid ? "invalid" : "error",
									error: outcome.error
								})
								result = failureResult(outcome)

								if (outcome.candidates !== undefined) {
									guard.record({
										tool,
										args,
										returned: outcome.candidates
									})
								}
							}
						}
					}
				}

				// Also the check's errors: their candidates carry item names.
				if (tool.untrustedOutput !== "none") {
					result = untrusted(result)
				}
			}

			if (record.status === "invalid") {
				invalidThisStep = true
			}

			processed.push(call)
			results.push({
				role: "tool",
				toolCallId: call.id,
				content: JSON.stringify(capToolResult(result))
			})

			if (declined) {
				break
			}
		}

		discardUnfinished(new Set(processed.map(processedCall => processedCall.id)))
		turn.push(
			{
				role: "assistant",
				content: generated.content.trim(),
				toolCalls: processed
			},
			...results
		)

		if (declined) {
			break
		}

		if (invalidThisStep) {
			if (retryUsed) {
				unrecovered = true

				break
			}

			retryUsed = true
		}

		// Repeating a call the guard refused gets nowhere; stop and let the model explain.
		if (blockedThisStep === stepCalls.length) {
			blockedSteps++

			if (blockedSteps >= 2) {
				break
			}
		}

		if (step === MAX_STEPS) {
			stepCapHit = true
		}
	}

	// The user always gets a reply: after the step cap, a refusal or a failed retry, one tool-less turn.
	if (answer.length === 0) {
		checkAborted(signal)
		onEvent({
			type: "phase",
			phase: "replying"
		})

		const closing = await generateSettled(
			model,
			{
				messages: fitMessages({
					system,
					history: past,
					// Without it, a model whose turn so far was all tool calls tends to write another one.
					turn:
						calls.length > 0
							? [
									...turn,
									{
										role: "user",
										content: CLOSING_NUDGE
									}
								]
							: turn,
					tools: [],
					maxTokens: STEP_MAX_TOKENS,
					contextTokens: model.contextTokens
				}),
				maxTokens: STEP_MAX_TOKENS
			},
			{
				signal,
				onText: streamReply,
				onProgress
			}
		)

		answer = MALFORMED_CALL.test(closing.content) ? "" : closing.content.trim()
	}

	const reply = screen(answer)

	onEvent({
		type: "reply",
		text: reply,
		final: true
	})

	return {
		reply,
		route,
		intent,
		calls,
		unrecovered,
		stepCapHit
	}
}

// Runs the extraction over a result's content fields with no tools; only its answer reaches the main loop.
async function quarantine({
	model,
	request,
	result,
	signal,
	onProgress
}: {
	model: AssistantModel
	request: string
	result: unknown
	signal: AssistantSignal
	onProgress: (progress: AssistantPromptProgress) => void
}): Promise<unknown> {
	if (typeof result !== "object" || result === null) {
		return result
	}

	const record = result as Record<string, unknown>
	const data = CONTENT_FIELDS.filter(field => field in record)
		.map(field => (typeof record[field] === "string" ? String(record[field]) : JSON.stringify(record[field])))
		.join("\n")

	if (data.length === 0) {
		return result
	}

	const { system, user } = quarantinedExtractionPrompt({
		request,
		data
	})
	const extracted = await generateSettled(
		model,
		{
			messages: [
				{
					role: "system",
					content: system
				},
				{
					role: "user",
					content: user
				}
			],
			maxTokens: QUARANTINE_MAX_TOKENS
		},
		{
			signal,
			onProgress
		}
	)

	return {
		...Object.fromEntries(Object.entries(record).filter(([key]) => !CONTENT_FIELDS.includes(key))),
		extract: extracted.content.trim()
	}
}
