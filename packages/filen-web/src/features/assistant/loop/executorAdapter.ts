import { isConsequentialTool, type AssistantSignal, type AssistantTool, type AssistantToolExecutor } from "@filen/shared/assistant"
import type { AssistantToolExecution, AssistantToolPreview, ToolExecutor } from "@/features/assistant/runtime/executor"
import type { AssistantProposal, AssistantProposalAction } from "@/features/assistant/types"
import type { UserCalendar } from "@/features/assistant/tools/dates"

// The confirmation card a consequential tool gets. A consequential tool without one cannot run at all.
export function proposalAction({ tool, args }: { tool: string; args: Record<string, unknown> }): AssistantProposalAction | null {
	switch (tool) {
		case "draft_chat_message":
			return "draft-message"
		case "move_items":
			return "move"
		case "copy_items":
			return "copy"
		case "trash_items":
			return "trash"
		case "rename_item":
			return "rename"
		case "set_favorite":
			return args["favorite"] === true ? "favorite" : "unfavorite"
		case "create_directory":
			return "create-directory"
		case "create_note":
			return "create-note"
		default:
			return null
	}
}

// The arguments naming the directory a move, copy or create lands in.
const DESTINATION_ARGS = ["destination", "parent"]

export function buildProposal({
	id,
	tool,
	args,
	preview
}: {
	id: string
	tool: AssistantTool
	args: Record<string, unknown>
	preview: Extract<AssistantToolPreview, { ok: true }>
}): AssistantProposal | null {
	const action = proposalAction({ tool: tool.name, args })

	if (action === null) {
		return null
	}

	const destinationHandles = new Set(DESTINATION_ARGS.map(key => args[key]).filter(value => typeof value === "string"))
	const destination = preview.targets.find(target => destinationHandles.has(target.handle))

	return {
		id,
		action,
		targets: preview.targets.filter(target => !destinationHandles.has(target.handle)),
		destination: destination?.name ?? null,
		text: preview.text
	}
}

export interface TurnExecutor {
	shared: AssistantToolExecutor
	// Runs a confirmed call now, inside the confirming click; the loop's execute() then awaits this run.
	start(call: { callId: string; tool: AssistantTool; args: Record<string, unknown> }): Promise<AssistantToolExecution>
}

// The app's ToolExecutor for one turn: every call gets the turn's signal, a confirmed call can be started
// before the loop resumes, and a consequential tool the UI has no card for is refused at the check.
export function turnExecutor({
	executor,
	signal,
	calendar
}: {
	executor: ToolExecutor
	signal: AssistantSignal
	calendar: UserCalendar
}): TurnExecutor {
	const started = new Map<string, Promise<AssistantToolExecution>>()

	const run = (call: { callId: string; tool: AssistantTool; args: Record<string, unknown>; confirmed: boolean }) =>
		executor.execute({ ...call, signal, calendar })

	return {
		shared: {
			context: () => executor.context(),

			name: handle => executor.name?.(handle),

			// A consequential tool also needs a card the UI can show.
			supports: tool =>
				(executor.supports?.(tool) ?? true) &&
				(!isConsequentialTool(tool) || proposalAction({ tool: tool.name, args: {} }) !== null),

			check: async call => {
				const checked = await executor.check(call)

				if (checked.ok && isConsequentialTool(call.tool) && proposalAction({ tool: call.tool.name, args: checked.args }) === null) {
					const error = `"${call.tool.title}" is not available yet; the user can do it in the app`

					return { ok: false, invalid: false, error, result: { error } }
				}

				return checked
			},

			execute: async call => {
				const running = started.get(call.callId)

				started.delete(call.callId)

				return running ?? run(call)
			}
		},

		start: call => {
			const running = run({ ...call, confirmed: true })

			started.set(call.callId, running)

			return running
		}
	}
}
