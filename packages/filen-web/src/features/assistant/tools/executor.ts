import {
	isConsequentialTool,
	type AssistantSignal,
	type AssistantTool,
	type AssistantToolFailure,
	type AssistantToolName,
	type AssistantUntrustedOutput
} from "@filen/shared/assistant"
import { log } from "@/lib/log"
import { errorLabel } from "@/lib/i18n/errorLabel"
import { findCachedListingItem } from "@/features/drive/queries/drive"
import type { AssistantHandleRef, ToolExecutor } from "@/features/assistant/runtime/executor"
import { AssistantHandleRegistry, ROOT_HANDLE, type AssistantEntityKind } from "@/features/assistant/tools/handles"
import { AssistantToolError, sanitizeStrings } from "@/features/assistant/tools/format"
import {
	chatName,
	driveName,
	entityName,
	noteTitle,
	resolveHandle,
	type ToolAdapter,
	type ToolScope
} from "@/features/assistant/tools/entities"
import { readAssistantContext } from "@/features/assistant/tools/context"
import { assistantRouter } from "@/features/assistant/tools/navigation"
import { navigateToRoute, type AssistantRouteName } from "@/features/assistant/tools/routes"
import { getItemInfoTool, listDirectoryTool, readTextFileTool, revealItemTool, searchItemsTool } from "@/features/assistant/tools/drive"
import { listNotesTool, openNoteTool, readNoteTool } from "@/features/assistant/tools/notes"
import { draftChatMessageTool, listChatsTool, openChatTool, readChatTool } from "@/features/assistant/tools/chats"
import { audioControlTool, listPlaylistsTool, openPlaylistTool, playTool } from "@/features/assistant/tools/audio"

// Which items each handle argument may name.
const HANDLE_KINDS: Partial<Record<AssistantToolName, Record<string, readonly AssistantEntityKind[]>>> = {
	list_directory: { handle: ["root", "drive"] },
	get_item_info: { handle: ["root", "drive"] },
	read_text_file: { handle: ["drive"] },
	reveal_item: { handle: ["root", "drive"] },
	set_favorite: { handles: ["drive"] },
	create_directory: { parent: ["root", "drive"] },
	rename_item: { handle: ["drive"] },
	move_items: { handles: ["drive"], destination: ["root", "drive"] },
	copy_items: { handles: ["drive"], destination: ["root", "drive"] },
	trash_items: { handles: ["drive", "note"] },
	read_note: { handle: ["note"] },
	open_note: { handle: ["note"] },
	set_note_flags: { handle: ["note"] },
	read_chat: { handle: ["chat"] },
	open_chat: { handle: ["chat"] },
	draft_chat_message: { chat: ["chat"] },
	open_playlist: { handle: ["playlist"] },
	play: { handle: ["drive", "playlist", "track"] },
	add_tracks_to_playlist: { playlist: ["playlist"], handles: ["drive", "track"] }
}

function describeContext(scope: ToolScope): string {
	const { screen, directory, selection, note, chat, playing } = readAssistantContext()
	const parts = [`screen: ${screen}`]

	if (directory !== null) {
		if (directory.uuid === null) {
			parts.push(`current directory: / (handle ${ROOT_HANDLE})`)
		} else if (directory.item !== undefined) {
			const handle = scope.show(
				directory.path === undefined
					? { kind: "drive", item: directory.item }
					: { kind: "drive", item: directory.item, path: directory.path },
				driveName(directory.item)
			)

			parts.push(`current directory: ${directory.path ?? driveName(directory.item)} (handle ${handle})`)
		}
	}

	if (selection.length > 0) {
		parts.push(
			`selected: ${selection
				.slice(0, 10)
				.map(item => `${driveName(item)} (handle ${scope.show({ kind: "drive", item }, driveName(item))})`)
				.join(", ")}${selection.length > 10 ? ` and ${String(selection.length - 10)} more` : ""}`
		)
	} else if (directory !== null) {
		parts.push("nothing selected")
	}

	if (note !== undefined) {
		parts.push(`open note: ${noteTitle(note)} (handle ${scope.show({ kind: "note", note }, noteTitle(note))})`)
	}

	if (chat !== undefined) {
		parts.push(`open chat: ${chatName(chat)} (handle ${scope.show({ kind: "chat", chat }, chatName(chat))})`)
	}

	if (playing === null) {
		parts.push("nothing playing")
	} else {
		const item = findCachedListingItem(playing.track.uuid)
		const handle =
			item === undefined
				? scope.show({ kind: "track", track: playing.track }, playing.track.name)
				: scope.show({ kind: "drive", item }, driveName(item))

		parts.push(`${playing.status === "paused" ? "paused" : "playing"}: ${playing.track.name} (handle ${handle})`)
	}

	return parts.join("; ")
}

const openRouteTool: ToolAdapter = async input => {
	const router = assistantRouter()
	const route = input["route"] as AssistantRouteName

	if (router === null) {
		throw new AssistantToolError("The app cannot navigate right now")
	}

	if (!(await navigateToRoute(router, route))) {
		throw new AssistantToolError(`"${route}" is not available in the web app`)
	}

	return { value: { opened: route }, uiEffect: { type: "navigated", pathname: router.state.location.pathname } }
}

// Phase-2 writes (favorite, create, rename, move, copy, trash, note flags, playlist adds) have no adapter
// yet: each will run the user's own action module after a confirmed card.
const ADAPTERS: Partial<Record<AssistantToolName, ToolAdapter>> = {
	get_context: (_input, scope) => Promise.resolve({ value: { context: describeContext(scope) } }),
	search_items: searchItemsTool,
	list_directory: listDirectoryTool,
	get_item_info: getItemInfoTool,
	read_text_file: readTextFileTool,
	reveal_item: revealItemTool,
	list_notes: listNotesTool,
	read_note: readNoteTool,
	open_note: openNoteTool,
	list_chats: listChatsTool,
	read_chat: readChatTool,
	open_chat: openChatTool,
	draft_chat_message: draftChatMessageTool,
	list_playlists: listPlaylistsTool,
	open_playlist: openPlaylistTool,
	play: playTool,
	audio_control: audioControlTool,
	open_route: openRouteTool
}

interface ResolvedArgs {
	args: Record<string, unknown>
	targets: AssistantHandleRef[]
	text: Record<string, string>
}

function failureOf(tool: AssistantTool, error: AssistantToolError): AssistantToolFailure & { untrusted: AssistantUntrustedOutput } {
	return {
		ok: false,
		invalid: error.invalid,
		error: error.message,
		result: sanitizeStrings({ error: error.message, ...(error.candidates.length > 0 ? { candidates: error.candidates } : {}) }, []),
		...(error.candidates.length > 0 ? { candidates: error.candidates } : {}),
		untrusted: tool.untrustedOutput === "none" ? "none" : "names"
	}
}

function notAvailable(tool: AssistantTool): AssistantToolError {
	return new AssistantToolError(`"${tool.title}" is not available yet; the user can do it in the app`)
}

export function createAssistantToolExecutor(): ToolExecutor {
	const registry = new AssistantHandleRegistry()

	function scopeFor(signal: AssistantSignal | undefined, shown: Map<string, AssistantHandleRef>): ToolScope {
		return {
			registry,
			signal,
			show: (entity, name) => {
				const ref = registry.register(entity, name)

				shown.set(ref.handle, ref)

				return ref.handle
			}
		}
	}

	// Every handle argument resolved, a path where a handle belongs replaced by the item's handle. The
	// loop validated the schema already.
	async function resolveArgs(tool: AssistantTool, input: Record<string, unknown>, scope: ToolScope): Promise<ResolvedArgs> {
		const kinds = HANDLE_KINDS[tool.name as AssistantToolName] ?? {}
		const args: Record<string, unknown> = { ...input }
		const targets: AssistantHandleRef[] = []
		const text: Record<string, string> = {}

		for (const [key, value] of Object.entries(args)) {
			const allowed = kinds[key]

			if (allowed === undefined) {
				if (typeof value === "string") {
					text[key] = value
				}

				continue
			}

			const handles: string[] = []

			for (const handle of Array.isArray(value) ? value.map(String) : [String(value)]) {
				const entity = await resolveHandle(scope, handle, allowed)
				const name = entityName(entity)
				const resolved = scope.show(entity, name)

				handles.push(resolved)
				targets.push({ handle: resolved, name })
			}

			args[key] = Array.isArray(value) ? handles : handles[0]
		}

		return { args, targets, text }
	}

	return {
		context: () => {
			const shown = new Map<string, AssistantHandleRef>()
			const description = describeContext(scopeFor(undefined, shown))

			return Promise.resolve({ description, handles: [ROOT_HANDLE, ...shown.keys()] })
		},

		check: async ({ tool, args }) => {
			const shown = new Map<string, AssistantHandleRef>()

			if (ADAPTERS[tool.name as AssistantToolName] === undefined) {
				return failureOf(tool, notAvailable(tool))
			}

			try {
				return { ok: true, args: (await resolveArgs(tool, args, scopeFor(undefined, shown))).args }
			} catch (e) {
				if (e instanceof AssistantToolError) {
					return failureOf(tool, e)
				}

				throw e
			}
		},

		preview: async ({ tool, args }) => {
			try {
				const { targets, text } = await resolveArgs(tool, args, scopeFor(undefined, new Map()))

				return { ok: true, title: tool.title, targets, text }
			} catch (e) {
				return { ok: false, error: e instanceof AssistantToolError ? e.message : errorLabel(e) }
			}
		},

		execute: async ({ tool, args, confirmed, signal }) => {
			const shown = new Map<string, AssistantHandleRef>()
			const scope = scopeFor(signal, shown)
			const adapter = ADAPTERS[tool.name as AssistantToolName]

			if (adapter === undefined) {
				return failureOf(tool, notAvailable(tool))
			}

			if (isConsequentialTool(tool) && !confirmed) {
				return failureOf(tool, new AssistantToolError("This needs the user's confirmation first"))
			}

			if (signal.aborted) {
				return failureOf(tool, new AssistantToolError("Stopped"))
			}

			try {
				const output = await adapter((await resolveArgs(tool, args, scope)).args, scope)

				return {
					ok: true,
					result: tool.untrustedOutput === "none" ? output.value : sanitizeStrings(output.value, []),
					returned: [...shown.values()],
					untrusted: tool.untrustedOutput,
					...(output.uiEffect === undefined ? {} : { uiEffect: output.uiEffect })
				}
			} catch (e) {
				if (e instanceof AssistantToolError) {
					return failureOf(tool, e)
				}

				log.warn("assistant", `tool ${tool.name} failed`, e)

				return failureOf(tool, new AssistantToolError(errorLabel(e)))
			}
		},

		name: handle => registry.get(handle)?.ref.name,

		reset: () => {
			registry.clear()
		}
	}
}
