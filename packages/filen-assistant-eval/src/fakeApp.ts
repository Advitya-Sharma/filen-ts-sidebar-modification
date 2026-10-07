import { assistantToolByName } from "@filen/shared/assistant"
import { createFixtureState, type AppContext, type AppState, type DriveItem, type Note, type NoteType } from "./fixture"

export type Candidate = {
	handle: string
	name: string
}

export type ToolOutcome =
	| {
			ok: true
			result: unknown
	  }
	| {
			ok: false
			// A malformed call (schema, unknown or wrong-kind handle) as opposed to a valid call the app refused.
			invalid: boolean
			error: string
			// Items whose names match a name the model passed where a handle belongs.
			candidates?: Candidate[]
	  }

class CallError extends Error {
	readonly invalid: boolean
	readonly candidates: Candidate[] | undefined

	constructor(message: string, invalid: boolean, candidates?: Candidate[]) {
		super(message)
		this.invalid = invalid
		this.candidates = candidates
	}
}

type EntityKind = "drive" | "note" | "chat" | "playlist"

const LOOKUP_TOOL: Record<EntityKind, string> = {
	drive: "search_items or list_directory",
	note: "list_notes",
	chat: "list_chats",
	playlist: "list_playlists"
}

const EXTENSIONS: Record<string, readonly string[]> = {
	image: ["jpg", "jpeg", "png", "heic", "gif", "webp", "svg"],
	video: ["mp4", "mov", "mkv", "webm"],
	audio: ["mp3", "flac", "m4a", "wav", "ogg"],
	document: ["pdf", "doc", "docx", "odt", "txt", "md", "xlsx", "csv", "pptx"]
}

const TEXT_EXTENSIONS = ["txt", "md", "csv", "json"]

const SEARCH_LIMIT = 12

export function normalizeText(value: string): string {
	return value.normalize("NFKD").replace(/\p{M}/gu, "").toLowerCase()
}

function tokens(value: string): string[] {
	return normalizeText(value)
		.split(/[^\p{L}\p{N}]+/u)
		.filter(token => token.length > 0)
}

// Crude plural folding so "invoices" finds "Invoice_…"; the real search is name-based too.
function tokenMatches({ haystack, token }: { haystack: string; token: string }): boolean {
	if (haystack.includes(token)) {
		return true
	}

	if (token.length > 4 && token.endsWith("es") && haystack.includes(token.slice(0, -2))) {
		return true
	}

	return token.length > 3 && token.endsWith("s") && haystack.includes(token.slice(0, -1))
}

function extension(name: string): string {
	const dot = name.lastIndexOf(".")

	return dot === -1 ? "" : name.slice(dot + 1).toLowerCase()
}

function sameName(a: string, b: string): boolean {
	return normalizeText(a.trim()) === normalizeText(b.trim())
}

export class FakeApp {
	readonly state: AppState
	readonly exposed = new Set<string>()

	constructor(context: Partial<AppContext> = {}) {
		this.state = createFixtureState(context)

		const visible = [
			"d0",
			this.state.context.directory,
			...this.state.context.selection,
			...(this.state.context.playing === null
				? []
				: [
						this.state.context.playing.handle,
						...(this.state.context.playing.playlist === null ? [] : [this.state.context.playing.playlist])
					])
		]

		for (const handle of visible) {
			this.exposed.add(handle)
		}
	}

	// One line for the system prompt, so "this file" or "here" resolve without a get_context round trip.
	describeContext(): string {
		const { screen, directory, selection, playing } = this.state.context
		const parts = [`screen: ${screen}`, `current directory: ${this.pathOf(this.drive(directory))} (handle ${directory})`]

		parts.push(
			selection.length === 0
				? "nothing selected"
				: `selected: ${selection.map(handle => `${this.drive(handle).name} (handle ${handle})`).join(", ")}`
		)

		if (playing === null) {
			parts.push("nothing playing")
		} else {
			parts.push(
				`${playing.paused ? "paused" : "playing"}: ${this.drive(playing.handle).name} (handle ${playing.handle})${
					playing.playlist === null ? "" : ` from playlist handle ${playing.playlist}`
				}`
			)
		}

		return parts.join("; ")
	}

	execute({ name, args }: { name: string; args: unknown }): ToolOutcome {
		const tool = assistantToolByName(name)

		if (tool === undefined) {
			return {
				ok: false,
				invalid: true,
				error: `Unknown tool "${name}"`
			}
		}

		const parsed = tool.input.safeParse(args ?? {})

		if (!parsed.success) {
			return {
				ok: false,
				invalid: true,
				error: parsed.error.issues.map(issue => `${issue.path.join(".") || "input"}: ${issue.message}`).join("; ")
			}
		}

		try {
			return {
				ok: true,
				result: this.run({
					name,
					input: parsed.data as Record<string, unknown>
				})
			}
		} catch (e) {
			if (e instanceof CallError) {
				return {
					ok: false,
					invalid: e.invalid,
					error: e.message,
					...(e.candidates === undefined || e.candidates.length === 0
						? {}
						: {
								candidates: e.candidates
							})
				}
			}

			throw e
		}
	}

	private directoryAt(path: string): DriveItem {
		const wanted = normalizeText(path.trim().replace(/\/+$/, "") || "/")
		const found = this.state.drive.find(
			item => item.kind === "directory" && !item.trashed && normalizeText(this.pathOf(item)) === wanted
		)

		if (found === undefined) {
			throw new CallError(`No directory at "${path}"`, false)
		}

		return found
	}

	// Paths are unambiguous, so a path where a handle belongs is resolved instead of rejected.
	resolvePaths(args: Record<string, unknown>): Record<string, unknown> {
		const resolve = (value: unknown): unknown => {
			if (typeof value !== "string" || !value.startsWith("/")) {
				return value
			}

			const wanted = normalizeText(value.trim().replace(/\/+$/, "") || "/")
			const found = this.state.drive.find(item => !item.trashed && normalizeText(this.pathOf(item)) === wanted)

			if (found === undefined) {
				return value
			}

			this.exposed.add(found.handle)

			return found.handle
		}

		return Object.fromEntries(
			Object.entries(args).map(([key, value]) => [
				key,
				["handle", "handles", "parent", "destination"].includes(key)
					? Array.isArray(value)
						? value.map(resolve)
						: resolve(value)
					: value
			])
		)
	}

	// Rejects handles the conversation has not seen before the guard judges the call, so a
	// hallucinated handle counts as an invalid output rather than a blocked one.
	checkHandles({ tool, args }: { tool: string; args: Record<string, unknown> }): ToolOutcome | null {
		const kindOf = (key: string, value: string): EntityKind => {
			if (key === "chat" || tool.includes("chat")) {
				return "chat"
			}

			if (key === "playlist" || tool === "open_playlist" || (tool === "play" && value.startsWith("p"))) {
				return "playlist"
			}

			return tool.includes("note") || (tool === "trash_items" && value.startsWith("n")) ? "note" : "drive"
		}

		try {
			for (const key of ["handle", "handles", "parent", "destination", "chat", "playlist"]) {
				const value = args[key]

				for (const handle of typeof value === "string" ? [value] : Array.isArray(value) ? value.map(String) : []) {
					this.known(handle, kindOf(key, handle))
				}
			}
		} catch (e) {
			if (e instanceof CallError) {
				return {
					ok: false,
					invalid: true,
					error: e.message,
					...(e.candidates === undefined || e.candidates.length === 0
						? {}
						: {
								candidates: e.candidates
							})
				}
			}

			throw e
		}

		return null
	}

	pathOf(item: DriveItem): string {
		const names: string[] = []
		let current: DriveItem | undefined = item

		while (current !== undefined && current.parent !== null) {
			names.unshift(current.name)
			current = this.state.drive.find(candidate => candidate.handle === current?.parent)
		}

		return `/${names.join("/")}`
	}

	private drive(handle: string): DriveItem {
		const item = this.state.drive.find(candidate => candidate.handle === handle)

		if (item === undefined) {
			throw new CallError(`Unknown handle "${handle}"`, true)
		}

		return item
	}

	// Small models often pass a name where a handle belongs; answering with the matching handles
	// turns that into a recoverable error instead of a dead end.
	private known(handle: string, kind: EntityKind): void {
		if (this.exposed.has(handle)) {
			return
		}

		const wanted = normalizeText(handle.trim())
		const pool: Candidate[] =
			kind === "drive"
				? this.state.drive
						.filter(item => !item.trashed && item.parent !== null)
						.map(item => ({
							handle: item.handle,
							name: item.name
						}))
				: kind === "note"
					? this.state.notes
							.filter(item => !item.trashed)
							.map(item => ({
								handle: item.handle,
								name: item.title
							}))
					: (kind === "chat" ? this.state.chats : this.state.playlists).map(item => ({
							handle: item.handle,
							name: item.name
						}))
		const candidates = /^[a-z]\d+$/.test(handle)
			? []
			: pool
					.filter(
						item =>
							wanted.length > 1 && (normalizeText(item.name).includes(wanted) || wanted.includes(normalizeText(item.name)))
					)
					.slice(0, 5)

		for (const candidate of candidates) {
			this.exposed.add(candidate.handle)
		}

		throw new CallError(
			candidates.length > 0
				? `"${handle}" is a name, not a handle. Matching handles: ${candidates.map(item => `${item.handle} "${item.name}"`).join(", ")}. Call again with the handle.`
				: `Unknown handle "${handle}". Find it first with ${LOOKUP_TOOL[kind]}.`,
			true,
			candidates
		)
	}

	private driveItem(handle: string): DriveItem {
		this.known(handle, "drive")

		const item = this.state.drive.find(candidate => candidate.handle === handle && !candidate.trashed)

		if (item === undefined) {
			throw new CallError(`"${handle}" is not a file or directory`, true)
		}

		return item
	}

	private directory(handle: string): DriveItem {
		const item = this.driveItem(handle)

		if (item.kind !== "directory") {
			throw new CallError(`"${handle}" is a file, not a directory`, true)
		}

		return item
	}

	private noteItem(handle: string): Note {
		this.known(handle, "note")

		const item = this.state.notes.find(candidate => candidate.handle === handle && !candidate.trashed)

		if (item === undefined) {
			throw new CallError(`"${handle}" is not a note`, true)
		}

		return item
	}

	private summary(item: DriveItem): Record<string, unknown> {
		this.exposed.add(item.handle)

		return {
			handle: item.handle,
			name: item.name,
			kind: item.kind,
			path: this.pathOf(item),
			...(item.sharedBy === undefined
				? {}
				: {
						sharedBy: item.sharedBy
					})
		}
	}

	private noteSummary(item: Note): Record<string, unknown> {
		this.exposed.add(item.handle)

		return {
			handle: item.handle,
			title: item.title,
			type: item.type,
			...(item.tags.length > 0
				? {
						tags: item.tags
					}
				: {}),
			...(item.pinned
				? {
						pinned: true
					}
				: {}),
			...(item.archived
				? {
						archived: true
					}
				: {})
		}
	}

	private children(parent: string): DriveItem[] {
		return this.state.drive.filter(item => item.parent === parent && !item.trashed)
	}

	private assertFreeName({ parent, name, except }: { parent: string; name: string; except?: string }): void {
		if (name.trim().length === 0 || /[/\\]/.test(name)) {
			throw new CallError(`"${name}" is not a valid name`, false)
		}

		if (this.children(parent).some(item => item.handle !== except && sameName(item.name, name))) {
			throw new CallError(`An item named "${name}" already exists there`, false)
		}
	}

	private nextHandle(prefix: string): string {
		const pool = prefix === "n" ? this.state.notes : this.state.drive
		const highest = Math.max(
			...pool.filter(item => item.handle.startsWith(prefix)).map(item => Number(item.handle.slice(prefix.length)))
		)

		return `${prefix}${highest + 1}`
	}

	private search({ query, kind }: { query: string; kind: string | undefined }): unknown {
		const queryTokens = tokens(query)
		const scored = this.state.drive
			.filter(item => item.parent !== null && !item.trashed)
			.filter(item => {
				if (kind === undefined) {
					return true
				}

				if (kind === "directory" || kind === "file") {
					return item.kind === kind
				}

				return item.kind === "file" && (EXTENSIONS[kind] ?? []).includes(extension(item.name))
			})
			.map(item => {
				const haystack = normalizeText(item.name)

				return {
					item,
					score: queryTokens.filter(token =>
						tokenMatches({
							haystack,
							token
						})
					).length
				}
			})
			.filter(({ score }) => score > 0)

		const all = scored.filter(({ score }) => score === queryTokens.length)
		const ranked = (all.length > 0 ? all : scored.sort((a, b) => b.score - a.score)).slice(0, SEARCH_LIMIT)

		return {
			results: ranked.map(({ item }) => this.summary(item)),
			...(scored.length > SEARCH_LIMIT
				? {
						more: true
					}
				: {})
		}
	}

	private run({ name, input }: { name: string; input: Record<string, unknown> }): unknown {
		const str = (key: string): string => String(input[key])
		const optionalStr = (key: string): string | undefined => (typeof input[key] === "string" ? input[key] : undefined)
		const list = (key: string): string[] => (Array.isArray(input[key]) ? input[key].map(String) : [])

		switch (name) {
			case "get_context": {
				return {
					context: this.describeContext()
				}
			}

			case "search_items": {
				return this.search({
					query: str("query"),
					kind: optionalStr("kind")
				})
			}

			case "list_directory": {
				// Models fill both optional fields or put a path into "handle"; take whichever resolves.
				const handleArg = optionalStr("handle")
				const path = optionalStr("path") ?? (handleArg?.startsWith("/") === true ? handleArg : undefined)
				const target =
					path !== undefined
						? this.directoryAt(path)
						: handleArg !== undefined
							? this.directory(handleArg)
							: this.directoryAt("/")

				this.exposed.add(target.handle)

				return {
					directory: {
						handle: target.handle,
						path: this.pathOf(target)
					},
					items: this.children(target.handle).map(item => {
						this.exposed.add(item.handle)

						return {
							handle: item.handle,
							name: item.name,
							kind: item.kind,
							...(item.sharedBy === undefined
								? {}
								: {
										sharedBy: item.sharedBy
									})
						}
					})
				}
			}

			case "get_item_info": {
				const item = this.driveItem(str("handle"))

				return {
					...this.summary(item),
					size: item.kind === "file" ? item.size : this.children(item.handle).length + " items",
					modified: item.modified,
					favorite: item.favorite,
					shared: item.shared
				}
			}

			case "read_text_file": {
				const item = this.driveItem(str("handle"))

				if (item.kind !== "file" || !TEXT_EXTENSIONS.includes(extension(item.name)) || item.content === undefined) {
					throw new CallError(`"${item.name}" is not a text file`, false)
				}

				return {
					handle: item.handle,
					name: item.name,
					content: item.content
				}
			}

			case "reveal_item": {
				const item = this.driveItem(str("handle"))

				this.state.context.screen = "drive"
				this.state.context.directory = item.parent ?? "d0"
				this.state.context.selection = [item.handle]

				return {
					shown: this.pathOf(item)
				}
			}

			case "set_favorite": {
				const items = list("handles").map(handle => this.driveItem(handle))

				for (const item of items) {
					item.favorite = input["favorite"] === true
				}

				return {
					updated: items.map(item => item.name)
				}
			}

			case "create_directory": {
				const parent = this.directory(str("parent"))
				const directoryName = str("name").trim()

				this.assertFreeName({
					parent: parent.handle,
					name: directoryName
				})

				const created: DriveItem = {
					handle: this.nextHandle("d"),
					kind: "directory",
					name: directoryName,
					parent: parent.handle,
					size: 0,
					modified: "2025-10-07T12:00:00Z",
					favorite: false,
					trashed: false,
					shared: false
				}

				this.state.drive.push(created)

				return {
					created: this.summary(created)
				}
			}

			case "rename_item": {
				const item = this.driveItem(str("handle"))
				const newName = str("new_name").trim()

				if (item.parent === null) {
					throw new CallError("The drive root cannot be renamed", false)
				}

				this.assertFreeName({
					parent: item.parent,
					name: newName,
					except: item.handle
				})

				item.name = newName

				return {
					renamed: this.summary(item)
				}
			}

			case "move_items":
			case "copy_items": {
				const items = list("handles").map(handle => this.driveItem(handle))
				const destination = this.directory(str("destination"))

				for (const item of items) {
					let cursor: DriveItem | undefined = destination

					while (cursor !== undefined) {
						if (cursor.handle === item.handle) {
							throw new CallError(`Cannot put "${item.name}" inside itself`, false)
						}

						cursor = this.state.drive.find(candidate => candidate.handle === cursor?.parent)
					}

					this.assertFreeName({
						parent: destination.handle,
						name: item.name,
						except: item.handle
					})
				}

				if (name === "move_items") {
					for (const item of items) {
						item.parent = destination.handle
					}

					return {
						moved: items.map(item => item.name),
						to: this.pathOf(destination)
					}
				}

				const copies = items.map(item => {
					const copy: DriveItem = {
						...structuredClone(item),
						handle: this.nextHandle(item.kind === "directory" ? "d" : "f"),
						parent: destination.handle,
						favorite: false,
						shared: false
					}

					this.state.drive.push(copy)

					return copy
				})

				return {
					copied: copies.map(copy => this.summary(copy))
				}
			}

			case "trash_items": {
				const trashed = list("handles").map(handle => {
					if (handle.startsWith("n")) {
						const item = this.noteItem(handle)

						item.trashed = true

						return item.title
					}

					const item = this.driveItem(handle)

					if (item.parent === null) {
						throw new CallError("The drive root cannot be trashed", false)
					}

					item.trashed = true

					return item.name
				})

				return {
					trashed,
					restorable: true
				}
			}

			case "list_notes": {
				const query = optionalStr("query")
				const queryTokens = query === undefined ? [] : tokens(query)
				const matches = this.state.notes.filter(item => {
					if (item.trashed) {
						return false
					}

					const haystack = normalizeText(`${item.title} ${item.tags.join(" ")}`)

					return (
						queryTokens.length === 0 ||
						queryTokens.some(token =>
							tokenMatches({
								haystack,
								token
							})
						)
					)
				})

				return {
					notes: matches.map(item => this.noteSummary(item))
				}
			}

			case "read_note": {
				const item = this.noteItem(str("handle"))

				return {
					handle: item.handle,
					title: item.title,
					content: item.content
				}
			}

			case "open_note": {
				const item = this.noteItem(str("handle"))

				this.state.context.screen = "note"
				this.state.opened = item.handle

				return {
					opened: item.title
				}
			}

			case "create_note": {
				const created: Note = {
					handle: this.nextHandle("n"),
					title: str("title").trim(),
					type: (optionalStr("type") ?? "text") as NoteType,
					content: optionalStr("content") ?? "",
					tags: [],
					pinned: false,
					favorite: false,
					archived: false,
					trashed: false,
					edited: "2025-10-07T12:00:00Z"
				}

				this.state.notes.push(created)

				return {
					created: this.noteSummary(created)
				}
			}

			case "set_note_flags": {
				const item = this.noteItem(str("handle"))
				const flags = (["pinned", "favorite", "archived"] as const).filter(flag => typeof input[flag] === "boolean")

				if (flags.length === 0) {
					throw new CallError("Pass at least one of pinned, favorite or archived", true)
				}

				for (const flag of flags) {
					item[flag] = input[flag] === true
				}

				return {
					updated: this.noteSummary(item)
				}
			}

			case "list_chats": {
				const query = optionalStr("query")
				const queryTokens = query === undefined ? [] : tokens(query)
				const matches = this.state.chats.filter(item => {
					const haystack = normalizeText(`${item.name} ${item.participants.join(" ")}`)

					return (
						queryTokens.length === 0 ||
						queryTokens.some(token =>
							tokenMatches({
								haystack,
								token
							})
						)
					)
				})

				return {
					chats: matches.map(item => {
						this.exposed.add(item.handle)

						return {
							handle: item.handle,
							name: item.name,
							participants: item.participants
						}
					})
				}
			}

			case "read_chat": {
				const item = this.chat(str("handle"))
				const limit = typeof input["limit"] === "number" && input["limit"] > 0 ? input["limit"] : 20

				return {
					chat: item.name,
					messages: item.messages.slice(-limit)
				}
			}

			case "open_chat": {
				const item = this.chat(str("handle"))

				this.state.context.screen = "chat"
				this.state.opened = item.handle

				return {
					opened: item.name
				}
			}

			case "draft_chat_message": {
				const item = this.chat(str("chat"))

				this.state.context.screen = "chat"
				this.state.opened = item.handle
				this.state.draft = {
					chat: item.handle,
					text: str("text")
				}

				return {
					drafted: true,
					chat: item.name,
					note: "The message is in the input field; the user sends it."
				}
			}

			case "list_playlists": {
				return {
					playlists: this.state.playlists.map(item => {
						this.exposed.add(item.handle)

						return {
							handle: item.handle,
							name: item.name,
							tracks: item.tracks.length
						}
					})
				}
			}

			case "open_playlist": {
				const item = this.playlist(str("handle"))

				this.state.context.screen = "playlist"
				this.state.opened = item.handle

				return {
					opened: item.name,
					tracks: item.tracks.map(handle => this.summary(this.drive(handle)))
				}
			}

			case "play": {
				const handle = str("handle")

				if (handle.startsWith("p")) {
					const item = this.playlist(handle)
					const first = item.tracks[0]

					if (first === undefined) {
						throw new CallError(`Playlist "${item.name}" is empty`, false)
					}

					this.state.context.playing = {
						handle: first,
						playlist: item.handle,
						paused: false,
						shuffle: false,
						repeat: false
					}

					return {
						playing: this.drive(first).name,
						playlist: item.name
					}
				}

				const item = this.driveItem(handle)

				if (item.kind !== "file" || !(EXTENSIONS["audio"] ?? []).includes(extension(item.name))) {
					throw new CallError(`"${item.name}" is not an audio file`, true)
				}

				this.state.context.playing = {
					handle: item.handle,
					playlist: null,
					paused: false,
					shuffle: false,
					repeat: false
				}

				return {
					playing: item.name
				}
			}

			case "audio_control": {
				const playing = this.state.context.playing
				const action = str("action")

				if (playing === null) {
					throw new CallError("Nothing is playing", false)
				}

				if (action === "pause" || action === "resume") {
					playing.paused = action === "pause"
				} else if (action === "stop") {
					this.state.context.playing = null
				} else if (action === "shuffle_on" || action === "shuffle_off") {
					playing.shuffle = action === "shuffle_on"
				} else if (action === "repeat_on" || action === "repeat_off") {
					playing.repeat = action === "repeat_on"
				} else if (playing.playlist !== null) {
					const tracks = this.playlist(playing.playlist).tracks
					const index = tracks.indexOf(playing.handle)
					const next = tracks[(index + (action === "next" ? 1 : tracks.length - 1)) % tracks.length]

					if (next !== undefined) {
						playing.handle = next
					}
				}

				return {
					done: action
				}
			}

			case "add_tracks_to_playlist": {
				const item = this.playlist(str("playlist"))
				const tracks = list("handles").map(handle => {
					const track = this.driveItem(handle)

					if (track.kind !== "file" || !(EXTENSIONS["audio"] ?? []).includes(extension(track.name))) {
						throw new CallError(`"${track.name}" is not an audio file`, true)
					}

					return track
				})
				const added = tracks.filter(track => !item.tracks.includes(track.handle))

				item.tracks.push(...added.map(track => track.handle))

				return {
					added: added.map(track => track.name),
					alreadyThere: tracks.length - added.length,
					playlist: item.name
				}
			}

			case "open_route": {
				this.state.context.screen = str("route")
				this.state.opened = null

				return {
					opened: str("route")
				}
			}

			default: {
				throw new CallError(`Tool "${name}" is not implemented`, true)
			}
		}
	}

	private chat(handle: string) {
		this.known(handle, "chat")

		const item = this.state.chats.find(candidate => candidate.handle === handle)

		if (item === undefined) {
			throw new CallError(`"${handle}" is not a chat`, true)
		}

		return item
	}

	private playlist(handle: string) {
		this.known(handle, "playlist")

		const item = this.state.playlists.find(candidate => candidate.handle === handle)

		if (item === undefined) {
			throw new CallError(`"${handle}" is not a playlist`, true)
		}

		return item
	}
}
