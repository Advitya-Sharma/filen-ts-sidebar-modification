import { driveItemName, fastLocaleCompare, formatBytes } from "@filen/shared"
import { MAX_ARRAY_ITEMS, type AssistantSignal } from "@filen/shared/assistant"
import { log } from "@/lib/log"
import { queryClient } from "@/queries/client"
import { cachedListing, directorySizeQueryOptions, fetchItemInfo, itemInfoQueryKey } from "@/features/drive/queries/drive"
import { asDirectoryOrFile, driveItemMime, isDirectoryItem, lastModifiedOf, narrowItem, type DriveItem } from "@/features/drive/lib/item"
import { decodeUtf8, extensionOf, looksBinary, previewType, type PreviewCategory } from "@/features/drive/lib/preview.logic"
import { defaultRevealDeps, resolveContainingDirectoryTarget } from "@/features/drive/lib/reveal"
import { searchDriveOnce, type DriveSearchOnceResult } from "@/features/drive/lib/searchQuery"
import { WATCHDOG_MS } from "@/features/drive/lib/searchStatus.logic"
import { useDriveStore } from "@/features/drive/store/useDriveStore"
import { spreadsheetFileKind } from "@/features/spreadsheet/lib/fileKind"
import { loadDriveFileBytes } from "@/features/preview/lib/fileBytes"
import { assistantRouter, navigateChecked } from "@/features/assistant/tools/navigation"
import { navigateToRoute } from "@/features/assistant/tools/routes"
import {
	abortSignalOf,
	cachedDriveRows,
	driveKind,
	driveName,
	freshDriveItem,
	joinPath,
	optionalString,
	readDriveListing,
	requiredString,
	resolveDrivePath,
	resolveHandle,
	showDrive,
	type ToolAdapter,
	type ToolScope
} from "@/features/assistant/tools/entities"
import {
	AssistantToolError,
	clipText,
	engineTokens,
	isoDate,
	matchScore,
	normalizeText,
	rawSearchTokens,
	searchTokens,
	throwIfStopped
} from "@/features/assistant/tools/format"
import { ROOT_HANDLE, type AssistantEntity } from "@/features/assistant/tools/handles"

const SEARCH_LIMIT = 12
// How long a turn's searches may wait for the drive's index in all: a first index of a very large drive
// takes minutes, and the step shows the progress meanwhile. Stop ends it at once.
export const SEARCH_CAP_MS = 300_000
// A search once the turn's budget is spent still reads what the index already holds.
const SEARCH_FLOOR_MS = 2_000
// A resync silent this long ends the wait: as long as the search box's own watchdog, since a long apply
// reports nothing and the messages are lossy.
const SEARCH_STALL_MS = WATCHDOG_MS
// The engine keeps the drive registered this long after a search, so a follow-up costs no new resync.
const SEARCH_KEEP_MS = 10 * 60_000
// The words the engine may try in turn while one matches more names than its window holds.
const ENGINE_WORDS = 3
// The loop shows the model at most this many items of a list: a longer listing would be cut twice.
const LIST_LIMIT = MAX_ARRAY_ITEMS
// read_text_file downloads the whole file, so only small ones; the model sees far less of it.
export const TEXT_FILE_MAX_BYTES = 1_048_576n
export const TEXT_FILE_MAX_CHARS = 4_000

const TEXT_CATEGORIES: ReadonlySet<PreviewCategory> = new Set(["text", "code", "markdown"])

const KIND_CATEGORIES: Record<string, ReadonlySet<PreviewCategory>> = {
	image: new Set(["image", "rawImage"]),
	video: new Set(["video"]),
	audio: new Set(["audio"]),
	document: new Set(["pdf", "docx", "spreadsheet", "text", "markdown"])
}

function matchesKind(item: DriveItem, kind: string | undefined): boolean {
	if (kind === undefined) {
		return true
	}

	if (kind === "directory" || kind === "file") {
		return driveKind(item) === kind
	}

	return !isDirectoryItem(item) && (KIND_CATEGORIES[kind]?.has(previewType(item)) ?? false)
}

function hitPath(parentPath: string | undefined, name: string): string {
	const parent = (parentPath ?? "").replace(/^\/+|\/+$/g, "")

	return parent === "" ? `/${name}` : `/${parent}/${name}`
}

function engineItemType(kind: string | undefined): "all" | "file" | "dir" {
	return kind === undefined ? "all" : kind === "directory" ? "dir" : "file"
}

// Each turn's deadline: a model retrying other words would otherwise wait out the same index build again
// on every step, holding the loaded model all along.
const turnDeadlines = new WeakMap<AssistantSignal, number>()

function searchBudget(signal: AssistantSignal | undefined): number {
	const now = Date.now()

	if (signal === undefined) {
		return SEARCH_CAP_MS
	}

	let deadline = turnDeadlines.get(signal)

	if (deadline === undefined) {
		deadline = now + SEARCH_CAP_MS
		turnDeadlines.set(signal, deadline)
	}

	return Math.max(SEARCH_FLOOR_MS, deadline - now)
}

// The engine's hits for the most selective of `words`; null when the engine failed.
async function engineSearch(words: readonly string[], kind: string | undefined, scope: ToolScope): Promise<DriveSearchOnceResult | null> {
	const { progress } = scope
	const result = await searchDriveOnce(words, {
		capMs: searchBudget(scope.signal),
		stallMs: SEARCH_STALL_MS,
		keepMs: SEARCH_KEEP_MS,
		signal: abortSignalOf(scope.signal),
		itemType: engineItemType(kind),
		onProgress:
			progress === undefined
				? undefined
				: ({ indexed, candidates }) => {
						progress({ share: indexed, found: Number(candidates) })
					}
	}).catch((e: unknown) => {
		if (scope.signal?.aborted !== true) {
			log.warn("assistant", "drive search failed", e)
		}

		return null
	})

	throwIfStopped(scope)

	return result
}

// The engine returns a window of its hits: more than that means some were never seen.
function truncated(result: DriveSearchOnceResult): boolean {
	return result.total > BigInt(result.items.length)
}

function searchNote(engine: DriveSearchOnceResult | null): string | null {
	if (engine === null) {
		return "The drive search failed, so only the directories already loaded were searched; more matches may exist."
	}

	if (!engine.settled) {
		return "The drive's search index is still being built, so only part of the drive was searched; more matches may exist. The user can ask again in a few minutes."
	}

	if (truncated(engine)) {
		return `Too many names match: ${String(engine.total)} names contain "${engine.name}" and only the first ${String(engine.items.length)} were checked, so matches are likely missing. Ask the user for more specific words from the name.`
	}

	return null
}

// One engine handle of its own beside the user's search, so the user's search box is never disturbed;
// while the index converges, the listings the drive has already loaded too. Words are matched one by one,
// so "invoice 2024" finds "Invoice_2024_03.pdf": the engine is asked for the most selective word as
// written (it folds case only), and for the next ones while a word matches more names than it returns.
export const searchItemsTool: ToolAdapter = async (input, scope) => {
	const kind = optionalString(input, "kind")
	const query = requiredString(input, "query")
	const tokens = searchTokens(query)
	const words = engineTokens(rawSearchTokens(query), ENGINE_WORDS)

	if (words.length === 0 || tokens.length === 0) {
		throw new AssistantToolError("Pass words from the item name", { invalid: true })
	}

	const engine = await engineSearch(words, kind, scope)
	const pool = new Map<string, { item: DriveItem; path: string | undefined }>()

	for (const item of engine?.items ?? []) {
		pool.set(item.data.uuid, { item, path: hitPath(engine?.parentPaths.get(item.data.uuid), driveItemName(item)) })
	}

	if (engine?.settled !== true) {
		for (const row of cachedDriveRows(item => matchScore(driveItemName(item), tokens) > 0)) {
			if (!pool.has(row.item.data.uuid)) {
				pool.set(row.item.data.uuid, row)
			}
		}
	}

	const wanted = normalizeText(query.trim())
	const scored = [...pool.values()]
		.filter(({ item }) => !item.data.undecryptable && matchesKind(item, kind))
		.map(row => ({ ...row, score: matchScore(driveItemName(row.item), tokens) }))
		.filter(({ score }) => score > 0)
	const all = scored.filter(({ score }) => score === tokens.length)
	// The exact name first, then the shortest names (the fewest extra words), then the newest.
	const ranked = (all.length > 0 ? all : scored)
		.toSorted(
			(a, b) =>
				b.score - a.score ||
				Number(normalizeText(driveItemName(b.item)) === wanted) - Number(normalizeText(driveItemName(a.item)) === wanted) ||
				driveItemName(a.item).length - driveItemName(b.item).length ||
				Number(lastModifiedOf(b.item) - lastModifiedOf(a.item))
		)
		.slice(0, SEARCH_LIMIT)
	const shown = all.length > 0 ? all.length : scored.length
	const note = searchNote(engine)

	return {
		value: {
			results: ranked.map(({ item, path }) => ({
				handle: showDrive(scope, item, path),
				name: driveName(item),
				kind: driveKind(item),
				...(path === undefined ? {} : { path })
			})),
			...(shown > SEARCH_LIMIT ? { more: shown - SEARCH_LIMIT } : {}),
			...(note === null ? {} : { note })
		}
	}
}

// Item-info entries read this page, by uuid: the entry is persisted, and a move keeps the uuid it is keyed
// by, so one restored from disk may name where the item used to be.
const infoReadThisPage = new Set<string>()

// The path the info dialog's Location row reads, from the same item-info query: read again once a page, as
// the dialog does (its staleTime is 0), then from the cache.
async function itemPath(item: DriveItem): Promise<string | undefined> {
	const base = asDirectoryOrFile(item)
	const fresh = infoReadThisPage.has(base.data.uuid)
	const info = await queryClient.query({
		queryKey: itemInfoQueryKey(base.data.uuid),
		queryFn: () => fetchItemInfo(base.data),
		staleTime: fresh ? "static" : 0
	})

	infoReadThisPage.add(base.data.uuid)

	if (info.path === null) {
		return undefined
	}

	return joinPath(
		info.ancestors.reduce((path, ancestor) => joinPath(path, driveItemName(narrowItem(ancestor))), "/"),
		driveItemName(item)
	)
}

function isTrashed(item: DriveItem): boolean {
	return cachedListing("trash", null)?.some(row => row.data.uuid === item.data.uuid) ?? false
}

async function directoryEntity(input: Record<string, unknown>, scope: ToolScope): Promise<AssistantEntity> {
	const handle = optionalString(input, "handle")
	// Models fill both optional fields or put a path into "handle"; a path wins, being unambiguous.
	const path = optionalString(input, "path") ?? (handle?.trim().startsWith("/") === true ? handle : undefined)

	if (path !== undefined) {
		return resolveDrivePath(path)
	}

	return handle === undefined ? { kind: "root" } : resolveHandle(scope, handle, ["root", "drive"])
}

export const listDirectoryTool: ToolAdapter = async (input, scope) => {
	const entity = await directoryEntity(input, scope)
	let uuid: string | null = null
	let directory: { handle: string; path: string | undefined } = { handle: ROOT_HANDLE, path: "/" }

	if (entity.kind === "drive") {
		const item = freshDriveItem(entity.item)

		if (!isDirectoryItem(item)) {
			throw new AssistantToolError(`"${driveName(item)}" is a file, not a directory`, { invalid: true })
		}

		// Shared directories list through their own variant and resolution path, not My Drive's.
		if (item.type !== "directory") {
			throw new AssistantToolError(`"${driveName(item)}" is a shared directory; open it with reveal_item instead`)
		}

		if (item.data.undecryptable) {
			throw new AssistantToolError("This directory cannot be decrypted")
		}

		const path = entity.path ?? (await itemPath(item).catch(() => undefined))

		uuid = item.data.uuid
		directory = { handle: showDrive(scope, item, path), path }
	} else {
		scope.show({ kind: "root" }, "/")
	}

	const items = [...(await readDriveListing(uuid))].sort(
		(a, b) => Number(isDirectoryItem(b)) - Number(isDirectoryItem(a)) || fastLocaleCompare(driveItemName(a), driveItemName(b))
	)

	throwIfStopped(scope)

	return {
		value: {
			directory: { handle: directory.handle, ...(directory.path === undefined ? {} : { path: directory.path }) },
			items: items.slice(0, LIST_LIMIT).map(item => ({
				handle: showDrive(scope, item, directory.path === undefined ? undefined : joinPath(directory.path, driveItemName(item))),
				name: driveName(item),
				kind: driveKind(item)
			})),
			...(items.length > LIST_LIMIT ? { more: items.length - LIST_LIMIT } : {})
		}
	}
}

export const getItemInfoTool: ToolAdapter = async (input, scope) => {
	const entity = await resolveHandle(scope, requiredString(input, "handle"), ["root", "drive"])

	if (entity.kind !== "drive") {
		return { value: { handle: scope.show(entity, "/"), name: "/", kind: "directory", path: "/" } }
	}

	const item = freshDriveItem(entity.item)

	if (item.data.undecryptable) {
		throw new AssistantToolError("This item cannot be decrypted")
	}

	// Like the info dialog, a trashed item gets no remote rows: its ancestry has nothing to walk.
	const trashed = isTrashed(item)
	const path = entity.path ?? (trashed ? undefined : await itemPath(item).catch(() => undefined))
	let size: string | undefined = isDirectoryItem(item) ? undefined : formatBytes(Number(item.data.size))

	if (isDirectoryItem(item) && !trashed) {
		const totals = await queryClient.query(directorySizeQueryOptions(item)).catch(() => undefined)

		size =
			totals === undefined
				? undefined
				: `${formatBytes(Number(totals.size))} in ${String(totals.files)} files and ${String(totals.dirs)} directories`
	}

	const mime = driveItemMime(item)
	// Only what the row itself says: an owned item's outgoing shares and links would cost a request each, so
	// its sharing state is left out rather than reported as "not shared".
	const sharedWithYou = item.type !== "directory" && item.type !== "file"

	return {
		value: {
			handle: showDrive(scope, item, path),
			name: driveName(item),
			kind: driveKind(item),
			...(path === undefined ? {} : { path }),
			...(size === undefined ? {} : { size }),
			// The uploader's own MIME type: untrusted text, under a key the loop sanitizes.
			...(mime === undefined || mime === "" ? {} : { mime }),
			created: isoDate(item.data.decryptedMeta?.created ?? item.data.timestamp),
			modified: isoDate(lastModifiedOf(item)),
			favorite: item.data.favorited,
			...(sharedWithYou ? { sharedWithYou: true } : {}),
			...(trashed ? { trashed: true } : {})
		}
	}
}

export const readTextFileTool: ToolAdapter = async (input, scope) => {
	const entity = await resolveHandle(scope, requiredString(input, "handle"), ["drive"])

	if (entity.kind !== "drive") {
		throw new AssistantToolError("Not a file", { invalid: true })
	}

	const item = freshDriveItem(entity.item)
	const name = driveName(item)

	if (isDirectoryItem(item)) {
		throw new AssistantToolError(`"${name}" is a directory, not a file`, { invalid: true })
	}

	if (item.data.undecryptable) {
		throw new AssistantToolError("This file cannot be decrypted")
	}

	const csv = spreadsheetFileKind(extensionOf(name).toLowerCase())

	// CSV and TSV preview as spreadsheets but are text.
	if (!TEXT_CATEGORIES.has(previewType(item)) && csv !== "csv" && csv !== "tsv") {
		throw new AssistantToolError(`"${name}" is not a text file`)
	}

	if (item.data.size > TEXT_FILE_MAX_BYTES) {
		throw new AssistantToolError(`"${name}" is too large to read (${formatBytes(Number(item.data.size))})`)
	}

	const bytes = await loadDriveFileBytes(item, abortSignalOf(scope.signal))

	if (looksBinary(bytes)) {
		throw new AssistantToolError(`"${name}" is not a text file`)
	}

	const { text, truncated } = clipText(decodeUtf8(bytes), TEXT_FILE_MAX_CHARS)

	return {
		value: {
			handle: showDrive(scope, item, entity.path),
			name,
			content: text,
			...(truncated ? { truncated: true } : {})
		}
	}
}

// A file: "Open containing directory", with the same resolver and armed reveal the item menu uses. One of the
// user's own directories opens itself, as a double click does. Nothing happens once the turn was stopped.
export const revealItemTool: ToolAdapter = async (input, scope) => {
	const entity = await resolveHandle(scope, requiredString(input, "handle"), ["root", "drive"])
	const router = assistantRouter()

	if (router === null) {
		throw new AssistantToolError("The app cannot navigate right now")
	}

	if (entity.kind !== "drive") {
		await navigateToRoute(router, "drive", scope.signal)

		return { value: { shown: "/" }, uiEffect: { type: "navigated", pathname: "/drive" } }
	}

	const item = freshDriveItem(entity.item)
	const outcome = await resolveContainingDirectoryTarget(defaultRevealDeps, item)

	throwIfStopped(scope)

	if (outcome.status === "error") {
		throw new AssistantToolError(`Could not open the location of "${driveName(item)}"`)
	}

	const into = item.type === "directory" && !item.data.undecryptable
	const splat = into ? [outcome.target.params._splat, item.data.uuid].filter(part => part !== "").join("/") : outcome.target.params._splat
	const target = { ...outcome.target, params: { _splat: splat } }
	const pathname = target.to.replace("$", splat)

	if (!into) {
		useDriveStore.getState().requestReveal({ uuid: item.data.uuid, splat })
	}

	await navigateChecked({
		router,
		navigate: () => router.navigate(target),
		landed: current => current.replace(/\/+$/, "") === pathname.replace(/\/+$/, ""),
		signal: scope.signal
	})

	return {
		value: { [into ? "opened" : "shown"]: driveName(item), ...(entity.path === undefined ? {} : { path: entity.path }) },
		uiEffect: { type: "navigated", pathname }
	}
}
