import { driveItemName, fastLocaleCompare, formatBytes } from "@filen/shared"
import { log } from "@/lib/log"
import { queryClient } from "@/queries/client"
import { cachedListing, directorySizeQueryOptions, fetchItemInfo, itemInfoQueryKey } from "@/features/drive/queries/drive"
import { asDirectoryOrFile, driveItemMime, isDirectoryItem, lastModifiedOf, narrowItem, type DriveItem } from "@/features/drive/lib/item"
import { decodeUtf8, looksBinary, previewType, type PreviewCategory } from "@/features/drive/lib/preview.logic"
import { defaultRevealDeps, runOpenContainingDirectory } from "@/features/drive/lib/reveal"
import { searchDriveOnce } from "@/features/drive/lib/searchQuery"
import { loadDriveFileBytes } from "@/features/preview/lib/fileBytes"
import { assistantRouter } from "@/features/assistant/tools/navigation"
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
import { AssistantToolError, clipText, isoDate, matchScore, searchTokens } from "@/features/assistant/tools/format"
import { ROOT_HANDLE, type AssistantEntity } from "@/features/assistant/tools/handles"

const SEARCH_LIMIT = 12
// How long a search may wait for the index to settle before answering with what it has.
const SEARCH_CAP_MS = 8_000
const LIST_LIMIT = 60
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

function throwIfAborted(scope: ToolScope): void {
	if (scope.signal?.aborted === true) {
		throw new AssistantToolError("Stopped")
	}
}

// The search box's own engine and projection when the user isn't searching, else (and while the index
// still converges) the listings the drive has already loaded. Words are matched one by one, so
// "invoice 2024" finds "Invoice_2024_03.pdf"; the engine is asked for the longest word only.
export const searchItemsTool: ToolAdapter = async (input, scope) => {
	const kind = optionalString(input, "kind")
	const tokens = searchTokens(requiredString(input, "query"))
	const longest = tokens.reduce((best, token) => (token.length > best.length ? token : best), "")

	if (longest === "") {
		throw new AssistantToolError("Pass words from the item name", { invalid: true })
	}

	const engine = await searchDriveOnce(longest, SEARCH_CAP_MS).catch((e: unknown) => {
		log.warn("assistant", "drive search failed", e)

		return null
	})

	throwIfAborted(scope)

	const pool = new Map<string, { item: DriveItem; path: string | undefined }>()

	for (const item of engine?.items ?? []) {
		pool.set(item.data.uuid, { item, path: hitPath(engine?.parentPaths.get(item.data.uuid), driveItemName(item)) })
	}

	const complete = engine?.settled === true

	if (!complete) {
		for (const row of cachedDriveRows()) {
			if (!pool.has(row.item.data.uuid)) {
				pool.set(row.item.data.uuid, row)
			}
		}
	}

	const scored = [...pool.values()]
		.filter(({ item }) => !item.data.undecryptable && matchesKind(item, kind))
		.map(row => ({ ...row, score: matchScore(driveItemName(row.item), tokens) }))
		.filter(({ score }) => score > 0)
	const all = scored.filter(({ score }) => score === tokens.length)
	const ranked = (all.length > 0 ? all : scored.sort((a, b) => b.score - a.score)).slice(0, SEARCH_LIMIT)

	return {
		value: {
			results: ranked.map(({ item, path }) => ({
				handle: showDrive(scope, item, path),
				name: driveName(item),
				kind: driveKind(item),
				...(path === undefined ? {} : { path })
			})),
			...(scored.length > SEARCH_LIMIT ? { more: true } : {}),
			...(complete ? {} : { note: "The search index is not complete yet; more matches may exist." })
		}
	}
}

// The path the info dialog's Location row reads, from the same cached item-info query.
async function itemPath(item: DriveItem): Promise<string | undefined> {
	const base = asDirectoryOrFile(item)
	const info = await queryClient.query({
		queryKey: itemInfoQueryKey(base.data.uuid),
		queryFn: () => fetchItemInfo(base.data),
		staleTime: "static"
	})

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

	throwIfAborted(scope)

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
	// Only what the row itself says: an owned item's outgoing shares and links would cost a request each.
	const sharedArm = item.type !== "directory" && item.type !== "file"

	return {
		value: {
			handle: showDrive(scope, item, path),
			name: driveName(item),
			kind: driveKind(item),
			...(path === undefined ? {} : { path }),
			...(size === undefined ? {} : { size }),
			...(mime === undefined || mime === "" ? {} : { type: mime }),
			created: isoDate(item.data.decryptedMeta?.created ?? item.data.timestamp),
			modified: isoDate(lastModifiedOf(item)),
			favorite: item.data.favorited,
			shared: sharedArm,
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

	if (!TEXT_CATEGORIES.has(previewType(item))) {
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

// "Open containing directory": the same resolver, armed reveal and toast the item menu uses.
export const revealItemTool: ToolAdapter = async (input, scope) => {
	const entity = await resolveHandle(scope, requiredString(input, "handle"), ["root", "drive"])
	const router = assistantRouter()

	if (router === null) {
		throw new AssistantToolError("The app cannot navigate right now")
	}

	if (entity.kind !== "drive") {
		await navigateToRoute(router, "drive")

		return { value: { shown: "/" }, uiEffect: { type: "navigated", pathname: "/drive" } }
	}

	const item = freshDriveItem(entity.item)
	const target = await runOpenContainingDirectory(defaultRevealDeps, item, next => {
		void router.navigate(next)
	})

	if (target === null) {
		throw new AssistantToolError(`Could not open the location of "${driveName(item)}"`)
	}

	return {
		value: { shown: driveName(item), ...(entity.path === undefined ? {} : { path: entity.path }) },
		uiEffect: { type: "navigated", pathname: target.to.replace("$", target.params._splat) }
	}
}
