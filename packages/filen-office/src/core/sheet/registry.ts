import type { Slicing } from "../slicing"
import type {
	SheetAbortSignal,
	SheetAxis,
	SheetCell,
	SheetCellDetailMap,
	SheetCellDetailName,
	SheetCommandOf,
	SheetCommandType,
	SheetId,
	SheetLayoutExtensionMap,
	SheetLayoutExtensionName,
	SheetProtectedAction,
	SheetQueryName,
	SheetQueryRequest,
	SheetQueryResult,
	SheetRange,
	SheetRefusal,
	SheetRenderStyle,
	SheetSelectionState,
	SheetSessionSettings,
	SheetTileLayerMap,
	SheetTileLayerName
} from "./protocol"

// Where commands, queries and per-tile, per-layout and per-cell data come from. Each feature is one module
// exporting a SheetFeature: its handlers for the names it declared (protocol.ts, by declaration merging). A session
// is built from a list of features (sheetRegistry); nothing dispatches on a command's type by a switch, and two
// features never share a file. A name registered twice is an error when the registry is built.
//
// D is the document a session's features work on (the xlsx session's own document, or a text file's): a feature
// is written for one kind of document.

// What an edit tells the session it changed, so the session can recalculate, patch the client's tiles and
// layouts, and say where the selection goes. Undo and redo report the same way. Over-reporting costs only work.
export interface SheetChangeSink {
	// Values or formulas changed: their dependents are recalculated and their tiles patched.
	content(sheet: SheetId, range: SheetRange): void
	// Only how the cells look changed.
	look(sheet: SheetId, range: SheetRange): void
	// Rows or columns at and after `at` moved by `count` (negative when deleted): every tile past `at` is invalid.
	moved(sheet: SheetId, axis: SheetAxis, at: number, count: number): void
	// The sheet's layout changed (sizes, merges, pane, protection, a layout extension).
	layout(sheet: SheetId): void
	// The sheet list changed (added, deleted, moved, renamed, hidden, tab colour).
	sheets(): void
	select(sheet: SheetId, selection: SheetSelectionState): void
}

export interface SheetReadContext<D> {
	readonly doc: D
	readonly settings: SheetSessionSettings
	// The in-thread host's slices: long work yields at each one (absent in a worker host).
	readonly slicing: Slicing | undefined
	readonly signal: SheetAbortSignal | undefined
	// The session's version (SheetDocState.version) when the call runs: what a query's answer reflects.
	readonly version: number
	// Work a feature does after a call, in turns of its own (statistics read again after a large edit, while its tiles
	// show what it last knew): what it resolves with patches those tiles, as a ready hook's result does. Absent where
	// the session runs nothing later; the feature then does the work at once.
	readonly later?: (sheet: SheetId, work: () => Promise<SheetReadyResult>) => void
}

export interface SheetCommandContext<D> extends SheetReadContext<D> {
	readonly changes: SheetChangeSink
}

// One undoable edit, kept in the session's history (budgeted by `bytes`). Undo and redo report their changes
// as the edit did.
export interface SheetEditStep {
	// What the step retains, estimated.
	readonly bytes: number
	undo(changes: SheetChangeSink): void | Promise<void>
	redo(changes: SheetChangeSink): void | Promise<void>
}

// A step, a refusal, or null when the command ran and changed nothing.
export type SheetHandled = SheetEditStep | { readonly refused: SheetRefusal } | null

// What the session checks before any handler runs, so no feature repeats it.
export interface SheetCommandInfo {
	// Changes the file: refused while the session is view-only, and waits for the proof.
	readonly write: boolean
	// The sheet protection attribute that allows it (§18.3.1.85); without one, a write to a protected sheet is
	// allowed only on unlocked cells, as the handler's check decides.
	readonly protection?: SheetProtectedAction
	// Changes the workbook's structure: refused under workbookProtection lockStructure.
	readonly structure?: boolean
}

// Where a menu or toolbar would apply a command.
export interface SheetTarget {
	readonly sheet: SheetId
	readonly ranges: readonly SheetRange[]
}

export interface SheetCommandHandler<D, T extends SheetCommandType> {
	readonly info: SheetCommandInfo
	// Whether it could run on `target` (commandState): cheap, as menus ask often. Null when it could.
	check?(context: SheetReadContext<D>, target: SheetTarget): SheetRefusal | null
	run(context: SheetCommandContext<D>, command: SheetCommandOf<T>): SheetHandled | Promise<SheetHandled>
}

export interface SheetQueryHandler<D, Q extends SheetQueryName> {
	run(context: SheetReadContext<D>, request: SheetQueryRequest<Q>): SheetQueryResult<Q> | Promise<SheetQueryResult<Q>>
}

// A tile while the session builds it: the base entries are in, each feature's producer then adds its flags,
// styles and layers. Rows and columns are the sheet's.
export interface SheetTileBuilder {
	readonly sheet: SheetId
	readonly firstRow: number
	readonly firstCol: number
	readonly count: number
	// The entry of a cell, or -1.
	find(row: number, col: number): number
	// The entry of a cell, added empty when it has none (a note on a blank cell).
	entry(row: number, col: number): number
	style(entry: number): number
	text(entry: number): string | null
	flags(entry: number): number
	addFlags(entry: number, flags: number): void
	setStyle(entry: number, style: number): void
	// A render style by its index (an entry's, as earlier features left it).
	styleOf(index: number): SheetRenderStyle
	// What the entry shows, and what editing it starts from (null: its text): a format a feature applies over the
	// cell's own (a conditional format's number format, a value hidden behind a data bar).
	setText(entry: number, text: string, input: string | null): void
	// The session's index of a look (interned; new ones reach the client with the tile).
	intern(style: SheetRenderStyle): number
	// Built entries are final: a layer refers to entries by the indices find and entry return.
	layer<L extends SheetTileLayerName>(name: L, value: SheetTileLayerMap[L]): void
	// A layer an earlier feature set on this tile (its entry indices as built).
	layerOf<L extends SheetTileLayerName>(name: L): SheetTileLayerMap[L] | undefined
}

export type SheetTileProducer<D> = (context: SheetReadContext<D>, tile: SheetTileBuilder) => void

// The cells whose look follows other cells' content (a conditional format over its range's values, a rule's formula
// reading cells elsewhere): told that `range` of `sheet` changed (its content, or cells moved there by a structure
// edit), a feature adds the ranges whose tiles must be built again, on any sheet. Called before the change's tiles
// are patched, also for undo, redo and recalculated results; a feature drops what it cached for those cells here.
export type SheetDependentsHook<D> = (
	context: SheetReadContext<D>,
	sheet: SheetId,
	range: SheetRange,
	add: (sheet: SheetId, range: SheetRange) => void
) => void

export type SheetLayoutProducers<D> = {
	readonly [K in SheetLayoutExtensionName]?: (context: SheetReadContext<D>, sheet: SheetId) => SheetLayoutExtensionMap[K] | undefined
}

// What a feature's ready hook found on a sheet: the cells whose tiles show it, and whether the sheet's layout does.
export interface SheetReadyResult {
	readonly ranges: readonly SheetRange[]
	readonly layout?: boolean
}

// Loads what a feature reads beside a sheet's own part (another part of the package): tile and layout producers are
// synchronous, so they show it once it has landed. Tiles and layouts never wait for it; the session patches those
// built before it landed. Rejects (with an Error saying why) when the feature cannot read what it needs: the feature
// then shows nothing on the sheet and refuses its commands there, and the session reports it (featureUnavailable).
export type SheetReadyHook<D> = (context: SheetReadContext<D>, sheet: SheetId) => Promise<SheetReadyResult>

export type SheetCellDetailProducers<D> = {
	readonly [K in SheetCellDetailName]?: (context: SheetReadContext<D>, cell: SheetCell) => SheetCellDetailMap[K] | undefined
}

export type SheetCommandHandlers<D> = { readonly [T in SheetCommandType]?: SheetCommandHandler<D, T> }

export type SheetQueryHandlers<D> = { readonly [Q in SheetQueryName]?: SheetQueryHandler<D, Q> }

export interface SheetFeature<D> {
	// Unique in a registry; names the feature in errors and logs.
	readonly id: string
	readonly commands?: SheetCommandHandlers<D>
	readonly queries?: SheetQueryHandlers<D>
	// Run on every tile built, in feature order.
	readonly tiles?: SheetTileProducer<D>
	// Started once per sheet, when the sheet is first opened or asked for.
	readonly ready?: SheetReadyHook<D>
	readonly dependents?: SheetDependentsHook<D>
	readonly layout?: SheetLayoutProducers<D>
	readonly cell?: SheetCellDetailProducers<D>
}

// Infers a feature's document from its handlers.
export function defineSheetFeature<D>(feature: SheetFeature<D>): SheetFeature<D> {
	return feature
}

export interface SheetRegistry<D> {
	readonly features: readonly SheetFeature<D>[]
	command<T extends SheetCommandType>(type: T): SheetCommandHandler<D, T> | undefined
	// The feature a command or query belongs to.
	owner(name: SheetCommandType | SheetQueryName): string | undefined
	query<Q extends SheetQueryName>(name: Q): SheetQueryHandler<D, Q> | undefined
	readonly tiles: readonly SheetTileProducer<D>[]
	// Each ready hook with its feature's id.
	readonly ready: readonly (readonly [string, SheetReadyHook<D>])[]
	// Each dependents hook with its feature's id.
	readonly dependents: readonly (readonly [string, SheetDependentsHook<D>])[]
	readonly layout: readonly (readonly [SheetLayoutExtensionName, (context: SheetReadContext<D>, sheet: SheetId) => unknown])[]
	readonly cell: readonly (readonly [SheetCellDetailName, (context: SheetReadContext<D>, cell: SheetCell) => unknown])[]
}

function claim(owners: Map<string, string>, kind: string, name: string, feature: string): void {
	const owner = owners.get(name)

	if (owner !== undefined) throw new Error(`Sheet ${kind} "${name}" is registered by both "${owner}" and "${feature}"`)

	owners.set(name, feature)
}

// A registry over `features`. Throws when two features register the same command, query, layout extension or
// cell detail, or share an id.
export function sheetRegistry<D>(features: readonly SheetFeature<D>[]): SheetRegistry<D> {
	const ids = new Set<string>()
	const owners = new Map<string, string>()
	const commands = new Map<string, unknown>()
	const queries = new Map<string, unknown>()
	const extensionOwners = new Map<string, string>()
	const tiles: SheetTileProducer<D>[] = []
	const ready: (readonly [string, SheetReadyHook<D>])[] = []
	const dependents: (readonly [string, SheetDependentsHook<D>])[] = []
	const layout: (readonly [SheetLayoutExtensionName, (context: SheetReadContext<D>, sheet: SheetId) => unknown])[] = []
	const cell: (readonly [SheetCellDetailName, (context: SheetReadContext<D>, cell: SheetCell) => unknown])[] = []

	for (const feature of features) {
		if (ids.has(feature.id)) throw new Error(`Sheet feature "${feature.id}" is registered twice`)

		ids.add(feature.id)

		for (const [type, handler] of Object.entries(feature.commands ?? {})) {
			if (handler === undefined) continue

			claim(owners, "command", type, feature.id)
			commands.set(type, handler)
		}

		for (const [name, handler] of Object.entries(feature.queries ?? {})) {
			if (handler === undefined) continue

			claim(owners, "query", name, feature.id)
			queries.set(name, handler)
		}

		if (feature.tiles !== undefined) tiles.push(feature.tiles)
		if (feature.ready !== undefined) ready.push([feature.id, feature.ready])
		if (feature.dependents !== undefined) dependents.push([feature.id, feature.dependents])

		for (const [name, produce] of Object.entries(feature.layout ?? {})) {
			if (produce === undefined) continue

			claim(extensionOwners, "layout extension", name, feature.id)
			layout.push([name as SheetLayoutExtensionName, produce as (context: SheetReadContext<D>, sheet: SheetId) => unknown])
		}

		for (const [name, produce] of Object.entries(feature.cell ?? {})) {
			if (produce === undefined) continue

			claim(extensionOwners, "cell detail", `cell:${name}`, feature.id)
			cell.push([name as SheetCellDetailName, produce as (context: SheetReadContext<D>, cell: SheetCell) => unknown])
		}
	}

	return {
		features,
		command: <T extends SheetCommandType>(type: T) => commands.get(type) as SheetCommandHandler<D, T> | undefined,
		owner: name => owners.get(name),
		query: <Q extends SheetQueryName>(name: Q) => queries.get(name) as SheetQueryHandler<D, Q> | undefined,
		tiles,
		ready,
		dependents,
		layout,
		cell
	}
}
