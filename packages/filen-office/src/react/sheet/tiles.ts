import {
	TILE_CELLS,
	TILE_COLS,
	TILE_ROWS,
	tileKey,
	tilePosition,
	type SheetId,
	type SheetRegion,
	type SheetTile,
	type SheetTileLayers,
	type SheetTilePatch
} from "../../core/sheet/protocol"

// The tiles a grid holds: least recently used first out, never more than its capacity once a frame has said what
// it needs. Tiles stay as they came (typed arrays and a string table); a cell lookup indexes a tile on first use.

export interface GridTile extends SheetTile {
	readonly key: number
}

// What a lookup answers for one cell: its entry in its tile.
export interface GridEntry {
	readonly tile: GridTile
	readonly entry: number
}

// Positions to entry + 1 (0: none), built when a tile is first looked into.
const indexes = new WeakMap<SheetTile, Uint16Array>()

export function tileIndex(tile: SheetTile): Uint16Array {
	let index = indexes.get(tile)

	if (index === undefined) {
		index = new Uint16Array(TILE_CELLS)

		for (let i = 0; i < tile.count; i++) index[tile.positions[i] ?? 0] = i + 1

		indexes.set(tile, index)
	}

	return index
}

// The entry of a cell in its tile, or -1.
export function entryOf(tile: SheetTile, row: number, col: number): number {
	return (tileIndex(tile)[tilePosition(row, col)] ?? 0) - 1
}

export function entryText(tile: SheetTile, entry: number): string {
	const index = tile.texts[entry] ?? -1

	return index < 0 ? "" : (tile.strings[index] ?? "")
}

// What editing the cell starts from: its input, else its text.
export function entryInput(tile: SheetTile, entry: number): string {
	const index = tile.inputs[entry] ?? -1

	return index < 0 ? entryText(tile, entry) : (tile.strings[index] ?? "")
}

// The runs of a RICH entry as [start, style] pairs, or null.
export function entryRuns(tile: SheetTile, entry: number): number[] | null {
	const runs = tile.layers?.runs

	if (runs === undefined) return null

	let found: number[] | null = null

	for (let i = 0; i < runs.runEntries.length; i++) {
		if (runs.runEntries[i] !== entry) continue

		found ??= []
		found.push(runs.runStarts[i] ?? 0, runs.runStyles[i] ?? 0)
	}

	return found
}

// Where a REPEAT entry's fill goes and what fills it, or null.
export function entryRepeat(tile: SheetTile, entry: number): { readonly offset: number; readonly char: string } | null {
	const repeats = tile.layers?.repeats

	if (repeats === undefined) return null

	for (let i = 0; i < repeats.entries.length; i++)
		if (repeats.entries[i] === entry) return { offset: repeats.offsets[i] ?? 0, char: repeats.chars[i] ?? "" }

	return null
}

function keyed(tile: SheetTile): GridTile {
	return { ...tile, key: tileKey(tile.sheet, tile.tileRow, tile.tileCol) }
}

// A "merge" patch folded into the tile it changes: the patch's entries replace those at their positions, `cleared`
// positions go, strings and layers are rebuilt for what is left.
export function mergeTile(base: SheetTile, patch: SheetTilePatch): SheetTile {
	const replaced = new Uint8Array(TILE_CELLS)

	for (let i = 0; i < patch.count; i++) replaced[patch.positions[i] ?? 0] = 1

	for (const position of patch.cleared ?? []) replaced[position] = 1

	// Sources: [tile, entry], ascending by position.
	const sources: [SheetTile, number][] = []

	for (let i = 0; i < base.count; i++) if (replaced[base.positions[i] ?? 0] === 0) sources.push([base, i])
	for (let i = 0; i < patch.count; i++) sources.push([patch, i])

	sources.sort((a, b) => (a[0].positions[a[1]] ?? 0) - (b[0].positions[b[1]] ?? 0))

	const count = sources.length
	const positions = new Uint16Array(count)
	const texts = new Int32Array(count)
	const inputs = new Int32Array(count)
	const styles = new Uint32Array(count)
	const flags = new Uint32Array(count)
	const strings: string[] = []
	const interned = new Map<string, number>()
	const intern = (tile: SheetTile, index: number): number => {
		if (index < 0) return -1

		const text = tile.strings[index] ?? ""
		let at = interned.get(text)

		if (at === undefined) {
			at = strings.length
			strings.push(text)
			interned.set(text, at)
		}

		return at
	}
	const runEntries: number[] = []
	const runStarts: number[] = []
	const runStyles: number[] = []
	const repeatEntries: number[] = []
	const repeatOffsets: number[] = []
	const repeatChars: string[] = []

	sources.forEach(([tile, i], k) => {
		positions[k] = tile.positions[i] ?? 0
		texts[k] = intern(tile, tile.texts[i] ?? -1)
		inputs[k] = intern(tile, tile.inputs[i] ?? -1)
		styles[k] = tile.styles[i] ?? 0
		flags[k] = tile.flags[i] ?? 0

		const runs = entryRuns(tile, i)

		for (let r = 0; runs !== null && r + 1 < runs.length; r += 2) {
			runEntries.push(k)
			runStarts.push(runs[r] ?? 0)
			runStyles.push(runs[r + 1] ?? 0)
		}

		const repeat = entryRepeat(tile, i)

		if (repeat !== null) {
			repeatEntries.push(k)
			repeatOffsets.push(repeat.offset)
			repeatChars.push(repeat.char)
		}
	})

	const layers: { -readonly [L in keyof SheetTileLayers]: SheetTileLayers[L] } = {}

	if (runEntries.length > 0)
		layers.runs = {
			runEntries: Uint16Array.from(runEntries),
			runStarts: Uint32Array.from(runStarts),
			runStyles: Uint32Array.from(runStyles)
		}
	if (repeatEntries.length > 0)
		layers.repeats = { entries: Uint16Array.from(repeatEntries), offsets: Uint32Array.from(repeatOffsets), chars: repeatChars }

	return {
		sheet: base.sheet,
		tileRow: base.tileRow,
		tileCol: base.tileCol,
		version: patch.version,
		count,
		positions,
		texts,
		inputs,
		styles,
		flags,
		strings,
		...(Object.keys(layers).length > 0 && { layers })
	}
}

function intersects(region: SheetRegion, tile: SheetTile): boolean {
	const top = tile.tileRow * TILE_ROWS
	const left = tile.tileCol * TILE_COLS

	return (
		region.sheet === tile.sheet &&
		region.startRow <= top + TILE_ROWS - 1 &&
		region.endRow >= top &&
		region.startCol <= left + TILE_COLS - 1 &&
		region.endCol >= left
	)
}

export function regionHoldsTile(region: SheetRegion, sheet: SheetId, tileRow: number, tileCol: number): boolean {
	return intersects(region, { sheet, tileRow, tileCol } as SheetTile)
}

export class TileCache {
	private readonly tiles = new Map<number, GridTile>()
	// Keys the session no longer patches for us: handed to `release` in one batch.
	private released: number[] = []

	get size(): number {
		return this.tiles.size
	}

	// A lookup that does not count as a use (painting looks up neighbours many times a frame).
	peek(key: number): GridTile | undefined {
		return this.tiles.get(key)
	}

	// Marks a tile as used now.
	touch(key: number): GridTile | undefined {
		const tile = this.tiles.get(key)

		if (tile !== undefined) {
			this.tiles.delete(key)
			this.tiles.set(key, tile)
		}

		return tile
	}

	has(key: number): boolean {
		return this.tiles.has(key)
	}

	put(tile: SheetTile): GridTile {
		const entry = keyed(tile)

		this.tiles.delete(entry.key)
		this.tiles.set(entry.key, entry)

		return entry
	}

	// A patch for a tile held: a merge needs the tile (else the patch is dropped and false answered, so the tile is
	// fetched again); a patch older than the tile held changes nothing.
	patch(patch: SheetTilePatch): boolean {
		const key = tileKey(patch.sheet, patch.tileRow, patch.tileCol)
		const held = this.tiles.get(key)

		if (held !== undefined && held.version > patch.version) return true
		if (patch.mode === "replace") {
			this.tiles.set(key, keyed(patch))

			return true
		}

		if (held === undefined) return false

		this.tiles.set(key, keyed(mergeTile(held, patch)))

		return true
	}

	// Drops the tiles a region covers (they are fetched again when seen), except those a patch of the same change
	// already replaced.
	invalidate(region: SheetRegion, keep: ReadonlySet<number>): void {
		for (const [key, tile] of this.tiles) {
			if (!keep.has(key) && intersects(region, tile)) {
				this.tiles.delete(key)
				this.released.push(key)
			}
		}
	}

	// Least recently used tiles out until at most `capacity` are left, never one in `needed`.
	evict(capacity: number, needed: ReadonlySet<number>): void {
		if (this.tiles.size <= capacity) return

		for (const key of this.tiles.keys()) {
			if (this.tiles.size <= capacity) break
			if (needed.has(key)) continue

			this.tiles.delete(key)
			this.released.push(key)
		}
	}

	// Every tile of a sheet out (the sheet was deleted).
	dropSheet(sheet: SheetId): void {
		for (const [key, tile] of this.tiles) {
			if (tile.sheet === sheet) {
				this.tiles.delete(key)
				this.released.push(key)
			}
		}
	}

	// Keys about to be fetched again: the session must keep patching them, so their release is not sent.
	keep(keys: ReadonlySet<number>): void {
		if (this.released.length > 0) this.released = this.released.filter(key => !keys.has(key))
	}

	takeReleased(): Float64Array | null {
		if (this.released.length === 0) return null

		const keys = Float64Array.from(this.released)

		this.released = []

		return keys
	}

	keys(): IterableIterator<number> {
		return this.tiles.keys()
	}
}
