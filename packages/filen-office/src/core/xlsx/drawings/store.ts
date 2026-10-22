import { partKey } from "../../opc/partNames"
import type { Relationship } from "../../opc/relationships"
import type { SheetId } from "../../sheet/protocol"
import { parseXml, parseXmlSliced, type Slicing } from "../../xml/index"
import { heldOrSource, relationshipKind } from "../comments/parts"
import type { XlsxDocument } from "../edit/document"
import type { XlsxWorkbookModel } from "../model/types"
import { RefRefused } from "../refs/catalogue"
import { liveRelationships, refState } from "../refs/state"
import { readDrawingPart, type DrawingPartRecord } from "./read"

// What the drawings feature reads beside a sheet: the drawing part its <drawing> names, and that part's
// relationships (pictures, charts). Source bytes are read once (load, the feature's ready hook); each version is
// parsed once, by its bytes (an edit or a structure edit holds new ones in the refs catalogue's state), so the
// layout producer and queries read them synchronously. Only the version each part now has stays parsed: the history
// holds every earlier version's bytes (within its budget), and a parse of each would be held beside them uncounted.

export const DRAWING_TYPE = "application/vnd.openxmlformats-officedocument.drawing+xml"

const RELATIONSHIP_ID = /\s(?:[A-Za-z_][\w.-]*:)?id\s*=\s*"([^"]*)"/

export interface SheetDrawingPart {
	readonly name: string
	// The sheet's relationship to it (its <drawing r:id>).
	readonly relationship: string
}

export class DrawingStore {
	readonly workbook: XlsxWorkbookModel
	private readonly source = new Map<string, Uint8Array>()
	private readonly loads = new Map<SheetId, Promise<void>>()
	// Why a part does not read, by partKey.
	private readonly failures = new Map<string, string>()
	// Relationships as read, by partKey of their source (sheets and drawing parts).
	private readonly readRelationships = new Map<string, readonly Relationship[]>()
	// Each part's current bytes parsed, by partKey.
	private readonly parsed = new Map<string, { readonly bytes: Uint8Array; readonly record: DrawingPartRecord }>()

	constructor(workbook: XlsxWorkbookModel) {
		this.workbook = workbook
	}

	// Reads a sheet's drawing part once (concurrent calls share it). The sheet is parsed first: its <drawing> names the
	// part. A part that does not read is the sheet's failure().
	load(document: XlsxDocument, sheet: SheetId, slicing?: Slicing): Promise<void> {
		let load = this.loads.get(sheet)

		if (load === undefined) {
			load = this.read(document, sheet, slicing)
			this.loads.set(sheet, load)
		}

		return load
	}

	private async read(document: XlsxDocument, sheet: SheetId, slicing: Slicing | undefined): Promise<void> {
		const entry = this.workbook.sheet(sheet)

		if (entry === undefined || entry.kind !== "worksheet") return

		await document.ready(sheet, slicing)
		await this.relationshipsOf(entry.part)

		const part = this.part(sheet)

		if (part === null || this.failures.has(partKey(part.name))) return

		try {
			await this.relationshipsOf(part.name)

			const bytes = await this.readSource(part.name, slicing)

			if (bytes === null || bytes === undefined) return

			const parsed = readDrawingPart(slicing === undefined ? parseXml(bytes) : await parseXmlSliced(bytes, slicing))

			this.parsed.set(partKey(part.name), { bytes, record: parsed })
		} catch (error) {
			this.failures.set(partKey(part.name), `${part.name}: ${error instanceof Error ? error.message : String(error)}`)
		}
	}

	// A part's relationships, read once (an absent part's: none, so an edit may add to them).
	async relationshipsOf(source: string): Promise<void> {
		const key = partKey(source)

		if (!this.readRelationships.has(key)) this.readRelationships.set(key, (await this.workbook.parts.relationships(source)).all)
	}

	// The relationships of a part as they now are (none before they were read).
	relationships(source: string): Relationship[] {
		return liveRelationships(this.workbook, source, this.readRelationships.get(partKey(source)) ?? [])
	}

	private async readSource(name: string, slicing: Slicing | undefined): Promise<Uint8Array | null | undefined> {
		const current = heldOrSource(this.workbook, name, this.source)

		if (current !== undefined) return current

		const part = this.workbook.parts.get(name)

		if (part?.source === null || part?.source === undefined) return undefined

		this.source.set(partKey(name), await this.workbook.parts.read(name, slicing))

		return heldOrSource(this.workbook, name, this.source)
	}

	// The bytes an edit starts from, read when they were not yet (null: the part is gone or never was).
	async current(name: string): Promise<Uint8Array | null> {
		return (await this.readSource(name, undefined)) ?? null
	}

	// Why the sheet's drawing does not read; null when it does or there is none.
	failure(sheet: SheetId): string | null {
		const part = this.part(sheet)

		return part === null ? null : (this.failures.get(partKey(part.name)) ?? null)
	}

	loaded(sheet: SheetId): boolean {
		return this.loads.has(sheet)
	}

	// The drawing part the sheet's <drawing> names through its relationships, null without one (or before the sheet
	// is parsed).
	part(sheet: SheetId): SheetDrawingPart | null {
		const entry = this.workbook.sheet(sheet)
		const model = entry?.model ?? null

		if (entry === undefined || model === null) return null

		let markup: string | undefined

		try {
			markup = refState(this.workbook).sectionMarkup(model, "drawing")[0]
		} catch (error) {
			if (!(error instanceof RefRefused)) throw error
		}

		const id = markup === undefined ? null : (RELATIONSHIP_ID.exec(markup)?.[1] ?? null)
		const relationship = id === null ? undefined : this.relationships(entry.part).find(candidate => candidate.id === id)

		if (id === null || relationship === undefined || relationship.external || relationship.partName === undefined) return null
		if (relationshipKind(relationship.type) !== "drawing") return null

		return { name: relationship.partName, relationship: id }
	}

	// The sheet's drawing part as it now is, parsed; null when it has none, it is gone or does not read.
	drawing(sheet: SheetId): (DrawingPartRecord & { readonly part: string }) | null {
		const part = this.part(sheet)

		if (part === null || this.failures.has(partKey(part.name))) return null

		const bytes = heldOrSource(this.workbook, part.name, this.source)

		if (bytes === null || bytes === undefined) return null

		const key = partKey(part.name)
		const held = this.parsed.get(key)
		let parsed = held?.bytes === bytes ? held.record : undefined

		if (parsed === undefined) {
			try {
				parsed = readDrawingPart(parseXml(bytes))
			} catch (error) {
				this.failures.set(key, `${part.name}: ${error instanceof Error ? error.message : String(error)}`)

				return null
			}

			this.parsed.set(key, { bytes, record: parsed })
		}

		return { ...parsed, part: part.name }
	}

	// The parsed record of a part's current bytes (what drawing() read last), undefined before.
	record(name: string): DrawingPartRecord | undefined {
		const bytes = heldOrSource(this.workbook, name, this.source)

		const held = this.parsed.get(partKey(name))

		return bytes === null || bytes === undefined || held?.bytes !== bytes ? undefined : held.record
	}

	// A media part's bytes for a renderer, never kept here (pictures can be large): an edit's held bytes copied, or the
	// source's inflated again; null when the part is gone or not in the package.
	async media(name: string): Promise<Uint8Array | null> {
		const current = heldOrSource(this.workbook, name, this.source)

		if (current === null) return null
		if (current !== undefined) return current.slice()

		const part = this.workbook.parts.get(name)

		return part?.source === null || part?.source === undefined ? null : this.workbook.parts.read(name)
	}

	// A part's bytes as they now are: null when removed, undefined when never read.
	bytes(name: string): Uint8Array | null | undefined {
		return heldOrSource(this.workbook, name, this.source)
	}
}

const stores = new WeakMap<XlsxWorkbookModel, DrawingStore>()

export function drawingStore(workbook: XlsxWorkbookModel): DrawingStore {
	let store = stores.get(workbook)

	if (store === undefined) {
		store = new DrawingStore(workbook)
		stores.set(workbook, store)
	}

	return store
}
