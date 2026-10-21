import { partKey } from "../../opc/partNames"
import type { Relationship } from "../../opc/relationships"
import type { SheetId } from "../../sheet/protocol"
import { parseXml, parseXmlSliced, type Slicing, type XmlDocument } from "../../xml/index"
import type { XlsxDocument } from "../edit/document"
import type { XlsxWorkbookModel } from "../model/types"
import { RefRefused } from "../refs/catalogue"
import { liveRelationships, refState } from "../refs/state"
import { readPersons, readThreaded, type PersonsPart, type ThreadedPart } from "../threaded/threads"
import { readComments, type CommentsPart } from "./notes"
import { heldOrSource, PERSON_RELATIONSHIP, relationshipKind } from "./parts"
import { readVml, type VmlPart } from "./vml"

// What notes and threaded comments read beside a sheet: its comments part, the legacy drawing its notes' shapes are
// in, its threaded comments and the workbook's persons. Their source bytes are read once per sheet (load, the
// features' ready hook) and each version parsed once (by its bytes: an edit or a structure edit holds new ones in
// the refs catalogue's state), so tile producers read them synchronously. A part that does not read leaves its
// feature unavailable on the sheet, with the reason.

export type AnnotationKind = "notes" | "drawing" | "threads"

export interface SheetAnnotationParts {
	readonly comments: string | null
	// The VML part of the sheet's legacyDrawing (not legacyDrawingHF), and that relationship's id.
	readonly vml: string | null
	readonly vmlRelationship: string | null
	readonly threaded: string | null
}

const NONE: SheetAnnotationParts = { comments: null, vml: null, vmlRelationship: null, threaded: null }

const RELATIONSHIP_ID = /\s(?:[A-Za-z_][\w.-]*:)?id\s*=\s*"([^"]*)"/

export class AnnotationStore {
	readonly workbook: XlsxWorkbookModel
	private readonly source = new Map<string, Uint8Array>()
	private readonly loads = new Map<SheetId, Promise<void>>()
	// Why a part does not read, by partKey.
	private readonly failures = new Map<string, string>()
	// Each loaded sheet's relationships as its part's relationships part holds them (a parsed sheet's model holds them
	// too).
	private readonly sheetRelationships = new Map<SheetId, readonly Relationship[]>()
	private readonly documents = new WeakMap<Uint8Array, XmlDocument>()
	private readonly comments = new WeakMap<XmlDocument, CommentsPart>()
	private readonly vmls = new WeakMap<XmlDocument, VmlPart>()
	private readonly threads = new WeakMap<XmlDocument, ThreadedPart>()
	private readonly persons = new WeakMap<XmlDocument, PersonsPart>()
	// The workbook part's relationships as read (the persons part's is among them).
	private workbookRelationships: readonly Relationship[] = []

	constructor(workbook: XlsxWorkbookModel) {
		this.workbook = workbook
	}

	// Reads a sheet's annotation parts once (concurrent calls share it), without parsing the sheet: its relationships
	// name them. A part that does not read is a failure of its kind (failure()).
	load(document: XlsxDocument, sheet: SheetId, slicing?: Slicing): Promise<void> {
		let load = this.loads.get(sheet)

		if (load === undefined) {
			load = this.read(document, sheet, slicing)
			this.loads.set(sheet, load)
		}

		return load
	}

	loaded(sheet: SheetId): boolean {
		return this.loads.has(sheet)
	}

	// Why the sheet's notes (its comments part), their drawing (any legacy drawing of the sheet: the refs catalogue
	// rewrites every one on a structure edit) or its threads (its threaded comments, the workbook's persons) do not
	// read; null when they do or are not there.
	failure(sheet: SheetId, kind: AnnotationKind): string | null {
		const relationships = this.relationships(sheet)
		const names =
			kind === "notes"
				? this.targets(relationships, "comments")
				: kind === "drawing"
					? this.targets(relationships, "vmlDrawing")
					: [
							...this.targets(relationships, "threadedComment"),
							...(this.personsPart() === null ? [] : [this.personsPart() ?? ""])
						]

		for (const name of names) {
			const failed = this.failures.get(partKey(name))

			if (failed !== undefined) return failed
		}

		return null
	}

	private targets(relationships: readonly Relationship[], kind: string): string[] {
		return relationships.flatMap(relationship =>
			!relationship.external && relationship.partName !== undefined && relationshipKind(relationship.type) === kind
				? [relationship.partName]
				: []
		)
	}

	private async read(_document: XlsxDocument, sheet: SheetId, slicing: Slicing | undefined): Promise<void> {
		const entry = this.workbook.sheet(sheet)

		if (entry === undefined || entry.kind !== "worksheet") return

		this.workbookRelationships = (await this.workbook.parts.relationships(this.workbook.workbookPart)).all

		if (entry.part !== "") this.sheetRelationships.set(sheet, (await this.workbook.parts.relationships(entry.part)).all)

		const relationships = this.relationships(sheet)
		const persons = this.personsPart()
		const checks: [string, (document: XmlDocument) => unknown][] = [
			...this.targets(relationships, "comments").map(name => [name, readComments] as [string, (document: XmlDocument) => unknown]),
			...this.targets(relationships, "vmlDrawing").map(name => [name, readVml] as [string, (document: XmlDocument) => unknown]),
			...this.targets(relationships, "threadedComment").map(
				name => [name, readThreaded] as [string, (document: XmlDocument) => unknown]
			),
			...(persons === null ? [] : [[persons, readPersons] as [string, (document: XmlDocument) => unknown]])
		]

		for (const [name, check] of checks) {
			if (this.failures.has(partKey(name))) continue

			try {
				const bytes = await this.readSource(name, slicing)

				if (bytes === null || bytes === undefined) continue

				check(await this.parse(bytes, slicing))
			} catch (error) {
				this.failures.set(partKey(name), `${name}: ${error instanceof Error ? error.message : String(error)}`)
			}
		}
	}

	private async readSource(name: string, slicing: Slicing | undefined): Promise<Uint8Array | null | undefined> {
		const current = heldOrSource(this.workbook, name, this.source)

		if (current !== undefined) return current

		const part = this.workbook.parts.get(name)

		if (part?.source === null || part?.source === undefined) return undefined

		const bytes = await this.workbook.parts.read(name, slicing)

		this.source.set(partKey(name), bytes)

		return heldOrSource(this.workbook, name, this.source)
	}

	private async parse(bytes: Uint8Array, slicing: Slicing | undefined): Promise<XmlDocument> {
		let document = this.documents.get(bytes)

		if (document === undefined) {
			document = slicing === undefined ? parseXml(bytes) : await parseXmlSliced(bytes, slicing)
			this.documents.set(bytes, document)
		}

		return document
	}

	private parsed(bytes: Uint8Array): XmlDocument {
		let document = this.documents.get(bytes)

		if (document === undefined) {
			document = parseXml(bytes)
			this.documents.set(bytes, document)
		}

		return document
	}

	// The sheet's relationships as they now are (none known before its load).
	relationships(sheet: SheetId): Relationship[] {
		const entry = this.workbook.sheet(sheet)
		const read = entry?.model?.relationships.all ?? this.sheetRelationships.get(sheet)

		if (entry === undefined || read === undefined) return []

		return liveRelationships(this.workbook, entry.part, read)
	}

	// The sheet's annotation parts; its legacy drawing only once the sheet is parsed (its legacyDrawing names it).
	parts(sheet: SheetId): SheetAnnotationParts {
		const entry = this.workbook.sheet(sheet)

		if (entry === undefined) return NONE

		const model = entry.model
		const relationships = this.relationships(sheet)
		const of = (kind: string): string | null =>
			relationships.find(relationship => !relationship.external && relationshipKind(relationship.type) === kind)?.partName ?? null
		let vmlRelationship: string | null = null

		try {
			const markup = model === null ? undefined : refState(this.workbook).sectionMarkup(model, "legacyDrawing")[0]

			vmlRelationship = markup === undefined ? null : (RELATIONSHIP_ID.exec(markup)?.[1] ?? null)
		} catch (error) {
			if (!(error instanceof RefRefused)) throw error
		}

		const vml = vmlRelationship === null ? undefined : relationships.find(relationship => relationship.id === vmlRelationship)

		return {
			comments: of("comments"),
			vml: vml === undefined || vml.external ? null : (vml.partName ?? null),
			vmlRelationship: vml === undefined ? null : vmlRelationship,
			threaded: of("threadedComment")
		}
	}

	// The workbook's persons part, null when it has none.
	personsPart(): string | null {
		return (
			liveRelationships(this.workbook, this.workbook.workbookPart, this.workbookRelationships).find(
				relationship => !relationship.external && relationship.type === PERSON_RELATIONSHIP
			)?.partName ?? null
		)
	}

	// A part's bytes as they now are: null when it was removed, undefined when they were never read.
	bytes(name: string): Uint8Array | null | undefined {
		return heldOrSource(this.workbook, name, this.source)
	}

	private document(name: string | null): XmlDocument | null {
		if (name === null) return null

		const bytes = this.bytes(name)

		return bytes === null || bytes === undefined ? null : this.parsed(bytes)
	}

	// The parsed parts, null when the sheet has none (or it failed: callers check failure() first).
	commentsOf(sheet: SheetId): CommentsPart | null {
		if (this.failure(sheet, "notes") !== null) return null

		return this.read$(this.document(this.parts(sheet).comments), this.comments, readComments)
	}

	vmlOf(sheet: SheetId): VmlPart | null {
		if (this.failure(sheet, "drawing") !== null) return null

		return this.read$(this.document(this.parts(sheet).vml), this.vmls, readVml)
	}

	threadedOf(sheet: SheetId): ThreadedPart | null {
		if (this.failure(sheet, "threads") !== null) return null

		return this.read$(this.document(this.parts(sheet).threaded), this.threads, readThreaded)
	}

	personsOf(): PersonsPart | null {
		return this.read$(this.document(this.personsPart()), this.persons, readPersons)
	}

	private read$<T>(document: XmlDocument | null, cache: WeakMap<XmlDocument, T>, read: (document: XmlDocument) => T): T | null {
		if (document === null) return null

		let found = cache.get(document)

		if (found === undefined) {
			found = read(document)
			cache.set(document, found)
		}

		return found
	}

	// The bytes an edit starts from, read when they were not yet.
	async current(name: string): Promise<Uint8Array | null> {
		const bytes = await this.readSource(name, undefined)

		return bytes ?? null
	}
}

const stores = new WeakMap<XlsxWorkbookModel, AnnotationStore>()

export function annotationStore(workbook: XlsxWorkbookModel): AnnotationStore {
	let store = stores.get(workbook)

	if (store === undefined) {
		store = new AnnotationStore(workbook)
		stores.set(workbook, store)
	}

	return store
}
