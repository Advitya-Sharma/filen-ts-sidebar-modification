import type { NewRelationship, Relationship, Relationships } from "../../opc/relationships"
import { resolvePartName } from "../../opc/partNames"
import type { SheetId } from "../../sheet/protocol"
import { XlsxReadError } from "../read/errors"
import type {
	WorksheetSection,
	XlsxChangeLog,
	XlsxPartWrite,
	XlsxRelationshipChanges,
	XlsxSectionWriter,
	XlsxSheetChanges,
	XlsxWorkbookSection,
	XlsxWorkbookSectionWriter
} from "./types"

interface SheetChanges extends XlsxSheetChanges {
	rows: boolean
	readonly sections: Map<WorksheetSection | (string & {}), XlsxSectionWriter | null>
}

interface RelationshipChanges extends XlsxRelationshipChanges {
	readonly added: Relationship[]
	readonly removed: Set<string>
}

// What changed since open, as the edits record it. `relationshipsOf` answers a part's source relationships once
// they were read (the part map caches them), so ids are allocated without a wait.
export class ChangeLog implements XlsxChangeLog {
	readonly sheets = new Map<SheetId, SheetChanges>()
	readonly workbookSections = new Map<XlsxWorkbookSection | (string & {}), XlsxWorkbookSectionWriter | null>()
	readonly parts = new Map<string, XlsxPartWrite | null>()
	readonly relationships = new Map<string, RelationshipChanges>()
	sharedStrings = false
	styles = false
	dropCalcChain = false
	fullCalcOnLoad = false
	// Highest numeric rId handed out per source part (removed ones included), so an id is never reused.
	private readonly highest = new Map<string, number>()
	// Added relationships removed again, by source and identity: added once more (a redo), each gets its id back.
	private readonly retired = new Map<string, Relationship>()
	private readonly relationshipsOf: (source: string) => Relationships | undefined

	constructor(relationshipsOf: (source: string) => Relationships | undefined) {
		this.relationshipsOf = relationshipsOf
	}

	get empty(): boolean {
		return (
			this.sheets.size === 0 &&
			this.workbookSections.size === 0 &&
			this.parts.size === 0 &&
			this.relationships.size === 0 &&
			!this.sharedStrings &&
			!this.styles &&
			!this.dropCalcChain &&
			!this.fullCalcOnLoad
		)
	}

	private sheet(id: SheetId): SheetChanges {
		let changes = this.sheets.get(id)

		if (changes === undefined) {
			changes = { rows: false, sections: new Map() }
			this.sheets.set(id, changes)
		}

		return changes
	}

	sheetRows(sheet: SheetId): void {
		this.sheet(sheet).rows = true
	}

	sheetSection(sheet: SheetId, section: WorksheetSection | (string & {}), writer?: XlsxSectionWriter): void {
		this.sheet(sheet).sections.set(section, writer ?? null)
	}

	workbookSection(section: XlsxWorkbookSection | (string & {}), writer?: XlsxWorkbookSectionWriter): void {
		this.workbookSections.set(section, writer ?? null)
	}

	sharedStringsAppended(): void {
		this.sharedStrings = true
	}

	stylesAppended(): void {
		this.styles = true
	}

	writePart(name: string, write: XlsxPartWrite): void {
		this.parts.set(name, write)
	}

	removePart(name: string): void {
		this.parts.set(name, null)
	}

	restorePart(name: string): void {
		this.parts.delete(name)
	}

	private relationshipChanges(source: string): RelationshipChanges {
		let changes = this.relationships.get(source)

		if (changes === undefined) {
			changes = { added: [], removed: new Set() }
			this.relationships.set(source, changes)
		}

		return changes
	}

	addRelationship(source: string, relationship: NewRelationship): Relationship {
		const identity = retiredKey(source, relationship.type, relationship.target, relationship.external === true)
		const retired = this.retired.get(identity)

		if (retired !== undefined) {
			this.retired.delete(identity)
			this.relationshipChanges(source).added.push(retired)

			return retired
		}

		let highest = this.highest.get(source)

		if (highest === undefined) {
			const existing = this.relationshipsOf(source)

			if (existing === undefined) {
				throw new XlsxReadError(
					"relationshipsNotRead",
					"A relationship was added before the part's relationships were read",
					source
				)
			}

			highest = 0

			for (const { id } of existing.all) {
				const numbered = /^rId(\d+)$/.exec(id)

				if (numbered !== null) highest = Math.max(highest, Number(numbered[1]))
			}
		}

		const taken = this.relationshipsOf(source)
		let id = `rId${++highest}`

		while (taken?.byId(id) !== undefined) id = `rId${++highest}`

		this.highest.set(source, highest)

		const external = relationship.external === true
		const added: Relationship = {
			id,
			type: relationship.type,
			target: relationship.target,
			external,
			partName: external ? undefined : resolvePartName(source, relationship.target)
		}

		this.relationshipChanges(source).added.push(added)

		return added
	}

	removeRelationship(source: string, id: string): void {
		const changes = this.relationshipChanges(source)
		const index = changes.added.findIndex(relationship => relationship.id === id)

		const added = changes.added[index]

		if (added !== undefined) {
			changes.added.splice(index, 1)
			this.retired.set(retiredKey(source, added.type, added.target, added.external), added)
		} else {
			changes.removed.add(id)
		}

		if (changes.added.length === 0 && changes.removed.size === 0) this.relationships.delete(source)
	}

	// The undo of removeRelationship for a source relationship: it stays.
	restoreRelationship(source: string, id: string): void {
		const changes = this.relationships.get(source)

		if (changes === undefined) return

		changes.removed.delete(id)

		if (changes.added.length === 0 && changes.removed.size === 0) this.relationships.delete(source)
	}

	formulasChanged(): void {
		this.dropCalcChain = true
	}

	needsFullCalc(): void {
		this.fullCalcOnLoad = true
	}
}

function retiredKey(source: string, type: string, target: string, external: boolean): string {
	return `${source}\n${type}\n${target}\n${external ? "1" : "0"}`
}
