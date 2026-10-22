import type { NewRelationship, Relationship } from "../../opc/relationships"
import type { SheetChangeSink, SheetEditStep } from "../../sheet/registry"
import { AnnotationWrites } from "../comments/parts"
import type { ChangeLog } from "../model/changes"
import type { XlsxSheetModel, XlsxWorkbookModel } from "../model/types"
import { refState } from "../refs/state"

// The writes of one drawing edit: parts, relationships and the sheet's sections as an annotation edit writes them
// (through the refs catalogue's state, so a structure edit moves what this wrote), and parts removed. Made forward
// once, then undone and redone as recorded.
export class DrawingWrites {
	readonly workbook: XlsxWorkbookModel
	private readonly writes: AnnotationWrites
	// Parts removed, with the bytes an earlier edit held for them (undefined: the source's).
	private readonly removed: { name: string; contentType: string; held: Uint8Array | undefined }[] = []

	constructor(workbook: XlsxWorkbookModel) {
		this.workbook = workbook
		this.writes = new AnnotationWrites(workbook)
	}

	get bytes(): number {
		return this.writes.bytes + this.removed.reduce((sum, part) => sum + (part.held?.length ?? 0), 0)
	}

	setPart(name: string, contentType: string, bytes: Uint8Array): void {
		this.writes.setPart(name, contentType, bytes)
	}

	// A part taken out of the package (its relationships are the caller's to remove).
	removePart(name: string, contentType: string): void {
		const held = refState(this.workbook).heldPart(name)

		this.removed.push({ name, contentType, held })
		this.remove(name)
	}

	// Back to the source, then out of it: a part an edit added has no source and is simply gone.
	private remove(name: string): void {
		refState(this.workbook).setPart(name, "", undefined)

		const source = this.workbook.parts.get(name)?.source

		if (source !== null && source !== undefined) (this.workbook.changes as ChangeLog).removePart(name)
	}

	addRelationship(source: string, relationship: NewRelationship): Relationship {
		return this.writes.addRelationship(source, relationship)
	}

	removeRelationship(source: string, id: string): void {
		this.writes.removeRelationship(source, id)
	}

	setSection(model: XlsxSheetModel, name: string, occurrences: readonly string[]): void {
		this.writes.setSection(model, name, occurrences)
	}

	undo(): void {
		const state = refState(this.workbook)

		for (let k = this.removed.length - 1; k >= 0; k--) {
			const part = this.removed[k]

			// The bytes it held before, or back to the source.
			if (part !== undefined) state.setPart(part.name, part.contentType, part.held)
		}

		this.writes.undo()
	}

	redo(): void {
		this.writes.redo()

		for (const part of this.removed) this.remove(part.name)
	}

	step(report: (changes: SheetChangeSink) => void): SheetEditStep {
		return {
			bytes: this.bytes,
			undo: changes => {
				this.undo()
				report(changes)
			},
			redo: changes => {
				this.redo()
				report(changes)
			}
		}
	}
}
