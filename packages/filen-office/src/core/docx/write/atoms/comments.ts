import { docxSchema } from "../../edit/schema"
import { WRITTEN_REFERENCES, referenceKey } from "../newStory"
import type { StoryWriting } from "../story"
import { DocxWriteError } from "../values"
import type { AtomWriter } from "./index"

// A new comment reference: w:commentReference with the comment's id, in the run written for it; and the
// annotation mark Word puts at the start of a comment's text (a run_object of kind AnnotationMark, the only
// run_object the editor makes: any other new one has nothing to be written from).
export const commentAtoms: readonly AtomWriter[] = [
	{
		type: docxSchema.nodes.run_object,
		write(w, out, node) {
			if (node.attrs["kind"] !== "AnnotationMark") throw new DocxWriteError("missingSource", "A new run_object")

			out.text(w.names.empty("annotationRef"))
		}
	},
	{
		type: docxSchema.nodes.comment_ref,
		write(w, out, node) {
			const id = referenceId(node.attrs["id"], "comment")

			recordReference(w, "comment", id)
			out.text(w.names.empty("commentReference", [["id", id]]))
		}
	}
]

// A reference's w:id: a decimal number (ST_DecimalNumber).
export function referenceId(id: unknown, what: string): string {
	if (typeof id !== "string" || !/^-?\d{1,10}$/.test(id) || !Number.isSafeInteger(Number(id)) || Math.abs(Number(id)) >= 2 ** 31)
		throw new DocxWriteError("missingSource", `A new ${what} reference without an id`)

	return id
}

// Notes the save that a new reference to comment or note `id` was written (see WRITTEN_REFERENCES).
export function recordReference(w: StoryWriting, kind: string, id: string): void {
	w.env.parts?.state(WRITTEN_REFERENCES, () => new Set<string>()).add(referenceKey(kind, id))
}
