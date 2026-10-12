import type { Node } from "prosemirror-model"
import { docxSchema } from "../../edit/schema"
import type { Out } from "../markup"
import type { StoryWriting } from "../story"
import { recordReference, referenceId } from "./comments"
import type { AtomWriter } from "./index"

// A new footnote or endnote reference, in the run written for it (followed by its custom mark where it has
// one); and the number mark a new note starts with.
export const noteAtoms: readonly AtomWriter[] = [
	{
		type: docxSchema.nodes.note_ref,
		write(w, out, node) {
			const kind = node.attrs["kind"] === "endnote" ? "endnote" : "footnote"
			const id = referenceId(node.attrs["id"], kind)

			recordReference(w, kind, id)
			out.text(
				w.names.empty(`${kind}Reference`, [
					["customMarkFollows", node.attrs["custom"] === true ? "1" : undefined],
					["id", id]
				])
			)
		}
	}
]

// The number mark of a note the editor made (a run_object of kind FootnoteMark or EndnoteMark): w:footnoteRef
// or w:endnoteRef. False for any other run_object.
export function writeNoteMark(w: StoryWriting, out: Out, node: Node): boolean {
	const kind = node.attrs["kind"]

	if (kind !== "FootnoteMark" && kind !== "EndnoteMark") return false

	out.text(w.names.empty(kind === "FootnoteMark" ? "footnoteRef" : "endnoteRef"))

	return true
}
