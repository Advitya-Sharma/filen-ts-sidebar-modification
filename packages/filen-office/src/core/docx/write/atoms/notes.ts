import { docxSchema } from "../../edit/schema"
import { recordReference, referenceId } from "./comments"
import type { AtomWriter } from "./index"

// A new footnote or endnote reference, in the run written for it.
export const noteAtoms: readonly AtomWriter[] = [
	{
		type: docxSchema.nodes.note_ref,
		write(w, out, node) {
			const kind = node.attrs["kind"] === "endnote" ? "endnote" : "footnote"
			const id = referenceId(node.attrs["id"], kind)

			recordReference(w, kind, id)
			out.text(w.names.empty(`${kind}Reference`, [["id", id]]))
		}
	}
]
