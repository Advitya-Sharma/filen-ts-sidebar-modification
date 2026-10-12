import type { Node } from "prosemirror-model"
import { NO_SOURCE, docxSchema } from "../../edit/schema"
import { src } from "../source"
import { DocxWriteError } from "../values"
import type { AtomWriter } from "./index"

// New drawings and pictures, and source ones a command changed (size, position, wrap, alt text), are not
// written yet: the save is refused rather than a change dropped.
export const drawingAtoms: readonly AtomWriter[] = [
	{
		type: docxSchema.nodes.drawing,
		edited: (node: Node) => (node.attrs["overrides"] ?? null) !== null,
		write(_w, _out, node) {
			throw new DocxWriteError("missingSource", src(node) === NO_SOURCE ? "A new drawing" : "A changed drawing")
		}
	}
]
