import type { Node } from "prosemirror-model"
import { NO_SOURCE, docxSchema } from "../../edit/schema"
import { src } from "../source"
import { DocxWriteError } from "../values"
import type { AtomWriter } from "./index"

// New equations, and source ones given new OMML, are not written yet: the save is refused rather than a
// change dropped.
export const mathAtoms: readonly AtomWriter[] = [
	{
		type: docxSchema.nodes.math,
		edited: (node: Node) => (node.attrs["omml"] ?? null) !== null,
		write(_w, _out, node) {
			throw new DocxWriteError("missingSource", src(node) === NO_SOURCE ? "A new math" : "A changed math")
		}
	}
]
