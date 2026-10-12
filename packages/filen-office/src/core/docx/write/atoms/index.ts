import type { Node, NodeType } from "prosemirror-model"
import { NO_SOURCE, docxSchema } from "../../edit/schema"
import type { Out } from "../markup"
import { writeMarkerAtom } from "../markers"
import { src } from "../source"
import type { StoryWriting } from "../story"
import { DocxWriteError } from "../values"
import { commentAtoms } from "./comments"
import { drawingAtoms } from "./drawings"
import { fieldAtoms } from "./fields"
import { mathAtoms } from "./math"
import { noteAtoms } from "./notes"

const nodes = docxSchema.nodes

// Writes an inline atom the writer has no source bytes for: a new one, or one whose attributes override its
// source (`edited`). A run-level atom is written inside the w:r opened for it, a paragraph-level one in the
// paragraph.
export interface AtomWriter {
	readonly type: NodeType
	// Whether a node that has a source must be written by `write` all the same (absent: never).
	edited?(node: Node): boolean
	write(w: StoryWriting, out: Out, node: Node): void
}

// New run content Word keeps no state in.
const RUN_ATOMS: readonly AtomWriter[] = [
	{
		type: nodes.tab,
		write(w, out) {
			out.text(w.names.empty("tab"))
		}
	},
	{
		type: nodes.break,
		write(w, out, node) {
			if (node.attrs["cr"] === true) {
				out.text(w.names.empty("cr"))

				return
			}

			const type = node.attrs["type"] as string
			const clear = node.attrs["clear"] as string

			out.text(
				w.names.empty("br", [
					["type", type === "textWrapping" ? null : type],
					["clear", clear === "none" ? null : clear]
				])
			)
		}
	},
	{
		type: nodes.hyphen,
		write(w, out, node) {
			out.text(w.names.empty(node.attrs["soft"] === true ? "softHyphen" : "noBreakHyphen"))
		}
	},
	{
		type: nodes.symbol,
		write(w, out, node) {
			out.text(
				w.names.empty("sym", [
					["font", node.attrs["font"] as string | null],
					["char", (node.attrs["char"] as number).toString(16).toUpperCase().padStart(4, "0")]
				])
			)
		}
	}
]

const ATOMS = new Map<NodeType, AtomWriter>()
// Writers that can find a node with a source edited.
const EDITABLE = new Map<NodeType, AtomWriter>()

for (const writer of [...RUN_ATOMS, ...fieldAtoms, ...drawingAtoms, ...noteAtoms, ...commentAtoms, ...mathAtoms]) {
	ATOMS.set(writer.type, writer)

	if (writer.edited !== undefined) EDITABLE.set(writer.type, writer)
}

// Whether `node` has a source but must not be written as its bytes.
export function atomEdited(node: Node): boolean {
	if (EDITABLE.size === 0) return false

	return EDITABLE.get(node.type)?.edited?.(node) === true
}

// An atom standing at paragraph level (a field, a marker, an equation, a proofing mark).
export function writeAtom(w: StoryWriting, out: Out, node: Node): void {
	const part = w.part
	const own = src(node)

	switch (node.type) {
		case nodes.proof_err: {
			const xmlNode = node.attrs["node"] as number

			if (xmlNode >= 0) w.copyNode(out, xmlNode)

			return
		}
		case nodes.marker:
			writeMarkerAtom(w, out, node)

			return
	}

	if (own !== NO_SOURCE && !atomEdited(node)) {
		// A field is its runs from w:fldChar begin to end.
		if (node.type === nodes.field) w.copySource(out, part.start(own), part.byteEnd(node.attrs["end"] as number), false)
		else w.copyNode(out, part.node(own))

		return
	}

	write(w, out, node)
}

// An atom inside a run (a tab, a drawing, a note reference); text is the caller's.
export function writeRunAtom(w: StoryWriting, out: Out, node: Node): void {
	const own = src(node)

	if (own !== NO_SOURCE && !atomEdited(node)) {
		w.copyNode(out, w.part.node(own))

		return
	}

	write(w, out, node)
}

function write(w: StoryWriting, out: Out, node: Node): void {
	const writer = ATOMS.get(node.type)

	if (writer === undefined) throw new DocxWriteError("missingSource", `A new ${node.type.name}`)

	writer.write(w, out, node)
}
