import { escapeAttribute } from "../../../xml/index"
import { docxSchema } from "../../edit/schema"
import { textElement } from "../markup"
import type { AtomWriter } from "./index"

// A new field: w:fldSimple with its instruction and cached result.
const FIELD: AtomWriter = {
	type: docxSchema.nodes.field,
	write(w, out, node) {
		const names = w.names
		const instruction = escapeAttribute(node.attrs["instruction"] as string)
		const result = node.attrs["result"] as string

		if (result === "") {
			out.text(`<${names.el("fldSimple")} ${names.attr("instr")}="${instruction}"/>`)

			return
		}

		out.text(
			`<${names.el("fldSimple")} ${names.attr("instr")}="${instruction}"><${names.el("r")}>${textElement(names, result, false)}</${names.el("r")}></${names.el("fldSimple")}>`
		)
	}
}

export const fieldAtoms: readonly AtomWriter[] = [FIELD]
