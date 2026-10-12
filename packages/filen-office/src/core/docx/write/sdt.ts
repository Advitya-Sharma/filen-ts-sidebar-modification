import type { Mark, Node } from "prosemirror-model"
import { NO_SOURCE } from "../edit/schema"
import { Item } from "../model/index"
import { copyNode, openTag, spliceChildren, type Out } from "./markup"
import { src } from "./source"
import type { StoryWriting } from "./story"
import { DocxWriteError } from "./values"

// A block-level content control: its w:sdtPr and w:sdtEndPr as they were, its content written as blocks.
export function* writeSdt(w: StoryWriting, out: Out, node: Node): Generator<void, void> {
	refuseState(node)

	const part = w.part
	const document = part.document
	const names = w.names

	const own = src(node)

	if (!w.owns(node) && own !== NO_SOURCE && part.kinds[own] === Item.Sdt) {
		const row = part.aux[own] ?? -1

		out.text(`<${names.el("sdt")}>`)
		w.copyProperties(out, [part.sdts.sdtPr[row] ?? -1, part.sdts.sdtEndPr[row] ?? -1], ["id"])
		out.text(`<${names.el("sdtContent")}>`)
		yield* w.writeChildren(out, node, NO_SOURCE, (child, blockOut) => w.writeBlock(blockOut, child))
		out.text(`</${names.el("sdtContent")}></${names.el("sdt")}>`)

		return
	}

	if (!w.owns(node)) {
		out.text(`<${names.el("sdt")}><${names.el("sdtPr")}/><${names.el("sdtContent")}>`)
		yield* w.writeChildren(out, node, NO_SOURCE, (child, blockOut) => w.writeBlock(blockOut, child))
		out.text(`</${names.el("sdtContent")}></${names.el("sdt")}>`)

		return
	}

	const item = src(node)
	const content = part.sdts.content[part.aux[item] ?? -1] ?? -1

	if (content < 0) throw new DocxWriteError("missingSource", "A content control without content")

	const element = part.node(item)

	out.bytes(w.bytes, document.start(element), document.openEnd(content))
	yield* w.writeChildren(out, node, item, (child, blockOut) => w.writeBlock(blockOut, child), {
		from: document.openEnd(content),
		to: document.contentEnd(content),
		tail: []
	})
	out.bytes(w.bytes, document.contentEnd(content), document.end(element))
}

// Whether a command set a control's state (checked, chosen, dated): not written yet.
export function changedControl(holder: Node | Mark): boolean {
	return (holder.attrs["state"] ?? null) !== null
}

function refuseState(holder: Node | Mark): void {
	if (changedControl(holder)) throw new DocxWriteError("missingSource", "A changed content control")
}

// The start of source content control `own` (`mark`) around inline content, up to its w:sdtContent start
// tag. `again`: a content control split in two, whose second part has no w:id of its own.
export function openSdtWrapper(w: StoryWriting, out: Out, mark: Mark, own: number, again: boolean): void {
	refuseState(mark)

	const part = w.part
	const document = part.document
	const element = part.node(own)
	const row = part.aux[own] ?? -1
	const content = part.sdts.content[row] ?? -1
	const sdtPr = part.sdts.sdtPr[row] ?? -1

	if (content < 0) throw new DocxWriteError("missingSource", "A content control without content")

	if (!again) {
		out.bytes(w.bytes, document.start(element), document.openEnd(content))

		return
	}

	openTag(out, document, element)

	for (let child = document.firstChildElement(element); child >= 0 && child !== content; child = document.nextSiblingElement(child)) {
		if (child === sdtPr) spliceChildren(out, part.reader, sdtPr, new Map([["id", ""]]), [])
		else copyNode(out, document, child)
	}

	openTag(out, document, content)
}

// The end of source content control `own` around inline content, from its w:sdtContent end tag.
export function closeSdtWrapper(w: StoryWriting, out: Out, own: number): void {
	const part = w.part
	const document = part.document
	const content = part.sdts.content[part.aux[own] ?? -1] ?? -1

	out.bytes(w.bytes, document.contentEnd(content), document.end(part.node(own)))
}
