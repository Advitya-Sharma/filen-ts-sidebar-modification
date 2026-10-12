import { Namespaces, XmlEditor } from "../../../xml/index"
import { partReader } from "../../model/reader"
import { BindingStores, boundNode, changedBindings, type BoundControl } from "../../forms/binding"
import { DocxWriteError } from "../values"
import type { PartContributor, SaveParts, WrittenStories } from "./index"

// The data bound content controls show (ECMA-376 Part 1 §17.5.2.6): Word replaces a bound control's content
// with its node's text on open, so a control whose value changed has its node updated (a custom XML item, or
// the core or extended properties part). Only controls whose value differs from the story as built are
// written; two of them bound to one node must agree. A binding whose store or node does not exist binds
// nothing (§17.5.2.6): there is nothing to update.

// Changed bound controls of every written story, by store item and bound node.
export function changedStores(written: Omit<WrittenStories, "stories">): Map<string, Map<string, BoundControl>> {
	const out = new Map<string, Map<string, BoundControl>>()

	written.sessions.forEach((session, k) => {
		const doc = written.docs[k] ?? session.state.doc

		for (const [key, controls] of changedBindings(session.story.part, session.story.doc, doc)) {
			const first = controls[0]

			if (first === undefined) continue
			if (controls.some(control => control.value !== first.value))
				throw new DocxWriteError("unsupported", "Content controls bound to one node hold different values")

			let store = out.get(first.binding.store)

			if (store === undefined) {
				store = new Map()
				out.set(first.binding.store, store)
			}

			const other = store.get(key)

			if (other !== undefined && other.value !== first.value)
				throw new DocxWriteError("unsupported", "Content controls bound to one node hold different values")

			store.set(key, first)
		}
	})

	return out
}

async function contribute(parts: SaveParts, written: WrittenStories): Promise<void> {
	const changed = changedStores(written)

	if (changed.size === 0) return

	const stores = new BindingStores(parts.pkg)

	for (const [store, controls] of changed) {
		const name = await stores.part(store)
		const document = name === undefined ? undefined : await parts.xml(name)

		if (name === undefined || document === undefined) continue

		const namespaces = new Namespaces(document, { lenient: true })
		const reader = partReader(document)
		const editor = new XmlEditor(document)

		for (const control of controls.values()) {
			const node = boundNode(document, namespaces, control.binding)

			if (node >= 0 && (document.firstChildElement(node) >= 0 || reader.text(node) !== control.value))
				editor.setText(node, control.value)
		}

		if (editor.changed) parts.replace(name, editor.serialize())
	}
}

export const customXmlParts: PartContributor = { contribute }
