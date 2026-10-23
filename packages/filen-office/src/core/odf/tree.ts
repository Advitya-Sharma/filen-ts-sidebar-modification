import { NodeKind, type Namespaces, type XmlDocument } from "../xml/index"
import { canonicalName } from "./namespaces"

// A small element as read: canonical names (see canonicalName), attributes in document order, child elements
// and character data. Used for the style parts only, which are small; document bodies are never built as trees.
export interface OdfElement {
	readonly name: string
	readonly attributes: ReadonlyMap<string, string>
	readonly children: readonly OdfElement[]
	// Character data directly inside the element, concatenated (number:text, text:bullet-char content).
	readonly text: string
}

// Elements nested deeper are refused: style parts are shallow, and recursion depth stays bounded.
const MAX_DEPTH = 64

export class OdfTreeError extends Error {
	constructor(message: string) {
		super(message)

		this.name = "OdfTreeError"
	}
}

export function elementName(document: XmlDocument, namespaces: Namespaces, node: number): string {
	return canonicalName(namespaces.elementUri(node), document.localName(node))
}

// The element's attributes by canonical name; namespace declarations are left out.
export function canonicalAttributes(document: XmlDocument, namespaces: Namespaces, node: number): Map<string, string> {
	const out = new Map<string, string>()

	for (const { name, value } of document.attributes(node)) {
		if (name === "xmlns" || name.startsWith("xmlns:")) continue

		const colon = name.indexOf(":")

		if (colon < 0) {
			out.set(name, value)

			continue
		}

		const prefix = name.slice(0, colon)
		const local = name.slice(colon + 1)
		const uri = prefix === "xml" ? "http://www.w3.org/XML/1998/namespace" : namespaces.uri(node, prefix)

		out.set(uri === "http://www.w3.org/XML/1998/namespace" ? `xml:${local}` : canonicalName(uri, local), value)
	}

	return out
}

export function readElement(document: XmlDocument, namespaces: Namespaces, node: number, depth: number = 0): OdfElement {
	if (depth > MAX_DEPTH) throw new OdfTreeError("Style markup nested too deeply")

	const children: OdfElement[] = []
	let text = ""

	for (let child = document.firstChild(node); child >= 0; child = document.nextSibling(child)) {
		const kind = document.kind(child)

		if (kind === NodeKind.Element) children.push(readElement(document, namespaces, child, depth + 1))
		else if (kind === NodeKind.Text || kind === NodeKind.CData) text += document.text(child)
	}

	return { name: elementName(document, namespaces, node), attributes: canonicalAttributes(document, namespaces, node), children, text }
}

export function childNamed(element: OdfElement, name: string): OdfElement | undefined {
	return element.children.find(child => child.name === name)
}
