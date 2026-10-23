import { canonicalAttributes, elementName, readElement, type OdfElement } from "../../odf/tree"
import { escapeAttribute, Namespaces, NodeKind, parseXml, XmlError, type XmlDocument } from "../../xml/index"
import { OdsReadError } from "./errors"

const encoder = new TextEncoder()

// A small element captured from content.xml (validations, names, an annotation...), parsed under the namespace
// declarations in force where it was captured.
export interface Captured {
	readonly document: XmlDocument
	readonly namespaces: Namespaces
	readonly node: number
}

export function parseCaptured(markup: Uint8Array, prefixes: ReadonlyMap<string, string>): Captured {
	let declarations = ""

	for (const [prefix, uri] of prefixes) declarations += ` ${prefix === "" ? "xmlns" : `xmlns:${prefix}`}="${escapeAttribute(uri)}"`

	const head = encoder.encode(`<captured${declarations}>`)
	const tail = encoder.encode("</captured>")
	const bytes = new Uint8Array(head.length + markup.length + tail.length)

	bytes.set(head)
	bytes.set(markup, head.length)
	bytes.set(tail, head.length + markup.length)

	try {
		const document = parseXml(bytes)
		const node = document.firstChildElement(document.documentElement())

		if (node < 0) throw new OdsReadError("malformed", "An empty element was captured")

		return { document, namespaces: new Namespaces(document, { lenient: true }), node }
	} catch (error) {
		if (error instanceof XmlError) throw new OdsReadError("malformed", `content.xml: ${error.message}`)

		throw error
	}
}

// The captured element as a tree with canonical names.
export function capturedElement(markup: Uint8Array, prefixes: ReadonlyMap<string, string>): OdfElement {
	const { document, namespaces, node } = parseCaptured(markup, prefixes)

	return readElement(document, namespaces, node)
}

// The text of an element in document order: paragraphs (text:p, text:h) on lines of their own, text:s, text:tab and
// text:line-break as the characters they stand for, other inline elements (spans, links, fields) by their text.
export function capturedText({ document, namespaces, node }: Captured, skip: ReadonlySet<string> = new Set()): string {
	const lines: string[] = []
	let line: string[] | null = null

	const walk = (parent: number): void => {
		for (let child = document.firstChild(parent); child >= 0; child = document.nextSibling(child)) {
			const kind = document.kind(child)

			if (kind === NodeKind.Text || kind === NodeKind.CData) {
				line?.push(document.text(child))

				continue
			}

			if (kind !== NodeKind.Element) continue

			const name = elementName(document, namespaces, child)

			if (skip.has(name)) continue

			if (name === "text:p" || name === "text:h") {
				const outer = line

				line = []
				walk(child)
				lines.push(line.join(""))
				line = outer
			} else if (name === "text:s") {
				line?.push(" ".repeat(Math.max(1, Number(canonicalAttributes(document, namespaces, child).get("text:c") ?? "1") || 1)))
			} else if (name === "text:tab") {
				line?.push("\t")
			} else if (name === "text:line-break") {
				line?.push("\n")
			} else {
				walk(child)
			}
		}
	}

	walk(node)

	return lines.join("\n")
}

// Each content validation's help and error messages (by table:name) as their text in document order (capturedText):
// the element tree keeps an element's character data apart from its children, which loses where a text:s was.
// Validations captured between yields.
const VALIDATIONS_PER_STEP = 256

// table:content-validations parsed once, in steps: the element, and each validation's messages, their text in
// document order.
export function* capturedValidationsSteps(
	markup: Uint8Array,
	prefixes: ReadonlyMap<string, string>
): Generator<void, { readonly element: OdfElement; readonly messages: Map<string, { help: string; error: string }> }> {
	const captured = parseCaptured(markup, prefixes)
	const { document, namespaces, node } = captured
	const children: OdfElement[] = []
	const messages = new Map<string, { help: string; error: string }>()
	let text = ""
	let step = 0

	yield

	for (let child = document.firstChild(node); child >= 0; child = document.nextSibling(child)) {
		const kind = document.kind(child)

		if (kind === NodeKind.Text || kind === NodeKind.CData) text += document.text(child)
		if (kind !== NodeKind.Element) continue
		if (++step % VALIDATIONS_PER_STEP === 0) yield

		children.push(readElement(document, namespaces, child, 1))

		if (elementName(document, namespaces, child) !== "table:content-validation") continue

		const name = canonicalAttributes(document, namespaces, child).get("table:name")

		if (name === undefined) continue

		const own = { help: "", error: "" }

		for (let message = document.firstChildElement(child); message >= 0; message = document.nextSiblingElement(message)) {
			const which = elementName(document, namespaces, message)

			if (which === "table:help-message") own.help = capturedText({ document, namespaces, node: message })
			else if (which === "table:error-message") own.error = capturedText({ document, namespaces, node: message })
		}

		messages.set(name, own)
	}

	return {
		element: {
			name: elementName(document, namespaces, node),
			attributes: canonicalAttributes(document, namespaces, node),
			children,
			text
		},
		messages
	}
}
