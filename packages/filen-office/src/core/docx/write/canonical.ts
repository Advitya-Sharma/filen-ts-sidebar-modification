import { NodeKind, type XmlDocument } from "../../xml/index"

// An element subtree in canonical form, for comparing what the edit path writes with the source at the
// level of XML meaning: attributes in name order with decoded values, character references resolved,
// empty-element form ignored, whitespace between elements ignored (inside text elements it counts). What the
// edit path may legitimately change is left out too: w:proofErr and w:lastRenderedPageBreak (Word
// regenerates both), runs and inline wrappers without content (or only empty text), empty property
// containers, xml:space as written (text is compared as Word reads it), and the split of one run's text
// into several w:t.

const SKIPPED = new Set(["proofErr", "lastRenderedPageBreak"])
const TEXT = new Set(["t", "delText", "instrText", "delInstrText"])
const PROPERTIES = new Set(["rPr", "pPr"])
// Inline wrappers: without content the editor has nothing of them to keep.
const WRAPPERS = new Set(["hyperlink", "smartTag", "customXml", "ins", "del", "moveFrom", "moveTo", "dir", "bdo"])

const TRIM = /^[ \t\r\n]+|[ \t\r\n]+$/g

function preserves(document: XmlDocument, node: number): boolean {
	for (let at = node; at > 0; at = document.parent(at)) {
		const space = document.attribute(at, "xml:space")

		if (space !== undefined) return space === "preserve"
	}

	return false
}

// A text element's text as Word reads it: edge whitespace counts only under xml:space="preserve".
function readText(document: XmlDocument, node: number): string {
	const text = document.textContent(node)

	return preserves(document, node) ? text : text.replace(TRIM, "")
}

function local(name: string): string {
	return name.slice(name.indexOf(":") + 1)
}

function isWhitespace(text: string): boolean {
	return /^[ \t\r\n]*$/.test(text)
}

function hasContent(document: XmlDocument, run: number): boolean {
	for (let child = document.firstChildElement(run); child >= 0; child = document.nextSiblingElement(child)) {
		const name = local(document.name(child))

		// An empty w:t is no content either: the editor has nothing to keep of it.
		if (name !== "rPr" && !SKIPPED.has(name) && !(TEXT.has(name) && readText(document, child) === "")) return true
	}

	return false
}

function attributes(document: XmlDocument, node: number): string {
	return document
		.attributes(node)
		.filter(attribute => attribute.name !== "xml:space" && attribute.name !== "xmlns" && !attribute.name.startsWith("xmlns:"))
		.map(attribute => `${attribute.name}=${JSON.stringify(attribute.value)}`)
		.sort()
		.join(" ")
}

// With `shell`, paragraphs stand for their byte length and nested tables for nothing: a table's own
// markup is compared, its content is proved block by block.
function visit(
	document: XmlDocument,
	node: number,
	out: string[],
	shell: boolean = false,
	depth: number = 0,
	skip: ReadonlySet<string> | undefined = undefined
): void {
	const name = document.name(node)
	const localName = local(name)

	if (SKIPPED.has(localName)) return
	if (depth === 1 && skip?.has(localName) === true) return

	if (shell && depth > 0 && (localName === "p" || localName === "tbl")) {
		out.push(localName === "p" ? `P${document.end(node) - document.start(node)}` : "TBL")

		return
	}
	if (localName === "r" && !hasContent(document, node)) return
	if (PROPERTIES.has(localName) && document.firstChildElement(node) < 0) return

	if (TEXT.has(localName)) {
		out.push(`T${name} ${attributes(document, node)}>${readText(document, node)}`)

		return
	}

	out.push(`<${name} ${attributes(document, node)}>`)

	for (let child = document.firstChild(node); child >= 0; child = document.nextSibling(child)) {
		const kind = document.kinds[child]

		if (kind === NodeKind.Element) {
			const before = out.length

			visit(document, child, out, shell, depth + 1, skip)

			// One text element right after another of its kind (with the same attributes) is one text.
			const added = out[before]
			const previous = out[before - 1]

			if (
				added !== undefined &&
				previous !== undefined &&
				out.length === before + 1 &&
				added.startsWith("T") &&
				previous.startsWith("T")
			) {
				const split = added.indexOf(">")

				if (previous.slice(0, previous.indexOf(">")) === added.slice(0, split)) {
					out[before - 1] = previous + added.slice(split + 1)
					out.length = before
				}
			}
		} else if (kind === NodeKind.Text || kind === NodeKind.CData) {
			const text = document.text(child)

			if (!isWhitespace(text)) out.push(`#${text}`)
		}
	}

	if (WRAPPERS.has(localName) && out[out.length - 1] === `<${name} ${attributes(document, node)}>`) {
		out.pop()

		return
	}

	out.push(`</${name}>`)
}

// `skip`: local names of the element's own children to leave out.
export function canonicalElement(document: XmlDocument, node: number, shell: boolean = false, skip?: ReadonlySet<string>): string {
	const out: string[] = []

	visit(document, node, out, shell, 0, skip)

	return out.join("\n")
}

// The canonical form of every top-level element in a parsed fragment, joined.
export function canonicalFragment(document: XmlDocument, container: number, shell: boolean = false): string {
	const out: string[] = []

	for (let child = document.firstChildElement(container); child >= 0; child = document.nextSiblingElement(child))
		visit(document, child, out, shell)

	return out.join("\n")
}
