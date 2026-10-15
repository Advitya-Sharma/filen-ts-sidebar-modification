import type { Node } from "prosemirror-model"
import { OMML_NAMESPACE } from "../../../math/ommlWriter"
import { Namespaces, NodeKind, encodeUtf8, parseXml, type XmlDocument } from "../../../xml/index"
import { NO_SOURCE, docxSchema } from "../../edit/schema"
import type { StoryPart } from "../../model/index"
import { src } from "../source"
import { DocxWriteError } from "../values"
import type { AtomWriter } from "./index"

// Equations with new OMML (`omml`, as writeOmml writes it): a new one, or one from the file edited as LaTeX.
// The math namespace is the one the part's root binds, under its prefix; where the root binds none the
// equation declares it itself (as LibreOffice writes equations), which every part and every reader takes as
// it is. An edited equation keeps the range markers and comment references it held where they were (the edit
// applies around them: ommlTrack.ts).

export const STRICT_OMML_NAMESPACE = "http://purl.oclc.org/ooxml/officeDocument/math"
const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const STRICT_W_NAMESPACE = "http://purl.oclc.org/ooxml/wordprocessingml/main"
const W_URIS: ReadonlySet<string> = new Set([W_NAMESPACE, STRICT_W_NAMESPACE])
// Tracked changes: their wrappers, and their marks in m:ctrlPr.
const CHANGES: ReadonlySet<string> = new Set(["ins", "del", "moveFrom", "moveTo"])

// Per part: the prefix its root binds the math namespace to (null: none).
const rootPrefixes = new WeakMap<StoryPart, Map<string, string | null>>()

function rootPrefix(part: StoryPart, uri: string): string | null {
	let known = rootPrefixes.get(part)

	if (known === undefined) {
		known = new Map()
		rootPrefixes.set(part, known)
	}

	let prefix = known.get(uri)

	if (prefix === undefined) {
		prefix = part.reader.namespaces.prefixFor(part.document.documentElement(), uri) ?? null
		known.set(uri, prefix)
	}

	return prefix
}

// `omml` (m: and w: prefixed as the editor keeps it, no declarations) under the prefixes `prefix` (math) and
// `word`; a namespace the part's root does not bind (`declare`, `declareWord`) declared on the equation's root
// element. Text is escaped and the editor's markup has no other attributes of the two namespaces, so "<m:",
// " m:val=" and the like appear in tags alone.
export function ommlWithPrefix(
	omml: string,
	prefix: string,
	uri: string,
	declare: boolean,
	word: { readonly prefix: string; readonly uri: string; readonly declare: boolean } = { prefix: "w", uri: W_NAMESPACE, declare: false }
): string {
	let out = prefix === "m" ? omml : omml.replace(/<(\/?)m:/g, `<$1${prefix}:`).replace(/ m:([A-Za-z]+)="/g, ` ${prefix}:$1="`)

	if (word.prefix !== "w") out = out.replace(/<(\/?)w:/g, `<$1${word.prefix}:`).replace(/ w:([A-Za-z]+)="/g, ` ${word.prefix}:$1="`)
	if (declare) out = out.replace(/^<([^\s>/]+)/, `<$1 xmlns:${prefix}="${uri}"`)
	if (word.declare) out = out.replace(/^<([^\s>/]+)/, `<$1 xmlns:${word.prefix}="${word.uri}"`)

	return out
}

// The editor's equation XML without the declarations on its root element of namespaces the part's root binds
// alike (Word 365's w14, w16du... a source equation's markup uses): the part's own declarations hold.
function withoutRootDeclarations(omml: string, part: StoryPart): string {
	const end = omml.indexOf(">")

	if (end < 0 || !omml.slice(0, end).includes(" xmlns:")) return omml

	const bound = part.reader.namespaces.inScope(part.document.documentElement())
	const start = omml
		.slice(0, end)
		.replace(/ xmlns:([^\s=]+)="([^"]*)"/g, (declaration, prefix: string, uri: string) =>
			bound.get(prefix) === uri ? "" : declaration
		)

	return start + omml.slice(end)
}

// Each tracked change of the editor's equation XML numbered anew (w:id: each save's own, unique in the package).
function renumbered(omml: string, next: () => number): string {
	return omml.replace(/<w:(ins|del|moveFrom|moveTo)(\s[^>]*)?>/g, tag => tag.replace(/ w:id="[^"]*"/, ` w:id="${String(next())}"`))
}

export const mathAtoms: readonly AtomWriter[] = [
	{
		type: docxSchema.nodes.math,
		edited: (node: Node) => (node.attrs["omml"] ?? null) !== null,
		write(w, out, node) {
			const omml = node.attrs["omml"] as string | null

			if (omml === null) throw new DocxWriteError("missingSource", "A new math")

			const uri = w.env.context.model.strict ? STRICT_OMML_NAMESPACE : OMML_NAMESPACE
			// A default namespace binding is no prefix to write the equation under.
			const bound = rootPrefix(w.part, uri) || null
			const wordUri = w.env.context.model.strict ? STRICT_W_NAMESPACE : W_NAMESPACE
			const wordBound = rootPrefix(w.part, wordUri) || null
			const written = ommlWithPrefix(
				withoutRootDeclarations(
					renumbered(omml, () => w.env.nextRevisionId()),
					w.part
				),
				bound ?? "m",
				uri,
				bound === null,
				{
					prefix: wordBound ?? "w",
					uri: wordUri,
					declare: wordBound === null
				}
			)
			out.text(written)
		}
	}
]

// An equation as local names, m:val values and the text of its math runs, its tracked changes as their kind,
// author and date (their ids are each save's own), range markers and comment references left out (they are the
// node's markers): what it is written as, to compare with what reads back. A new or edited equation is its OMML,
// one from the file its element in `part`.
export function mathSignature(node: Node, part: StoryPart | undefined): string {
	const omml = node.attrs["omml"] as string | null
	let document: XmlDocument
	let element: number
	let reader: Namespaces

	if (omml !== null) {
		document = parseXml(encodeUtf8(`<math xmlns:m="${OMML_NAMESPACE}" xmlns:w="${W_NAMESPACE}">${omml}</math>`))
		element = document.firstChildElement(document.documentElement())
		reader = new Namespaces(document, { lenient: true })
	} else {
		const own = src(node)

		if (part === undefined || own === NO_SOURCE) return ""

		document = part.document
		element = part.node(own)
		reader = part.reader.namespaces
	}

	const out: string[] = []
	// Iterative (an equation may nest deeper than the stack goes): a node to read, or a closing mark to write.
	const stack: (readonly [number, boolean] | string)[] = element >= 0 ? [[element, false]] : []

	for (let next = stack.pop(); next !== undefined; next = stack.pop()) {
		if (typeof next === "string") {
			out.push(next)
			continue
		}

		const [at, text] = next
		const kind = document.kinds[at]

		if (kind === NodeKind.Text || kind === NodeKind.CData) {
			if (text) out.push(`#${document.text(at)}`)

			continue
		}

		if (kind !== NodeKind.Element) continue

		const uri = reader.elementUri(at)
		const local = document.localName(at)
		const children = (inText: boolean): void => {
			const list: number[] = []

			for (let child = document.firstChild(at); child >= 0; child = document.nextSibling(child)) list.push(child)

			for (let k = list.length - 1; k >= 0; k--) stack.push([list[k] ?? -1, inText])
		}

		if (uri === OMML_NAMESPACE || uri === STRICT_OMML_NAMESPACE) {
			const val = reader.attribute(at, uri, "val")

			out.push(`<${local}${val === undefined ? "" : `=${val}`}>`)
			stack.push("/")
			children(local === "t")
		} else if (uri !== undefined && W_URIS.has(uri)) {
			if (!CHANGES.has(local)) {
				children(false)
				continue
			}

			out.push(`[${local}:${reader.attribute(at, uri, "author") ?? ""}:${reader.attribute(at, uri, "date") ?? ""}`)
			stack.push("]")
			children(false)
		}
	}

	return out.join("")
}
