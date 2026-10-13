import { NodeKind, escapeAttribute } from "../../xml/index"
import type { StoryPart } from "../model/index"
import { Ns } from "../model/namespaces"
import { DocxWriteError } from "../write/values"

// Property elements (w:tblPr, w:trPr, w:tcPr, w:sdtPr) of another part's source items, which a copy into a
// part keeps as data: that part's items are not this one's, and its prefixes may differ. Written back with the
// prefixes of the part the copy goes into.

// An element: its namespace (MAIN: WordprocessingML's, whichever of its URIs the part binds), local name,
// attributes ([namespace, local name, value]; "" for an unprefixed one) and child elements.
export interface ElementData {
	readonly uri: string
	readonly name: string
	readonly attrs: readonly (readonly [string, string, string])[]
	readonly children: readonly ElementData[]
	// The namespaces the source part declared ignorable (mc:Ignorable) that the element uses: they stay so.
	readonly ignorable?: readonly string[]
}

export const MAIN = "w"
// A prefix the source part does not bind (its markup is not namespace-well-formed there): such data cannot be
// written, and a copy holding it is refused (writableData).
const UNBOUND = "?"

const XML_NAME = /^[A-Za-z_][\w.-]*$/
const XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace"
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
// Markup compatibility attributes whose value names namespaces by prefix (ECMA-376 Part 3 §7): their prefixes
// are carried as "{uri}" (with the rest of a qualified name after it) and written with the part's own.
const MC_PREFIXED: ReadonlySet<string> = new Set([
	"Ignorable",
	"ProcessContent",
	"MustUnderstand",
	"PreserveElements",
	"PreserveAttributes"
])

function prefixed(elementUri: string, element: string, uri: string, local: string): boolean {
	return (uri === MC && MC_PREFIXED.has(local)) || (elementUri === MC && element === "Choice" && uri === "" && local === "Requires")
}

// WordprocessingML's URIs (Transitional, Strict).
const MAIN_URIS: ReadonlySet<string> = new Set([
	"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
	"http://purl.oclc.org/ooxml/wordprocessingml/main"
])

// Element `node` of `part` as data, without its WordprocessingML children named in `leaveOut`.
export function elementData(part: StoryPart, node: number, leaveOut: ReadonlySet<string> = new Set()): ElementData {
	const data = readElement(part, node, leaveOut)
	const root = part.document.documentElement()
	const namespaces = part.reader.namespaces
	const declared = root < 0 ? undefined : namespaces.attribute(root, MC, "Ignorable")
	const ignorable = new Set<string>()

	for (const token of (declared ?? "").split(/\s+/)) {
		const uri = token === "" ? undefined : namespaces.uri(root, token)

		if (uri !== undefined) ignorable.add(uri)
	}

	const used = [...usedUris(data, new Set())].filter(uri => ignorable.has(uri))

	return used.length === 0 ? data : { ...data, ignorable: used }
}

// The namespaces of `data`'s elements and attributes.
function usedUris(data: ElementData, out: Set<string>): Set<string> {
	if (data.uri !== MAIN && data.uri !== "") out.add(data.uri)

	for (const [uri] of data.attrs) if (uri !== MAIN && uri !== "") out.add(uri)
	for (const child of data.children) usedUris(child, out)

	return out
}

// Whether every name and prefix of `data` can be written (see UNBOUND).
export function writableData(data: ElementData): boolean {
	if (data.uri === UNBOUND) return false

	for (const [uri, , value] of data.attrs) if (uri === UNBOUND || value.includes(`{${UNBOUND}}`)) return false

	return data.children.every(writableData)
}

function readElement(part: StoryPart, node: number, leaveOut: ReadonlySet<string>): ElementData {
	const document = part.document
	const reader = part.reader
	const namespaces = reader.namespaces
	const uriOf = (element: number): string => (reader.group(element) === Ns.W ? MAIN : (namespaces.elementUri(element) ?? UNBOUND))
	const attrs: [string, string, string][] = []
	const children: ElementData[] = []

	for (const { name, value } of document.attributes(node)) {
		const colon = name.indexOf(":")
		const prefix = colon < 0 ? "" : name.slice(0, colon)

		if (prefix === "xmlns" || name === "xmlns") continue

		const uri = prefix === "" ? "" : (namespaces.uri(node, prefix) ?? UNBOUND)

		const local = name.slice(colon + 1)
		const tokens = (): string =>
			value
				.split(/\s+/)
				.filter(token => token !== "")
				.map(token => {
					const at = token.indexOf(":")
					const bound = namespaces.uri(node, at < 0 ? token : token.slice(0, at)) ?? UNBOUND

					return `{${MAIN_URIS.has(bound) ? MAIN : bound}}${at < 0 ? "" : token.slice(at + 1)}`
				})
				.join(" ")

		attrs.push([MAIN_URIS.has(uri) ? MAIN : uri, local, prefixed(uriOf(node), document.localName(node), uri, local) ? tokens() : value])
	}

	for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) {
		if (document.kinds[child] !== NodeKind.Element) continue

		const uri = uriOf(child)

		if (uri === MAIN && leaveOut.has(document.localName(child))) continue

		children.push(readElement(part, child, new Set()))
	}

	return { uri: uriOf(node), name: document.localName(node), attrs, children }
}

// The value of WordprocessingML attribute w:`name` of `data`.
export function dataAttr(data: ElementData, name: string): string | undefined {
	return data.attrs.find(([uri, local]) => uri === MAIN && local === name)?.[2]
}

// Child elements of `data` as markup for `part` by WordprocessingML local name (other namespaces' by
// "{uri}name", which property splicing puts last), for property writers to splice their changes into.
export function dataChildren(part: StoryPart, w: string, data: ElementData): Map<string, string> {
	const out = new Map<string, string>()

	for (const child of data.children) {
		const key = child.uri === MAIN ? child.name : `{${child.uri}}${child.name}`

		out.set(
			key,
			(out.get(key) ?? "") + dataMarkup(part, w, data.ignorable === undefined ? child : { ...child, ignorable: data.ignorable })
		)
	}

	return out
}

// `data` as markup for `part` (`w`: the prefix the part binds WordprocessingML to; a namespace it does not bind
// is declared on the element using it).
export function dataMarkup(part: StoryPart, w: string, data: ElementData): string {
	const root = part.document.documentElement()
	// A namespace the part does not bind is declared on each element using it.
	const prefixOf = (uri: string, declared: Map<string, string>): string => {
		if (uri === MAIN) return w
		if (uri === XML_NAMESPACE) return "xml"

		const bound = root < 0 ? undefined : part.reader.namespaces.prefixFor(root, uri)

		if (bound !== undefined && bound !== "") return bound

		let prefix = declared.get(uri)

		// One the part binds to nothing, so nothing it holds is shadowed.
		for (let k = declared.size + 1; prefix === undefined; k++) {
			const candidate = `ns${String(k)}`

			if ((root < 0 || part.reader.namespaces.uri(root, candidate) === undefined) && ![...declared.values()].includes(candidate)) {
				prefix = candidate
				declared.set(uri, prefix)
			}
		}

		return prefix
	}
	// "{uri}rest" tokens of a markup compatibility attribute with this part's prefixes.
	const tokens = (value: string, declared: Map<string, string>): string =>
		value
			.split(" ")
			.filter(token => token !== "")
			.map(token => {
				const close = token.indexOf("}")
				const uri = token.slice(1, close)

				if (!token.startsWith("{") || close < 0 || uri === "" || uri === UNBOUND)
					throw new DocxWriteError("invalidValue", `Copied markup compatibility prefix ${JSON.stringify(token)}`)

				const prefix = prefixOf(uri, declared)
				const rest = token.slice(close + 1)

				return rest === "" ? prefix : `${prefix}:${rest}`
			})
			.join(" ")
	// What the source declared ignorable and this part does not: the written element says so itself (an
	// application that does not understand it still ignores it, ECMA-376 Part 3 §7.2).
	const rootIgnorable = new Set<string>()
	const declaredIgnorable = root < 0 ? undefined : part.reader.namespaces.attribute(root, MC, "Ignorable")

	for (const token of (declaredIgnorable ?? "").split(/\s+/)) {
		const uri = token === "" ? undefined : part.reader.namespaces.uri(root, token)

		if (uri !== undefined) rootIgnorable.add(uri)
	}

	const ignorable = (data.ignorable ?? []).filter(uri => !rootIgnorable.has(uri) && usedUris(data, new Set()).has(uri))
	const write = (element: ElementData, top = false): string => {
		if (element.uri === UNBOUND || element.attrs.some(([uri]) => uri === UNBOUND))
			throw new DocxWriteError("invalidValue", `Copied property element ${JSON.stringify(element.name)} with an unbound prefix`)
		if (!XML_NAME.test(element.name))
			throw new DocxWriteError("invalidValue", `Copied property element ${JSON.stringify(element.name)}`)

		const declared = new Map<string, string>()
		const prefix = prefixOf(element.uri, declared)
		const qualified = prefix === "" ? element.name : `${prefix}:${element.name}`
		let out = `<${qualified}`

		const own = top ? element.attrs.find(([uri, local]) => uri === MC && local === "Ignorable") : undefined
		const attrs =
			own === undefined || ignorable.length === 0
				? element.attrs
				: element.attrs.map(attr =>
						attr === own
							? ([
									MC,
									"Ignorable",
									[attr[2], ...ignorable.filter(uri => !attr[2].includes(`{${uri}}`)).map(uri => `{${uri}}`)]
										.join(" ")
										.trim()
								] as const)
							: attr
					)

		for (const [uri, local, value] of attrs) {
			if (!XML_NAME.test(local)) throw new DocxWriteError("invalidValue", `Copied property attribute ${JSON.stringify(local)}`)

			const at = uri === "" ? "" : prefixOf(uri, declared)
			const written = prefixed(element.uri, element.name, uri, local) ? tokens(value, declared) : value

			out += ` ${at === "" ? local : `${at}:${local}`}="${escapeAttribute(written)}"`
		}

		if (top && ignorable.length > 0 && own === undefined)
			out += ` ${prefixOf(MC, declared)}:Ignorable="${ignorable.map(uri => prefixOf(uri, declared)).join(" ")}"`

		for (const [uri, declaredPrefix] of declared) out += ` xmlns:${declaredPrefix}="${escapeAttribute(uri)}"`

		if (element.children.length === 0) return `${out}/>`

		return `${out}>${element.children.map(child => write(child)).join("")}</${qualified}>`
	}

	return write(data, true)
}
