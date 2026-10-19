import { writeSync } from "node:fs"

// Markup Compatibility (ECMA-376 Part 3 §9-10) for a consumer that understands only the ECMA namespaces, as a
// streaming byte filter so a 100 MB part costs no tree: an element or attribute of a namespace an in-scope
// mc:Ignorable lists is dropped (an element's content is kept when an in-scope mc:ProcessContent names it),
// AlternateContent becomes its Fallback's content (every Choice requires an extension), mc attributes go.
// Character data, comments and processing instructions are copied byte for byte. Independent of the package's
// XML parser on purpose: it feeds the schema oracle.

export const MC_NAMESPACE = "http://schemas.openxmlformats.org/markup-compatibility/2006"

const LT = 0x3c
const GT = 0x3e
const SLASH = 0x2f
const QUOTE = 0x22
const APOSTROPHE = 0x27

interface Frame {
	readonly qname: string
	// Namespace declarations this element makes.
	readonly bindings: Map<string, string> | undefined
	readonly ignorable: Set<string> | undefined
	// "uri local" or "uri *" entries of its mc:ProcessContent.
	readonly processContent: Set<string> | undefined
	// Written to the output (an unwrapped element's tags are not).
	readonly emitted: boolean
}

export class NotWellFormed extends Error {}

const decoder = new TextDecoder("utf-8", { fatal: true })
const encoder = new TextEncoder()

function find(input: Uint8Array, needle: string, from: number): number {
	const first = needle.charCodeAt(0)

	for (let at = input.indexOf(first, from); at !== -1; at = input.indexOf(first, at + 1)) {
		let index = 1

		while (index < needle.length && input[at + index] === needle.charCodeAt(index)) index++

		if (index === needle.length) return at
	}

	throw new NotWellFormed(`Unterminated markup: ${needle} missing`)
}

// The end of the start tag at `start` ("<"): the index of its ">", honouring quoted attribute values.
function tagEnd(input: Uint8Array, start: number): number {
	let quote = 0

	for (let at = start + 1; at < input.length; at++) {
		const byte = input[at] as number

		if (quote !== 0) {
			if (byte === quote) quote = 0
		} else if (byte === QUOTE || byte === APOSTROPHE) quote = byte
		else if (byte === GT) return at
	}

	throw new NotWellFormed("Unterminated start tag")
}

const ATTRIBUTE = /([^\s=]+)\s*=\s*("([^"]*)"|'([^']*)')/g

function parseTag(text: string): { qname: string; attributes: [string, string, string][] } {
	const nameEnd = text.search(/[\s/>]/)
	const qname = text.slice(1, nameEnd)
	const attributes: [string, string, string][] = []

	ATTRIBUTE.lastIndex = nameEnd

	for (let match = ATTRIBUTE.exec(text); match !== null; match = ATTRIBUTE.exec(text))
		attributes.push([match[1] as string, match[3] ?? match[4] ?? "", match[0]])

	return { qname, attributes }
}

function prefixOf(qname: string): string {
	const colon = qname.indexOf(":")

	return colon === -1 ? "" : qname.slice(0, colon)
}

function localOf(qname: string): string {
	return qname.slice(qname.indexOf(":") + 1)
}

// Buffered output: one write per MiB.
class Writer {
	private readonly buffer = new Uint8Array(1 << 20)
	private size = 0
	private readonly fd: number

	constructor(fd: number) {
		this.fd = fd
	}

	bytes(bytes: Uint8Array): void {
		if (this.size + bytes.length > this.buffer.length) this.flush()
		if (bytes.length > this.buffer.length) {
			writeSync(this.fd, bytes)
			return
		}

		this.buffer.set(bytes, this.size)
		this.size += bytes.length
	}

	text(text: string): void {
		if (this.size + text.length * 3 > this.buffer.length) this.flush()
		if (text.length * 3 > this.buffer.length) {
			writeSync(this.fd, encoder.encode(text))
			return
		}

		this.size += encoder.encodeInto(text, this.buffer.subarray(this.size)).written
	}

	flush(): void {
		if (this.size > 0) writeSync(this.fd, this.buffer, 0, this.size)

		this.size = 0
	}
}

// UTF-16 input as UTF-8 without its XML declaration (which names the old encoding).
function utf8(input: Uint8Array): Uint8Array {
	const le = input[0] === 0xff && input[1] === 0xfe
	const be = input[0] === 0xfe && input[1] === 0xff

	if (!le && !be) return input[0] === 0xef && input[1] === 0xbb && input[2] === 0xbf ? input.subarray(3) : input

	const text = new TextDecoder(le ? "utf-16le" : "utf-16be").decode(input)

	return encoder.encode(text.replace(/^<\?xml[^?]*\?>/, ""))
}

// Writes `source` with Markup Compatibility applied to `fd`. Throws NotWellFormed on markup it cannot follow.
export function applyMarkupCompatibility(source: Uint8Array, fd: number): void {
	const input = utf8(source)
	const out = new Writer(fd)
	const stack: Frame[] = []
	// Depth inside a dropped subtree (0: not dropping).
	let skip = 0
	let at = 0

	const resolve = (prefix: string): string | undefined => {
		if (prefix === "xml") return "http://www.w3.org/XML/1998/namespace"

		for (let index = stack.length - 1; index >= 0; index--) {
			const uri = stack[index]?.bindings?.get(prefix)

			if (uri !== undefined) return uri
		}

		return prefix === "" ? "" : undefined
	}
	const ignorable = (uri: string): boolean => stack.some(frame => frame.ignorable?.has(uri) === true)
	const processed = (uri: string, local: string): boolean =>
		stack.some(frame => frame.processContent?.has(`${uri} ${local}`) === true || frame.processContent?.has(`${uri} *`) === true)
	// The frame that declares `prefix`, if it was not written: its children redeclare the binding.
	const unwritten = (prefix: string): string | undefined => {
		for (let index = stack.length - 1; index >= 0; index--) {
			const frame = stack[index] as Frame
			const uri = frame.bindings?.get(prefix)

			if (uri !== undefined) return frame.emitted ? undefined : uri
		}

		return undefined
	}

	while (at < input.length) {
		const open = input.indexOf(LT, at)
		const textEnd = open === -1 ? input.length : open

		if (skip === 0) out.bytes(input.subarray(at, textEnd))
		if (open === -1) break

		const next = input[open + 1]

		if (next === 0x3f) {
			const end = find(input, "?>", open) + 2

			if (skip === 0) out.bytes(input.subarray(open, end))
			at = end
			continue
		}

		if (next === 0x21) {
			const end =
				input[open + 2] === 0x2d
					? find(input, "-->", open) + 3
					: input[open + 2] === 0x5b
						? find(input, "]]>", open) + 3
						: find(input, ">", open) + 1

			if (skip === 0) out.bytes(input.subarray(open, end))
			at = end
			continue
		}

		if (next === SLASH) {
			const end = find(input, ">", open) + 1

			at = end

			if (skip > 0) {
				skip--
				continue
			}

			const frame = stack.pop()

			if (frame === undefined) throw new NotWellFormed("End tag without a start tag")
			if (frame.emitted) out.text(`</${frame.qname}>`)
			continue
		}

		const end = tagEnd(input, open)
		const selfClosing = input[end - 1] === SLASH

		at = end + 1

		if (skip > 0) {
			if (!selfClosing) skip++
			continue
		}

		let text: string

		try {
			text = decoder.decode(input.subarray(open, end + 1))
		} catch {
			throw new NotWellFormed("Invalid UTF-8 in a tag")
		}

		const { qname, attributes } = parseTag(text)
		let bindings: Map<string, string> | undefined

		for (const [name, value] of attributes) {
			if (name === "xmlns") (bindings ??= new Map()).set("", value)
			else if (name.startsWith("xmlns:")) (bindings ??= new Map()).set(name.slice(6), value)
		}

		// Bindings first: the element's own declarations are in scope for its name and attributes.
		stack.push({ qname, bindings, ignorable: undefined, processContent: undefined, emitted: false })

		const uri = resolve(prefixOf(qname))

		if (uri === undefined) throw new NotWellFormed(`Undeclared prefix of ${qname}`)

		const local = localOf(qname)
		let ignorableSet: Set<string> | undefined
		let processContent: Set<string> | undefined

		for (const [name, value] of attributes) {
			if (resolve(prefixOf(name)) !== MC_NAMESPACE || prefixOf(name) === "") continue

			const tokens = value.split(/\s+/).filter(token => token !== "")

			if (localOf(name) === "Ignorable") {
				ignorableSet = new Set(tokens.map(prefix => resolve(prefix)).filter((value): value is string => value !== undefined))
			} else if (localOf(name) === "ProcessContent") {
				processContent = new Set(
					tokens.flatMap(token => {
						const namespace = resolve(prefixOf(token))

						return namespace === undefined ? [] : [`${namespace} ${localOf(token)}`]
					})
				)
			}
		}

		stack.pop()

		const dropped = uri !== MC_NAMESPACE && ignorable(uri)
		const unwrap =
			(uri === MC_NAMESPACE && (local === "AlternateContent" || local === "Fallback")) || (dropped && processed(uri, local))

		if ((dropped && !unwrap) || (uri === MC_NAMESPACE && !unwrap)) {
			// An ignored element, a Choice, or MC markup out of place: the whole subtree goes.
			if (!selfClosing) skip = 1
			continue
		}

		const frame: Frame = { qname, bindings, ignorable: ignorableSet, processContent, emitted: !unwrap }

		if (!unwrap) {
			let tag = `<${qname}`
			const redeclared = new Set<string>()

			for (const name of [qname, ...attributes.map(([attribute]) => attribute)]) {
				const prefix = prefixOf(name)

				if (name.startsWith("xmlns") || redeclared.has(prefix) || bindings?.has(prefix) === true) continue
				if (prefix === "" && name !== qname) continue

				const binding = unwritten(prefix)

				if (binding !== undefined) {
					redeclared.add(prefix)
					tag += prefix === "" ? ` xmlns="${binding}"` : ` xmlns:${prefix}="${binding}"`
				}
			}

			stack.push(frame)

			for (const [name, , raw] of attributes) {
				const prefix = prefixOf(name)

				if (prefix !== "" && prefix !== "xmlns" && name !== "xmlns") {
					const namespace = resolve(prefix)

					if (namespace === undefined) throw new NotWellFormed(`Undeclared prefix of ${name}`)
					if (namespace === MC_NAMESPACE || ignorable(namespace)) continue
				}

				tag += ` ${raw}`
			}

			stack.pop()
			out.text(selfClosing ? `${tag}/>` : `${tag}>`)
		}

		if (!selfClosing) stack.push(frame)
	}

	if (stack.length > 0 || skip > 0) throw new NotWellFormed("Unclosed elements")

	out.flush()
}
