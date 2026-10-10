import type { XmlDocument } from "../../xml/index"
import { Ns, groupOf } from "../model/namespaces"
import { clampInt32, parseUniversalMeasure, partReader, type PartReader as ModelReader } from "../model/reader"

export { clampInt32, parseUniversalMeasure }

const readers = new WeakMap<XmlDocument, PropsReader>()

// WordprocessingML property reads over the docx layer's shared part reader (model/reader.ts): names
// compared as (namespace, local name), so Strict OOXML and generator-chosen prefixes work, and nothing
// throws on malformed markup (unreadable start tags land in `malformedNodes` and read as attribute-less).
export class PropsReader {
	readonly document: XmlDocument
	readonly model: ModelReader
	private readonly span = new Int32Array(2)
	private readonly signatures = new Map<ReadonlyMap<string, string>, string>()

	private constructor(document: XmlDocument) {
		this.document = document
		this.model = partReader(document)
	}

	// The WordprocessingML and relationship prefix bindings in scope at `node`: only those change what a
	// container parses to, so equal bytes are shared only between equal bindings. One string per scope map.
	signature(node: number): string {
		const scope = this.model.namespaces.inScope(node)
		let signature = this.signatures.get(scope)

		if (signature === undefined) {
			const bindings: string[] = []

			for (const [prefix, uri] of scope) {
				const group = groupOf(uri)

				if (group === Ns.W || group === Ns.R) bindings.push(`${prefix}=${uri}`)
			}

			signature = bindings.sort().join("\u0000")
			this.signatures.set(scope, signature)
		}

		return signature
	}

	static for(document: XmlDocument): PropsReader {
		let reader = readers.get(document)

		if (reader === undefined) {
			reader = new PropsReader(document)
			readers.set(document, reader)
		}

		return reader
	}

	get malformedNodes(): ReadonlySet<number> {
		return this.model.malformedNodes
	}

	// The local name of a WordprocessingML element; "" for anything else.
	local(node: number): string {
		return this.model.group(node) === Ns.W ? this.document.localName(node) : ""
	}

	// First WordprocessingML child element with the given local name.
	child(node: number, local: string): number {
		return this.model.child(node, Ns.W, local)
	}

	// Whether the element's start tag reads cleanly (a malformed one is recorded).
	wellFormed(node: number): boolean {
		return this.model.wellFormed(node)
	}

	has(node: number, local: string): boolean {
		return this.model.has(node, Ns.W, local)
	}

	string(node: number, local: string): string | undefined {
		return this.model.attr(node, Ns.W, local)
	}

	// The value as a plain number (ST_DecimalNumber and friends); undefined when absent or not numeric.
	number(node: number, local: string): number | undefined {
		const value = this.model.number(node, Ns.W, local, Number.NaN)

		return Number.isNaN(value) ? undefined : value
	}

	// Rounded and clamped to int32.
	integer(node: number, local: string): number | undefined {
		const value = this.number(node, local)

		return value === undefined ? undefined : clampInt32(value)
	}

	// A length: a bare number is already in the target unit; an ST_UniversalMeasure ("1.5cm", "12pt") is
	// converted with `perPoint` target units per point (20 for twips, 2 for half-points, 8 for eighths).
	measure(node: number, local: string, perPoint: number): number | undefined {
		const value = this.model.measure(node, Ns.W, local, perPoint, Number.NaN)

		return Number.isNaN(value) ? undefined : value
	}

	// ST_OnOff of an attribute that is present; `fallback` when it is absent.
	onOff(node: number, local: string, fallback: boolean): boolean {
		return this.optionalOnOff(node, local) ?? fallback
	}

	// An attribute holding ST_OnOff; undefined when absent.
	optionalOnOff(node: number, local: string): boolean | undefined {
		return this.model.find(node, Ns.W, local, this.span) ? this.model.flag(node, Ns.W, local) : undefined
	}

	// A toggle or on/off property element: present without w:val means on.
	toggle(node: number): boolean {
		return this.model.onOff(node)
	}

	// ST_HexColor: "auto" or six hex digits, upper-cased; undefined when invalid.
	hexColor(node: number, local: string): string | undefined {
		const text = this.string(node, local)

		return text === undefined ? undefined : normalizeHexColor(text)
	}

	// ST_UcharHexNumber (themeTint and friends): 0..255.
	hexByte(node: number, local: string): number | undefined {
		const value = this.model.hex(node, Ns.W, local, -1)

		return value >= 0 && value <= 255 ? value : undefined
	}

	// A relationship id attribute (r:id) in either the transitional or the Strict namespace.
	relationshipId(node: number, local: string = "id"): string | undefined {
		return this.model.attr(node, Ns.R, local)
	}
}

export function normalizeHexColor(text: string): string | undefined {
	if (text === "auto") return "auto"
	if (/^[0-9a-fA-F]{6}$/.test(text)) return text.toUpperCase()

	// Some producers write "#RRGGBB" or 8-digit ARGB.
	if (/^#[0-9a-fA-F]{6}$/.test(text)) return text.slice(1).toUpperCase()
	if (/^[0-9a-fA-F]{8}$/.test(text)) return text.slice(2).toUpperCase()

	return undefined
}
