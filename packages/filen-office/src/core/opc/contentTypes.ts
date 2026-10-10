import { Namespaces, XmlEditor, escapeAttribute, parseXml, type XmlDocument } from "../xml/index"
import { OpcError } from "./errors"
import { partKey, partNameUri } from "./partNames"

export const CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types"
export const CONTENT_TYPES_PART = "[Content_Types].xml"

export class ContentTypes {
	// Extension (lower case, no dot) -> media type.
	readonly defaults: ReadonlyMap<string, string>
	// Part key (see partKey) -> media type.
	readonly overrides: ReadonlyMap<string, string>

	constructor(defaults: ReadonlyMap<string, string>, overrides: ReadonlyMap<string, string>) {
		this.defaults = defaults
		this.overrides = overrides
	}

	// Media type of a part: its Override, else the Default for its extension.
	typeOf(partName: string): string | undefined {
		const key = partKey(partName)
		const override = this.overrides.get(key)

		if (override !== undefined) return override

		const dot = key.lastIndexOf(".")

		return dot > key.lastIndexOf("/") ? this.defaults.get(key.slice(dot + 1)) : undefined
	}
}

function typesRoot(bytes: Uint8Array): { document: XmlDocument; namespaces: Namespaces; root: number } {
	const document = parseXml(bytes)
	const namespaces = new Namespaces(document)
	const root = document.documentElement()

	if (!namespaces.is(root, CONTENT_TYPES_NAMESPACE, "Types")) {
		throw new OpcError("invalidContentTypes", "[Content_Types].xml has no Types root")
	}

	return { document, namespaces, root }
}

export function parseContentTypes(bytes: Uint8Array): ContentTypes {
	const { document, namespaces, root } = typesRoot(bytes)

	const defaults = new Map<string, string>()
	const overrides = new Map<string, string>()

	for (let node = document.firstChildElement(root); node >= 0; node = document.nextSiblingElement(node)) {
		const contentType = document.attribute(node, "ContentType")

		if (namespaces.is(node, CONTENT_TYPES_NAMESPACE, "Default")) {
			const extension = document.attribute(node, "Extension")

			if (extension === undefined || contentType === undefined) {
				throw new OpcError("invalidContentTypes", "Default without Extension or ContentType")
			}

			if (!defaults.has(partKey(extension))) defaults.set(partKey(extension), contentType)
		} else if (namespaces.is(node, CONTENT_TYPES_NAMESPACE, "Override")) {
			const partName = document.attribute(node, "PartName")

			if (partName === undefined || contentType === undefined) {
				throw new OpcError("invalidContentTypes", "Override without PartName or ContentType")
			}

			if (!overrides.has(partKey(partName))) overrides.set(partKey(partName), contentType)
		}
	}

	return new ContentTypes(defaults, overrides)
}

export interface ContentTypeChanges {
	// Add, or change the type of, a Default by extension (without the dot).
	readonly defaults?: readonly { readonly extension: string; readonly contentType: string }[]
	// Add, or change the type of, an Override by part name.
	readonly overrides?: readonly { readonly partName: string; readonly contentType: string }[]
	readonly removeOverrides?: readonly string[]
}

// [Content_Types].xml with `changes` spliced in: existing entries change in place, new Defaults go before
// the first Override (where Word keeps them), new Overrides go last, all with the part's own prefix for the
// namespace. Everything else stays byte for byte.
export function editContentTypes(bytes: Uint8Array, changes: ContentTypeChanges): Uint8Array {
	const { document, namespaces, root } = typesRoot(bytes)
	const editor = new XmlEditor(document)
	const defaults = new Map<string, number>()
	const overrides = new Map<string, number>()
	let firstOverride = -1

	for (let node = document.firstChildElement(root); node >= 0; node = document.nextSiblingElement(node)) {
		if (namespaces.is(node, CONTENT_TYPES_NAMESPACE, "Default")) {
			const key = partKey(document.attribute(node, "Extension") ?? "")

			if (!defaults.has(key)) defaults.set(key, node)
		} else if (namespaces.is(node, CONTENT_TYPES_NAMESPACE, "Override")) {
			const key = partKey(document.attribute(node, "PartName") ?? "")

			if (firstOverride < 0) firstOverride = node
			if (!overrides.has(key)) overrides.set(key, node)
		}
	}

	const prefix = namespaces.prefixFor(root, CONTENT_TYPES_NAMESPACE) ?? ""
	const qualified = (local: string): string => (prefix === "" ? local : `${prefix}:${local}`)

	// New entries are keyed like existing ones, so a repeat in `changes` updates the pending one.
	const newDefaults = new Map<string, { extension: string; contentType: string }>()
	const newOverrides = new Map<string, { partName: string; contentType: string }>()

	for (const { extension, contentType } of changes.defaults ?? []) {
		const key = partKey(extension)
		const existing = defaults.get(key)

		if (existing !== undefined) {
			editor.setAttribute(existing, "ContentType", contentType)
		} else {
			newDefaults.set(key, { extension: newDefaults.get(key)?.extension ?? extension, contentType })
		}
	}

	const removed = new Set((changes.removeOverrides ?? []).map(partKey))

	for (const key of removed) {
		const existing = overrides.get(key)

		if (existing !== undefined) editor.remove(existing)
	}

	for (const { partName, contentType } of changes.overrides ?? []) {
		const key = partKey(partName)
		const existing = overrides.get(key)

		if (existing !== undefined && !removed.has(key)) {
			editor.setAttribute(existing, "ContentType", contentType)
		} else {
			newOverrides.set(key, { partName: newOverrides.get(key)?.partName ?? partName, contentType })
		}
	}

	for (const { extension, contentType } of newDefaults.values()) {
		const markup = `<${qualified("Default")} Extension="${escapeAttribute(extension)}" ContentType="${escapeAttribute(contentType)}"/>`

		if (firstOverride >= 0) {
			editor.insertBefore(firstOverride, markup)
		} else {
			editor.append(root, markup)
		}
	}

	for (const { partName, contentType } of newOverrides.values()) {
		editor.append(
			root,
			`<${qualified("Override")} PartName="${escapeAttribute(partNameUri(partName))}" ContentType="${escapeAttribute(contentType)}"/>`
		)
	}

	return editor.serialize()
}
