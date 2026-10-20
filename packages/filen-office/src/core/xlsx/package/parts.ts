import { CONTENT_TYPES_PART, type ContentTypes } from "../../opc/contentTypes"
import type { OpcPackage } from "../../opc/package"
import { partKey } from "../../opc/partNames"
import type { Relationships } from "../../opc/relationships"
import type { Slicing } from "../../slicing"
import type { XlsxPart, XlsxPartMap } from "../model/types"

// The source package's parts, read lazily: an entry is inflated only when asked for and never kept here.
// Relationships, small and read often, are parsed once and then also answer synchronously (`loadedRelationships`).
export class XlsxParts implements XlsxPartMap {
	readonly opc: OpcPackage
	readonly contentTypes: ContentTypes
	private readonly byKey = new Map<string, XlsxPart>()
	private readonly relationshipCache = new Map<string, Relationships>()

	constructor(opc: OpcPackage, contentTypes: ContentTypes) {
		this.opc = opc
		this.contentTypes = contentTypes
	}

	get(name: string): XlsxPart | undefined {
		const key = partKey(name)
		let part = this.byKey.get(key)

		if (part === undefined) {
			const entry = this.opc.entry(name)

			if (entry === undefined || key === partKey(CONTENT_TYPES_PART)) return undefined

			part = { name: entry.name, contentType: this.contentTypes.typeOf(entry.name) ?? "", source: entry }
			this.byKey.set(key, part)
		}

		return part
	}

	*parts(): Iterable<XlsxPart> {
		for (const entry of this.opc.zip.entries) {
			if (entry.name.endsWith("/") && entry.uncompressedSize === 0) continue

			const part = this.get(entry.name)

			if (part !== undefined) yield part
		}
	}

	read(name: string, slicing?: Slicing): Promise<Uint8Array> {
		return this.opc.read(name, slicing)
	}

	async relationships(source: string): Promise<Relationships> {
		const relationships = await this.opc.relationships(source)

		this.relationshipCache.set(partKey(source), relationships)

		return relationships
	}

	// A part's relationships once relationships() read them.
	loadedRelationships(source: string): Relationships | undefined {
		return this.relationshipCache.get(partKey(source))
	}
}
