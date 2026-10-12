import type { Slicing } from "../slicing"
import { parseXml, parseXmlSliced, type XmlDocument } from "../xml/index"
import { CONTENT_TYPES_PART, parseContentTypes, type ContentTypes } from "./contentTypes"
import { OpcError } from "./errors"
import type { ByteSource } from "./io"
import { joinBytes } from "./pieces"
import { isValidPartName, partKey, relationshipsPartName } from "./partNames"
import { Relationships, parseRelationships } from "./relationships"
import { ZipArchive, type ZipEntry, type ZipOpenOptions } from "./zipReader"
import { planZip, type PartChange, type ZipChanges, type ZipPlan } from "./zipWriter"

// An OPC package (docx, xlsx, pptx) over a lazily read zip. Parts are inflated on request and not kept:
// whoever parses a part owns the cache for it. Content types and relationships, small and read often,
// are parsed once.
export class OpcPackage {
	readonly zip: ZipArchive
	private readonly parts = new Map<string, ZipEntry>()
	private contentTypesPromise: Promise<ContentTypes> | null = null
	private readonly relationshipsPromises = new Map<string, Promise<Relationships>>()

	private constructor(zip: ZipArchive) {
		this.zip = zip

		for (const entry of zip.entries) {
			// Folder entries some writers add are not parts.
			if (entry.name.endsWith("/") && entry.uncompressedSize === 0) continue

			if (!isValidPartName(entry.name)) throw new OpcError("invalidPartName", `Invalid part name ${entry.name}`)

			const key = partKey(entry.name)

			if (this.parts.has(key)) throw new OpcError("duplicatePart", `Part ${entry.name} appears twice`)

			this.parts.set(key, entry)
		}

		if (!this.parts.has(partKey(CONTENT_TYPES_PART))) {
			throw new OpcError("missingPart", "The package has no [Content_Types].xml")
		}
	}

	static async open(source: ByteSource, options: ZipOpenOptions = {}): Promise<OpcPackage> {
		return new OpcPackage(await ZipArchive.open(source, options))
	}

	// The zip entry of a part, matched as OPC compares names (case-insensitive, percent-decoded).
	entry(partName: string): ZipEntry | undefined {
		return this.parts.get(partKey(partName))
	}

	has(partName: string): boolean {
		return this.parts.has(partKey(partName))
	}

	// With `slicing`, the host is given a turn whenever the slice is used up (inflating and parsing).
	async read(partName: string, slicing?: Slicing): Promise<Uint8Array> {
		const entry = this.entry(partName)

		if (entry === undefined) throw new OpcError("missingPart", `No part ${partName}`)

		return slicing === undefined ? this.zip.read(entry) : this.zip.readSliced(entry, slicing)
	}

	async xml(partName: string, slicing?: Slicing): Promise<XmlDocument> {
		const bytes = await this.read(partName, slicing)

		return slicing === undefined ? parseXml(bytes) : parseXmlSliced(bytes, slicing)
	}

	contentTypes(): Promise<ContentTypes> {
		this.contentTypesPromise ??= this.read(CONTENT_TYPES_PART).then(parseContentTypes)

		return this.contentTypesPromise
	}

	// Relationships whose source is `sourcePart` ("" for the package); none when it has no .rels part.
	relationships(sourcePart: string): Promise<Relationships> {
		const key = partKey(sourcePart)
		let promise = this.relationshipsPromises.get(key)

		if (promise === undefined) {
			const relsPart = relationshipsPartName(this.entry(sourcePart)?.name ?? sourcePart)

			promise = this.has(relsPart)
				? this.read(relsPart).then(bytes => parseRelationships(bytes, sourcePart))
				: Promise.resolve(new Relationships([]))
			this.relationshipsPromises.set(key, promise)
		}

		return promise
	}

	// A save plan: untouched entries copied raw, replaced parts rewritten in place (names matched as OPC
	// does), new parts appended. New names must not collide with existing parts in any letter case, and
	// each must have a content type, from [Content_Types].xml as replaced in the same changes if it is.
	async plan(changes: ZipChanges): Promise<ZipPlan> {
		const toEntryName = (change: PartChange): PartChange => {
			const entry = this.entry(change.name)

			if (entry === undefined) throw new OpcError("missingPart", `No part ${change.name} to replace`)

			return { ...change, name: entry.name }
		}

		const added = new Set<string>()

		for (const change of changes.add ?? []) {
			const key = partKey(change.name)

			if (!isValidPartName(change.name)) throw new OpcError("invalidPartName", `Invalid part name ${change.name}`)
			if (this.parts.has(key) || added.has(key)) throw new OpcError("duplicatePart", `Part ${change.name} already exists`)

			added.add(key)
		}

		const contentTypesKey = partKey(CONTENT_TYPES_PART)
		const newTypes = [...(changes.replace ?? []), ...(changes.add ?? [])].find(change => partKey(change.name) === contentTypesKey)
		const types = newTypes !== undefined ? parseContentTypes(await joinBytes(newTypes.bytes)) : await this.contentTypes()

		for (const change of changes.add ?? []) {
			if (types.typeOf(change.name) === undefined) {
				throw new OpcError("missingContentType", `Part ${change.name} has no content type`)
			}
		}

		return planZip(this.zip, {
			replace: (changes.replace ?? []).map(toEntryName),
			remove: (changes.remove ?? []).map(name => this.entry(name)?.name ?? name),
			add: changes.add ?? []
		})
	}
}
