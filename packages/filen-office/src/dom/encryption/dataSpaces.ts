import { corrupt } from "./errors"

// The \x06DataSpaces storage ([MS-OFFCRYPTO] §2.1): which transforms stand between the EncryptedPackage stream
// and the package. Password encryption is the one strong-encryption transform; a DRM transform is IRM
// (§2.2.10, "rights management"); anything else is a transform this backend cannot undo.

export const STRONG_ENCRYPTION_TRANSFORM_ID = "{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}"
export const DRM_TRANSFORM_ID = "{C73DFACD-061F-43B0-8B64-0C620D2A8B50}"

// The streams of \x06DataSpaces the caller reads out of the compound file.
export interface DataSpacesSource {
	// \x06DataSpaces\DataSpaceMap; undefined when the storage or stream is missing.
	readonly dataSpaceMap: Uint8Array | undefined
	// \x06DataSpaces\DataSpaceInfo\<name>
	dataSpaceDefinition(name: string): Uint8Array | undefined
	// \x06DataSpaces\TransformInfo\<name>\x06Primary
	transformPrimary(name: string): Uint8Array | undefined
}

//   password       EncryptedPackage passes through exactly the strong-encryption transform
//   rightsManaged  a DRM transform or data space anywhere in the map
//   unsupported    any other transform chain
//   absent         no data space map: the EncryptionInfo stream alone decides
export type DataSpacesVerdict = "password" | "rightsManaged" | "unsupported" | "absent"

const MAX_ENTRIES = 256
const MAX_TRANSFORMS = 16

class Reader {
	private offset = 0
	private readonly bytes: Uint8Array
	private readonly view: DataView

	constructor(bytes: Uint8Array) {
		this.bytes = bytes
		this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
	}

	get position(): number {
		return this.offset
	}

	seek(offset: number): void {
		if (offset < this.offset || offset > this.bytes.length) throw corrupt("DataSpaces structure is out of bounds")

		this.offset = offset
	}

	uint32(): number {
		if (this.offset + 4 > this.bytes.length) throw corrupt("DataSpaces structure is truncated")

		const value = this.view.getUint32(this.offset, true)

		this.offset += 4

		return value
	}

	// UNICODE-LP-P4 (§2.1.2): byte length, UTF-16LE, padded to 4 bytes.
	string(): string {
		const length = this.uint32()

		if (length % 2 !== 0 || length > this.bytes.length - this.offset) throw corrupt("DataSpaces string is out of bounds")

		let out = ""

		for (let k = 0; k < length; k += 2)
			out += String.fromCharCode((this.bytes[this.offset + k] ?? 0) | ((this.bytes[this.offset + k + 1] ?? 0) << 8))

		this.offset = Math.min(this.bytes.length, this.offset + length + ((4 - (length % 4)) % 4))

		return out
	}
}

interface MapEntry {
	readonly components: readonly string[]
	readonly dataSpace: string
}

// DataSpaceMap (§2.1.6): HeaderLength 8, EntryCount, then DataSpaceMapEntry (§2.1.6.1) each.
function mapEntries(bytes: Uint8Array): MapEntry[] {
	const reader = new Reader(bytes)
	const headerLength = reader.uint32()
	const count = reader.uint32()

	if (headerLength !== 8) throw corrupt("DataSpaceMap header length is not 8")
	if (count > MAX_ENTRIES) throw corrupt("DataSpaceMap has too many entries")

	const entries: MapEntry[] = []

	for (let k = 0; k < count; k++) {
		const start = reader.position
		const length = reader.uint32()
		const references = reader.uint32()

		if (length < 12 || references > MAX_ENTRIES) throw corrupt("DataSpaceMap entry is out of range")

		const components: string[] = []

		for (let r = 0; r < references; r++) {
			reader.uint32()
			components.push(reader.string())
		}

		entries.push({ components, dataSpace: reader.string() })
		reader.seek(start + length)
	}

	return entries
}

// DataSpaceDefinition (§2.1.7): HeaderLength 8, TransformReferenceCount, transform names.
function transformNames(bytes: Uint8Array): string[] {
	const reader = new Reader(bytes)
	const headerLength = reader.uint32()
	const count = reader.uint32()

	if (headerLength !== 8) throw corrupt("DataSpaceDefinition header length is not 8")
	if (count > MAX_TRANSFORMS) throw corrupt("DataSpaceDefinition has too many transforms")

	const names: string[] = []

	for (let k = 0; k < count; k++) names.push(reader.string())

	return names
}

// TransformInfoHeader (§2.1.8): TransformLength, TransformType, TransformID, ...
function transformId(bytes: Uint8Array): string {
	const reader = new Reader(bytes)

	reader.uint32()
	reader.uint32()

	return reader.string().toUpperCase()
}

const isDrmName = (name: string): boolean => /drm/iu.test(name)

export function classifyDataSpaces(source: DataSpacesSource): DataSpacesVerdict {
	if (source.dataSpaceMap === undefined) return "absent"

	const entries = mapEntries(source.dataSpaceMap)
	let verdict: DataSpacesVerdict = "unsupported"

	for (const entry of entries) {
		const definition = source.dataSpaceDefinition(entry.dataSpace)
		const transforms = definition === undefined ? [] : transformNames(definition)
		const ids = transforms.map(name => {
			const primary = source.transformPrimary(name)

			return primary === undefined ? "" : transformId(primary)
		})

		// Any DRM content (\x09DRMContent, a DRMEncryptedDataSpace) makes the document rights-managed.
		if (isDrmName(entry.dataSpace) || ids.includes(DRM_TRANSFORM_ID) || transforms.some(isDrmName)) return "rightsManaged"

		if (entry.components.length === 1 && entry.components[0] === "EncryptedPackage") {
			verdict = ids.length === 1 && ids[0] === STRONG_ENCRYPTION_TRANSFORM_ID ? "password" : "unsupported"
		}
	}

	return verdict
}
