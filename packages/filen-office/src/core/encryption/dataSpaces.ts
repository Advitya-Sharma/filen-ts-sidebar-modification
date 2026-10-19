// The \x06DataSpaces storage of an encrypted package ([MS-OFFCRYPTO] §2.1), written as Office writes it for
// password encryption (§2.3.4.4 and the streams Word 2013+ stores): one data space mapping EncryptedPackage
// through the strong-encryption transform. Builders only; reading and classifying is the backend's.

export const DATA_SPACES = "\u0006DataSpaces"
export const STRONG_ENCRYPTION_DATA_SPACE = "StrongEncryptionDataSpace"
export const STRONG_ENCRYPTION_TRANSFORM = "StrongEncryptionTransform"
export const STRONG_ENCRYPTION_TRANSFORM_GUID = "{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}"
export const ENCRYPTION_TRANSFORM_NAME = "Microsoft.Container.EncryptionTransform"
const DATA_SPACES_FEATURE = "Microsoft.Container.DataSpaces"

class Writer {
	private readonly parts: number[] = []

	uint32(value: number): this {
		this.parts.push(value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff, (value >>> 24) & 0xff)

		return this
	}

	// Major and minor (§2.1.4).
	version(major: number, minor: number): this {
		this.parts.push(major & 0xff, major >>> 8, minor & 0xff, minor >>> 8)

		return this
	}

	// UNICODE-LP-P4 (§2.1.2): byte length, UTF-16LE code units, zero-padded to 4 bytes.
	unicode(text: string): this {
		this.uint32(text.length * 2)

		for (let k = 0; k < text.length; k++) {
			const code = text.charCodeAt(k)

			this.parts.push(code & 0xff, code >>> 8)
		}

		return this.pad()
	}

	get length(): number {
		return this.parts.length
	}

	bytes(): Uint8Array {
		return Uint8Array.from(this.parts)
	}

	private pad(): this {
		while (this.parts.length % 4 !== 0) this.parts.push(0)

		return this
	}
}

// DataSpaceVersionInfo (§2.1.5): reader, updater and writer version 1.0.
export function dataSpaceVersionStream(): Uint8Array {
	return new Writer().unicode(DATA_SPACES_FEATURE).version(1, 0).version(1, 0).version(1, 0).bytes()
}

export interface DataSpaceMapEntry {
	// Stream names from the root (ReferenceComponent type 0, a stream).
	readonly components: readonly string[]
	readonly dataSpace: string
}

// DataSpaceMap (§2.1.6): HeaderLength 8, EntryCount, then each DataSpaceMapEntry (§2.1.6.1) with its length.
export function dataSpaceMapStream(entries: readonly DataSpaceMapEntry[]): Uint8Array {
	const out = new Writer().uint32(8).uint32(entries.length)
	const bytes: Uint8Array[] = [out.bytes()]

	for (const entry of entries) {
		const body = new Writer().uint32(entry.components.length)

		for (const component of entry.components) body.uint32(0).unicode(component)

		body.unicode(entry.dataSpace)
		bytes.push(new Writer().uint32(body.length + 4).bytes(), body.bytes())
	}

	return join(bytes)
}

// DataSpaceDefinition (§2.1.7): HeaderLength 8, the transforms' names in order.
export function dataSpaceDefinitionStream(transforms: readonly string[]): Uint8Array {
	const out = new Writer().uint32(8).uint32(transforms.length)

	for (const name of transforms) out.unicode(name)

	return out.bytes()
}

// TransformInfoHeader (§2.1.8), TransformType 1: TransformLength counts the bytes up to the end of TransformID.
// For the strong-encryption transform an EncryptionTransformInfo (§2.1.9) follows; Word 2013+ writes it with
// an empty name, block size and cipher mode 0 (the EncryptionInfo stream holds the real parameters).
export function transformPrimaryStream(id: string, name: string, encryptionTransform: boolean): Uint8Array {
	const head = new Writer().uint32(1).unicode(id)
	const out = new Writer().uint32(head.length + 4)

	return join([
		out.bytes(),
		head.bytes(),
		new Writer().unicode(name).version(1, 0).version(1, 0).version(1, 0).bytes(),
		encryptionTransform ? new Writer().uint32(0).uint32(0).uint32(0).uint32(4).bytes() : new Uint8Array(0)
	])
}

// Path and bytes of every stream of the storage for password encryption.
export function passwordDataSpaces(): readonly { readonly path: readonly string[]; readonly bytes: Uint8Array }[] {
	return [
		{ path: [DATA_SPACES, "Version"], bytes: dataSpaceVersionStream() },
		{
			path: [DATA_SPACES, "DataSpaceMap"],
			bytes: dataSpaceMapStream([{ components: ["EncryptedPackage"], dataSpace: STRONG_ENCRYPTION_DATA_SPACE }])
		},
		{
			path: [DATA_SPACES, "DataSpaceInfo", STRONG_ENCRYPTION_DATA_SPACE],
			bytes: dataSpaceDefinitionStream([STRONG_ENCRYPTION_TRANSFORM])
		},
		{
			path: [DATA_SPACES, "TransformInfo", STRONG_ENCRYPTION_TRANSFORM, "\u0006Primary"],
			bytes: transformPrimaryStream(STRONG_ENCRYPTION_TRANSFORM_GUID, ENCRYPTION_TRANSFORM_NAME, true)
		}
	]
}

function join(parts: readonly Uint8Array[]): Uint8Array {
	const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
	let at = 0

	for (const part of parts) {
		out.set(part, at)
		at += part.length
	}

	return out
}
