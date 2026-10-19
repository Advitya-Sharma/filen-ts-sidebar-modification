import { CfbError, CfbFile, cfbBytes, isCompoundFile, type CfbEntry, type CfbWriteEntry } from "../cfb/index"
import type { ByteSource } from "../opc/index"
import { DATA_SPACES, passwordDataSpaces } from "./dataSpaces"

// An encrypted OOXML package as a compound file ([MS-OFFCRYPTO] §2.3.4.4 and §2.1): the EncryptionInfo and
// EncryptedPackage streams at the root and the \x06DataSpaces storage saying which transform the package went
// through. Format glue only: no crypto here.

export const ENCRYPTION_INFO = "EncryptionInfo"
export const ENCRYPTED_PACKAGE = "EncryptedPackage"

// Larger than any EncryptionInfo or DataSpaces stream Office writes (a few KiB); a hostile size is refused
// before its bytes are read.
const MAX_INFO_BYTES = 1024 * 1024
const MAX_DATA_SPACE_BYTES = 64 * 1024
const MAX_DATA_SPACE_STREAMS = 64

// The DataSpaces streams, read ahead so a classifier can look them up synchronously.
export interface DataSpacesStreams {
	// \x06DataSpaces\DataSpaceMap; undefined without the storage or the stream.
	readonly dataSpaceMap: Uint8Array | undefined
	// \x06DataSpaces\DataSpaceInfo\<name>
	dataSpaceDefinition(name: string): Uint8Array | undefined
	// \x06DataSpaces\TransformInfo\<name>\x06Primary
	transformPrimary(name: string): Uint8Array | undefined
}

export interface EncryptedContainer {
	// Undefined when the file has none (an IRM package is described by its DataSpaces alone).
	readonly encryptionInfo: Uint8Array | undefined
	// The whole stream: StreamSize, then the ciphertext. Read on demand (a classifier needs its first 8 bytes); a
	// read inside one run of the file is a view of its bytes.
	readonly encryptedPackage: ByteSource
	readonly dataSpaces: DataSpacesStreams
}

function stream(entry: CfbEntry | undefined): CfbEntry | undefined {
	return entry?.kind === "stream" ? entry : undefined
}

// An EncryptedPackage stream at the root, with EncryptionInfo or a DataSpaces storage to say how it was
// transformed. Checked before any legacy stream name: Word 2007 and POI also leave a WordDocument stream (and an
// ObjectPool storage) beside an encrypted package.
export function isEncryptedContainer(cfb: CfbFile): boolean {
	return (
		stream(cfb.entry([ENCRYPTED_PACKAGE])) !== undefined &&
		(stream(cfb.entry([ENCRYPTION_INFO])) !== undefined || cfb.entry([DATA_SPACES])?.kind === "storage")
	)
}

async function readBounded(cfb: CfbFile, entry: CfbEntry, max: number): Promise<Uint8Array> {
	if (entry.size > max) throw new CfbError("tooLarge", `${entry.name} is larger than its limit`)

	return cfb.read(entry)
}

// Reads the streams of an encrypted container but the package itself (isEncryptedContainer first). Throws
// CfbError for a malformed or oversized structure.
export async function readEncryptedContainer(cfb: CfbFile): Promise<EncryptedContainer> {
	const packageEntry = stream(cfb.entry([ENCRYPTED_PACKAGE]))

	if (packageEntry === undefined) throw new CfbError("notFound", "No EncryptedPackage stream")

	const infoEntry = stream(cfb.entry([ENCRYPTION_INFO]))
	const encryptionInfo = infoEntry === undefined ? undefined : await readBounded(cfb, infoEntry, MAX_INFO_BYTES)
	const encryptedPackage = cfb.stream(packageEntry)
	const mapEntry = stream(cfb.entry([DATA_SPACES, "DataSpaceMap"]))
	const dataSpaceMap = mapEntry === undefined ? undefined : await readBounded(cfb, mapEntry, MAX_DATA_SPACE_BYTES)
	const definitions = new Map<string, Uint8Array>()
	const primaries = new Map<string, Uint8Array>()
	let budget = MAX_DATA_SPACE_STREAMS

	// Keys upper-cased: §2.6.4 compares names that way, and so does a lookup by the names the map holds.
	for (const child of cfb.entry([DATA_SPACES, "DataSpaceInfo"])?.children ?? []) {
		if (child.kind !== "stream") continue
		if (--budget < 0) throw new CfbError("tooManyEntries", "Too many DataSpaces streams")

		definitions.set(child.name.toUpperCase(), await readBounded(cfb, child, MAX_DATA_SPACE_BYTES))
	}

	for (const transform of cfb.entry([DATA_SPACES, "TransformInfo"])?.children ?? []) {
		const primary = stream(transform.children.find(child => child.name === "\u0006Primary"))

		if (primary === undefined) continue
		if (--budget < 0) throw new CfbError("tooManyEntries", "Too many DataSpaces streams")

		primaries.set(transform.name.toUpperCase(), await readBounded(cfb, primary, MAX_DATA_SPACE_BYTES))
	}

	return {
		encryptionInfo,
		encryptedPackage,
		dataSpaces: {
			dataSpaceMap,
			dataSpaceDefinition: name => definitions.get(name.toUpperCase()),
			transformPrimary: name => primaries.get(name.toUpperCase())
		}
	}
}

// The compound file of a sealed package: the password DataSpaces Word writes, EncryptionInfo and
// EncryptedPackage, nothing else (a source's other root streams could only disagree with the new package).
// Version 3 (512-byte sectors): LibreOffice 7.5 reads no other.
export function sealedContainer(encryptionInfo: Uint8Array, encryptedPackage: Uint8Array): Promise<Uint8Array> {
	const entries: CfbWriteEntry[] = [
		...passwordDataSpaces(),
		{ path: [ENCRYPTION_INFO], bytes: encryptionInfo },
		{ path: [ENCRYPTED_PACKAGE], bytes: encryptedPackage }
	]

	return cfbBytes(entries, { version: 3 })
}

// Whether bytes are fit to store for an encrypted package: a compound file, never a zip.
export function isSealedOfficeFile(bytes: Uint8Array): boolean {
	return isCompoundFile(bytes)
}
