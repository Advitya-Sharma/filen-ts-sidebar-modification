import { cfbBytes } from "../../core/cfb/index"
import { passwordDataSpaces } from "../../core/encryption/index"
import type { EncryptedStreams } from "./generate"

// Test support for apps (@filen/office/testing/encryption): encrypted packages built with Node's crypto,
// independently of the package's own WebCrypto backend, and the compound file Word writes around them. Node only.

export { agileInfo, agileStreams, decryptAgileStreams, standardStreams, type AgileOptions, type EncryptedStreams } from "./generate"

export interface EncryptedFileStreams {
	readonly encryptionInfo?: EncryptedStreams["encryptionInfo"]
	readonly encryptedPackage: EncryptedStreams["encryptedPackage"]
}

// A compound file holding `streams` with the password DataSpaces Word writes (or `dataSpaces`), plus any extra
// streams (an IRM licence, a stale summary).
export function encryptedOfficeFile(
	streams: EncryptedFileStreams,
	extra: readonly { readonly path: readonly string[]; readonly bytes: Uint8Array }[] = [],
	dataSpaces: readonly { readonly path: readonly string[]; readonly bytes: Uint8Array }[] = passwordDataSpaces()
): Promise<Uint8Array> {
	return cfbBytes([
		...dataSpaces,
		...(streams.encryptionInfo === undefined ? [] : [{ path: ["EncryptionInfo"], bytes: streams.encryptionInfo }]),
		{ path: ["EncryptedPackage"], bytes: streams.encryptedPackage },
		...extra
	])
}
