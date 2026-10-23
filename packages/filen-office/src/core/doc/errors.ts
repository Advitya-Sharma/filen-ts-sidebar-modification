// Why a Word binary file cannot be read. Every refusal carries one of these codes, never a bare exception.
export type DocErrorCode =
	// Not a compound file, or no WordDocument stream in it.
	| "notWordBinary"
	// An encrypted OOXML package in a compound file ([MS-OFFCRYPTO] §2.3.4.4), not a Word binary document.
	| "encryptedPackage"
	// FibBase.wIdent is not 0xA5EC.
	| "badSignature"
	// Word 6/95 or earlier: the Fib layout of [MS-DOC] starts at nFib 0x00C1.
	| "legacyVersion"
	// FibBase.fEncrypted ([MS-DOC] §2.5.1) and no password given, or a scheme this reader does not decrypt
	// (`encryption` says which).
	| "encrypted"
	// The password given does not match the encrypted file's verifier.
	| "wrongPassword"
	// A structure points outside its stream or contradicts itself where the text could not be recovered.
	| "corrupt"
	// A count or size beyond what this reader allocates for.
	| "tooLarge"

// [MS-DOC] §2.2.6: XOR obfuscation (§2.3.7 of [MS-OFFCRYPTO]), RC4 (§2.3.6) or CryptoAPI RC4 (§2.3.5).
export type DocEncryptionScheme = "xor" | "rc4" | "cryptoApiRc4" | "unknown"

export class DocError extends Error {
	readonly code: DocErrorCode
	readonly encryption: DocEncryptionScheme | undefined
	// The file's nFib, set for legacyVersion.
	readonly nFib: number | undefined

	constructor(code: DocErrorCode, message: string, details: { readonly encryption?: DocEncryptionScheme; readonly nFib?: number } = {}) {
		super(message)

		this.name = "DocError"
		this.code = code
		this.encryption = details.encryption
		this.nFib = details.nFib
	}
}

export function corrupt(message: string): DocError {
	return new DocError("corrupt", message)
}
