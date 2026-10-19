// Why an encrypted package could not be opened or sealed. Messages never carry a password, key or salt.
//   wrongPassword  the password verifier did not match
//   unsupported    a scheme or parameter this backend does not implement (Extensible, RC4, non-AES, CFB
//                  chaining, a certificate-only key, a key size the platform's AES lacks), or no SubtleCrypto
//   rightsManaged  IRM / DRM: a data space with a transform other than password encryption
//   integrity      the Agile HMAC over the encrypted package did not match: the file was changed after it was sealed
//   corrupt        a malformed or out-of-bounds structure
//   aborted        the caller's signal fired
export type OfficeCryptoErrorCode = "wrongPassword" | "unsupported" | "rightsManaged" | "integrity" | "corrupt" | "aborted"

export class OfficeCryptoError extends Error {
	readonly code: OfficeCryptoErrorCode

	constructor(code: OfficeCryptoErrorCode, message: string) {
		super(message)

		this.name = "OfficeCryptoError"
		this.code = code
	}
}

export function corrupt(message: string): OfficeCryptoError {
	return new OfficeCryptoError("corrupt", message)
}

export function unsupported(message: string): OfficeCryptoError {
	return new OfficeCryptoError("unsupported", message)
}
