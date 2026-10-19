import { MAX_PROTECTION_SALT_BYTES, protectionHashInputs, type ProtectionHash } from "../core/docx/protection/hash"
import { constantTimeEqual, decodeBase64, utf16LE } from "./encryption/bytes"
import { OfficeCryptoError } from "./encryption/errors"
import { subtle } from "./encryption/subtle"
import { turn, turnsOf, type OfficeCryptoOptions, type Turns } from "./encryption/turns"

// Checks a password against a protection hash (core/docx/protection/hash.ts) on the platform's WebCrypto:
// "password to modify" lifts a document's write protection for the session. Not a security boundary (the
// hash only gates an editor, [MS-OFFCRYPTO] §4.1.4), but the password is still never logged or kept, and the
// hashes compare in constant time. The spin gives the host a turn every few milliseconds (turns.ts), so it
// may run on a UI thread; the caller's signal stops it.

export type ProtectionVerdict = "correct" | "wrong" | "unsupported"

const DIGEST_BYTES: Readonly<Record<ProtectionHash["algorithm"], number>> = { "SHA-1": 20, "SHA-256": 32, "SHA-384": 48, "SHA-512": 64 }
// Spin iterations between looks at the clock and the abort signal.
const SPIN_CHECK = 64

// Rejects with an OfficeCryptoError "aborted" when the signal fires; "unsupported" covers a hash or salt that
// is not base64 of the right size, and a platform without SubtleCrypto.
export async function verifyProtectionPassword(
	hash: ProtectionHash,
	password: string,
	options: OfficeCryptoOptions = {}
): Promise<ProtectionVerdict> {
	let expected: Uint8Array
	let salt: Uint8Array

	try {
		expected = decodeBase64(hash.hash)
		salt = decodeBase64(hash.salt)
	} catch (error) {
		if (error instanceof OfficeCryptoError) return "unsupported"

		throw error
	}

	if (expected.length !== DIGEST_BYTES[hash.algorithm] || salt.length > MAX_PROTECTION_SALT_BYTES) return "unsupported"

	const turns = turnsOf(options)

	try {
		for (const input of protectionHashInputs(hash, password)) {
			const encoded = utf16LE(input)
			const first = new Uint8Array(salt.length + encoded.length)

			first.set(salt)
			first.set(encoded, salt.length)
			encoded.fill(0)

			const found = await spin(hash, first, turns)
			const match = constantTimeEqual(found, expected)

			found.fill(0)

			if (match) return "correct"
		}
	} catch (error) {
		if (error instanceof OfficeCryptoError) {
			if (error.code === "aborted") throw error

			return "unsupported"
		}

		throw error
	}

	return "wrong"
}

// H0 = H(first), then Hn = H(Hn-1 + LE32(i)) for i < spinCount ([MS-OFFCRYPTO] §2.4.2.4). One buffer is
// reused: digest copies its input before it resolves.
async function spin(hash: ProtectionHash, first: Uint8Array, turns: Turns): Promise<Uint8Array> {
	const crypto = subtle()
	const size = DIGEST_BYTES[hash.algorithm]
	const buffer = new Uint8Array(size + 4)
	const view = new DataView(buffer.buffer)
	let h: Uint8Array

	try {
		h = new Uint8Array(await crypto.digest(hash.algorithm, first))
	} finally {
		first.fill(0)
	}

	try {
		for (let i = 0; i < hash.spinCount; i++) {
			buffer.set(h)
			view.setUint32(size, i, true)
			h.fill(0)
			h = new Uint8Array(await crypto.digest(hash.algorithm, buffer))

			if (i % SPIN_CHECK === SPIN_CHECK - 1) await turn(turns)
		}
	} catch (error) {
		h.fill(0)

		throw error
	} finally {
		buffer.fill(0)
	}

	return h
}
