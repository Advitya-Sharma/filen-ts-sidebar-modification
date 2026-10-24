import type { DocDecryption } from "../../core/doc/decrypt"
import { DocError } from "../../core/doc/errors"
import { OfficeCryptoError } from "./errors"
import { digest } from "./subtle"

// The password to open a Word or Excel 97–2003 file with ([MS-OFFCRYPTO] §2.3.5–§2.3.7). CryptoAPI RC4 keys hash
// with SHA-1 on the platform's WebCrypto; without it the file cannot be decrypted here.
export function decryptionOf(password: string): DocDecryption {
	return {
		password,
		sha1: data =>
			digest("SHA-1", data).catch((error: unknown) => {
				if (error instanceof OfficeCryptoError)
					throw new DocError("encrypted", "SHA-1 is not available here", { encryption: "cryptoApiRc4" })

				throw error
			})
	}
}
