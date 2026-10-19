import type { ByteSource } from "../../core/opc/index"
import type { OfficeKeyTransfer } from "../encryption/webCrypto"
import type { DocxHost } from "./host"

// Password-protected Word documents ([MS-OFFCRYPTO] §2.3.4) as a host opens them: where the document engine
// runs (web: the docx worker), so the password spin and the decrypted package stay there. The viewer opens the
// decrypted package through the unlocked document's own host; only an editor takes a copy of it, and the key,
// to its thread to seal saves (docxSealer in editing.ts). No crypto runs on the caller's thread.

// Why an encrypted document is not opened at all (no prompt):
//   rightsManaged  IRM / sensitivity label
//   unsupported    an encryption Filen does not implement (Extensible, RC4 or another non-AES cipher, CFB
//                  chaining, a certificate-only key)
//   damaged        a malformed encrypted container, or (once a password verified) an integrity check that failed
export type DocxEncryptionRefusal = "rightsManaged" | "unsupported" | "damaged"

// An AbortSignal, or anything with `aborted` (polled at the next answer when it has no events).
export interface DocxAbortSignal {
	readonly aborted: boolean
	addEventListener?(type: "abort", listener: () => void): void
	removeEventListener?(type: "abort", listener: () => void): void
}

export interface DocxProtectionOptions {
	readonly signal?: DocxAbortSignal
}

export type DocxProtectedOpen =
	// Not an encrypted package: open the file as it is (a legacy compound file is reported by the engine).
	| { readonly kind: "plain" }
	| { readonly kind: "refused"; readonly reason: DocxEncryptionRefusal }
	| { readonly kind: "needsPassword"; readonly locked: DocxLockedDocument }
	| { readonly kind: "aborted" }

export interface DocxLockedDocument {
	readonly scheme: "standard" | "agile"
	// Hash iterations an unlock costs (100,000 for Word's Agile default): what an "Unlocking…" state waits on.
	readonly spinCount: number
	// May be called again after a wrong password or an abort. Word has no default password: none is tried.
	unlock(password: string, options?: DocxProtectionOptions): Promise<DocxDocumentUnlock>
	// Drops what is held for the locked file (the prompt went without an unlock).
	dispose(): void
}

export type DocxDocumentUnlock =
	| { readonly ok: true; readonly document: DocxUnlockedDocument }
	| { readonly ok: false; readonly reason: "wrongPassword" | "aborted" | DocxEncryptionRefusal }

export interface DocxUnlockedDocument {
	readonly scheme: "standard" | "agile"
	// verified: the Agile HMAC matched; unchecked: Standard has none, or the Agile source omitted it (the document
	// may have been changed after it was encrypted: the app says so).
	readonly integrity: "verified" | "unchecked"
	// The viewer's source: { host, bytes, options }. The host opens the decrypted package where it is held; `bytes`
	// gives its size only (reads are refused on this thread).
	readonly host: DocxHost
	readonly bytes: ByteSource
	// The editor's copy of the decrypted package, at once; the key comes from its key() when the editor first seals
	// (a session that never saves never derives or moves it). Each call makes a new copy.
	take(options?: DocxProtectionOptions): Promise<DocxTakenPackage>
	// The editor has opened on its copy: the held copy goes (the key stays for key()); the host can open the package
	// no more.
	dropPackage(): void
	// Drops the decrypted package and the key where they are held; call once every session on them closed.
	release(): void
}

export interface DocxTakenPackage {
	readonly bytes: Uint8Array
	// The key moved to this thread, derived first where the document is held (a Standard source's upgrade spin,
	// sliced there): docxSealer asks at its first seal.
	key(options?: DocxProtectionOptions): Promise<OfficeKeyTransfer>
}

// Opens files that may be password-protected (docxWorkerPool(...).protection, or inThreadDocxProtection).
export interface DocxProtection {
	open(source: ByteSource, options?: DocxProtectionOptions): Promise<DocxProtectedOpen>
}

// A source whose bytes are held elsewhere: its size, nothing to read here.
export function heldSource(size: number): ByteSource {
	return {
		size,
		read: () => Promise.reject(new Error("The decrypted package is held by the document engine"))
	}
}
