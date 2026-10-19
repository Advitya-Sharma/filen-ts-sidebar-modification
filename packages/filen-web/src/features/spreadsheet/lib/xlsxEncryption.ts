import { isCompoundFile } from "@filen/office/core/cfb"
import { bytesSource } from "@filen/office/core/opc"
import {
	isSealedOfficeFile,
	openEncryptedPackage,
	type DecryptedPackage,
	type EncryptionRefusal,
	type LockedPackage
} from "@filen/office/dom/encryption"
import { packageSealer, type PackageSealer } from "@filen/office/dom/encryption/seal"

// Password-protected workbooks ([MS-OFFCRYPTO] §2.3.4) through the office package's encrypted-package open: an
// encrypted .xlsx is the same compound file around a zip as an encrypted .docx. Runs in the spreadsheet worker, so the password spin never holds the page: the key, the plaintext and the seal
// stay there, and the page only gets views, a verdict, or sealed bytes.

// Why a protected workbook is not opened, without asking for a password: rights-managed (IRM), an encryption
// Filen does not implement, or a damaged container (or, once a password verified, a failed integrity check).
export type WorkbookRefusal = EncryptionRefusal

// A protected workbook waiting for its password: `unlock` may be tried again after a wrong one.
export type WorkbookLocked = LockedPackage

// What a decrypted workbook's saves seal with, and the release of its key.
export interface WorkbookProtection {
	readonly sealer: PackageSealer
	readonly release: () => void
}

export interface WorkbookUnlocked {
	readonly plaintext: Uint8Array
	readonly protection: WorkbookProtection
	// verified: the Agile HMAC matched; unchecked: the package carries none (Standard, or an Agile file without it).
	readonly integrity: "verified" | "unchecked"
}

export type EncryptedWorkbookOpen =
	| { readonly kind: "plain" }
	| { readonly kind: "refused"; readonly reason: WorkbookRefusal }
	| ({ readonly kind: "unlocked" } & WorkbookUnlocked)
	| { readonly kind: "needsPassword"; readonly locked: WorkbookLocked }
	| { readonly kind: "aborted" }

export type WorkbookUnlock =
	({ readonly ok: true } & WorkbookUnlocked) | { readonly ok: false; readonly reason: "wrongPassword" | "aborted" | WorkbookRefusal }

export interface AbortFlag {
	readonly aborted: boolean
}

// Only the bytes, the sealer of its key and the release are kept: nothing else of the unlock outlives it.
function workbookOf(document: DecryptedPackage): WorkbookUnlocked {
	const held = packageSealer(document.key)

	return {
		plaintext: document.bytes,
		integrity: document.integrity,
		protection: {
			sealer: held.sealer,
			release: () => {
				held.release()
				document.release()
			}
		}
	}
}

// What a file named or sniffed as a spreadsheet is: plain (a zip, a legacy .xls, or a compound file that is
// not an encrypted package), refused from its structure alone, opened with Excel's default password
// ("VelvetSweatshop", which Excel writes for a workbook encrypted without one), or locked until the user's.
export async function openEncryptedWorkbook(bytes: Uint8Array, signal: AbortFlag): Promise<EncryptedWorkbookOpen> {
	if (!isCompoundFile(bytes)) {
		return { kind: "plain" }
	}

	const opened = await openEncryptedPackage(bytesSource(bytes), { tryDefault: true, signal })

	return opened.kind === "unlocked" ? { kind: "unlocked", ...workbookOf(opened.document) } : opened
}

export async function unlockWorkbook(locked: WorkbookLocked, password: string, signal: AbortFlag): Promise<WorkbookUnlock> {
	const result = await locked.unlock(password, { signal })

	return result.ok ? { ok: true, ...workbookOf(result.document) } : result
}

// What leaves a protected workbook's save for upload is a compound file, never its zip.
export function assertSealedUpload(bytes: Uint8Array): void {
	if (!isSealedOfficeFile(bytes)) {
		throw new Error("spreadsheet: a protected workbook's save is not sealed")
	}
}

// The bytes a save of a protected workbook hands out: `zip` sealed as Agile under the session's key and read
// back (PackageSealer.seal's post-seal check), then held to the upload guard. The zip is zeroed once sealed or
// failed: nothing else holds it.
export async function sealWorkbook(protection: WorkbookProtection, zip: Uint8Array): Promise<Uint8Array> {
	try {
		const sealed = await protection.sealer.seal(zip)

		assertSealedUpload(sealed)

		return sealed
	} finally {
		zip.fill(0)
	}
}
