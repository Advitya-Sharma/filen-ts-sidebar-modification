import { CfbError, CfbFile, isCompoundFile } from "../../core/cfb/index"
import { isEncryptedContainer, readEncryptedContainer, type EncryptedContainer } from "../../core/encryption/index"
import type { ByteSource } from "../../core/opc/index"
import type { Slicing } from "../../core/slicing"
import { classifyDataSpaces } from "./dataSpaces"
import { OfficeCryptoError, type OfficeCryptoErrorCode } from "./errors"
import {
	AGILE_DEFAULT_SPIN_COUNT,
	DEFAULT_OFFICE_PASSWORD,
	probeEncryptionInfo,
	releaseOfficeKey,
	streamSizeOf,
	unlockOfficePackage,
	type AbortSignalLike,
	type OfficeKey
} from "./webCrypto"

// Password-protected OOXML packages ([MS-OFFCRYPTO] §2.3.4), whatever the format inside (docx, xlsx, pptx): what
// the file is before anything is shown, and the unlock (the user's password; a workbook tries Excel's default
// first). The decrypted package is an ordinary zip in memory: the format's viewer and editor open it like any
// other. A save seals it again through seal.ts (packageSealer), which nothing here reaches.
//
// Runs where the file is opened (web: the docx worker, the spreadsheet worker): the key lives there. The password
// spin is cut into turns of the caller's slicing (8 ms by default) and stops at the caller's signal.

// Why an encrypted package is not opened at all (no prompt):
//   rightsManaged  IRM / sensitivity label: a DRM transform in its DataSpaces
//   unsupported    an encryption Filen does not implement (Extensible, RC4 or another non-AES cipher, CFB
//                  chaining, a certificate-only key)
//   damaged        a malformed encrypted container, or (once a password verified) an integrity check that failed
export type EncryptionRefusal = "rightsManaged" | "unsupported" | "damaged"

export interface PackageUnlockOptions {
	readonly signal?: AbortSignalLike
	// The host's turns; by default a macrotask turn every 8 ms.
	readonly slicing?: Slicing
}

export interface OpenEncryptedPackageOptions extends PackageUnlockOptions {
	// Try Excel's default password ("VelvetSweatshop") before asking: workbooks only, and only when the file's
	// spin is at most Excel's 100,000 (the try costs a full spin before the prompt). Word has no default password.
	readonly tryDefault?: boolean
}

export type EncryptedPackageOpen =
	// Not an encrypted package: open the file as it is.
	| { readonly kind: "plain" }
	| { readonly kind: "refused"; readonly reason: EncryptionRefusal }
	// The default password opened it (tryDefault): nothing to ask.
	| { readonly kind: "unlocked"; readonly document: DecryptedPackage }
	// Ask for the password; unlock may be called again after a wrong one.
	| { readonly kind: "needsPassword"; readonly locked: LockedPackage }
	| { readonly kind: "aborted" }

export interface LockedPackage {
	readonly scheme: "standard" | "agile"
	// Hash iterations an unlock costs (100,000 for Word's Agile default): what an "Unlocking…" state waits on.
	readonly spinCount: number
	unlock(password: string, options?: PackageUnlockOptions): Promise<PackageUnlockResult>
}

export type PackageUnlockResult =
	| { readonly ok: true; readonly document: DecryptedPackage }
	| {
			readonly ok: false
			readonly reason: "wrongPassword" | "aborted" | EncryptionRefusal
	  }

export interface DecryptedPackage {
	// The decrypted package (a zip), in memory only: open the viewer or the editor on bytesSource(bytes).
	readonly bytes: Uint8Array
	readonly scheme: "standard" | "agile"
	// verified: the Agile HMAC matched; unchecked: Standard has none, or the Agile source omitted it (the package
	// may have been changed after it was encrypted: say so before it is edited and sealed again).
	readonly integrity: "verified" | "unchecked"
	// Opened with the default password, without asking.
	readonly defaultPassword: boolean
	// What a save seals with (seal.ts packageSealer, or exportOfficeKey to seal on another thread).
	readonly key: OfficeKey
	// Drops the key (call once the viewer and editor on these bytes closed).
	release(): void
}

function refusal(code: OfficeCryptoErrorCode): EncryptionRefusal | "wrongPassword" | "aborted" {
	switch (code) {
		case "wrongPassword":
		case "aborted":
		case "rightsManaged":
		case "unsupported":
			return code
		case "integrity":
		case "corrupt":
			return "damaged"
	}
}

type Classified =
	| { readonly kind: "plain" }
	| { readonly kind: "refused"; readonly reason: EncryptionRefusal }
	| {
			readonly kind: "encrypted"
			readonly encryptionInfo: Uint8Array
			readonly encryptedPackage: ByteSource
			readonly scheme: "standard" | "agile"
			readonly spinCount: number
	  }

// What the compound file holds, from its structure alone (no password, no crypto): rights management and
// unsupported schemes are known before any prompt. Of the package only StreamSize is read.
async function classify(source: ByteSource): Promise<Classified> {
	if (source.size < 8 || !isCompoundFile(await source.read(0, 8))) return { kind: "plain" }

	let cfb: CfbFile
	let container: EncryptedContainer

	try {
		cfb = await CfbFile.open(source)
	} catch (error) {
		// Not a compound file this reader opens: the package reader reports it as such.
		if (error instanceof CfbError) return { kind: "plain" }

		throw error
	}

	if (!isEncryptedContainer(cfb)) return { kind: "plain" }

	try {
		container = await readEncryptedContainer(cfb)
	} catch (error) {
		if (error instanceof CfbError) return { kind: "refused", reason: "damaged" }

		throw error
	}

	let verdict: ReturnType<typeof classifyDataSpaces>

	try {
		verdict = classifyDataSpaces(container.dataSpaces)
	} catch (error) {
		if (error instanceof OfficeCryptoError) return { kind: "refused", reason: "damaged" }

		throw error
	}

	if (verdict === "rightsManaged" || verdict === "unsupported") return { kind: "refused", reason: verdict }
	if (container.encryptionInfo === undefined) return { kind: "refused", reason: "damaged" }

	const probe = probeEncryptionInfo(container.encryptionInfo)

	if (!probe.ok) {
		const reason = refusal(probe.reason)

		return {
			kind: "refused",
			reason: reason === "wrongPassword" || reason === "aborted" ? "damaged" : reason
		}
	}

	// The scheme decides first: one Filen does not implement is unsupported however its package looks.
	try {
		const stream = container.encryptedPackage

		streamSizeOf(stream.size < 8 ? new Uint8Array(0) : await stream.read(0, 8), stream.size)
	} catch (error) {
		if (error instanceof CfbError || error instanceof OfficeCryptoError) return { kind: "refused", reason: "damaged" }

		throw error
	}

	return {
		kind: "encrypted",
		encryptionInfo: container.encryptionInfo,
		encryptedPackage: container.encryptedPackage,
		scheme: probe.scheme,
		spinCount: probe.spinCount
	}
}

// How a format engine reports a file it cannot open because it is encrypted (no unlock attempted).
export async function encryptedPackageVerdict(source: ByteSource): Promise<"needsPassword" | EncryptionRefusal | undefined> {
	const classified = await classify(source)

	return classified.kind === "plain" ? undefined : classified.kind === "refused" ? classified.reason : "needsPassword"
}

type UnlockAs = (password: string, defaultPassword: boolean, options: PackageUnlockOptions) => Promise<PackageUnlockResult>

// The locked package. The whole EncryptedPackage stream is read at the first unlock and kept for retries until
// one succeeds.
function lockedOf(classified: Extract<Classified, { kind: "encrypted" }>): { readonly locked: LockedPackage; readonly unlockAs: UnlockAs } {
	let stream: Promise<Uint8Array> | undefined

	const unlockAs: UnlockAs = async (password, defaultPassword, options) => {
		let encryptedPackage: Uint8Array

		stream ??= classified.encryptedPackage.read(0, classified.encryptedPackage.size)

		try {
			encryptedPackage = await stream
		} catch (error) {
			stream = undefined

			if (error instanceof CfbError) return { ok: false, reason: "damaged" }

			throw error
		}

		const result = await unlockOfficePackage(classified.encryptionInfo, encryptedPackage, password, {
			...(options.signal === undefined ? {} : { signal: options.signal }),
			...(options.slicing === undefined ? {} : { slicing: options.slicing })
		})

		if (!result.ok) return { ok: false, reason: refusal(result.reason) }

		stream = undefined

		const key = result.key

		return {
			ok: true,
			document: {
				bytes: result.plaintext,
				scheme: key.scheme,
				integrity: result.integrity,
				defaultPassword,
				key,
				release: () => releaseOfficeKey(key)
			}
		}
	}

	return {
		locked: {
			scheme: classified.scheme,
			spinCount: classified.spinCount,
			unlock: (password, options = {}) => unlockAs(password, false, options)
		},
		unlockAs
	}
}

// Opens a file that may be password-protected: plain files say so at once; rights-managed and unsupported ones
// are refused from their structure; an encrypted one is left locked for the user's password (tried with Excel's
// default first when `tryDefault`). Never throws for the file's content (a read error of the source still rejects).
export async function openEncryptedPackage(source: ByteSource, options: OpenEncryptedPackageOptions = {}): Promise<EncryptedPackageOpen> {
	const classified = await classify(source)

	if (classified.kind !== "encrypted") return classified

	const { locked, unlockAs } = lockedOf(classified)

	if (options.tryDefault !== true || classified.spinCount > AGILE_DEFAULT_SPIN_COUNT) return { kind: "needsPassword", locked }

	const tried = await unlockAs(DEFAULT_OFFICE_PASSWORD, true, options)

	if (tried.ok) return { kind: "unlocked", document: tried.document }
	if (tried.reason === "wrongPassword") return { kind: "needsPassword", locked }
	if (tried.reason === "aborted") return { kind: "aborted" }

	return { kind: "refused", reason: tried.reason }
}
