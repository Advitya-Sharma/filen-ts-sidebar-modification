import type { ByteSource } from "../../core/opc/index"
import type { Slicing } from "../../core/slicing"
import type {
	SheetAbortSignal,
	SheetEncryption,
	SheetEncryptionRefusal,
	SheetSealer,
	SheetTakenPackage,
	SheetUnlockedPackage
} from "../../core/sheet/protocol"
import { openEncryptedPackage, type DecryptedPackage, type LockedPackage } from "../encryption/sealedPackage"
import { PackageSealError, packageSealer, type PackageSealer } from "../encryption/seal"
import { exportOfficeKey, importOfficeKey, type OfficeKeyTransfer } from "../encryption/webCrypto"

// Password-protected workbooks ([MS-OFFCRYPTO] §2.3.4) where the session runs: classified and unlocked there
// (Excel's default password tried first, so a file it opens never asks), the key kept there, and every save
// sealed under it. A session in another realm than the unlock takes a copy of the package and gets the key only
// at its first seal.

export type SheetProtectedOpen =
	| { readonly kind: "plain" }
	| { readonly kind: "refused"; readonly reason: SheetEncryptionRefusal }
	| { readonly kind: "unlocked"; readonly document: DecryptedPackage }
	| { readonly kind: "needsPassword"; readonly locked: LockedPackage }
	| { readonly kind: "aborted" }

export function openProtectedWorkbook(source: ByteSource, signal?: SheetAbortSignal, slicing?: Slicing): Promise<SheetProtectedOpen> {
	return openEncryptedPackage(source, {
		tryDefault: true,
		...(signal !== undefined && { signal }),
		...(slicing !== undefined && { slicing })
	})
}

// What a session opened on a package unlocked in its own realm seals with; release drops the key and the package.
export function sheetEncryption(document: DecryptedPackage): SheetEncryption {
	const held = packageSealer(document.key)

	return {
		sealer: held.sealer,
		integrity: document.integrity,
		release() {
			held.release()
			document.release()
		}
	}
}

// A package unlocked here, for a session in another realm: `take` copies it, `key` moves the key at the session's
// first seal (a Standard source's upgrade spin runs here then).
export function sheetUnlockedPackage(document: DecryptedPackage): SheetUnlockedPackage<OfficeKeyTransfer> {
	let held: Uint8Array | null = document.bytes
	let released = false

	return {
		scheme: document.scheme,
		integrity: document.integrity,
		size: document.bytes.length,
		take(options = {}) {
			if (options.signal?.aborted === true) return Promise.reject(new Error("Aborted"))
			if (released || held === null) return Promise.reject(new Error("The workbook is no longer unlocked"))

			return Promise.resolve({
				bytes: held.slice(),
				key: (keyOptions: { readonly signal?: SheetAbortSignal } = {}) => {
					if (released) return Promise.reject(new Error("The workbook is no longer unlocked"))

					return exportOfficeKey(document.key, keyOptions.signal === undefined ? {} : { signal: keyOptions.signal })
				}
			} satisfies SheetTakenPackage<OfficeKeyTransfer>)
		},
		dropPackage() {
			held = null
		},
		release() {
			if (released) return

			released = true
			held = null
			document.release()
		}
	}
}

// The encryption of a session on a taken package: the key is asked for at the first seal, imported once, and
// released with the session.
export function takenSheetEncryption(taken: SheetTakenPackage<OfficeKeyTransfer>, integrity: "verified" | "unchecked"): SheetEncryption {
	let released = false
	let imported: Promise<{ readonly sealer: PackageSealer; release(): void }> | undefined

	const sealerOf = (signal: SheetAbortSignal | undefined): Promise<{ readonly sealer: PackageSealer; release(): void }> => {
		imported ??= taken.key(signal === undefined ? {} : { signal }).then(
			key => {
				const sealer = packageSealer(importOfficeKey(key))

				// Released while the key was on its way: it goes at once.
				if (released) sealer.release()

				return sealer
			},
			(error: unknown) => {
				// Asked again by the next seal.
				imported = undefined

				throw error
			}
		)

		return imported
	}
	const sealer: SheetSealer = {
		get released() {
			return released
		},
		async seal(zip, options = {}) {
			if (released) throw new PackageSealError("released", "The package's key was released")

			let own: { readonly sealer: PackageSealer }

			try {
				own = await sealerOf(options.signal)
			} catch (error) {
				if (released) throw new PackageSealError("released", "The package's key was released")
				if (options.signal?.aborted === true) throw new PackageSealError("aborted", "The seal was aborted")

				throw new PackageSealError("check", error instanceof Error ? error.message : String(error))
			}

			if (released) throw new PackageSealError("released", "The package's key was released")

			return own.sealer.seal(zip, {
				...(options.signal !== undefined && { signal: options.signal }),
				...(options.slicing !== undefined && { slicing: options.slicing })
			})
		}
	}

	return {
		sealer,
		integrity,
		release() {
			if (released) return

			released = true
			void imported?.then(
				own => own.release(),
				() => undefined
			)
		}
	}
}
