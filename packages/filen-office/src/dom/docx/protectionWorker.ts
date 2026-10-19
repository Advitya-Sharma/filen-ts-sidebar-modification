import { bytesSource, type ByteSource } from "../../core/opc/index"
import { openEncryptedPackage, type DecryptedPackage, type LockedPackage } from "../encryption/sealedPackage"
import { exportOfficeKey, type OfficeKeyTransfer } from "../encryption/webCrypto"
import type { DocxEngineServices } from "./engine"
import { inThreadDocxHost, type DocxHost } from "./host"
import { heldSource, type DocxEncryptionRefusal, type DocxProtection, type DocxProtectionOptions } from "./protection"
import type { DocxOpenOptions } from "./protocol"

// The side of password protection that holds the document: classify, unlock, keep the decrypted package and the
// key, hand an editor its copy. Runs in the docx worker (serveDocxWorker) or, for a host with no worker, on the
// caller's thread (inThreadDocxProtection).

export type DocxProtectVerdict =
	| { readonly kind: "plain" }
	| { readonly kind: "refused"; readonly reason: DocxEncryptionRefusal }
	| { readonly kind: "needsPassword"; readonly scheme: "standard" | "agile"; readonly spinCount: number }

export type DocxUnlockAnswer =
	| { readonly ok: true; readonly scheme: "standard" | "agile"; readonly integrity: "verified" | "unchecked"; readonly size: number }
	| { readonly ok: false; readonly reason: "wrongPassword" | "aborted" | DocxEncryptionRefusal }

interface Held {
	locked: LockedPackage | undefined
	document: DecryptedPackage | undefined
	// The decrypted package while it is held here: dropped once the editor has its own copy.
	plaintext: Uint8Array | undefined
	closed: boolean
	readonly aborts: Map<number, { stopped: boolean }>
}

export interface DocxProtectionServer {
	protect(id: number, source: ByteSource): Promise<DocxProtectVerdict>
	// `outer`: the caller's own signal, read alongside abort(id, call).
	unlock(id: number, call: number, password: string, outer?: { readonly aborted: boolean }): Promise<DocxUnlockAnswer>
	// A copy of the decrypted package, for an editor.
	take(id: number): Uint8Array
	// The key as the editor's thread registers it: a Standard source's upgrade runs here first, sliced.
	key(id: number, call: number, outer?: { readonly aborted: boolean }): Promise<OfficeKeyTransfer>
	// Drops the decrypted package here (the editor has its copy); the key stays for key().
	drop(id: number): void
	abort(id: number, call: number): void
	// The decrypted package of an unlocked document, for an engine session opened on it.
	plaintext(id: number): Uint8Array | undefined
	// Drops a held document (its key released); false when `id` is not one.
	close(id: number): boolean
}

export function docxProtectionServer(): DocxProtectionServer {
	const held = new Map<number, Held>()
	const signalOf = (entry: Held, call: number, outer: { readonly aborted: boolean } | undefined): { readonly aborted: boolean } => {
		const own = { stopped: false }

		entry.aborts.set(call, own)

		return {
			get aborted() {
				return own.stopped || entry.closed || outer?.aborted === true
			}
		}
	}

	return {
		async protect(id, source) {
			const entry: Held = { locked: undefined, document: undefined, plaintext: undefined, closed: false, aborts: new Map() }

			held.set(id, entry)

			const opened = await openEncryptedPackage(source)

			if (entry.closed) return { kind: "plain" }

			switch (opened.kind) {
				case "needsPassword":
					entry.locked = opened.locked

					return { kind: "needsPassword", scheme: opened.locked.scheme, spinCount: opened.locked.spinCount }
				case "unlocked":
					// Not asked for (no default password for documents); never kept.
					opened.document.release()
					held.delete(id)

					return { kind: "plain" }
				case "refused":
					held.delete(id)

					return { kind: "refused", reason: opened.reason }
				case "plain":
				case "aborted":
					held.delete(id)

					return { kind: "plain" }
			}
		},
		async unlock(id, call, password, outer) {
			const entry = held.get(id)
			const locked = entry?.locked

			if (entry === undefined || locked === undefined || entry.document !== undefined) return { ok: false, reason: "aborted" }

			const signal = signalOf(entry, call, outer)

			try {
				const result = await locked.unlock(password, { signal })

				if (!result.ok) return result
				if (entry.closed || entry.document !== undefined) {
					result.document.release()

					return { ok: false, reason: "aborted" }
				}

				entry.document = result.document
				entry.plaintext = result.document.bytes
				entry.locked = undefined

				return {
					ok: true,
					scheme: result.document.scheme,
					integrity: result.document.integrity,
					size: result.document.bytes.length
				}
			} finally {
				entry.aborts.delete(call)
			}
		},
		take(id) {
			const entry = held.get(id)
			const plaintext = entry?.closed === false ? entry.plaintext : undefined

			if (plaintext === undefined) throw new Error("The document is not unlocked")

			return plaintext.slice()
		},
		async key(id, call, outer) {
			const entry = held.get(id)
			const document = entry?.document

			if (entry === undefined || document === undefined || entry.closed) throw new Error("The document is not unlocked")

			const signal = signalOf(entry, call, outer)

			try {
				const key = await exportOfficeKey(document.key, { signal })

				// Closed meanwhile (in the same turn too): nothing is handed out.
				if (entry.closed) throw new Error("The document is not unlocked")

				return key
			} finally {
				entry.aborts.delete(call)
			}
		},
		drop(id) {
			const entry = held.get(id)

			if (entry !== undefined) entry.plaintext = undefined
		},
		abort(id, call) {
			const signal = held.get(id)?.aborts.get(call)

			if (signal !== undefined) signal.stopped = true
		},
		plaintext(id) {
			const entry = held.get(id)

			return entry?.closed === false ? entry.plaintext : undefined
		},
		close(id) {
			const entry = held.get(id)

			if (entry === undefined) return false

			held.delete(id)
			entry.closed = true

			for (const signal of entry.aborts.values()) signal.stopped = true

			entry.document?.release()
			entry.document = undefined
			entry.plaintext = undefined
			entry.locked = undefined

			return true
		}
	}
}

// Password protection on the caller's thread, for hosts without a worker (a WebView, tests): the decrypted package
// lives on this thread, with the same interface as the worker pool's.
export function inThreadDocxProtection(services: (options: DocxOpenOptions) => DocxEngineServices): DocxProtection {
	const engines = inThreadDocxHost(services)

	return {
		async open(source, options = {}) {
			if (options.signal?.aborted === true) return { kind: "aborted" }

			const server = docxProtectionServer()
			const verdict = await server.protect(0, source)

			if (verdict.kind !== "needsPassword") return verdict

			let next = 1
			const locked = {
				scheme: verdict.scheme,
				spinCount: verdict.spinCount,
				unlock: async (password: string, unlockOptions: DocxProtectionOptions = {}) => {
					const answer = await server.unlock(0, next++, password, unlockOptions.signal)

					if (!answer.ok) return answer

					return {
						ok: true,
						document: {
							scheme: answer.scheme,
							integrity: answer.integrity,
							bytes: heldSource(answer.size),
							host: ((_source, listener, openOptions) => {
								const plaintext = server.plaintext(0)

								if (plaintext !== undefined) return engines(bytesSource(plaintext), listener, openOptions)

								listener({ type: "error", reason: "failed", message: "The document is no longer unlocked" })

								return {
									objects: () => Promise.resolve([]),
									media: () => Promise.resolve(null),
									locate: () => Promise.resolve(null),
									find: () => Promise.resolve(new Int32Array(0)),
									close: () => undefined
								}
							}) satisfies DocxHost,
							take: (takeOptions: DocxProtectionOptions = {}) => {
								if (takeOptions.signal?.aborted === true) return Promise.reject(new Error("Aborted"))

								try {
									const bytes = server.take(0)

									return Promise.resolve({
										bytes,
										key: (keyOptions: DocxProtectionOptions = {}) => server.key(0, next++, keyOptions.signal)
									})
								} catch (error) {
									return Promise.reject(error instanceof Error ? error : new Error(String(error)))
								}
							},
							dropPackage: () => server.drop(0),
							release: () => {
								server.close(0)
							}
						}
					} as const
				},
				dispose: () => {
					server.close(0)
				}
			}

			return { kind: "needsPassword", locked }
		}
	}
}
