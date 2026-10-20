import { readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it, vi } from "vitest"
import { CfbFile, cfbBytes } from "@filen/office/core/cfb"
import { bytesSource } from "@filen/office/core/opc"
import { releaseOfficeKey, unlockOfficePackage } from "@filen/office/dom/encryption"
import { cellKey, type SpreadsheetDoc } from "@filen/office/core/sheet"
import { assertSealedUpload } from "@/features/spreadsheet/lib/xlsxEncryption"
import type { SpreadsheetOpen, SpreadsheetOpened, SpreadsheetWorkerApi } from "@/features/spreadsheet/workers/spreadsheet.worker"
import { WORKBOOK_PASSWORD, encryptedWorkbook } from "@/tests/encryptedWorkbookFixtures"

// The spreadsheet worker's password-protected workbooks, with the real WebCrypto backend: refused or locked
// before anything is shown, Excel's default password tried silently, wrong passwords and cancelled tries asked
// again, every save sealed (a compound file, never the zip) and opened again, the key released on close.
// Fixtures: encryptedWorkbookFixtures.ts, fixtures/spreadsheet/lo-plain.xlsx encrypted at test time.

const { released, exposed } = vi.hoisted(() => {
	const api: { api: unknown } = { api: undefined }

	return { released: { count: 0 }, exposed: api }
})

vi.mock(import("comlink"), async importOriginal => ({
	...(await importOriginal()),
	expose: (api: unknown) => {
		exposed.api = api
	}
}))
// Counts the keys released: each opened workbook's, once.
vi.mock(import("@filen/office/dom/encryption"), async importOriginal => {
	const actual = await importOriginal()

	type Document = Extract<Awaited<ReturnType<typeof actual.openEncryptedPackage>>, { kind: "unlocked" }>["document"]

	const countedDocument = (document: Document): Document => ({
		...document,
		release: () => {
			released.count++
			document.release()
		}
	})

	return {
		...actual,
		openEncryptedPackage: async (...args: Parameters<typeof actual.openEncryptedPackage>) => {
			const opened = await actual.openEncryptedPackage(...args)

			if (opened.kind === "unlocked") {
				return { kind: "unlocked", document: countedDocument(opened.document) }
			}

			if (opened.kind !== "needsPassword") {
				return opened
			}

			const { locked } = opened

			return {
				kind: "needsPassword",
				locked: {
					scheme: locked.scheme,
					spinCount: locked.spinCount,
					unlock: async (...unlockArgs: Parameters<typeof locked.unlock>) => {
						const result = await locked.unlock(...unlockArgs)

						return result.ok ? { ok: true, document: countedDocument(result.document) } : result
					}
				}
			}
		}
	}
})

await import("@/features/spreadsheet/workers/spreadsheet.worker")

const api = exposed.api as SpreadsheetWorkerApi
const PASSWORD = WORKBOOK_PASSWORD
const FIXTURES = join(process.cwd(), "src/tests/fixtures/spreadsheet")

function fixture(name: string): Uint8Array {
	return new Uint8Array(readFileSync(join(FIXTURES, `${name}.xlsx`)))
}

function opened(result: SpreadsheetOpen | Awaited<ReturnType<SpreadsheetWorkerApi["unlock"]>>): SpreadsheetOpened {
	if (!("id" in result)) {
		throw new Error(`expected an opened workbook, got ${JSON.stringify(result)}`)
	}

	return result
}

function lockOf(result: SpreadsheetOpen): number {
	if (!("locked" in result)) {
		throw new Error(`expected a locked workbook, got ${JSON.stringify(result)}`)
	}

	return result.locked
}

function text(doc: SpreadsheetDoc, row: number, col: number): string | undefined {
	return doc.sheets[0]?.cells.get(cellKey(row, col))?.text
}

const DATA_SPACES = [
	["\u0006DataSpaces", "Version"],
	["\u0006DataSpaces", "DataSpaceMap"],
	["\u0006DataSpaces", "DataSpaceInfo", "StrongEncryptionDataSpace"],
	["\u0006DataSpaces", "TransformInfo", "StrongEncryptionTransform", "\u0006Primary"]
] as const

interface Streams {
	encryptionInfo: Uint8Array
	encryptedPackage: Uint8Array
}

async function streams(file: Uint8Array): Promise<Streams & { dataSpaces: { path: readonly string[]; bytes: Uint8Array }[] }> {
	const cfb = await CfbFile.open(bytesSource(file))

	return {
		encryptionInfo: await cfb.read(["EncryptionInfo"]),
		encryptedPackage: await cfb.read(["EncryptedPackage"]),
		dataSpaces: await Promise.all(DATA_SPACES.map(async path => ({ path, bytes: await cfb.read(path) })))
	}
}

// The agile fixture with its EncryptionInfo or EncryptedPackage replaced.
async function variant(change: (read: Streams) => Streams): Promise<Uint8Array> {
	const read = await streams(await encryptedWorkbook("agile"))
	const changed = change(read)

	return await cfbBytes([
		...read.dataSpaces,
		{ path: ["EncryptionInfo"], bytes: changed.encryptionInfo },
		{ path: ["EncryptedPackage"], bytes: changed.encryptedPackage }
	])
}

// A file's package, decrypted with its password.
async function decrypt(file: Uint8Array, password: string): Promise<Uint8Array> {
	const read = await streams(file)
	const result = await unlockOfficePackage(read.encryptionInfo, read.encryptedPackage, password)

	if (!result.ok) {
		throw new Error(`not decrypted: ${result.reason}`)
	}

	releaseOfficeKey(result.key)

	return result.plaintext
}

// Proves the workbook, makes an edit and saves it.
async function editAndSave(id: number, input: string): Promise<Uint8Array> {
	expect(await api.writability(id, null)).toBe(true)
	expect((await api.apply(id, { type: "setCells", sheet: 0, cells: [{ row: 20, col: 0, input }] })).type).not.toBe("refused")

	return (await api.serialize(id)).bytes
}

describe("spreadsheet worker: password-protected workbooks", () => {
	it("asks for the password, says when it is wrong, and opens with the right one", async () => {
		const lock = lockOf(await api.open(await encryptedWorkbook("agile"), "xlsx"))

		expect(await api.unlock(lock, "wrong")).toEqual({ wrongPassword: true })
		// NFC and NFD are different passwords (UTF-16 code units, no normalisation).
		expect(await api.unlock(lock, PASSWORD.normalize("NFD"))).toEqual({ wrongPassword: true })

		const workbook = opened(await api.unlock(lock, PASSWORD))
		const plain = opened(await api.open(fixture("lo-plain"), "xlsx"))

		expect(workbook.encrypted).toBe("verified")
		expect(plain.encrypted).toBeUndefined()
		expect(workbook.doc).toEqual(plain.doc)
		await expect(api.unlock(lock, PASSWORD)).rejects.toThrow("no locked workbook")
		api.close(workbook.id)
		api.close(plain.id)
	})

	it("opens a protected workbook named .xls too (the page keeps it read-only, as any .xls)", async () => {
		const workbook = opened(await api.unlock(lockOf(await api.open(await encryptedWorkbook("agile"), "xls")), PASSWORD))

		expect(workbook.doc.kind).toBe("xlsx")
		api.close(workbook.id)
	})

	it("seals every save, which opens again with the password and holds the edit", async () => {
		const workbook = opened(await api.unlock(lockOf(await api.open(await encryptedWorkbook("agile"), "xlsx")), PASSWORD))
		const sealed = await editAndSave(workbook.id, "Sealed edit")

		expect(() => {
			assertSealedUpload(sealed)
		}).not.toThrow()
		expect((await decrypt(sealed, PASSWORD)).subarray(0, 2)).toEqual(Uint8Array.of(0x50, 0x4b))
		api.close(workbook.id)

		const reopened = opened(await api.unlock(lockOf(await api.open(sealed, "xlsx")), PASSWORD))

		expect(text(reopened.doc, 20, 0)).toBe("Sealed edit")
		api.close(reopened.id)
	})

	it("opens Excel's default password silently and seals its saves with it", async () => {
		const workbook = opened(await api.open(await encryptedWorkbook("default"), "xlsx"))

		expect(workbook.encrypted).toBe("verified")

		const sealed = await editAndSave(workbook.id, "Default")

		api.close(workbook.id)

		const reopened = opened(await api.open(sealed, "xlsx"))

		expect(reopened.encrypted).toBe("verified")
		expect(text(reopened.doc, 20, 0)).toBe("Default")
		api.close(reopened.id)
	})

	it("seals exactly the workbook the unencrypted path writes", async () => {
		const plain = opened(await api.open(fixture("lo-plain"), "xlsx"))
		const workbook = opened(await api.unlock(lockOf(await api.open(await encryptedWorkbook("agile"), "xlsx")), PASSWORD))

		expect(await api.writability(plain.id, fixture("lo-plain"))).toBe(true)
		expect(await api.writability(workbook.id, null)).toBe(true)

		const written = (await api.serialize(plain.id)).bytes
		const sealed = (await api.serialize(workbook.id)).bytes

		expect(await decrypt(sealed, PASSWORD)).toEqual(written)
		api.close(plain.id)
		api.close(workbook.id)
	})

	// The upgrade's key encryptor costs a 100,000-round spin, once per session.
	it("upgrades a Standard workbook's saves to Agile", { timeout: 60_000 }, async () => {
		const workbook = opened(await api.unlock(lockOf(await api.open(await encryptedWorkbook("standard"), "xlsx")), PASSWORD))

		// Standard encryption carries no integrity check: the page is told, to say so.
		expect(workbook.encrypted).toBe("unchecked")
		const sealed = await editAndSave(workbook.id, "Upgraded")

		// EncryptionInfo version 4.4: Agile ([MS-OFFCRYPTO] §2.3.4.10).
		expect((await streams(sealed)).encryptionInfo.subarray(0, 4)).toEqual(Uint8Array.of(4, 0, 4, 0))
		expect((await decrypt(sealed, PASSWORD)).subarray(0, 2)).toEqual(Uint8Array.of(0x50, 0x4b))
		api.close(workbook.id)
	})

	it("releases the key on close, so a save still under way fails rather than sealing", async () => {
		const workbook = opened(await api.unlock(lockOf(await api.open(await encryptedWorkbook("agile"), "xlsx")), PASSWORD))

		expect(await api.writability(workbook.id, null)).toBe(true)
		await api.apply(workbook.id, { type: "setCells", sheet: 0, cells: [{ row: 20, col: 0, input: "x" }] })

		const before = released.count
		const saving = api.serialize(workbook.id)

		api.close(workbook.id)

		await expect(saving).rejects.toThrow()
		expect(released.count).toBe(before + 1)
		await expect(api.serialize(workbook.id)).rejects.toThrow("no open document")
	})

	it("stops a cancelled try and keeps the workbook locked; a discarded one is gone", { timeout: 60_000 }, async () => {
		const lock = lockOf(await api.open(await encryptedWorkbook("slow"), "xlsx"))
		const trying = api.unlock(lock, PASSWORD)

		api.cancelUnlock(lock)
		expect(await trying).toEqual({ aborted: true })

		const discarding = api.unlock(lock, PASSWORD)

		api.discardLocked(lock)
		expect(await discarding).toEqual({ aborted: true })
		await expect(api.unlock(lock, PASSWORD)).rejects.toThrow("no locked workbook")
	})

	it("stops the default password's try when the page cancels the open", { timeout: 60_000 }, async () => {
		const opening = api.open(await encryptedWorkbook("slow"), "xlsx", 41)

		api.cancelOpen(41)
		await expect(opening).rejects.toThrow("cancelled")
		// Unknown or finished opens: nothing to stop.
		api.cancelOpen(41)
	})

	it("refuses rights-managed, unsupported and damaged workbooks without a lock", async () => {
		// Extensible encryption (version 4.3 with fExternal): not implemented.
		const extensible = await variant(read => {
			const encryptionInfo = read.encryptionInfo.slice()

			encryptionInfo.set([4, 0, 3, 0, 0x10, 0, 0, 0])

			return { encryptionInfo, encryptedPackage: read.encryptedPackage }
		})
		const truncatedInfo = await variant(read => ({
			encryptionInfo: Uint8Array.of(4, 0, 4, 0),
			encryptedPackage: read.encryptedPackage
		}))

		expect(await api.open(await encryptedWorkbook("irm"), "xlsx")).toEqual({ refused: "rightsManaged" })
		expect(await api.open(extensible, "xlsx")).toEqual({ refused: "unsupported" })
		expect(await api.open(truncatedInfo, "xlsx")).toEqual({ refused: "damaged" })

		// A tampered package fails its integrity check once the password verified.
		const tampered = await variant(read => {
			const encryptedPackage = read.encryptedPackage.slice()

			encryptedPackage[encryptedPackage.length - 1] = (encryptedPackage[encryptedPackage.length - 1] ?? 0) ^ 1

			return { encryptionInfo: read.encryptionInfo, encryptedPackage }
		})

		expect(await api.unlock(lockOf(await api.open(tampered, "xlsx")), PASSWORD)).toEqual({ refused: "damaged" })
	})

	it("opens a plain workbook, and a compound file that is not an encrypted package, as before", async () => {
		const plain = opened(await api.open(fixture("lo-plain"), "xlsx"))

		expect(plain.encrypted).toBeUndefined()
		api.close(plain.id)

		const notEncrypted = await cfbBytes([{ path: ["Workbook"], bytes: new Uint8Array(64) }])

		// Named .xlsx, it is unreadable as it was; named .xls, it goes to the .xls reader.
		await expect(api.open(notEncrypted.slice(), "xlsx")).rejects.toThrow()
		await expect(api.open(notEncrypted.slice(), "xls")).rejects.toThrow()
	})
})
