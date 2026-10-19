import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs"
import { join, resolve } from "node:path"
import { describe, expect, it, vi } from "vitest"
import { CfbFile } from "@filen/office/core/cfb"
import { bytesSource } from "@filen/office/core/opc"
import { parseEncryptionInfo, releaseOfficeKey, unlockOfficePackage } from "@filen/office/dom/encryption"
import { cellKey } from "@/features/spreadsheet/lib/model"
import { assertSealedUpload } from "@/features/spreadsheet/lib/xlsxEncryption"
import type { SpreadsheetOpen, SpreadsheetOpened, SpreadsheetWorkerApi } from "@/features/spreadsheet/workers/spreadsheet.worker"
import { encryptedWorkbook } from "@/tests/encryptedWorkbookFixtures"

// The password-protected workbooks of the office corpus (never committed: `pnpm run corpus:fetch` in
// packages/filen-office, or FILEN_OFFICE_CORPUS) through the spreadsheet worker: Excel's default password
// opens silently, the others prompt, wrong passwords are asked again; each opens, saves sealed (unedited and
// edited) and opens again with its password, unless the same workbook unencrypted is view-only too (a save
// would lose part of it, as for poi__58616's cell styles). A save's package is the workbook the unencrypted path
// writes, byte for byte: sealing changes nothing about what is saved.
//
// FILEN_OFFICE_SEAL_OUT=<dir> also writes each sealed save with its password and expected package, in the
// manifest form the office package's encryption oracles read (packages/filen-office
// src/tests/encryption/oracles: FILEN_OFFICE_SEALED=<dir>/all, or <dir>/nosha1 for tools that misread Agile
// SHA-1).

const { exposed } = vi.hoisted(() => {
	const api: { api: unknown } = { api: undefined }

	return { exposed: api }
})

vi.mock(import("comlink"), async importOriginal => ({
	...(await importOriginal()),
	expose: (api: unknown) => {
		exposed.api = api
	}
}))

await import("@/features/spreadsheet/workers/spreadsheet.worker")

const api = exposed.api as SpreadsheetWorkerApi
const OFFICE = join(process.cwd(), "..", "filen-office")
const configured = process.env["FILEN_OFFICE_CORPUS"]
const CORPUS = configured !== undefined && configured !== "" ? resolve(configured) : join(OFFICE, ".corpus")
const manifest = JSON.parse(readFileSync(join(OFFICE, "corpus", "manifest.json"), "utf8")) as {
	files: { path: string; password?: string }[]
}
const SAMPLES = manifest.files
	.filter(file => /^(xlsx|xlsm|xltx|xltm)\//u.test(file.path) && file.password !== undefined)
	.filter(file => existsSync(join(CORPUS, file.path)))
const OUT = process.env["FILEN_OFFICE_SEAL_OUT"]
const MARKER = "Filen sealed edit"

const oracleEntries: { entry: { file: string; password: string; plaintext: string }; sha1: boolean }[] = []

function record(name: string, kind: string, sealed: Uint8Array, plaintext: Uint8Array, password: string, sha1: boolean): void {
	if (OUT === undefined || OUT === "") {
		return
	}

	mkdirSync(OUT, { recursive: true })

	const base = `${name.replace(/[^\w.-]+/gu, "_")}.${kind}`
	const file = join(OUT, `${base}.xlsx`)
	const plain = join(OUT, `${base}.zip`)

	writeFileSync(file, sealed)
	writeFileSync(plain, plaintext)
	oracleEntries.push({ entry: { file, password, plaintext: plain }, sha1 })

	for (const [dir, entries] of [
		["all", oracleEntries],
		["nosha1", oracleEntries.filter(item => !item.sha1)]
	] as const) {
		mkdirSync(join(OUT, dir), { recursive: true })
		writeFileSync(join(OUT, dir, "manifest.json"), JSON.stringify({ version: 1, entries: entries.map(item => item.entry) }, null, "\t"))
	}
}

function opened(result: SpreadsheetOpen | Awaited<ReturnType<SpreadsheetWorkerApi["unlock"]>>): SpreadsheetOpened {
	if (!("id" in result)) {
		throw new Error(`expected an opened workbook, got ${JSON.stringify(result)}`)
	}

	return result
}

async function container(file: Uint8Array): Promise<{ encryptionInfo: Uint8Array; encryptedPackage: Uint8Array }> {
	const cfb = await CfbFile.open(bytesSource(file))

	return { encryptionInfo: await cfb.read(["EncryptionInfo"]), encryptedPackage: await cfb.read(["EncryptedPackage"]) }
}

// A sealed save's package, decrypted with its password (the oracles decrypt it with other tools).
async function decryptSealed(file: Uint8Array, password: string): Promise<Uint8Array> {
	const read = await container(file)
	const result = await unlockOfficePackage(read.encryptionInfo, read.encryptedPackage, password)

	if (!result.ok) {
		throw new Error(`not decrypted: ${result.reason}`)
	}

	releaseOfficeKey(result.key)

	return result.plaintext
}

// Opens a protected file: silently with the default password, else with `password` after a wrong one.
async function openWith(file: Uint8Array, password: string): Promise<SpreadsheetOpened> {
	const result = await api.open(file.slice(), "xlsx")

	if (password === "VelvetSweatshop") {
		return opened(result)
	}

	if (!("locked" in result)) {
		throw new Error(`expected a password prompt, got ${JSON.stringify(result)}`)
	}

	expect(await api.unlock(result.locked, `${password}x`)).toEqual({ wrongPassword: true })

	return opened(await api.unlock(result.locked, password))
}

describe.skipIf(SAMPLES.length === 0)("encrypted xlsx corpus", () => {
	it.each(SAMPLES.map(sample => [sample.path, sample.password ?? ""] as const))(
		"%s: opens, saves sealed and reopens",
		{ timeout: 120_000 },
		async (path, password) => {
			const file = new Uint8Array(readFileSync(join(CORPUS, path)))
			const source = await container(file)
			const info = parseEncryptionInfo(source.encryptionInfo)
			const sha1 = info.kind === "agile" && info.keyData.hash === "SHA-1"
			const unlocked = await unlockOfficePackage(source.encryptionInfo, source.encryptedPackage, password)

			if (!unlocked.ok) {
				throw new Error(`the backend did not open ${path}: ${unlocked.reason}`)
			}

			const workbook = await openWith(file, password)

			expect(["verified", "unchecked"]).toContain(workbook.encrypted)
			expect(workbook.doc.sheets.length).toBeGreaterThan(0)

			// The unencrypted path over the same package: what every save must decrypt to.
			const plain = opened(await api.open(unlocked.plaintext.slice(), "xlsx"))

			const writable = await api.writability(plain.id, unlocked.plaintext.slice())

			// Encryption changes nothing about whether it saves intact.
			expect(await api.writability(workbook.id, null)).toBe(writable)

			if (!writable) {
				// View-only, as the same workbook unencrypted (a save would lose part of it): nothing is sealed.
				const edit = await api.apply(workbook.id, { type: "setCells", sheet: 0, cells: [{ row: 0, col: 0, input: MARKER }] })

				expect(edit.type).toBe("refused")
				await expect(api.serialize(workbook.id)).rejects.toThrow("cannot be saved")
				api.close(plain.id)
				api.close(workbook.id)
				releaseOfficeKey(unlocked.key)

				return
			}

			const written = (await api.serialize(plain.id)).bytes
			const unedited = (await api.serialize(workbook.id)).bytes

			assertSealedUpload(unedited)
			expect(await decryptSealed(unedited, password)).toEqual(written)
			record(path, "unedited", unedited, written, password, sha1)
			api.close(plain.id)

			const edit = await api.apply(workbook.id, { type: "setCells", sheet: 0, cells: [{ row: 0, col: 0, input: MARKER }] })

			expect(edit.type).not.toBe("refused")

			const edited = (await api.serialize(workbook.id)).bytes

			assertSealedUpload(edited)
			api.close(workbook.id)

			const reopened = await openWith(edited, password)

			expect(reopened.doc.sheets[0]?.cells.get(cellKey(0, 0))?.text).toBe(MARKER)
			expect(reopened.doc.sheets.map(sheet => sheet.name)).toEqual(workbook.doc.sheets.map(sheet => sheet.name))
			record(path, "edited", edited, await decryptSealed(edited, password), password, sha1)
			api.close(reopened.id)
			releaseOfficeKey(unlocked.key)
		}
	)
})

// Excel's default password (encryptedWorkbookFixtures.ts "default"): opened without a prompt, its save
// sealed with the same password (for the oracles too, as the corpus sample with it is view-only).
describe("encrypted xlsx with the default password", () => {
	it("opens silently, saves sealed and reopens silently", { timeout: 120_000 }, async () => {
		const file = await encryptedWorkbook("default")
		const workbook = opened(await api.open(file, "xlsx"))

		expect(workbook.encrypted).toBe("verified")
		expect(await api.writability(workbook.id, null)).toBe(true)
		await api.apply(workbook.id, { type: "setCells", sheet: 0, cells: [{ row: 2, col: 0, input: MARKER }] })

		const edited = (await api.serialize(workbook.id)).bytes

		assertSealedUpload(edited)
		api.close(workbook.id)

		const reopened = opened(await api.open(edited.slice(), "xlsx"))

		expect(reopened.encrypted).toBe("verified")
		expect(reopened.doc.sheets[0]?.cells.get(cellKey(2, 0))?.text).toBe(MARKER)
		api.close(reopened.id)
		record("synthetic_velvetsweatshop", "edited", edited, await decryptSealed(edited, "VelvetSweatshop"), "VelvetSweatshop", false)
	})
})
