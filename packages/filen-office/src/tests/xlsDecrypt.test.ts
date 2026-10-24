import { createHash } from "node:crypto"
import { describe, expect, it } from "vitest"
import { CfbFile } from "@filen/office/core/cfb"
import { bytesSource } from "@filen/office/core/opc"
import type { DocDecryption } from "../core/doc/decrypt"
import { decryptXlsWorkbook, xlsEncryption } from "../core/xls/decrypt"
import {
	encryptedXlsStream,
	xlsFile,
	xlsLabel,
	xlsNumber,
	xlsRecord,
	xlsWorkbookStream,
	type XlsEncryptOptions,
	type XlsSheetSpec
} from "./encryption/xls"

function decryption(password: string): DocDecryption {
	return { password, sha1: data => Promise.resolve(new Uint8Array(createHash("sha1").update(data).digest())) }
}

// Enough records for several 1024-byte RC4 blocks, records straddling their edges, and two sheets.
const SHEETS: readonly XlsSheetSpec[] = [
	{
		name: "First",
		records: Array.from({ length: 120 }, (_, k) =>
			k % 3 === 0 ? xlsLabel(k, 0, `text ${k} `.repeat(k % 7)) : xlsNumber(k, 1, k * 1.5)
		)
	},
	{ name: "Second", records: [xlsLabel(0, 0, "second sheet"), xlsNumber(1, 0, 42)] }
]

async function workbookOf(file: Uint8Array): Promise<Uint8Array> {
	const cfb = await CfbFile.open(bytesSource(file))

	return (await cfb.read("Workbook")).slice()
}

async function decrypted(options: XlsEncryptOptions, password = options.password) {
	const { plain, encrypted } = encryptedXlsStream(SHEETS, options)
	const file = await xlsFile(encrypted)
	const before = file.slice()
	const encryption = await xlsEncryption(file)

	if (encryption.kind !== "encrypted") throw new Error(encryption.kind)

	const result = await decryptXlsWorkbook(file, encryption, decryption(password))

	expect(file).toEqual(before)

	return { plain, encryption, result }
}

const SCHEMES: readonly XlsEncryptOptions[] = [
	{ scheme: "xor", password: "secret" },
	{ scheme: "xor", password: "123456789012345" },
	{ scheme: "rc4", password: "password1" },
	{ scheme: "rc4", password: "VelvetSweatshop" },
	{ scheme: "cryptoApiRc4", password: "Pässwort", keySize: 128 },
	{ scheme: "cryptoApiRc4", password: "Pässwort", keySize: 56 },
	{ scheme: "cryptoApiRc4", password: "Pässwort", keySize: 0 }
]

describe("encrypted Excel 97–2003 workbooks", () => {
	it.each(SCHEMES.map(options => [`${options.scheme} ${options.keySize ?? ""} ${options.password}`, options] as const))(
		"decrypts a %s Workbook stream to the plain one, the clear records and lbPlyPos kept",
		async (_, options) => {
			const { plain, encryption, result } = await decrypted(options)

			expect(encryption.scheme).toBe(options.scheme)

			if (result.kind !== "decrypted") throw new Error(result.kind)

			expect(Buffer.from(await workbookOf(result.file)).equals(plain)).toBe(true)
		}
	)

	it.each(SCHEMES.map(options => [options.scheme, options] as const))("answers a wrong %s password as wrong", async (_, options) => {
		expect((await decrypted(options, "not it")).result).toEqual({ kind: "wrongPassword" })
	})

	it("reads a workbook without a FilePass as plain, and a file that is no workbook as plain for its reader to refuse", async () => {
		expect(await xlsEncryption(await xlsFile(xlsWorkbookStream(SHEETS)))).toEqual({ kind: "plain" })
		expect(await xlsEncryption(new Uint8Array(600))).toEqual({ kind: "plain" })
	})

	it("finds the FilePass after a WriteProtect, and no later than that", async () => {
		const { encrypted } = encryptedXlsStream(SHEETS, { scheme: "rc4", password: "p" })
		const bof = encrypted.subarray(0, 4 + ((encrypted[2] ?? 0) | ((encrypted[3] ?? 0) << 8)))
		const rest = encrypted.subarray(bof.length)
		const protectedFirst = new Uint8Array([...bof, ...xlsRecord(0x0086), ...rest])
		const late = new Uint8Array([...bof, ...xlsRecord(0x00e1, new Uint8Array(2)), ...rest])

		expect((await xlsEncryption(await xlsFile(protectedFirst))).kind).toBe("encrypted")
		expect((await xlsEncryption(await xlsFile(late))).kind).toBe("plain")
	})

	it("refuses a cipher none of these is as unsupported, before any password", async () => {
		const otherType = xlsWorkbookStream(SHEETS, Uint8Array.of(2, 0, 0, 0))
		// An RC4 CryptoAPI header naming AES-128 (AlgID 0x660E).
		const { encrypted } = encryptedXlsStream(SHEETS, { scheme: "cryptoApiRc4", password: "p" })
		const aes = encrypted.slice()
		const filePassAt = 4 + ((aes[2] ?? 0) | ((aes[3] ?? 0) << 8))
		const algIdAt = filePassAt + 4 + 2 + 12 + 8

		aes.set([0x0e, 0x66, 0, 0], algIdAt)

		expect(await xlsEncryption(await xlsFile(otherType))).toEqual({ kind: "unsupported" })
		expect(await xlsEncryption(await xlsFile(aes))).toEqual({ kind: "unsupported" })
	})

	it("refuses a FilePass cut short as damaged", async () => {
		expect((await xlsEncryption(await xlsFile(xlsWorkbookStream(SHEETS, Uint8Array.of(0, 0, 1))))).kind).toBe("damaged")
		expect((await xlsEncryption(await xlsFile(xlsWorkbookStream(SHEETS, Uint8Array.of(1, 0, 4, 0, 2, 0, 4, 0))))).kind).toBe("damaged")
	})

	it("answers an RC4 header cut short as damaged when it is decrypted", async () => {
		const file = await xlsFile(xlsWorkbookStream(SHEETS, Uint8Array.of(1, 0, 1, 0, 1, 0, 9, 9)))
		const encryption = await xlsEncryption(file)

		if (encryption.kind !== "encrypted") throw new Error(encryption.kind)

		expect((await decryptXlsWorkbook(file, encryption, decryption("p"))).kind).toBe("damaged")
	})
})
