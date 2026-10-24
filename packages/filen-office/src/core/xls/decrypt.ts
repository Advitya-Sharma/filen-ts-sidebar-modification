import { CfbFile, cfbBytes } from "../cfb/index"
import {
	cryptoApiHeaderProblem,
	type DocCipherScheme,
	type DocDecryption,
	docCipherScheme,
	docKey,
	xorObfuscationArray
} from "../doc/decrypt"
import { DocError } from "../doc/errors"
import { bytesSource } from "../opc/io"

// Password to open on an Excel 97–2003 workbook ([MS-XLS] §2.2.10, §2.4.117 FilePass): XOR obfuscation ([MS-OFFCRYPTO]
// §2.3.7, Data Transformation Method 1), RC4 (§2.3.6) and CryptoAPI RC4 (§2.3.5), keyed as a Word document's
// (core/doc/decrypt.ts). The Workbook stream is decrypted into a new array (the file's bytes are never written) and
// handed on as a compound file of its own, for the BIFF8 reader. Read-only: nothing writes these ciphers.

export type XlsCipherScheme = DocCipherScheme

export type XlsEncryption =
	| { readonly kind: "plain" }
	// A FilePass of a cipher none of these is (another wEncryptionType or CryptoAPI algorithm).
	| { readonly kind: "unsupported" }
	| { readonly kind: "damaged"; readonly message: string }
	| { readonly kind: "encrypted"; readonly scheme: XlsCipherScheme; readonly filePass: Uint8Array }

export type XlsDecrypted =
	| { readonly kind: "decrypted"; readonly file: Uint8Array }
	| { readonly kind: "wrongPassword" }
	| { readonly kind: "unsupported" }
	| { readonly kind: "damaged"; readonly message: string }

const BOF = 0x0809
const FILEPASS = 0x002f
const BOUNDSHEET = 0x0085
const WRITEPROTECT = 0x0086
// [MS-XLS] §2.2.10: BOF, FilePass, UsrExcl, FileLock, InterfaceHdr, RRDInfo and RRDHead are never encrypted, nor is
// BoundSheet8's lbPlyPos (its first 4 bytes). Record headers never are.
const CLEAR = new Set([BOF, FILEPASS, 0x0194, 0x0195, 0x00e1, 0x0196, 0x0138])
// §2.2.10: RC4 rekeys every 1024 bytes of the stream.
const RC4_BLOCK = 1024
// The globals' first records read for a FilePass: it follows BOF and at most WriteProtect ([MS-XLS] §2.1.7.20.1).
const HEAD_WINDOW = 4096

function le16(bytes: Uint8Array, at: number): number {
	return (bytes[at] ?? 0) | ((bytes[at + 1] ?? 0) << 8)
}

// The Workbook stream of a BIFF8 compound file, else undefined (not one: the BIFF8 reader says why).
async function workbookStream(file: Uint8Array): Promise<{ readonly cfb: CfbFile; readonly size: number } | undefined> {
	let cfb: CfbFile

	try {
		cfb = await CfbFile.open(bytesSource(file))
	} catch {
		return undefined
	}

	const entry = cfb.entry("Workbook")

	return entry?.kind === "stream" ? { cfb, size: entry.size } : undefined
}

// Whether the workbook has a password to open, and which cipher. Reads the compound file's directory and the first
// records of the Workbook stream only.
export async function xlsEncryption(file: Uint8Array): Promise<XlsEncryption> {
	const workbook = await workbookStream(file)

	if (workbook === undefined) return { kind: "plain" }

	const head = await workbook.cfb.stream("Workbook").read(0, Math.min(HEAD_WINDOW, workbook.size))

	for (let at = 0, index = 0; at + 4 <= head.length; index++) {
		const id = le16(head, at)
		const size = le16(head, at + 2)

		if (index === 0) {
			if (id !== BOF) return { kind: "plain" }
		} else if (id === FILEPASS) {
			if (at + 4 + size > head.length) return { kind: "damaged", message: "the FilePass record is cut short" }

			return filePassEncryption(head.slice(at + 4, at + 4 + size))
		} else if (id !== WRITEPROTECT) {
			return { kind: "plain" }
		}

		at += 4 + size
	}

	return { kind: "plain" }
}

// FilePass ([MS-XLS] §2.4.117): wEncryptionType 0 is XOR (XorKey, verificationBytes), 1 an RC4 EncryptionHeader.
function filePassEncryption(filePass: Uint8Array): XlsEncryption {
	if (filePass.length < 2) return { kind: "damaged", message: "the FilePass record is cut short" }

	const type = le16(filePass, 0)

	if (type === 0)
		return filePass.length < 6
			? { kind: "damaged", message: "the XOR FilePass is cut short" }
			: { kind: "encrypted", scheme: "xor", filePass }
	if (type !== 1) return { kind: "unsupported" }

	const header = filePass.subarray(2)
	const scheme = docCipherScheme(false, header)

	if (scheme === "unknown") return { kind: "unsupported" }

	if (scheme === "cryptoApiRc4") {
		const problem = cryptoApiHeaderProblem(header, header.length)

		if (problem !== undefined) return { kind: "damaged", message: problem }
	}

	return { kind: "encrypted", scheme, filePass }
}

// `encryption`'s Workbook stream decrypted with `decryption.password`, as a compound file holding only it (the
// FilePass record kept, which a reader skips). The plaintext stream is wiped once written into the file; the file
// is the caller's to wipe.
export async function decryptXlsWorkbook(
	file: Uint8Array,
	encryption: Extract<XlsEncryption, { kind: "encrypted" }>,
	decryption: DocDecryption
): Promise<XlsDecrypted> {
	const workbook = await workbookStream(file)

	if (workbook === undefined) return { kind: "damaged", message: "no Workbook stream" }

	let plain: Uint8Array

	try {
		const stream = await workbook.cfb.read("Workbook")

		plain =
			encryption.scheme === "xor"
				? xorDecrypt(stream, encryption.filePass, decryption.password)
				: await rc4Decrypt(stream, encryption, decryption)
	} catch (error) {
		if (!(error instanceof DocError)) throw error

		switch (error.code) {
			case "wrongPassword":
				return { kind: "wrongPassword" }
			case "encrypted":
				return { kind: "unsupported" }
			default:
				return { kind: "damaged", message: error.message }
		}
	}

	try {
		return { kind: "decrypted", file: await cfbBytes([{ path: ["Workbook"], bytes: plain }]) }
	} finally {
		plain.fill(0)
	}
}

// Calls `visit` with each encrypted run of `stream`: its offset, length and its record's data size.
function forEachEncryptedRun(stream: Uint8Array, visit: (at: number, length: number, size: number) => void): void {
	for (let at = 0; at + 4 <= stream.length;) {
		const id = le16(stream, at)
		const size = le16(stream, at + 2)
		const data = at + 4
		const end = Math.min(data + size, stream.length)

		if (!CLEAR.has(id)) {
			const from = id === BOUNDSHEET ? Math.min(data + 4, end) : data

			if (end > from) visit(from, end - from, size)
		}

		at = data + size
	}
}

// RC4 runs its key stream over the whole stream, record headers and clear records included ([MS-XLS] §2.2.10): the
// stream is decrypted whole, then the bytes that were never encrypted are put back.
async function rc4Decrypt(
	stream: Uint8Array,
	encryption: Extract<XlsEncryption, { kind: "encrypted" }>,
	decryption: DocDecryption
): Promise<Uint8Array> {
	const header = encryption.filePass.subarray(2)
	const key = await docKey(encryption.scheme, header.length, header, decryption, { block: RC4_BLOCK })

	try {
		const decrypted = await key.decrypt(stream, 0)
		const out = stream.slice()

		forEachEncryptedRun(stream, (at, length) => {
			out.set(decrypted.subarray(at, at + length), at)
		})
		decrypted.fill(0)

		return out
	} finally {
		key.wipe()
	}
}

// XOR Data Transformation Method 1 ([MS-OFFCRYPTO] §2.3.7.7 with [MS-XLS] §2.2.10): each encrypted byte XORed with
// XorArray at (its stream offset + its record's size) mod 16, then rotated right by 5.
function xorDecrypt(stream: Uint8Array, filePass: Uint8Array, password: string): Uint8Array {
	const array = xorObfuscationArray(((le16(filePass, 2) << 16) | le16(filePass, 4)) >>> 0, password)
	const out = stream.slice()

	forEachEncryptedRun(stream, (from, length, size) => {
		for (let at = from; at < from + length; at++) {
			const value = ((stream[at] ?? 0) ^ (array[(at + size) & 15] ?? 0)) & 0xff

			out[at] = ((value >>> 5) | (value << 3)) & 0xff
		}
	})
	array.fill(0)

	return out
}
