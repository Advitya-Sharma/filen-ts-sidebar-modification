import { cfbBytes } from "../../core/cfb/index"
import { Rc4 } from "../../core/crypto/legacy"
import { legacyPasswordVerifier } from "../../core/docx/protection/hash"
import { cryptoApiKeys, encryptionHeader, rc4Keys } from "../docDecryptFixtures"

// Excel 97–2003 workbooks ([MS-XLS] BIFF8) built record by record, and encrypted the way Excel writes a password to
// open ([MS-XLS] §2.2.10, §2.4.117 FilePass): XOR obfuscation ([MS-OFFCRYPTO] §2.3.7, Method 1), RC4 (§2.3.6) or
// CryptoAPI RC4 (§2.3.5). Hashes come from node:crypto, not the reader's MD5. Node only.

export const XLS_BOF = 0x0809
export const XLS_EOF = 0x000a
const BOUNDSHEET = 0x0085
const FILEPASS = 0x002f
const NUMBER = 0x0203
const LABEL = 0x0204
const WINDOW1 = 0x003d

// BOF substream types ([MS-XLS] §2.4.21 dt).
export const XLS_GLOBALS = 0x0005
export const XLS_WORKSHEET = 0x0010
export const XLS_CHART = 0x0020

function u16(value: number): Uint8Array {
	return Uint8Array.of(value & 0xff, (value >>> 8) & 0xff)
}

function concat(...parts: readonly Uint8Array[]): Uint8Array {
	const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
	let at = 0

	for (const part of parts) {
		out.set(part, at)
		at += part.length
	}

	return out
}

export function xlsRecord(id: number, data: Uint8Array = new Uint8Array(0)): Uint8Array {
	return concat(u16(id), u16(data.length), data)
}

// BOF of a BIFF8 substream of type `dt`.
export function xlsBof(dt: number): Uint8Array {
	return xlsRecord(XLS_BOF, concat(u16(0x0600), u16(dt), u16(0x0dbb), u16(0x07cc), new Uint8Array(8)))
}

export function xlsNumber(row: number, col: number, value: number): Uint8Array {
	const data = new Uint8Array(14)
	const view = new DataView(data.buffer)

	view.setUint16(0, row, true)
	view.setUint16(2, col, true)
	view.setUint16(4, 0x0f, true)
	view.setFloat64(6, value, true)

	return xlsRecord(NUMBER, data)
}

// A LABEL of compressed (single-byte) characters.
export function xlsLabel(row: number, col: number, text: string): Uint8Array {
	return xlsRecord(
		LABEL,
		concat(
			u16(row),
			u16(col),
			u16(0x0f),
			u16(text.length),
			Uint8Array.of(0),
			Uint8Array.from(text, c => c.charCodeAt(0))
		)
	)
}

// A chart substream (an embedded chart inside a worksheet, or a chart sheet) holding `records`.
export function xlsChart(records: readonly Uint8Array[]): Uint8Array {
	return concat(xlsBof(XLS_CHART), ...records, xlsRecord(XLS_EOF))
}

export interface XlsSheetSpec {
	readonly name: string
	// The substream's records between its BOF and EOF.
	readonly records: readonly Uint8Array[]
	// BOF dt (default a worksheet); BoundSheet8.dt follows it (2 for a chart sheet).
	readonly dt?: number
}

// A Workbook stream: globals (after BOF: `filePass` when given, then WINDOW1 and a BoundSheet8 per sheet), EOF, then
// each sheet's substream at the offset its BoundSheet8 names.
export function xlsWorkbookStream(sheets: readonly XlsSheetSpec[], filePass?: Uint8Array): Uint8Array {
	const boundSheet = (sheet: XlsSheetSpec, position: number): Uint8Array =>
		xlsRecord(
			BOUNDSHEET,
			concat(
				Uint8Array.of(position & 0xff, (position >>> 8) & 0xff, (position >>> 16) & 0xff, position >>> 24),
				Uint8Array.of(0, sheet.dt === XLS_CHART ? 2 : 0),
				Uint8Array.of(sheet.name.length, 0),
				Uint8Array.from(sheet.name, c => c.charCodeAt(0))
			)
		)
	const head = [
		xlsBof(XLS_GLOBALS),
		...(filePass === undefined ? [] : [xlsRecord(FILEPASS, filePass)]),
		xlsRecord(WINDOW1, new Uint8Array(18))
	]
	const bodies = sheets.map(sheet => concat(xlsBof(sheet.dt ?? XLS_WORKSHEET), ...sheet.records, xlsRecord(XLS_EOF)))
	const globalsLength =
		head.reduce((sum, part) => sum + part.length, 0) +
		sheets.reduce((sum, sheet) => sum + boundSheet(sheet, 0).length, 0) +
		xlsRecord(XLS_EOF).length
	let position = globalsLength
	const bound = sheets.map((sheet, k) => {
		const record = boundSheet(sheet, position)

		position += bodies[k]?.length ?? 0

		return record
	})

	return concat(...head, ...bound, xlsRecord(XLS_EOF), ...bodies)
}

// A compound file holding `workbook` as its Workbook stream (and any extra streams).
export function xlsFile(
	workbook: Uint8Array,
	extra: readonly { readonly path: readonly string[]; readonly bytes: Uint8Array }[] = []
): Promise<Uint8Array> {
	return cfbBytes([{ path: ["Workbook"], bytes: workbook }, ...extra])
}

export type XlsScheme = "xor" | "rc4" | "cryptoApiRc4"

export interface XlsEncryptOptions {
	readonly scheme: XlsScheme
	readonly password: string
	// CryptoAPI KeySize in bits (0 stands for 40).
	readonly keySize?: number
}

// Records [MS-XLS] §2.2.10 leaves in the clear: BOF, FilePass, UsrExcl, FileLock, InterfaceHdr, RRDInfo, RRDHead.
const CLEAR = new Set([XLS_BOF, FILEPASS, 0x0194, 0x0195, 0x00e1, 0x0196, 0x0138])

// CreateXorArray_Method1 ([MS-OFFCRYPTO] §2.3.7.2), transcribed for the encrypting side.
function xorArray(password: Uint8Array, key: number): Uint8Array {
	const pad = [0xbb, 0xff, 0xff, 0xba, 0xff, 0xff, 0xb9, 0x80, 0x00, 0xbe, 0x0f, 0x00, 0xbf, 0x0f, 0x00]
	const ror = (a: number, b: number): number => {
		const v = (a ^ b) & 0xff

		return ((v >> 1) | (v << 7)) & 0xff
	}
	const out = new Uint8Array(16)
	let index = password.length

	if (index % 2 === 1) {
		out[index] = ror(pad[0] ?? 0, key >> 8)
		index -= 1
		out[index] = ror(password[password.length - 1] ?? 0, key & 0xff)
	}

	while (index > 0) {
		index -= 1
		out[index] = ror(password[index] ?? 0, key >> 8)
		index -= 1
		out[index] = ror(password[index] ?? 0, key & 0xff)
	}

	for (let at = 15, padIndex = 15 - password.length; padIndex > 0;) {
		out[at--] = ror(pad[padIndex--] ?? 0, key >> 8)
		out[at--] = ror(pad[padIndex--] ?? 0, key & 0xff)
	}

	return out
}

// The FilePass record's data for `options` (wEncryptionType, then the scheme's key and verifier).
function filePassData(options: XlsEncryptOptions): Uint8Array {
	if (options.scheme === "xor") {
		const verifier = legacyPasswordVerifier(Uint8Array.from(options.password, c => c.charCodeAt(0) & 0xff))

		return concat(u16(0), u16(verifier >>> 16), u16(verifier & 0xffff))
	}

	return concat(
		u16(1),
		encryptionHeader({
			scheme: options.scheme,
			password: options.password,
			...(options.keySize !== undefined && { keySize: options.keySize })
		})
	)
}

// The bytes of `stream` that are encrypted, by stream offset: every record's data but the records left in the clear
// and BoundSheet8's lbPlyPos. `size` is the record's data length (XOR's index uses it).
function forEachEncrypted(stream: Uint8Array, visit: (at: number, size: number) => void): void {
	for (let at = 0; at + 4 <= stream.length;) {
		const id = (stream[at] ?? 0) | ((stream[at + 1] ?? 0) << 8)
		const size = (stream[at + 2] ?? 0) | ((stream[at + 3] ?? 0) << 8)
		const data = at + 4

		if (!CLEAR.has(id)) for (let k = id === BOUNDSHEET ? 4 : 0; k < size; k++) visit(data + k, size)

		at = data + size
	}
}

// `sheets` as an encrypted Workbook stream, and the plain one it was made from (identical but the encrypted bytes).
export function encryptedXlsStream(
	sheets: readonly XlsSheetSpec[],
	options: XlsEncryptOptions
): { readonly plain: Uint8Array; readonly encrypted: Uint8Array } {
	const plain = xlsWorkbookStream(sheets, filePassData(options))
	const encrypted = plain.slice()

	if (options.scheme === "xor") {
		const bytes = Uint8Array.from(options.password, c => c.charCodeAt(0) & 0xff)
		const array = xorArray(bytes, legacyPasswordVerifier(bytes) >>> 16)

		// XOR Data Transformation Method 1 (§2.3.7.7 for the decrypting side): rotate left 5, then XOR with the array at
		// (stream offset + record size) mod 16.
		forEachEncrypted(plain, (at, size) => {
			const value = plain[at] ?? 0

			encrypted[at] = (((value << 5) | (value >>> 3)) & 0xff) ^ (array[(at + size) & 15] ?? 0)
		})

		return { plain, encrypted }
	}

	// RC4: the key stream runs over the whole stream, rekeyed every 1024 bytes ([MS-XLS] §2.2.10); bytes left in the
	// clear still use their place in it.
	const keyOf = options.scheme === "rc4" ? rc4Keys(options.password) : cryptoApiKeys(options.password, options.keySize ?? 128)
	const stream = new Uint8Array(plain.length)

	for (let at = 0; at < plain.length; at += 1024)
		new Rc4(keyOf(at / 1024)).apply(plain, stream, at, at, Math.min(1024, plain.length - at))

	forEachEncrypted(plain, at => {
		encrypted[at] = stream[at] ?? 0
	})

	return { plain, encrypted }
}
