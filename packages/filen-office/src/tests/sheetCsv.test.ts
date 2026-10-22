import { describe, expect, it } from "vitest"
import { canEncodeWindows1252, parseCsvFile, serializeCsv } from "../core/sheet/index"

const encoder = new TextEncoder()

function hexBytes(hex: string): Uint8Array {
	const bytes = new Uint8Array(hex.length / 2)

	for (let i = 0; i < bytes.length; i++) {
		bytes[i] = Number.parseInt(hex.slice(i * 2, i * 2 + 2), 16)
	}

	return bytes
}

function utf16Bytes(text: string, littleEndian: boolean, bom: readonly number[]): Uint8Array {
	const body = new Uint8Array(text.length * 2)

	for (let i = 0; i < text.length; i++) {
		const code = text.charCodeAt(i)
		const high = (code >> 8) & 0xff
		const low = code & 0xff

		if (littleEndian) {
			body[i * 2] = low
			body[i * 2 + 1] = high
		} else {
			body[i * 2] = high
			body[i * 2 + 1] = low
		}
	}

	const out = new Uint8Array(bom.length + body.length)

	out.set(bom)
	out.set(body, bom.length)

	return out
}

describe("CSV encodings", () => {
	it("reads and writes back a UTF-16LE file with its BOM", () => {
		const source = "name,city\r\ncafé,björk\r\n"
		const bytes = utf16Bytes(source, true, [0xff, 0xfe])
		const { rows, format } = parseCsvFile(bytes, false)

		expect(rows).toEqual([
			["name", "city"],
			["café", "björk"]
		])
		expect(format.encoding).toBe("utf-16le")
		expect(Array.from(serializeCsv(rows, format))).toEqual(Array.from(bytes))
	})

	it("reads and writes back a UTF-16BE file with its BOM", () => {
		const source = "a,b\r\n1,2\r\n"
		const bytes = utf16Bytes(source, false, [0xfe, 0xff])
		const { rows, format } = parseCsvFile(bytes, false)

		expect(format.encoding).toBe("utf-16be")
		expect(Array.from(serializeCsv(rows, format))).toEqual(Array.from(bytes))
	})

	it("round-trips a windows-1252 file byte for byte", () => {
		// "café,€" in windows-1252: é = 0xe9, € = 0x80.
		const bytes = new Uint8Array([0x63, 0x61, 0x66, 0xe9, 0x2c, 0x80])
		const { rows, format } = parseCsvFile(bytes, false)

		expect(rows).toEqual([["café", "€"]])
		expect(format.encoding).toBe("windows-1252")
		expect(Array.from(serializeCsv(rows, format))).toEqual(Array.from(bytes))
	})

	it("throws rather than fall back to UTF-8 when asked to write a character windows-1252 cannot hold", () => {
		// serializeCsv has no UTF-8 fallback: the text session refuses such an edit first (sheetSessionText.test.ts), so
		// reaching this directly bypasses that guard, which must fail loudly, not silently re-encode the file as UTF-8.
		const { rows, format } = parseCsvFile(new Uint8Array([0x63, 0x61, 0x66, 0xe9]), false)

		rows[0]?.push("日本語")

		expect(() => serializeCsv(rows, format)).toThrow()
	})

	it("keeps a genuine windows-1252 export writable across German, French and Spanish text", () => {
		// "Größe,Gewicht\nKäfer,Öl\nMüller,Fön\n" and "Name,City\nJosé,Málaga\nFrançois,Genève\n" in
		// windows-1252: accented letters sit isolated inside otherwise-ASCII words, well under the
		// CJK/Cyrillic run and ratio thresholds below.
		for (const hex of [
			"4772f6df652c476577696368740a4be46665722cd66c0a4dfc6c6c65722c46f66e0a",
			"4e616d652c436974790a4a6f73e92c4de16c6167610a4672616ee76f69732c47656ee876650a"
		]) {
			const { format } = parseCsvFile(hexBytes(hex), false)

			expect(format.encoding).toBe("windows-1252")
			expect(format.writable).toBe(true)
		}
	})

	it("opens read-only when non-UTF-8 bytes are CJK or Cyrillic text that windows-1252 decodes without error", () => {
		// Excel's default export encoding for "name,city" + one data row in a CJK/Cyrillic locale: GBK,
		// EUC-KR, windows-1251 and Shift-JIS. None of these bytes are undefined in windows-1252, so only
		// the run/ratio heuristic — not hasWin1252UndefinedByte — tells them apart from genuine Western text.
		for (const hex of [
			"d0d5c3fb2cb3c7cad00ad5c5c8fd2cb1b1bea90a",
			"c0ccb8a72cb5b5bdc30ab1e8c3b6bcf62cbcadbfef0a",
			"c8ecff2cc3eef0eee40ac8e2e0ed2ccceef1eae2e00a",
			"96bc914f2c93738e730a936392862c938c8b9e0a"
		]) {
			const { format } = parseCsvFile(hexBytes(hex), false)

			expect(format.encoding).toBe("windows-1252")
			expect(format.writable).toBe(false)
		}
	})

	it("keeps a UTF-8 file's BOM presence as read", () => {
		const withBom = parseCsvFile(encoder.encode("﻿a,b\n"), false)
		const withoutBom = parseCsvFile(encoder.encode("a,b\n"), false)

		expect(withBom.format.bom).toBe(true)
		expect(withoutBom.format.bom).toBe(false)
		expect(Array.from(serializeCsv(withBom.rows, withBom.format).subarray(0, 3))).toEqual([0xef, 0xbb, 0xbf])
		expect(serializeCsv(withoutBom.rows, withoutBom.format)[0]).not.toBe(0xef)
	})

	it("canEncodeWindows1252 accepts Latin-1 text and the windows-1252 special block, rejects Cyrillic, Han and Ł", () => {
		expect(canEncodeWindows1252("Müller — 100%")).toBe(true)
		expect(canEncodeWindows1252("Łódź")).toBe(false)
		expect(canEncodeWindows1252("шт")).toBe(false)
		expect(canEncodeWindows1252("新")).toBe(false)
	})
})

describe("CSV line endings", () => {
	it("preserves lone-CR line separators and a trailing CR", () => {
		const source = "a,b\rc,d\r"
		const { rows, format } = parseCsvFile(encoder.encode(source), false)

		expect(rows).toEqual([
			["a", "b"],
			["c", "d"]
		])
		expect(format.lineSeparator).toBe("\r")
		expect(format.trailingNewline).toBe(true)
		expect(new TextDecoder().decode(serializeCsv(rows, format))).toBe(source)
	})

	it("keeps a file with no trailing newline from gaining one", () => {
		const source = "a,b\nc,d"
		const { rows, format } = parseCsvFile(encoder.encode(source), false)

		expect(format.trailingNewline).toBe(false)
		expect(new TextDecoder().decode(serializeCsv(rows, format))).toBe(source)
	})
})

describe("CSV baseline parsing", () => {
	it("detects the delimiter among comma, semicolon, tab and pipe", () => {
		expect(parseCsvFile(encoder.encode("a|b|c\n1|2|3\n"), false).format.delimiter).toBe("|")
		expect(parseCsvFile(encoder.encode("a;b;c\n1;2;3\n"), false).format.delimiter).toBe(";")
	})

	it("keeps embedded newlines and quotes inside a quoted field", () => {
		const { rows } = parseCsvFile(encoder.encode('a,"line1\nline2",c\n'), false)

		expect(rows).toEqual([["a", "line1\nline2", "c"]])
	})

	it('keeps a doubled quote inside a quoted field as one literal "', () => {
		const { rows } = parseCsvFile(encoder.encode('a,"say ""hi""",c\n'), false)

		expect(rows).toEqual([["a", 'say "hi"', "c"]])
	})

	it("keeps ragged rows as given, padding nothing", () => {
		const { rows } = parseCsvFile(encoder.encode("a,b,c\n1,2\n"), false)

		expect(rows).toEqual([
			["a", "b", "c"],
			["1", "2"]
		])
	})

	it("keeps values starting with =, +, -, @ exactly as typed, with no formula-injection escaping", () => {
		const source = "=SUM(A1),+1,-1,@cmd\n"
		const { rows, format } = parseCsvFile(encoder.encode(source), false)

		expect(rows).toEqual([["=SUM(A1)", "+1", "-1", "@cmd"]])
		expect(new TextDecoder().decode(serializeCsv(rows, format))).toBe(source)
	})
})
