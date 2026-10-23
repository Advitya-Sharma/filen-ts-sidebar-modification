import { describe, expect, it } from "vitest"
import { constantTimeEqual, md5, Rc4 } from "../core/crypto/legacy"

function hex(bytes: Uint8Array): string {
	return [...bytes].map(byte => byte.toString(16).padStart(2, "0")).join("")
}

function bytesOf(hexText: string): Uint8Array {
	return Uint8Array.from(hexText.match(/../g) ?? [], pair => parseInt(pair, 16))
}

const ascii = (text: string): Uint8Array => Uint8Array.from(text, char => char.charCodeAt(0))

// RFC 1321 Appendix A.5, the test suite.
const RFC1321: readonly (readonly [string, string])[] = [
	["", "d41d8cd98f00b204e9800998ecf8427e"],
	["a", "0cc175b9c0f1b6a831c399e269772661"],
	["abc", "900150983cd24fb0d6963f7d28e17f72"],
	["message digest", "f96b697d7cb7938d525a2f31aaf161d0"],
	["abcdefghijklmnopqrstuvwxyz", "c3fcd3d76192e4007dfb496cca67e13b"],
	["ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789", "d174ab98d277d9f5a5611c2c9f419d9f"],
	["12345678901234567890123456789012345678901234567890123456789012345678901234567890", "57edf4a22be3c955ac49da2e2107b67a"]
]

// RFC 6229 §2, verbatim: key, then "DEC offset HEX offset: 16 key-stream bytes" lines.
const RFC6229: readonly (readonly [key: string, lines: readonly string[]])[] = [
	[
		"0102030405",
		[
			"DEC    0 HEX    0:  b2 39 63 05  f0 3d c0 27   cc c3 52 4a  0a 11 18 a8",
			"DEC   16 HEX   10:  69 82 94 4f  18 fc 82 d5   89 c4 03 a4  7a 0d 09 19",
			"DEC  240 HEX   f0:  28 cb 11 32  c9 6c e2 86   42 1d ca ad  b8 b6 9e ae",
			"DEC  256 HEX  100:  1c fc f6 2b  03 ed db 64   1d 77 df cf  7f 8d 8c 93",
			"DEC  496 HEX  1f0:  42 b7 d0 cd  d9 18 a8 a3   3d d5 17 81  c8 1f 40 41",
			"DEC  512 HEX  200:  64 59 84 44  32 a7 da 92   3c fb 3e b4  98 06 61 f6",
			"DEC  752 HEX  2f0:  ec 10 32 7b  de 2b ee fd   18 f9 27 76  80 45 7e 22",
			"DEC  768 HEX  300:  eb 62 63 8d  4f 0b a1 fe   9f ca 20 e0  5b f8 ff 2b",
			"DEC 1008 HEX  3f0:  45 12 90 48  e6 a0 ed 0b   56 b4 90 33  8f 07 8d a5",
			"DEC 1024 HEX  400:  30 ab bc c7  c2 0b 01 60   9f 23 ee 2d  5f 6b b7 df",
			"DEC 1520 HEX  5f0:  32 94 f7 44  d8 f9 79 05   07 e7 0f 62  e5 bb ce ea",
			"DEC 1536 HEX  600:  d8 72 9d b4  18 82 25 9b   ee 4f 82 53  25 f5 a1 30",
			"DEC 2032 HEX  7f0:  1e b1 4a 0c  13 b3 bf 47   fa 2a 0b a9  3a d4 5b 8b",
			"DEC 2048 HEX  800:  cc 58 2f 8b  a9 f2 65 e2   b1 be 91 12  e9 75 d2 d7",
			"DEC 3056 HEX  bf0:  f2 e3 0f 9b  d1 02 ec bf   75 aa ad e9  bc 35 c4 3c",
			"DEC 3072 HEX  c00:  ec 0e 11 c4  79 dc 32 9d   c8 da 79 68  fe 96 56 81",
			"DEC 4080 HEX  ff0:  06 83 26 a2  11 84 16 d2   1f 9d 04 b2  cd 1c a0 50",
			"DEC 4096 HEX 1000:  ff 25 b5 89  95 99 67 07   e5 1f bd f0  8b 34 d8 75"
		]
	],
	[
		"0102030405060708090a0b0c0d0e0f10",
		[
			"DEC    0 HEX    0:  9a c7 cc 9a  60 9d 1e f7   b2 93 28 99  cd e4 1b 97",
			"DEC   16 HEX   10:  52 48 c4 95  90 14 12 6a   6e 8a 84 f1  1d 1a 9e 1c",
			"DEC  240 HEX   f0:  06 59 02 e4  b6 20 f6 cc   36 c8 58 9f  66 43 2f 2b",
			"DEC  256 HEX  100:  d3 9d 56 6b  c6 bc e3 01   07 68 15 15  49 f3 87 3f",
			"DEC  496 HEX  1f0:  b6 d1 e6 c4  a5 e4 77 1c   ad 79 53 8d  f2 95 fb 11",
			"DEC  512 HEX  200:  c6 8c 1d 5c  55 9a 97 41   23 df 1d bc  52 a4 3b 89",
			"DEC  752 HEX  2f0:  c5 ec f8 8d  e8 97 fd 57   fe d3 01 70  1b 82 a2 59",
			"DEC  768 HEX  300:  ec cb e1 3d  e1 fc c9 1c   11 a0 b2 6c  0b c8 fa 4d",
			"DEC 1008 HEX  3f0:  e7 a7 25 74  f8 78 2a e2   6a ab cf 9e  bc d6 60 65",
			"DEC 1024 HEX  400:  bd f0 32 4e  60 83 dc c6   d3 ce dd 3c  a8 c5 3c 16",
			"DEC 1520 HEX  5f0:  b4 01 10 c4  19 0b 56 22   a9 61 16 b0  01 7e d2 97",
			"DEC 1536 HEX  600:  ff a0 b5 14  64 7e c0 4f   63 06 b8 92  ae 66 11 81",
			"DEC 2032 HEX  7f0:  d0 3d 1b c0  3c d3 3d 70   df f9 fa 5d  71 96 3e bd",
			"DEC 2048 HEX  800:  8a 44 12 64  11 ea a7 8b   d5 1e 8d 87  a8 87 9b f5",
			"DEC 3056 HEX  bf0:  fa be b7 60  28 ad e2 d0   e4 87 22 e4  6c 46 15 a3",
			"DEC 3072 HEX  c00:  c0 5d 88 ab  d5 03 57 f9   35 a6 3c 59  ee 53 76 23",
			"DEC 4080 HEX  ff0:  ff 38 26 5c  16 42 c1 ab   e8 d3 c2 fe  5e 57 2b f8",
			"DEC 4096 HEX 1000:  a3 6a 4c 30  1a e8 ac 13   61 0c cb c1  22 56 ca cc"
		]
	],
	[
		"0102030405060708090a0b0c0d0e0f101112131415161718191a1b1c1d1e1f20",
		[
			"DEC    0 HEX    0:  ea a6 bd 25  88 0b f9 3d   3f 5d 1e 4c  a2 61 1d 91",
			"DEC   16 HEX   10:  cf a4 5c 9f  7e 71 4b 54   bd fa 80 02  7c b1 43 80",
			"DEC  240 HEX   f0:  11 4a e3 44  de d7 1b 35   f2 e6 0f eb  ad 72 7f d8",
			"DEC  256 HEX  100:  02 e1 e7 05  6b 0f 62 39   00 49 64 22  94 3e 97 b6",
			"DEC  496 HEX  1f0:  91 cb 93 c7  87 96 4e 10   d9 52 7d 99  9c 6f 93 6b",
			"DEC  512 HEX  200:  49 b1 8b 42  f8 e8 36 7c   be b5 ef 10  4b a1 c7 cd",
			"DEC  752 HEX  2f0:  87 08 4b 3b  a7 00 ba de   95 56 10 67  27 45 b3 74",
			"DEC  768 HEX  300:  e7 a7 b9 e9  ec 54 0d 5f   f4 3b db 12  79 2d 1b 35",
			"DEC 1008 HEX  3f0:  c7 99 b5 96  73 8f 6b 01   8c 76 c7 4b  17 59 bd 90",
			"DEC 1024 HEX  400:  7f ec 5b fd  9f 9b 89 ce   65 48 30 90  92 d7 e9 58",
			"DEC 1520 HEX  5f0:  40 f2 50 b2  6d 1f 09 6a   4a fd 4c 34  0a 58 88 15",
			"DEC 1536 HEX  600:  3e 34 13 5c  79 db 01 02   00 76 76 51  cf 26 30 73",
			"DEC 2032 HEX  7f0:  f6 56 ab cc  f8 8d d8 27   02 7b 2c e9  17 d4 64 ec",
			"DEC 2048 HEX  800:  18 b6 25 03  bf bc 07 7f   ba bb 98 f2  0d 98 ab 34",
			"DEC 3056 HEX  bf0:  8a ed 95 ee  5b 0d cb fb   ef 4e b2 1d  3a 3f 52 f9",
			"DEC 3072 HEX  c00:  62 5a 1a b0  0e e3 9a 53   27 34 6b dd  b0 1a 9c 18",
			"DEC 4080 HEX  ff0:  a1 3a 7c 79  c7 e1 19 b5   ab 02 96 ab  28 c3 00 b9",
			"DEC 4096 HEX 1000:  f3 e4 c0 a2  e0 2d 1d 01   f7 f0 a7 46  18 af 2b 48"
		]
	],
	[
		"833222772a",
		[
			"DEC    0 HEX    0:  80 ad 97 bd  c9 73 df 8a   2e 87 9e 92  a4 97 ef da",
			"DEC   16 HEX   10:  20 f0 60 c2  f2 e5 12 65   01 d3 d4 fe  a1 0d 5f c0",
			"DEC  240 HEX   f0:  fa a1 48 e9  90 46 18 1f   ec 6b 20 85  f3 b2 0e d9",
			"DEC  256 HEX  100:  f0 da f5 ba  b3 d5 96 83   98 57 84 6f  73 fb fe 5a",
			"DEC  496 HEX  1f0:  1c 7e 2f c4  63 92 32 fe   29 75 84 b2  96 99 6b c8",
			"DEC  512 HEX  200:  3d b9 b2 49  40 6c c8 ed   ff ac 55 cc  d3 22 ba 12",
			"DEC  752 HEX  2f0:  e4 f9 f7 e0  06 61 54 bb   d1 25 b7 45  56 9b c8 97",
			"DEC  768 HEX  300:  75 d5 ef 26  2b 44 c4 1a   9c f6 3a e1  45 68 e1 b9",
			"DEC 1008 HEX  3f0:  6d a4 53 db  f8 1e 82 33   4a 3d 88 66  cb 50 a1 e3",
			"DEC 1024 HEX  400:  78 28 d0 74  11 9c ab 5c   22 b2 94 d7  a9 bf a0 bb",
			"DEC 1520 HEX  5f0:  ad b8 9c ea  9a 15 fb e6   17 29 5b d0  4b 8c a0 5c",
			"DEC 1536 HEX  600:  62 51 d8 7f  d4 aa ae 9a   7e 4a d5 c2  17 d3 f3 00",
			"DEC 2032 HEX  7f0:  e7 11 9b d6  dd 9b 22 af   e8 f8 95 85  43 28 81 e2",
			"DEC 2048 HEX  800:  78 5b 60 fd  7e c4 e9 fc   b6 54 5f 35  0d 66 0f ab",
			"DEC 3056 HEX  bf0:  af ec c0 37  fd b7 b0 83   8e b3 d7 0b  cd 26 83 82",
			"DEC 3072 HEX  c00:  db c1 a7 b4  9d 57 35 8c   c9 fa 6d 61  d7 3b 7c f0",
			"DEC 4080 HEX  ff0:  63 49 d1 26  a3 7a fc ba   89 79 4f 98  04 91 4f dc",
			"DEC 4096 HEX 1000:  bf 42 c3 01  8c 2f 7c 66   bf de 52 49  75 76 81 15"
		]
	]
]

describe("MD5", () => {
	it("gives the RFC 1321 test suite digests", () => {
		expect(RFC1321.map(([text]) => hex(md5(ascii(text))))).toEqual(RFC1321.map(([, digest]) => digest))
	})

	it("hashes parts as their concatenation, across block and padding boundaries", () => {
		const text = "12345678901234567890123456789012345678901234567890123456789012345678901234567890"

		for (const cut of [0, 1, 55, 56, 63, 64, 65, 80]) {
			expect(hex(md5(ascii(text.slice(0, cut)), ascii(text.slice(cut))))).toBe("57edf4a22be3c955ac49da2e2107b67a")
		}
	})

	it("leaves its input untouched", () => {
		const input = ascii("abc")

		md5(input)

		expect(hex(input)).toBe("616263")
	})
})

describe("RC4", () => {
	it("gives the RFC 6229 key streams, read in one call", () => {
		for (const [key, lines] of RFC6229) {
			const stream = new Uint8Array(4112)

			new Rc4(bytesOf(key)).apply(stream, stream)

			for (const line of lines) {
				const [, offset, bytes] = /^DEC\s+(\d+) HEX\s+\w+:\s+(.*)$/.exec(line) ?? []
				const at = Number(offset)

				expect([key, at, hex(stream.subarray(at, at + 16))]).toEqual([key, at, (bytes ?? "").replace(/\s+/g, "")])
			}
		}
	})

	it("continues the key stream across calls of any size", () => {
		const [key] = RFC6229[1] ?? ["", []]
		const whole = new Uint8Array(1000)
		const pieces = new Uint8Array(1000)
		const cipher = new Rc4(bytesOf(key))

		new Rc4(bytesOf(key)).apply(whole, whole)

		for (let at = 0, size = 1; at < 1000; at += size, size = size * 3 + 1)
			cipher.apply(pieces, pieces, at, at, Math.min(size, 1000 - at))

		expect(hex(pieces)).toBe(hex(whole))
	})

	it("decrypts what it encrypted, into another array", () => {
		const plain = ascii("Attack at dawn, with a 512-byte rekey")
		const sealed = new Uint8Array(plain.length)
		const opened = new Uint8Array(plain.length)

		new Rc4(ascii("Secret")).apply(plain, sealed)
		new Rc4(ascii("Secret")).apply(sealed, opened)

		expect([hex(sealed) === hex(plain), hex(opened)]).toEqual([false, hex(plain)])
	})

	it("refuses an empty or over-long key", () => {
		expect(() => new Rc4(new Uint8Array(0))).toThrow(RangeError)
		expect(() => new Rc4(new Uint8Array(257))).toThrow(RangeError)
	})
})

describe("constantTimeEqual", () => {
	it("compares by content and length", () => {
		expect([
			constantTimeEqual(Uint8Array.of(1, 2, 3), Uint8Array.of(1, 2, 3)),
			constantTimeEqual(Uint8Array.of(1, 2, 3), Uint8Array.of(1, 2, 4)),
			constantTimeEqual(Uint8Array.of(1, 2), Uint8Array.of(1, 2, 3))
		]).toEqual([true, false, false])
	})
})
