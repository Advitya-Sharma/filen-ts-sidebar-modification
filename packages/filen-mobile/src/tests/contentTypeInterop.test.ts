import { describe, it, expect } from "vitest"
import { createRequire } from "node:module"
import { readFileSync } from "node:fs"
import path from "node:path"
import { parseBuffer } from "music-metadata"

// Regression guard for the "music-metadata can't parse any audio file" bug (Hermes only).
//
// music-metadata 11 default-imported content-type 2 (`import ContentType from "content-type"`), a CJS
// module declaring `__esModule: true` with no `default` export. Metro/Babel's `_interopRequireDefault`
// bound that import to `undefined`, `ContentType.parse(...)` threw inside the MIME→parser lookup, and
// every file failed with "Guessed MIME-type not supported". Node's (and vitest's) interop hides it.
//
// music-metadata 12 imports `parse` BY NAME from content-type 3 (ESM), which Babel compiles to a plain
// `require("content-type").parse` — no default interop, so no patch. These tests fail if either side
// regresses to the broken shape.
const require = createRequire(import.meta.url)
const musicMetadataEntry = require.resolve("music-metadata")
const parserFactorySource = readFileSync(path.join(path.dirname(musicMetadataEntry), "ParserFactory.js"), "utf8")

// 44-byte PCM WAV header + 4 silent samples (8 kHz, mono, 16 bit).
function silentWav(): Uint8Array {
	const dataSize = 8
	const view = new DataView(new ArrayBuffer(44 + dataSize))
	const ascii = (offset: number, text: string) => {
		for (let i = 0; i < text.length; i++) {
			view.setUint8(offset + i, text.charCodeAt(i))
		}
	}

	ascii(0, "RIFF")
	view.setUint32(4, 36 + dataSize, true)
	ascii(8, "WAVE")
	ascii(12, "fmt ")
	view.setUint32(16, 16, true)
	view.setUint16(20, 1, true)
	view.setUint16(22, 1, true)
	view.setUint32(24, 8000, true)
	view.setUint32(28, 16000, true)
	view.setUint16(32, 2, true)
	view.setUint16(34, 16, true)
	ascii(36, "data")
	view.setUint32(40, dataSize, true)

	return new Uint8Array(view.buffer)
}

describe("content-type import interop (Metro/Babel → music-metadata)", () => {
	it("music-metadata imports content-type's parse by name, never as a default import", () => {
		expect(parserFactorySource).toMatch(/import\s*\{\s*parse(\s+as\s+\w+)?\s*\}\s*from\s*["']content-type["']/)
		expect(parserFactorySource).not.toMatch(/import\s+\w+\s*(,\s*\{[^}]*\})?\s*from\s*["']content-type["']/)
	})

	it("the content-type music-metadata resolves exposes a callable named parse()", () => {
		const contentType = createRequire(musicMetadataEntry)("content-type") as { parse?: (input: string) => { type: string } }

		expect(typeof contentType.parse).toBe("function")
		expect(contentType.parse?.("audio/mpeg; charset=binary").type).toBe("audio/mpeg")
	})

	it("selects the parser from the MIME type (the path that broke) and parses a WAV", async () => {
		const metadata = await parseBuffer(silentWav(), { mimeType: "audio/wav", size: 52 })

		expect(metadata.format.container).toBe("WAVE")
		expect(metadata.format.sampleRate).toBe(8000)
	})
})
