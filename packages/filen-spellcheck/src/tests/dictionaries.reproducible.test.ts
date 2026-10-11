import { createHash } from "node:crypto"
import { mkdtempSync, readdirSync, readFileSync, rmSync } from "node:fs"
import { tmpdir } from "node:os"
import { join, relative, sep } from "node:path"
import { describe, expect, it } from "vitest"
import { DICTIONARIES_DIR, ensureDictionaries, STAMP, type Manifest } from "../../scripts/build-dictionaries"

// Opt-in: downloads the pinned upstream files (cached under node_modules) and builds everything from scratch.
const ENABLED = process.env["FILEN_SPELLCHECK_DICTIONARIES"] === "1"

function sha256(bytes: Uint8Array): string {
	return createHash("sha256").update(bytes).digest("hex")
}

describe.skipIf(!ENABLED)("dictionaries rebuild (FILEN_SPELLCHECK_DICTIONARIES=1)", () => {
	it("reproduces every file the manifest expects, byte for byte", { timeout: 600_000 }, async () => {
		const manifestBytes = readFileSync(join(DICTIONARIES_DIR, "manifest.json"))
		const manifest = JSON.parse(manifestBytes.toString("utf8")) as Manifest
		const out = mkdtempSync(join(tmpdir(), "filen-spellcheck-dictionaries-"))

		try {
			expect(await ensureDictionaries({ offline: false, outDir: out })).toBe(true)

			const expected = manifest.dictionaries.flatMap(entry => [entry.aff, entry.dic])
			const files = readdirSync(out, { recursive: true, withFileTypes: true })
				.filter(entry => entry.isFile())
				.map(entry => relative(out, join(entry.parentPath, entry.name)).split(sep).join("/"))

			expect(files.sort()).toEqual([STAMP, ...expected.map(file => file.file)].sort())
			expect(readFileSync(join(out, STAMP), "utf8")).toBe(sha256(manifestBytes))

			for (const file of expected) {
				const bytes = readFileSync(join(out, file.file))

				expect([bytes.length, sha256(bytes)], file.file).toEqual([file.bytes, file.sha256])
			}

			// A second run trusts the stamp.
			expect(await ensureDictionaries({ offline: true, outDir: out })).toBe(false)
		} finally {
			rmSync(out, { recursive: true, force: true })
		}
	})
})
