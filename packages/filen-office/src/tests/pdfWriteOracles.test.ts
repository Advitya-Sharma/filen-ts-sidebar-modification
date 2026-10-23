import { execFileSync } from "node:child_process"
import { existsSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { afterAll, beforeAll, describe, expect, it } from "vitest"
import { buildSyntheticPdf } from "./pdfWriteFixtures"

// External checkers of the writer's output: qpdf --check (structure, streams, encryption), veraPDF's
// PDF/UA-1 profile (tagged output) and pypdf (reads, decrypts with the password). Opt-in:
// FILEN_OFFICE_PDF_TOOLS=<dir> holding qpdf/bin/qpdf, verapdf/verapdf, pypdf and pycryptodome; skips without.

const TOOLS = process.env["FILEN_OFFICE_PDF_TOOLS"] ?? ""
const QPDF = join(TOOLS, "qpdf", "bin", "qpdf")
const VERAPDF = join(TOOLS, "verapdf", "verapdf")
const PASSWORD = "pässwörd 1"

const PYPDF = `
import sys, json
sys.path[:0] = [sys.argv[1] + "/pypdf", sys.argv[1] + "/pycryptodome"]
from pypdf import PdfReader
r = PdfReader(sys.argv[2])
out = {"encrypted": r.is_encrypted}
if r.is_encrypted:
    out["decrypt"] = int(r.decrypt(sys.argv[3]))
out["pages"] = len(r.pages)
out["title"] = r.metadata.title
out["text"] = r.pages[0].extract_text()
out["images"] = [[str(x.get_object().get("/Filter")), len(x.get_object().get_data())] for x in r.pages[0]["/Resources"]["/XObject"].values()]
print(json.dumps(out))
`

describe.skipIf(TOOLS === "" || !existsSync(QPDF) || !existsSync(VERAPDF))("PDF writer output under qpdf, veraPDF and pypdf", () => {
	let dir = ""
	let plain = ""
	let encrypted = ""

	beforeAll(async () => {
		dir = mkdtempSync(join(tmpdir(), "pdfwrite-"))
		plain = join(dir, "tagged.pdf")
		encrypted = join(dir, "encrypted.pdf")
		writeFileSync(plain, (await buildSyntheticPdf({ pages: 4, deterministic: false })).bytes)
		writeFileSync(encrypted, (await buildSyntheticPdf({ pages: 4, password: PASSWORD, deterministic: false })).bytes)
	})

	afterAll(() => {
		if (dir !== "") rmSync(dir, { recursive: true, force: true })
	})

	it("passes qpdf --check, encrypted with the password included", () => {
		expect(execFileSync(QPDF, ["--check", plain]).toString()).toContain("No syntax or stream encoding errors found")

		const checked = execFileSync(QPDF, ["--check", `--password=${PASSWORD}`, encrypted]).toString()

		expect(checked).toContain("No syntax or stream encoding errors found")
		expect(checked).toContain("file encryption method: AESv3")
		expect(() => execFileSync(QPDF, ["--check", encrypted], { stdio: "pipe" })).toThrow()
	})

	it("passes veraPDF's PDF/UA-1 profile", { timeout: 120_000 }, () => {
		const report = execFileSync(VERAPDF, ["-f", "ua1", "--format", "xml", plain], { stdio: ["ignore", "pipe", "ignore"] }).toString()

		expect(report).toMatch(/failedRules="0"/)
		expect(report).toMatch(/isCompliant="true"/)
	})

	it("is read by pypdf, which decrypts it with the password", () => {
		for (const file of [plain, encrypted]) {
			const result = JSON.parse(execFileSync("python3", ["-I", "-c", PYPDF, TOOLS, file, PASSWORD]).toString()) as Record<
				string,
				unknown
			>

			expect(result).toMatchObject({ pages: 4, title: "Synthetic PDF – Ünïcode" })
			expect(result["text"]).toContain("Καλημέρα κόσμε")
			expect(result["images"]).toEqual([
				["/DCTDecode", 927],
				["/FlateDecode", 3600],
				["/FlateDecode", 256],
				["/FlateDecode", 1728]
			])

			if (file === encrypted) expect(result).toMatchObject({ encrypted: true, decrypt: 2 })
		}
	})
})
