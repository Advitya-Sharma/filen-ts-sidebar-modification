import { existsSync, readFileSync, readdirSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { CfbError, CfbFile, cfbBytes, isCompoundFile, type CfbEntry, type CfbWriteEntry } from "@filen/office/core/cfb"
import { bytesSource } from "@filen/office/core/opc"
import { CORPUS_DIR } from "./corpusDir"

// Every compound file of the corpus (legacy .doc/.xls/.ppt and the encrypted OOXML packages): each stream
// reads, and our writer's copy of the file (version 3 and 4) reads back the same tree and bytes.

const fetched = existsSync(CORPUS_DIR)

// Streams of corpus files that are malformed on purpose (POI regression samples); "<file> <path>" -> code.
const DAMAGED = new Map([["xls/poi__61300.xls \u0005SummaryInformation", "corrupt"]])

// The encrypted OOXML samples: EncryptionInfo version ([MS-OFFCRYPTO] §2.3.4.5 / §2.3.4.10) and whether
// the \u0006DataSpaces storage is present.
const ENCRYPTED: Record<string, { readonly version: string; readonly dataSpaces: boolean }> = {
	"docx/lo__Encrypted_LO_Standard_abc.docx": { version: "3.2", dataSpaces: false },
	"docx/oxsdk__document_with_password.docx": { version: "3.2", dataSpaces: true },
	"docx/poi__bug53475-password-is-pass.docx": { version: "4.4", dataSpaces: true },
	"docx/poi__bug53475-password-is-solrcell.docx": { version: "4.2", dataSpaces: true },
	"xlsx/poi__58616.xlsx": { version: "4.2", dataSpaces: true },
	"xlsx/poi__protected_passtika.xlsx": { version: "4.2", dataSpaces: true }
}

function compoundFiles(): string[] {
	if (!fetched) return []

	const out: string[] = []

	for (const family of readdirSync(CORPUS_DIR, { withFileTypes: true })) {
		if (!family.isDirectory()) continue

		for (const name of readdirSync(join(CORPUS_DIR, family.name))) {
			const path = `${family.name}/${name}`
			const head = new Uint8Array(8)
			const bytes = readFileSync(join(CORPUS_DIR, path))

			head.set(bytes.subarray(0, 8))

			if (isCompoundFile(head)) out.push(path)
		}
	}

	return out.sort()
}

function load(path: string): Uint8Array {
	const bytes = readFileSync(join(CORPUS_DIR, path))

	return new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength)
}

function same(a: Uint8Array, b: Uint8Array): boolean {
	return Buffer.from(a.buffer, a.byteOffset, a.byteLength).equals(Buffer.from(b.buffer, b.byteOffset, b.byteLength))
}

const key = (entry: CfbEntry): string => entry.path.join("/")

describe.skipIf(!fetched)("compound files of the corpus", () => {
	const files = compoundFiles()

	it("finds the legacy and encrypted samples", () => {
		expect(files.length).toBeGreaterThan(300)

		for (const extension of [".doc", ".xls", ".ppt"]) expect(files.some(path => path.endsWith(extension))).toBe(true)
		for (const path of Object.keys(ENCRYPTED)) expect(files).toContain(path)
	})

	it("reads every stream, and the writer's version 3 and 4 copies hold the same", { timeout: 600_000 }, async () => {
		const problems: string[] = []
		let streams = 0

		for (const path of files) {
			const file = await CfbFile.open(bytesSource(load(path)))
			const entries: CfbWriteEntry[] = []
			const kept: string[] = []

			for (const entry of file.walk()) {
				if (entry.kind === "root") continue

				if (entry.kind === "storage") {
					entries.push({ path: entry.path, storage: true, ...(entry.clsid === undefined ? {} : { clsid: entry.clsid }) })
					kept.push(key(entry))

					continue
				}

				const damaged = DAMAGED.get(`${path} ${key(entry)}`)

				try {
					const bytes = await file.read(entry)

					if (bytes.length !== entry.size) problems.push(`${path} ${key(entry)}: ${bytes.length} bytes of ${entry.size}`)
					if (damaged !== undefined) problems.push(`${path} ${key(entry)}: read, expected ${damaged}`)

					entries.push({ path: entry.path, bytes })
					kept.push(key(entry))
					streams++
				} catch (error) {
					const code = error instanceof CfbError ? error.code : `untyped ${String(error)}`

					if (code !== damaged) problems.push(`${path} ${key(entry)}: ${code}`)
				}
			}

			for (const version of [3, 4] as const) {
				const copy = await CfbFile.open(
					bytesSource(
						await cfbBytes(entries, { version, ...(file.root.clsid === undefined ? {} : { rootClsid: file.root.clsid }) })
					)
				)
				const walked = [...copy.walk()].slice(1)

				if (copy.version !== version) problems.push(`${path} v${version}: version ${copy.version}`)
				if (copy.root.clsid !== file.root.clsid) problems.push(`${path} v${version}: root CLSID ${copy.root.clsid}`)
				if (walked.map(key).join("\n") !== kept.join("\n")) problems.push(`${path} v${version}: the tree differs`)

				for (const entry of entries) {
					const ours = copy.entry(entry.path)

					if ("bytes" in entry) {
						if (!(entry.bytes instanceof Uint8Array) || !same(await copy.read(entry.path), entry.bytes))
							problems.push(`${path} v${version} ${entry.path.join("/")}: bytes differ`)
					} else if (ours?.clsid !== (entry.clsid ?? undefined)) {
						problems.push(`${path} v${version} ${entry.path.join("/")}: CLSID ${ours?.clsid}`)
					}
				}
			}
		}

		expect(problems).toEqual([])
		expect(streams).toBeGreaterThan(2_500)
	})

	it("reads the encrypted packages' streams", async () => {
		for (const [path, expected] of Object.entries(ENCRYPTED)) {
			const file = await CfbFile.open(bytesSource(load(path)))
			const info = await file.read("EncryptionInfo")
			const header = new DataView(info.buffer, info.byteOffset, info.byteLength)
			const encrypted = file.stream("EncryptedPackage")
			const size = new DataView((await encrypted.read(0, 8)).slice().buffer).getBigUint64(0, true)

			expect(`${header.getUint16(0, true)}.${header.getUint16(2, true)}`, path).toBe(expected.version)
			// StreamSize, then the plaintext's length rounded up to the cipher block.
			expect(Number(size), path).toBeLessThanOrEqual(encrypted.size - 8)
			expect(encrypted.size - 8 - Number(size), path).toBeLessThan(16)
			expect(file.entry("\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary")?.kind, path).toBe(
				expected.dataSpaces ? "stream" : undefined
			)

			if (expected.dataSpaces) expect((await file.read("\u0006dataspaces/version")).length, path).toBeGreaterThan(0)
		}
	})
})
