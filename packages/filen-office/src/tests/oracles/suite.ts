import { existsSync, readdirSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { bytesSource, OpcPackage } from "@filen/office/core/opc"
import { saveDocx } from "@filen/office/core/docx/write"
import { CORPUS_DIR } from "../corpusDir"
import { openBytes } from "../docxEditFixtures"
import { runEditScript } from "../docxWriteFixtures"

// The saves the oracles judge: the scripted edit suite over the corpus, and the fuzzers' saved outputs
// (FILEN_OFFICE_FUZZ_SAVES, the directory whose manifest.json the fuzz harness writes).

export interface Save {
	// "<corpus path>" for the edit suite, the manifest path for a fuzz save.
	readonly name: string
	// The corpus file it was edited from.
	readonly source: Uint8Array
	readonly sourceName: string
	readonly bytes: Uint8Array
	// Parts that are new or differ from the source's.
	readonly changedParts: readonly string[]
	readonly family: string
}

export function corpusFiles(): string[] {
	const files: string[] = []

	for (const extension of ["docx", "docm", "dotx", "dotm"]) {
		const directory = join(CORPUS_DIR, extension)

		if (existsSync(directory)) for (const name of readdirSync(directory).sort()) files.push(join(directory, name))
	}

	return files
}

// The scripted edit of every editable corpus file, saved; files that do not open or are gated are left out.
export async function* editSuite(): AsyncGenerator<Save> {
	for (const file of corpusFiles()) {
		const source = new Uint8Array(readFileSync(file))
		let opened

		try {
			opened = await openBytes(source)
		} catch {
			continue
		}

		if (opened.model.body === undefined || opened.context.gate.length > 0) continue

		const script = runEditScript(opened)
		const result = await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: script.sessions })

		if (result.bytes === undefined) throw new Error("No bytes")

		const name = file.slice(CORPUS_DIR.length + 1)
		const changed = await changedParts(opened.pkg.opc, await OpcPackage.open(bytesSource(result.bytes)))

		yield { name, source, sourceName: name, bytes: result.bytes, changedParts: changed, family: "edit-suite" }
	}
}

export interface FuzzManifestEntry {
	readonly source: string
	readonly family: string
	readonly path: string
}

export const FUZZ_SAVES = process.env["FILEN_OFFICE_FUZZ_SAVES"] ?? ""

// The fuzz saves' manifest entries, or undefined without a manifest.
export function fuzzManifest(): readonly FuzzManifestEntry[] | undefined {
	const manifest = join(FUZZ_SAVES, "manifest.json")

	if (FUZZ_SAVES === "" || !existsSync(manifest)) return undefined

	return (JSON.parse(readFileSync(manifest, "utf8")) as { entries: FuzzManifestEntry[] }).entries
}

// Parts of `saved` that are new or whose bytes differ from `source` (by CRC-32 and size, so unchanged parts are
// never inflated).
export async function changedParts(source: OpcPackage, saved: OpcPackage): Promise<string[]> {
	const out: string[] = []

	for (const entry of saved.zip.entries) {
		if (entry.name.endsWith("/")) continue

		const before = source.entry(entry.name)

		if (before === undefined || before.crc32 !== entry.crc32 || before.uncompressedSize !== entry.uncompressedSize) out.push(entry.name)
	}

	return out
}

// Every fuzz save of the manifest with the corpus file it came from.
export async function* fuzzSaves(entries: readonly FuzzManifestEntry[]): AsyncGenerator<Save> {
	for (const entry of entries) {
		const source = new Uint8Array(readFileSync(join(CORPUS_DIR, entry.source)))
		const bytes = new Uint8Array(readFileSync(join(FUZZ_SAVES, entry.path)))
		const changed = await changedParts(await OpcPackage.open(bytesSource(source)), await OpcPackage.open(bytesSource(bytes)))

		yield { name: entry.path, source, sourceName: entry.source, bytes, changedParts: changed, family: entry.family }
	}
}
