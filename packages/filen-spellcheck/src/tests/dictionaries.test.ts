import { createHash } from "node:crypto"
import { existsSync, readdirSync, readFileSync } from "node:fs"
import { join, relative, sep } from "node:path"
import { describe, expect, it } from "vitest"
import {
	DICTIONARIES,
	DICTIONARIES_DIR,
	DICTIONARIES_OUT,
	MARK,
	SOURCES,
	STAMP,
	TRANSFORM,
	UPSTREAM,
	sourceUrl,
	type Manifest,
	type ManifestFile
} from "../../scripts/build-dictionaries"

const MANIFEST_BYTES = readFileSync(join(DICTIONARIES_DIR, "manifest.json"))
const MANIFEST_TEXT = MANIFEST_BYTES.toString("utf8")
const manifest = JSON.parse(MANIFEST_TEXT) as Manifest
const utf8 = new TextDecoder("utf-8", { fatal: true, ignoreBOM: true })

function sha256(bytes: Uint8Array): string {
	return createHash("sha256").update(bytes).digest("hex")
}

const stampPath = join(DICTIONARIES_OUT, STAMP)
// Built outputs are verified only when `pnpm run build:dictionaries` completed against this manifest.
const BUILT = existsSync(stampPath) && readFileSync(stampPath, "utf8") === sha256(MANIFEST_BYTES)

// Reads a manifest file from dir and checks its size and hash.
function read(dir: string, entry: ManifestFile): Buffer {
	const bytes = readFileSync(join(dir, entry.file))

	expect(bytes.length, entry.file).toBe(entry.bytes)
	expect(sha256(bytes), entry.file).toBe(entry.sha256)

	return bytes
}

// Decodes strictly and checks the normalization the build promises: no BOM, LF only, a final newline.
function normalizedText(entry: ManifestFile, bytes: Buffer): string {
	const text = utf8.decode(bytes)

	expect(text.startsWith("﻿"), `${entry.file} has a BOM`).toBe(false)
	expect(text.includes("\r"), `${entry.file} has a CR`).toBe(false)
	expect(text.endsWith("\n"), `${entry.file} lacks a final newline`).toBe(true)

	return text
}

// Some upstream notices are Latin-1.
function noticeText(bytes: Buffer): string {
	try {
		return utf8.decode(bytes)
	} catch {
		return new TextDecoder("windows-1252").decode(bytes)
	}
}

// The words of a holder line that must occur in the notices: without "Copyright", years, symbols and a trailing role such
// as "(authors of the ... dictionary)".
function holderWords(line: string): string[] {
	return (line.replace(/\s*\([^()]* of the [^()]*\)$/, "").match(/\p{L}[\p{L}.'-]*/gu) ?? [])
		.map(word => word.replace(/\.$/, ""))
		.filter(word => word.length >= 3 && word !== "Copyright")
}

describe("dictionaries/manifest.json", () => {
	it("has the documented shape, in canonical formatting", () => {
		expect(MANIFEST_TEXT).toBe(`${JSON.stringify(manifest, null, "\t")}\n`)
		expect(Object.keys(manifest)).toEqual(["version", "transform", "upstream", "dictionaries"])
		expect(manifest.version).toBe(1)
		expect(manifest.transform).toBe(TRANSFORM)
		expect(manifest.upstream).toEqual({ repository: UPSTREAM.repository, commit: UPSTREAM.commit })
		expect(manifest.dictionaries.map(entry => entry.id)).toEqual(DICTIONARIES.map(source => source.id))

		for (const [index, entry] of manifest.dictionaries.entries()) {
			const source = DICTIONARIES[index]

			expect(Object.keys(entry)).toEqual([
				"id",
				"name",
				"upstreamVersion",
				"license",
				"copyright",
				"sources",
				"aff",
				"dic",
				"notices",
				"modifications"
			])
			expect([entry.name, entry.license, entry.copyright, entry.upstreamVersion]).toEqual([
				source?.name,
				source?.license,
				source?.copyright,
				source?.upstreamVersion
			])
			expect(entry.sources).toEqual(
				source === undefined
					? []
					: [source.aff, source.dic, ...source.notices].map(path => ({ url: sourceUrl(path), sha256: SOURCES[path] }))
			)
			expect(
				entry.sources.every(item =>
					item.url.startsWith(`https://raw.githubusercontent.com/LibreOffice/dictionaries/${UPSTREAM.commit}/`)
				)
			).toBe(true)
			expect(entry.aff.file).toBe(`${entry.id}/${entry.id}.aff`)
			expect(entry.dic.file).toBe(`${entry.id}/${entry.id}.dic`)

			for (const file of [entry.aff, entry.dic, ...entry.notices]) {
				expect(Object.keys(file)).toEqual(["file", "bytes", "sha256"])
				expect(file.file.startsWith(`${entry.id}/`)).toBe(true)
				expect(file.sha256).toMatch(/^[0-9a-f]{64}$/)
			}

			expect(entry.modifications.length).toBeGreaterThan(0)
			expect(entry.modifications.every(line => typeof line === "string" && line !== "")).toBe(true)
		}
	})

	it("lists every file in dictionaries/ besides itself and README.md, and no built output", () => {
		const listed = manifest.dictionaries.flatMap(entry => entry.notices.map(file => file.file))
		const present = readdirSync(DICTIONARIES_DIR, { recursive: true, withFileTypes: true })
			.filter(entry => entry.isFile())
			.map(entry => relative(DICTIONARIES_DIR, join(entry.parentPath, entry.name)).split(sep).join("/"))

		expect(new Set(listed).size).toBe(listed.length)
		expect(present.sort()).toEqual(["README.md", "manifest.json", ...listed].sort())
	})
})

describe.each(manifest.dictionaries)("dictionaries/$id", entry => {
	it("ships its README.md and upstream notices as the manifest lists them", () => {
		expect(entry.notices[0]?.file).toBe(`${entry.id}/README.md`)
		expect(entry.notices.length).toBeGreaterThan(1)

		for (const notice of entry.notices) read(DICTIONARIES_DIR, notice)
	})

	it("names copyright holders that its upstream notices state", () => {
		const notices = entry.notices.slice(1).map(notice => noticeText(readFileSync(join(DICTIONARIES_DIR, notice.file))))

		expect(entry.copyright.length).toBeGreaterThan(0)

		for (const line of entry.copyright) {
			const names = holderWords(line)

			expect(names.length, line).toBeGreaterThan(0)

			for (const name of names)
				expect(
					notices.some(text => text.includes(name)),
					`${line}: ${name}`
				).toBe(true)
		}
	})

	it.skipIf(!BUILT)("built into .dictionaries/ as the manifest expects, normalized UTF-8 with the modification notice", () => {
		const aff = normalizedText(entry.aff, read(DICTIONARIES_OUT, entry.aff))

		normalizedText(entry.dic, read(DICTIONARIES_OUT, entry.dic))

		expect(aff.startsWith(`${MARK}\n`)).toBe(true)
		expect(aff.match(/^SET[ \t].*$/gm)).toEqual(["SET UTF-8"])
	})
})
