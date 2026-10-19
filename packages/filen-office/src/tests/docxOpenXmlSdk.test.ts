import { afterAll, beforeAll, describe, expect, it } from "vitest"
import { existsSync, mkdtempSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { CORPUS_DIR } from "./corpusDir"
import { buildDocx } from "./docxSynthetic"
import { loadAllowlist, newSdkErrors, OpenXmlSdk, openXmlSdkMissing, sdkErrorCounts, type AllowEntry } from "./oracles/openXmlSdk"
import { editSuite, FUZZ_SAVES, fuzzManifest, fuzzSaves, type Save } from "./oracles/suite"

// Microsoft's Open XML SDK validator (FileFormatVersions.Microsoft365: schemas with Word's extensions, plus
// semantic checks such as unique ids and resolvable relationships) over the scripted edit suite and the fuzz
// saves: a save may not have more of any error than its source file, except the allowlisted false positives
// (tools/openxml-validate/allow.json, each with its reason). Opt-in: FILEN_OFFICE_OOXML_VALIDATOR names the built
// validator (tools/openxml-validate), FILEN_OFFICE_DOTNET the dotnet host; skips without them.

const MISSING = openXmlSdkMissing()

class Checker {
	saves = 0
	readonly byFamily: Record<string, number> = {}
	readonly failures: string[] = []
	readonly used = new Map<AllowEntry, number>()
	private readonly sources = new Map<string, Map<string, number>>()
	private readonly allow = loadAllowlist()
	private readonly sdk: OpenXmlSdk
	private readonly scratch: string

	constructor(sdk: OpenXmlSdk, scratch: string) {
		this.sdk = sdk
		this.scratch = scratch
	}

	async add(save: Save, savedPath: string | undefined): Promise<void> {
		let before = this.sources.get(save.sourceName)

		if (before === undefined) {
			before = sdkErrorCounts(await this.sdk.validate(join(CORPUS_DIR, save.sourceName)))
			this.sources.set(save.sourceName, before)
		}

		let path = savedPath

		if (path === undefined) {
			path = join(this.scratch, "save.docx")
			writeFileSync(path, save.bytes)
		}

		const after = sdkErrorCounts(await this.sdk.validate(path))

		this.saves++
		this.byFamily[save.family] = (this.byFamily[save.family] ?? 0) + 1

		for (const failure of newSdkErrors(before, after, this.allow, this.used)) this.failures.push(`${save.name}: ${failure}`)
	}

	report(label: string): void {
		console.log(
			`openxml-sdk, ${label}: ${this.saves} saves ${JSON.stringify(this.byFamily)}; allowlisted ${JSON.stringify(
				[...this.used].map(([entry, count]) => `${entry.id} ${entry.description}: ${count}`)
			)}`
		)
	}
}

describe.skipIf(MISSING !== undefined)("written packages against the Open XML SDK validator", () => {
	let sdk: OpenXmlSdk
	let scratch: string

	beforeAll(() => {
		sdk = new OpenXmlSdk()
		scratch = mkdtempSync(join(tmpdir(), "filen-office-openxml-"))
	})

	afterAll(async () => {
		await sdk.close()
		rmSync(scratch, { recursive: true, force: true })
	})

	it("reports what it should", { timeout: 60_000 }, async () => {
		const valid = join(scratch, "valid.docx")
		const invalid = join(scratch, "invalid.docx")

		writeFileSync(valid, buildDocx({ body: "<w:p><w:r><w:t>a</w:t></w:r></w:p>" }))
		writeFileSync(invalid, buildDocx({ body: '<w:p><w:pPr><w:jc w:val="sideways"/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>' }))

		const before = sdkErrorCounts(await sdk.validate(valid))
		const after = sdkErrorCounts(await sdk.validate(invalid))

		expect([...before.keys()]).toEqual([])
		expect(newSdkErrors(before, after, [])).toEqual([expect.stringMatching(/^\/word\/document\.xml Sch_\w+: .*sideways/)])
	})

	it("has a reason for every allowlist entry", () => {
		for (const entry of loadAllowlist()) expect(entry.reason.length, `${entry.id} ${entry.description}`).toBeGreaterThan(20)
	})

	it.skipIf(!existsSync(CORPUS_DIR))("adds no errors to the edit suite's saves", { timeout: 1_800_000 }, async () => {
		const checker = new Checker(sdk, scratch)

		for await (const save of editSuite()) await checker.add(save, undefined)

		checker.report("edit suite")
		expect(checker.failures).toEqual([])
		expect(checker.saves).toBeGreaterThanOrEqual(250)
	})

	const manifest = fuzzManifest()

	it.skipIf(manifest === undefined || !existsSync(CORPUS_DIR))("adds no errors to the fuzz saves", { timeout: 7_200_000 }, async () => {
		const entries = manifest ?? []
		const checker = new Checker(sdk, scratch)

		for await (const save of fuzzSaves(entries)) await checker.add(save, join(FUZZ_SAVES, save.name))

		checker.report("fuzz saves")
		expect(checker.saves).toBe(entries.length)
		expect(checker.failures).toEqual([])
	})
})
