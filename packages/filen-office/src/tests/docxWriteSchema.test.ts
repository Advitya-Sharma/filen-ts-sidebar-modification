import { afterAll, beforeAll, describe, expect, it } from "vitest"
import { existsSync } from "node:fs"
import { bytesSource, OpcPackage } from "@filen/office/core/opc"
import { encodeUtf8 } from "@filen/office/core/xml"
import { CORPUS_DIR } from "./corpusDir"
import { editSuite, fuzzManifest, fuzzSaves, type Save } from "./oracles/suite"
import { newValidityErrors, NOT_WELL_FORMED, opcSchemasMissing, schemaFamily, XsdOracle, xsdOracleMissing } from "./oracles/xsd"

// Every part a save rewrites or adds that an ECMA-376 schema declares (WordprocessingML, DrawingML, OPC
// relationships and content types, core/extended/custom properties), validated by xmllint against the
// Transitional schemas after Markup Compatibility, over the scripted edit suite and the fuzz saves
// (FILEN_OFFICE_FUZZ_SAVES). A rewritten part may not have more of any error than its source part; a new part
// has none. Runs where xmllint and the pinned schemas exist (oracles/xsd.ts says how to fetch them).

const MISSING = xsdOracleMissing()

interface Pending {
	readonly label: string
	// "<source file>\0<part>", or undefined for a part the source does not have.
	readonly key: string | undefined
	// The source part, when its errors are not known yet.
	readonly source: Uint8Array | undefined
	// Reads the source part again (for the exemptions, on a failure only: sources are not kept).
	readonly reread: (() => Promise<Uint8Array>) | undefined
	readonly written: Uint8Array
}

// Validates the changed parts of saves in batches: one xmllint run per batch, each source part once.
class Checker {
	saves = 0
	readonly byFamily: Record<string, number> = {}
	readonly byKind: Record<string, number> = {}
	readonly failures: string[] = []
	private readonly sources = new Map<string, Map<string, number>>()
	private pending: Pending[] = []
	private pendingBytes = 0

	private readonly oracle: XsdOracle

	constructor(oracle: XsdOracle) {
		this.oracle = oracle
	}

	async add(save: Save): Promise<void> {
		const written = await OpcPackage.open(bytesSource(save.bytes))
		const source = await OpcPackage.open(bytesSource(save.source))
		const types = await written.contentTypes()

		this.saves++
		this.byFamily[save.family] = (this.byFamily[save.family] ?? 0) + 1

		for (const part of save.changedParts) {
			const contentTypes = part === "[Content_Types].xml"
			const type = contentTypes ? undefined : types.typeOf(part)
			const family = schemaFamily(part, type)

			if (family === undefined || (family === "opc" && !this.oracle.opc)) continue

			const kind = contentTypes ? "contentTypes" : `${family}:${(type ?? "").replace(/^.*[./]([\w-]+)\+xml$/, "$1")}`
			const key = source.has(part) ? `${save.sourceName}\0${part}` : undefined
			const sourceBytes = key !== undefined && !this.sources.has(key) ? await source.read(part) : undefined
			const bytes = await written.read(part)

			this.byKind[kind] = (this.byKind[kind] ?? 0) + 1
			this.pending.push({
				label: `${save.name} ${part}`,
				key,
				source: sourceBytes,
				reread: key === undefined ? undefined : () => source.read(part),
				written: bytes
			})
			this.pendingBytes += bytes.length + (sourceBytes?.length ?? 0)
		}

		if (this.pending.length >= 100 || this.pendingBytes >= 64 << 20) await this.flush()
	}

	async flush(): Promise<void> {
		const batch = this.pending
		const sources = new Map<string, Uint8Array>()

		this.pending = []
		this.pendingBytes = 0

		for (const item of batch) if (item.key !== undefined && item.source !== undefined) sources.set(item.key, item.source)

		const results = this.oracle.validate([...sources.values(), ...batch.map(item => item.written)])
		let index = 0

		for (const key of sources.keys()) this.sources.set(key, results[index++] as Map<string, number>)

		for (const item of batch) {
			const after = results[index++] as Map<string, number>
			const before = (item.key === undefined ? undefined : this.sources.get(item.key)) ?? new Map<string, number>()
			let failures = newValidityErrors(before, after)

			if (failures.length > 0 && item.reread !== undefined) failures = newValidityErrors(before, after, await item.reread())

			for (const failure of failures) this.failures.push(`${item.label}: ${failure}`)
		}
	}
}

describe.skipIf(MISSING !== undefined)("written packages against the ECMA-376 schemas", () => {
	let oracle: XsdOracle

	beforeAll(() => {
		oracle = new XsdOracle()

		const opc = opcSchemasMissing()

		// Relationships, content types and core properties go unvalidated without them: say so in the log.
		if (opc !== undefined) console.warn(`schema: OPC parts not validated: ${opc}`)
	})

	afterAll(() => {
		oracle.dispose()
	})

	it("reports what it should", { timeout: 60_000 }, () => {
		const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
		const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
		const W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
		const documents = [
			`<w:document ${W}><w:body><w:p/></w:body></w:document>`,
			`<w:document ${W}><w:body><w:p><w:bogus/></w:p></w:body></w:document>`,
			// Ignorable markup and AlternateContent go; the error after them is still found.
			`<w:document ${W} ${MC} ${W14} mc:Ignorable="w14"><w:body><w:p w14:paraId="12345678"><w14:x/><mc:AlternateContent><mc:Choice Requires="w14"><w14:y/></mc:Choice><mc:Fallback><w:r><w:t>a</w:t></w:r></mc:Fallback></mc:AlternateContent><w:r><w:bogus/></w:r></w:p></w:body></w:document>`,
			`<w:document ${W}><w:body><w:p>`,
			'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="t" Target="x" TargetMode="Outside"/></Relationships>',
			'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dcterms:created xsi:type="dcterms:W3CDTF">yesterday</dcterms:created></cp:coreProperties>',
			// Properties after the run: the tree-mode report names the element too.
			`<w:document ${W}><w:body><w:p><w:r/><w:pPr/></w:p></w:body></w:document>`
		]
		const results = oracle.validate(documents.map(encodeUtf8)).map(errors => [...errors.keys()].join(" | "))

		expect(results[0]).toBe("")
		expect(results[1]).toMatch(/\}bogus': This element is not expected/)
		expect(results[2]).toMatch(/\}bogus': This element is not expected/)
		expect(results[2]).not.toMatch(/paraId|w14|wordml/)
		expect(results[3]).toBe(NOT_WELL_FORMED)
		if (oracle.opc) {
			expect(results[4]).toMatch(/TargetMode/)
			expect(results[5]).toMatch(/yesterday/)
		}
		expect(results[6]).toMatch(/\}pPr': This element is not expected/)
	})

	it.skipIf(!existsSync(CORPUS_DIR))("adds no validity errors to the parts the edit suite writes", { timeout: 1_800_000 }, async () => {
		const checker = new Checker(oracle)

		for await (const save of editSuite()) await checker.add(save)

		await checker.flush()
		console.log(`schema, edit suite: ${checker.saves} saves, parts ${JSON.stringify(checker.byKind)}`)
		expect(checker.failures).toEqual([])
		// Fewer means lost coverage: the corpus rewrites about 265 main, 115 numbering and 112 styles parts.
		expect(checker.byKind["wml:main"] ?? 0).toBeGreaterThanOrEqual(250)
		expect(checker.byKind["wml:numbering"] ?? 0).toBeGreaterThanOrEqual(105)
		expect(checker.byKind["wml:styles"] ?? 0).toBeGreaterThanOrEqual(105)
	})

	const manifest = fuzzManifest()

	it.skipIf(manifest === undefined || !existsSync(CORPUS_DIR))(
		"adds no validity errors to the fuzz saves",
		{ timeout: 7_200_000 },
		async () => {
			const entries = manifest ?? []
			const checker = new Checker(oracle)

			for await (const save of fuzzSaves(entries)) await checker.add(save)

			await checker.flush()
			console.log(
				`schema, fuzz saves: ${checker.saves} by family ${JSON.stringify(checker.byFamily)}, parts ${JSON.stringify(checker.byKind)}`
			)
			expect(checker.saves).toBe(entries.length)
			expect(checker.failures).toEqual([])
		}
	)
})
