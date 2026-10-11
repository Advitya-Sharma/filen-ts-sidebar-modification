import { spawnSync } from "node:child_process"
import { existsSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { afterAll, describe, expect, it } from "vitest"
import { bytesSource, OpcPackage } from "@filen/office/core/opc"
import { saveDocx } from "@filen/office/core/docx/write"
import { SCHEMA_CACHE, transitionalSchemas } from "../scripts/wmlOrder"
import { CORPUS_DIR as CORPUS } from "./corpusDir"
import { openBytes } from "./docxEditFixtures"
import { runEditScript } from "./docxWriteFixtures"

// The parts the edit suite rewrites, validated by xmllint against the ECMA-376 Transitional schemas. Runs where
// xmllint, the corpus and the pinned Part 4 archive exist (FILEN_OFFICE_WML_SCHEMAS, else the cache that
// `pnpm run generate:wml-order --fetch` fills).

const ARCHIVE = process.env["FILEN_OFFICE_WML_SCHEMAS"] || SCHEMA_CACHE
const XMLLINT = spawnSync("xmllint", ["--version"]).status === 0
// The XML namespace's attributes as XML 1.0 defines them; wml.xsd imports the namespace without a location.
const XML_SCHEMA =
	'<?xml version="1.0"?><xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema" targetNamespace="http://www.w3.org/XML/1998/namespace">' +
	'<xs:attribute name="space"><xs:simpleType><xs:restriction base="xs:NCName"><xs:enumeration value="default"/><xs:enumeration value="preserve"/></xs:restriction></xs:simpleType></xs:attribute>' +
	'<xs:attribute name="lang" type="xs:language"/></xs:schema>'
// Word's extensions (w14, w15, mc...) are outside the schema: errors about them say nothing of the writer.
const EXTENSIONS = /\{http:\/\/schemas\.microsoft\.com\/|\{http:\/\/schemas\.openxmlformats\.org\/markup-compatibility\//

function schemaDirectory(): string {
	const directory = mkdtempSync(join(tmpdir(), "filen-office-wml-"))

	for (const [name, bytes] of transitionalSchemas(new Uint8Array(readFileSync(ARCHIVE)))) {
		const text = new TextDecoder().decode(bytes)

		writeFileSync(
			join(directory, name),
			text.replace(
				'<xsd:import namespace="http://www.w3.org/XML/1998/namespace"/>',
				'<xsd:import namespace="http://www.w3.org/XML/1998/namespace" schemaLocation="xml.xsd"/>'
			)
		)
	}

	writeFileSync(join(directory, "xml.xsd"), XML_SCHEMA)

	return directory
}

// Validity errors by message (line numbers and xmllint's list of expected elements left out), counted.
function validityErrors(schema: string, bytes: Uint8Array): Map<string, number> {
	const result = spawnSync("xmllint", ["--noout", "--schema", schema, "-"], { input: bytes, maxBuffer: 1 << 30 })
	const errors = new Map<string, number>()

	for (const line of result.stderr.toString().split("\n")) {
		if (!line.includes("validity error") || EXTENSIONS.test(line)) continue

		const message = line.replace(/^-:\d+: /, "").replace(/ Expected is .*$/, "")

		errors.set(message, (errors.get(message) ?? 0) + 1)
	}

	return errors
}

function corpusFiles(): string[] {
	const files: string[] = []

	for (const extension of ["docx", "docm", "dotx", "dotm"]) {
		const directory = join(CORPUS, extension)

		if (existsSync(directory)) for (const name of readdirSync(directory).sort()) files.push(join(directory, name))
	}

	return files
}

describe.skipIf(!XMLLINT || !existsSync(ARCHIVE) || !existsSync(CORPUS))("written WordprocessingML against the ECMA-376 schemas", () => {
	const directory = XMLLINT && existsSync(ARCHIVE) ? schemaDirectory() : ""

	afterAll(() => {
		if (directory !== "") rmSync(directory, { recursive: true, force: true })
	})

	it("adds no validity errors to the parts the edit suite rewrites", { timeout: 1_800_000 }, async () => {
		const schema = join(directory, "wml.xsd")
		const failures: string[] = []
		let parts = 0

		for (const file of corpusFiles()) {
			const name = file.slice(CORPUS.length + 1)
			let opened

			try {
				opened = await openBytes(new Uint8Array(readFileSync(file)))
			} catch {
				continue
			}

			if (opened.model.body === undefined || opened.context.gate.length > 0) continue

			const script = runEditScript(opened)
			const result = await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: [script.session] })

			if (result.bytes === undefined) throw new Error("No bytes")

			const written = await OpcPackage.open(bytesSource(result.bytes))
			const checked = new Set([
				opened.pkg.mainPart,
				opened.pkg.part("numbering"),
				opened.pkg.part("styles"),
				"word/numbering.xml",
				"word/styles.xml"
			])

			for (const part of result.changedParts) {
				if (!checked.has(part)) continue

				parts++

				const before = opened.pkg.has(part) ? validityErrors(schema, await opened.pkg.opc.read(part)) : new Map<string, number>()
				const after = validityErrors(schema, await written.read(part))

				for (const [message, count] of after) {
					const was = before.get(message) ?? 0

					// A value the source already had wrong is copied with the run it is in (a split repeats it).
					if (was > 0 && message.includes("[facet ")) continue
					if (count > was) failures.push(`${name} ${part}: ${message} (${was} -> ${count})`)
				}
			}
		}

		console.log(`schema: ${parts} rewritten parts validated`)
		expect(failures).toEqual([])
	})
})
