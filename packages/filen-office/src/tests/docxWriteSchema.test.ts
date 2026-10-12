import { existsSync, readFileSync, readdirSync, rmSync } from "node:fs"
import { join } from "node:path"
import { afterAll, describe, expect, it } from "vitest"
import { bytesSource, OpcPackage } from "@filen/office/core/opc"
import { saveDocx } from "@filen/office/core/docx/write"
import { CORPUS_DIR as CORPUS } from "./corpusDir"
import { openBytes } from "./docxEditFixtures"
import { runEditScript } from "./docxWriteFixtures"
import { SCHEMA_ARCHIVE, XMLLINT, schemaDirectory, validityErrors } from "./fuzz/harness"

// The WordprocessingML parts the edit suite rewrites or adds (document, styles, numbering, comments, notes,
// headers, footers, settings...), validated by xmllint against the ECMA-376 Transitional schemas. Runs where
// xmllint, the corpus and the pinned Part 4 archive exist (FILEN_OFFICE_WML_SCHEMAS, else the cache that
// `pnpm run generate:wml-order --fetch` fills).

// Main document types (document, template, macro-enabled) and every other wordprocessingml.* XML part.
const WML_TYPE =
	/^application\/(vnd\.openxmlformats-officedocument\.wordprocessingml\.[\w.-]+\+xml|vnd\.ms-word\.(document|template)\.macroEnabled(Template)?\.main\+xml)$/

function corpusFiles(): string[] {
	const files: string[] = []

	for (const extension of ["docx", "docm", "dotx", "dotm"]) {
		const directory = join(CORPUS, extension)

		if (existsSync(directory)) for (const name of readdirSync(directory).sort()) files.push(join(directory, name))
	}

	return files
}

describe.skipIf(!XMLLINT || !existsSync(SCHEMA_ARCHIVE) || !existsSync(CORPUS))(
	"written WordprocessingML against the ECMA-376 schemas",
	() => {
		const directory = XMLLINT && existsSync(SCHEMA_ARCHIVE) ? schemaDirectory() : ""

		afterAll(() => {
			if (directory !== "") rmSync(directory, { recursive: true, force: true })
		})

		it("adds no validity errors to the parts the edit suite rewrites", { timeout: 1_800_000 }, async () => {
			const schema = join(directory, "wml.xsd")
			const failures: string[] = []
			const validated: Record<string, number> = {}

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
				const result = await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: script.sessions })

				if (result.bytes === undefined) throw new Error("No bytes")

				const written = await OpcPackage.open(bytesSource(result.bytes))
				const types = await written.contentTypes()

				for (const part of result.changedParts) {
					// A Default for "xml" can give the content types part itself a WordprocessingML type.
					const type = part === "[Content_Types].xml" ? undefined : types.typeOf(part)

					if (type === undefined || !WML_TYPE.test(type)) continue

					const kind = type.replace(/^.*[./]([\w-]+)\+xml$/, "$1")

					validated[kind] = (validated[kind] ?? 0) + 1

					const before = opened.pkg.has(part)
						? validityErrors(schema, await opened.pkg.opc.read(part))
						: new Map<string, number>()
					const after = validityErrors(schema, await written.read(part))

					for (const [message, count] of after) {
						const was = before.get(message) ?? 0

						// A value the source already had wrong is copied with the run it is in (a split repeats it).
						if (was > 0 && message.includes("[facet ")) continue
						if (count > was) failures.push(`${name} ${part}: ${message} (${was} -> ${count})`)
					}
				}
			}

			console.log(`schema: rewritten parts validated ${JSON.stringify(validated)}`)
			expect(failures).toEqual([])
			// The full corpus rewrites 265 main, 115 numbering and 112 styles parts: fewer means lost coverage.
			expect(validated["main"] ?? 0).toBeGreaterThanOrEqual(250)
			expect(validated["numbering"] ?? 0).toBeGreaterThanOrEqual(105)
			expect(validated["styles"] ?? 0).toBeGreaterThanOrEqual(105)
		})
	}
)
