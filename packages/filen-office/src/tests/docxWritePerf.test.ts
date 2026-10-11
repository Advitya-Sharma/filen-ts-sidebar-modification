import { existsSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { DocxEditorSession } from "@filen/office/core/docx/edit"
import { saveDocx, writerProof } from "@filen/office/core/docx/write"
import { CORPUS_DIR } from "./corpusDir"
import { openBytes } from "./docxEditFixtures"
import { runEditScript } from "./docxWriteFixtures"

const FIXTURE = join(import.meta.dirname, "fixtures", "08-large-300p.docx")
const TARIFF = join(CORPUS_DIR, "docx", "govuk__Tariff_of_the_United_Kingdom_v1.34.docx")

async function measure(path: string): Promise<string> {
	const opened = await openBytes(new Uint8Array(readFileSync(path)))
	let start = performance.now()
	const session = new DocxEditorSession(opened.context, { proof: null })
	const build = performance.now() - start

	start = performance.now()
	session.prove(writerProof(opened.context))

	const proof = performance.now() - start
	const unedited = await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: [session] })
	const script = runEditScript(opened)
	const peak = process.resourceUsage().maxRSS
	const edited = await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: [script.session] })
	const grown = (process.resourceUsage().maxRSS - peak) / 1024

	expect(unedited.changedParts).toEqual([])

	const t = edited.timings

	return (
		`build ${build.toFixed(0)} ms, gate 1 ${proof.toFixed(0)} ms; unedited save ${(unedited.timings.serializeMs + unedited.timings.zipMs).toFixed(0)} ms; ` +
		`edited (${script.done.join("+")}) serialise ${t.serializeMs} ms, gate 2 parts ${t.verifyPartsMs} ms, zip ${t.zipMs} ms, ` +
		`gate 2 zip ${t.verifyZipMs} ms, peak RSS +${grown.toFixed(0)} MB`
	)
}

describe("docx writer performance", () => {
	it("proves and saves the 300-page fixture", { timeout: 120_000 }, async () => {
		console.log(`300p: ${await measure(FIXTURE)}`)
	})

	it.skipIf(!existsSync(TARIFF))("proves and saves the Tariff document", { timeout: 600_000 }, async () => {
		console.log(`Tariff: ${await measure(TARIFF)}`)
	})
})
