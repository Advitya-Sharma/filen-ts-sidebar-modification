import { describe, expect, it } from "vitest"
import { build } from "vite"

// The PDF fonts as a worker bundles them: @filen/office is "sideEffects": false, so a module kept only for what its
// loading does would be dropped. The Indic shaper is registered by a named import; the bundle must hold it.

const ENTRY = "virtual:pdf-fonts-entry"

describe("the PDF font set in a production bundle", () => {
	it("keeps the Indic shaper (its reph and pre-base forms)", async () => {
		const result: unknown = await build({
			configFile: false,
			logLevel: "silent",
			root: import.meta.dirname,
			plugins: [
				{
					name: "entry",
					resolveId: (id: string) => (id === ENTRY ? `\0${ENTRY}` : null),
					load: (id: string) => (id === `\0${ENTRY}` ? 'export { PdfFontSet } from "@filen/office/core/pdfwrite/fonts"' : null)
				}
			],
			build: {
				write: false,
				minify: false,
				rolldownOptions: {
					input: ENTRY,
					preserveEntrySignatures: "strict",
					external: ["@filen/shared/linkPolicy"],
					output: { format: "es" }
				}
			}
		})
		const outputs = (Array.isArray(result) ? result : [result]) as { readonly output: readonly { readonly code?: string }[] }[]
		const code = outputs
			.flatMap(output => output.output)
			.map(chunk => chunk.code ?? "")
			.join("\n")

		expect(code).toContain("PdfFontSet")
		expect(code).toMatch(/["']rphf["']/)
		expect(code).toMatch(/["']pref["']/)
	}, 120_000)
})
