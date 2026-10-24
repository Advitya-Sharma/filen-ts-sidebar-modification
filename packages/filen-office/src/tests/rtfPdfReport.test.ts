import { describe, expect, it } from "vitest"
import type { FontExtents, TextMeasurer } from "../core/text/index"
import type { DocxEngineServices } from "../dom/docx/index"
import { exportDocxPdf } from "../dom/docx/pdf"
import { reportOf } from "../dom/docx/pdfClient"
import { workerHarness } from "./docxEncryptedFixtures"
import { pdfAssets } from "./docxPdfFixtures"
import { latin1 } from "./rtfWriteFixtures"

// What an RTF file shows that the package it reads as leaves out (an equation, a shape, a form field...) reaches the
// PDF's report as the conversions' reports have it, so the export says it was simplified; what no page shows (an
// unknown destination such as the generator's name) does not.

const measurer: TextMeasurer = {
	kerningControl: true,
	measure(font, _kerning, texts, count, widths) {
		const size = Number(/([\d.]+)px/.exec(font)?.[1] ?? "16")

		for (let i = 0; i < count; i++) widths[i] = (texts[i] ?? "").length * size * 0.5
	},
	fontExtents(font): FontExtents {
		const size = Number(/([\d.]+)px/.exec(font)?.[1] ?? "16")

		return { ascent: size * 0.8, descent: size * 0.2 }
	}
}

function services(): DocxEngineServices {
	return {
		fontHost: { addBytes: () => Promise.resolve(true), addLocal: () => Promise.resolve(false), remove: () => undefined },
		replayFonts: false,
		fontAssets: { read: () => Promise.resolve(new Uint8Array(4)) },
		measurer,
		yield: () => new Promise(resolve => setTimeout(resolve, 0)),
		sliceMs: 8
	}
}

const HEADER = "{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Times;}}{\\*\\generator Some Writer;}"
const MATH = latin1(
	`${HEADER}\\pard\\plain Before\\par\\pard\\plain {\\mmath{\\*\\moMathPara{\\*\\moMath{\\*\\mf{\\*\\mnum a}{\\*\\mden b}}}}}\\par\\pard\\plain After\\par}`
)
const PLAIN = latin1(`${HEADER}\\pard\\plain Only text\\par}`)

async function exported(rtf: Uint8Array): Promise<ReturnType<typeof reportOf>> {
	const harness = workerHarness()

	try {
		const opened = await harness.pool.rtf.open(rtf)

		if (opened.kind !== "opened") throw new Error(opened.kind)

		try {
			const result = await exportDocxPdf(await opened.document.take(), services(), pdfAssets, {
				options: { title: "letter", deterministic: true, untranslated: opened.document.untranslated },
				signal: { aborted: false }
			})

			return reportOf(result.report)
		} finally {
			opened.document.release()
		}
	} finally {
		harness.dispose()
	}
}

describe("the PDF of an RTF file", () => {
	it("reports an equation its package leaves out as content left out of the PDF", async () => {
		expect((await exported(MATH)).refused).toEqual([["untranslated", 1]])
	})

	it("reports nothing for a file whose package holds all it shows (unknown destinations aside)", async () => {
		expect((await exported(PLAIN)).refused).toEqual([])
	})
})
