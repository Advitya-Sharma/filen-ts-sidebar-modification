import { existsSync, readdirSync } from "node:fs"
import { join } from "node:path"
import { strFromU8, unzipSync } from "fflate"
import { describe, expect, it } from "vitest"
import { Item, type DocumentLayout, type LayoutOptions, type StyleSettings } from "@filen/office/core/docx"
import type { FontManager, FontRequest, ResolvedFont } from "@filen/office/core/fonts"
import { CORPUS_DIR } from "./corpusDir"
import { adjustedFonts, layOut, readBytes } from "./layoutFixtures"

// Page breaks Word recorded (w:lastRenderedPageBreak) against ours, on Word-authored files whose recorded
// breaks agree with docProps/app.xml <Pages> (the corpus README's page-count ground truth). Widths come from
// the bundled fonts' advance widths, so the comparison is deterministic and needs no browser.
//
// FILEN_OFFICE_LAYOUT_SWEEP=1 also runs the calibration variants and prints their agreement.

const DIRECTORY = join(CORPUS_DIR, "docx")
const present = existsSync(DIRECTORY)

interface Truth {
	readonly name: string
	readonly bytes: Uint8Array
	readonly pages: number
	readonly breaks: number
	readonly minorFont: string
}

// `loose`: every Word-authored file with recorded breaks, consistent with <Pages> or not.
function groundTruth(loose = false): Truth[] {
	const out: Truth[] = []

	for (const name of readdirSync(DIRECTORY).sort()) {
		if (!name.endsWith(".docx")) continue

		const bytes = readBytes(join(DIRECTORY, name))
		let files: Record<string, Uint8Array>

		try {
			files = unzipSync(bytes, {
				filter: file =>
					file.name === "docProps/app.xml" || file.name === "word/document.xml" || file.name === "word/theme/theme1.xml"
			})
		} catch {
			continue
		}

		const app = files["docProps/app.xml"]
		const document = files["word/document.xml"]

		if (app === undefined || document === undefined) continue

		const appXml = strFromU8(app)

		if (!/<Application>Microsoft (Office|Macintosh) Word/.test(appXml)) continue

		const pages = Number(/<Pages>(\d+)<\/Pages>/.exec(appXml)?.[1] ?? "NaN")
		const breaks = (strFromU8(document).match(/<w:lastRenderedPageBreak\/>/g) ?? []).length

		if (loose ? breaks === 0 : !(pages === breaks + 1) || pages < 2) continue

		const theme = files["word/theme/theme1.xml"]
		const minorFont = theme === undefined ? "" : (/<a:minorFont><a:latin typeface="([^"]*)"/.exec(strFromU8(theme))?.[1] ?? "")

		out.push({ name, bytes, pages, breaks, minorFont })
	}

	return out
}

interface FileResult {
	readonly name: string
	// The document's default font (docDefaults and Normal style).
	readonly font: string
	// Every ascii font family the text uses.
	readonly families: readonly string[]
	readonly truth: number
	readonly ours: number
	// Per Word page break: lines between our page start and Word's (positive: we start the page earlier).
	readonly offsets: number[]
	// Per Word page break: our page minus Word's.
	readonly pageOffsets: number[]
	readonly drift: number
	readonly ms: number
}

// Word's page number of each recorded break, by (paragraph item, ordinal within the paragraph).
function wordPages(layout: DocumentLayout): Map<string, number> {
	const out = new Map<string, number>()
	const story = layout.model.body
	const part = story === undefined ? undefined : layout.model.parts[story.part]

	if (part === undefined) return out

	let paragraph = -1
	let ordinal = 0
	let page = 1

	for (let item = 0; item < part.count; item++) {
		const kind = part.kinds[item]

		if (kind === Item.Paragraph) {
			paragraph = item
			ordinal = 0
		} else if (kind === Item.LastRenderedPageBreak) {
			page++
			out.set(`${paragraph}:${ordinal++}`, page)
		}
	}

	return out
}

async function measure(
	truth: Truth,
	options: LayoutOptions,
	fonts?: (manager: FontManager) => ReturnType<typeof adjustedFonts>,
	settings?: (settings: StyleSettings) => StyleSettings
): Promise<FileResult> {
	const started = performance.now()
	const layout = await layOut(truth.bytes, options, fonts, settings)
	const ms = performance.now() - started
	const word = wordPages(layout)
	const offsets: number[] = []
	const pageOffsets: number[] = []

	for (const found of layout.renderedBreaks) {
		const page = word.get(`${found.item}:${found.ordinal}`)

		if (page === undefined) continue

		pageOffsets.push(found.page + 1 - page)

		const first = layout.pageFirstLines[page - 1]

		if (first !== undefined) offsets.push(found.line - first)
	}

	let drift = 0

	for (let k = 1; k < offsets.length; k++) drift += (offsets[k] ?? 0) - (offsets[k - 1] ?? 0)

	const resolver = layout.loaded.resolver
	const font = resolver.run(resolver.paragraph({}), {}).fonts.ascii
	const families = [...new Set(layout.styleInfo.map(info => info?.run?.fonts.ascii.toLowerCase() ?? ""))].filter(name => name !== "")

	return {
		name: truth.name,
		font,
		families,
		truth: truth.pages,
		ours: layout.pages.length,
		offsets,
		pageOffsets,
		drift: offsets.length > 1 ? drift / (offsets.length - 1) : 0,
		ms
	}
}

interface Summary {
	readonly files: number
	readonly exact: number
	readonly withinOne: number
	readonly meanDelta: number
	readonly meanOffset: number
	readonly meanAbsOffset: number
	readonly meanDrift: number
	// Share of Word's breaks we break at exactly, and within one line.
	readonly zero: number
	readonly nearby: number
	// Share of Word's breaks that land on the page Word starts with them.
	readonly samePage: number
}

function summarize(results: readonly FileResult[]): Summary {
	const offsets = results.flatMap(result => result.offsets)

	return {
		files: results.length,
		exact: results.filter(result => result.ours === result.truth).length,
		withinOne: results.filter(result => Math.abs(result.ours - result.truth) <= 1).length,
		meanDelta: results.reduce((sum, result) => sum + (result.ours - result.truth), 0) / Math.max(1, results.length),
		meanOffset: offsets.reduce((sum, offset) => sum + offset, 0) / Math.max(1, offsets.length),
		meanAbsOffset: offsets.reduce((sum, offset) => sum + Math.abs(offset), 0) / Math.max(1, offsets.length),
		meanDrift: results.reduce((sum, result) => sum + result.drift, 0) / Math.max(1, results.length),
		zero: offsets.filter(offset => offset === 0).length / Math.max(1, offsets.length),
		nearby: offsets.filter(offset => Math.abs(offset) <= 1).length / Math.max(1, offsets.length),
		samePage:
			results.flatMap(result => result.pageOffsets).filter(offset => offset === 0).length /
			Math.max(1, results.flatMap(result => result.pageOffsets).length)
	}
}

function format(label: string, summary: Summary): string {
	return (
		`${label.padEnd(34)} files ${summary.files}  exact ${summary.exact}  ±1 ${summary.withinOne} (${((100 * summary.withinOne) / Math.max(1, summary.files)).toFixed(1)} %)  ` +
		`Δpages ${summary.meanDelta.toFixed(2)}  breaks on Word's page ${(100 * summary.samePage).toFixed(1)} %  exact ${(100 * summary.zero).toFixed(1)} % ±1 line ${(100 * summary.nearby).toFixed(1)} %  ` +
		`offset ${summary.meanOffset.toFixed(2)} (|${summary.meanAbsOffset.toFixed(2)}|)  drift/page ${summary.meanDrift.toFixed(2)}`
	)
}

function scaleFamily(prefix: string, scale: number): (manager: FontManager) => ReturnType<typeof adjustedFonts> {
	return manager =>
		adjustedFonts(manager, (request: FontRequest, font: ResolvedFont) =>
			request.family.toLowerCase().startsWith(prefix) ? { ...font, id: font.id + 100000, scale } : font
		)
}

function lineHeightOf(family: string, factor: number): (manager: FontManager) => ReturnType<typeof adjustedFonts> {
	return manager =>
		adjustedFonts(manager, (request: FontRequest, font: ResolvedFont) => {
			const m = font.metrics

			if (request.family.toLowerCase() !== family || m === null) return font

			return {
				...font,
				id: font.id + 200000,
				metrics: {
					...m,
					winAscent: m.winAscent * factor,
					winDescent: m.winDescent * factor,
					hheaAscender: m.hheaAscender * factor,
					hheaDescender: m.hheaDescender * factor,
					hheaLineGap: m.hheaLineGap * factor,
					typoAscender: m.typoAscender * factor,
					typoDescender: m.typoDescender * factor,
					typoLineGap: m.typoLineGap * factor
				}
			}
		})
}

describe.skipIf(!present)("layout against Word's recorded page breaks", () => {
	it("agrees on the page count of Word-authored files", async () => {
		const truths = groundTruth()
		const results: FileResult[] = []

		for (const truth of truths) results.push(await measure(truth, {}))

		const summary = summarize(results)
		const lines = [format("default", summary)]

		for (const result of results) {
			if (result.ours !== result.truth) {
				lines.push(
					`  ${result.name}: Word ${result.truth}, ours ${result.ours}, pages ${result.pageOffsets.join(" ")}, lines ${result.offsets.slice(0, 8).join(" ")}`
				)
			}
		}

		const slowest = results.reduce((max, result) => (result.ms > max.ms ? result : max), results[0] ?? { name: "", ms: 0 })

		lines.push(`slowest ${slowest.name} ${slowest.ms.toFixed(0)} ms`)
		console.log(lines.join("\n"))

		expect(summary.files).toBeGreaterThan(50)
		expect(summary.withinOne / summary.files).toBeGreaterThanOrEqual(0.9)
	}, 600000)

	it("lays out every corpus document", async () => {
		const failures: string[] = []
		const slow: string[] = []

		for (const name of readdirSync(DIRECTORY).sort()) {
			if (!name.endsWith(".docx")) continue

			const started = performance.now()

			try {
				const layout = await layOut(readBytes(join(DIRECTORY, name)), { maxPages: 5000 })

				expect(layout.final).toBe(true)
				expect(layout.pages.length).toBeGreaterThan(0)
			} catch (error) {
				// Encrypted and unreadable packages are the model's to report.
				if (!(error instanceof Error && /encrypted|notZip|unsupported|invalid/i.test(error.message)))
					failures.push(`${name}: ${String(error)}`)
			}

			const ms = performance.now() - started

			if (ms > 1000) slow.push(`${name} ${ms.toFixed(0)} ms`)
		}

		console.log(slow.join("\n"))
		expect(failures).toEqual([])
	}, 600000)

	it.runIf(process.env["FILEN_OFFICE_LAYOUT_COMPARE"] !== undefined)(
		"compares two settings per file",
		async () => {
			const truths = groundTruth()
			const variant = JSON.parse(process.env["FILEN_OFFICE_LAYOUT_COMPARE"] ?? "{}") as LayoutOptions
			const lines: string[] = []

			for (const truth of truths) {
				const a = await measure(truth, {})
				const b = await measure(truth, variant)

				if (a.ours !== b.ours || a.offsets.join() !== b.offsets.join()) {
					lines.push(
						`${truth.name.slice(0, 60).padEnd(60)} Word ${truth.pages}  base ${a.ours} [${a.offsets.slice(0, 6).join(" ")}]  variant ${b.ours} [${b.offsets.slice(0, 6).join(" ")}]`
					)
				}
			}

			console.log(lines.join("\n"))
		},
		3600000
	)

	it.runIf(process.env["FILEN_OFFICE_LAYOUT_SWEEP"] === "1")(
		"overrideTableStyleFontSizeAndJustification under compatibility mode 15",
		async () => {
			const truths = groundTruth(true)
			const on = (settings: StyleSettings): StyleSettings =>
				settings.compatibilityMode === 15 ? { ...settings, overrideTableStyleFontSizeAndJustification: true } : settings
			const lines: string[] = []
			const base: FileResult[] = []
			const variant: FileResult[] = []

			for (const truth of truths) {
				const a = await measure(truth, {})
				const b = await measure(truth, {}, undefined, on)

				base.push(a)
				variant.push(b)

				if (a.ours !== b.ours || a.pageOffsets.join() !== b.pageOffsets.join()) {
					lines.push(
						`  ${truth.name.slice(0, 60).padEnd(60)} Word ${truth.pages}  spec ${a.ours} [${a.pageOffsets.join(" ")}]  on ${b.ours} [${b.pageOffsets.join(" ")}]`
					)
				}
			}

			console.log(
				[format("as MS-DOCX specifies", summarize(base)), format("on under compat 15", summarize(variant)), ...lines].join("\n")
			)
		},
		3600000
	)

	it.runIf(process.env["FILEN_OFFICE_LAYOUT_SWEEP"] === "1")(
		"calibration sweep",
		async () => {
			const truths = groundTruth()
			const run = async (
				label: string,
				subset: readonly Truth[],
				options: LayoutOptions,
				fonts?: (manager: FontManager) => ReturnType<typeof adjustedFonts>
			) => {
				const results: FileResult[] = []

				for (const truth of subset) results.push(await measure(truth, options, fonts))

				console.log(format(label, summarize(results)))
			}

			// Each choice below is decided by one metric fixed in advance: the share of Word's recorded breaks our
			// layout breaks at on exactly the same line; ties go to the value closer to the default.
			await run("gdi", truths, { lineSpacingRule: "gdi" })
			await run("typoWhenFlagged", truths, { lineSpacingRule: "typoWhenFlagged" })

			for (const mode of ["overflow", "breaks", "compat", "never"] as const)
				await run(`page top space dropped: ${mode}`, truths, { tuning: { dropSpaceAtPageTop: mode } })

			for (const mode of ["sum", "max", "auto"] as const)
				await run(`paragraph spacing: ${mode}`, truths, { tuning: { paragraphSpacing: mode } })

			const aptos = groundTruth(true).filter(truth => truth.minorFont.startsWith("Aptos"))

			for (const scale of [1, 1.01, 1.015, 1.02, 1.025, 1.03, 1.04, 1.05, 1.06, 1.07]) {
				await run(`aptos ×${scale} (${aptos.length})`, aptos, {}, scaleFamily("aptos", scale))

				if (process.env["FILEN_OFFICE_LAYOUT_DETAIL"] === "1") {
					for (const truth of aptos) {
						const result = await measure(truth, {}, scaleFamily("aptos", scale))

						console.log(
							`    ${truth.name.slice(0, 50)} Word ${truth.pages} ours ${result.ours} pages [${result.pageOffsets.join(" ")}] lines [${result.offsets.join(" ")}]`
						)
					}
				}
			}

			// Line height per font family: which factor on our line metrics fits Word's breaks best, over the
			// files using the family anywhere.
			const byFamily = new Map<string, Truth[]>()

			for (const truth of truths) {
				for (const family of (await measure(truth, {})).families) byFamily.set(family, [...(byFamily.get(family) ?? []), truth])
			}

			const wanted = (
				process.env["FILEN_OFFICE_LAYOUT_FAMILIES"] ?? "calibri,cambria,georgia,segoe ui,arial,times new roman,tahoma"
			).split(",")

			for (const family of wanted) {
				const subset = byFamily.get(family) ?? []

				if (subset.length === 0) continue

				for (const factor of [0.88, 0.92, 0.96, 1, 1.04, 1.08])
					await run(`${family} line ×${factor} (${subset.length})`, subset, {}, lineHeightOf(family, factor))
			}
		},
		3600000
	)
})
