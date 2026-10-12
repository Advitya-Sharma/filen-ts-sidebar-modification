import { createHash } from "node:crypto"
import { readFileSync } from "node:fs"
import type { PageDisplayList } from "@filen/office/core/docx"
import { FRAGMENT_FLOATS, FRAGMENT_INTS, LINE_FLOATS, LINE_INTS, LineInt, type StyleTable } from "@filen/office/core/text"
import { layOut } from "../layoutFixtures"
import { pageDigest } from "../pagedFixtures"

// Page digests: the paged fuzz compares what the editor shows with a fresh layout of the saved file, and the
// viewer digest compares the corpus's unedited layout before and after a change.

function round(value: number | undefined): number {
	return Math.round((value ?? 0) * 10) / 10
}

// A page's lines (geometry to 0.1 pt, flags, region, text) and fragments (x, text range, font, size, info).
export function digestPage(styles: StyleTable, page: PageDisplayList): string {
	const text = page.text
	const out = [`#${page.number}|${page.numberText} ${round(page.width)}x${round(page.height)} lines ${text.lineCount}`]

	for (let line = 0; line < text.lineCount; line++) {
		const f = line * LINE_FLOATS
		const i = line * LINE_INTS
		const source = text.lineInts[i + LineInt.SOURCE] ?? -1
		const content = page.sourceTexts[source] ?? ""
		const start = text.lineInts[i + LineInt.TEXT_START] ?? 0
		const end = text.lineInts[i + LineInt.TEXT_END] ?? 0
		const parts: (string | number)[] = [
			round(text.lineFloats[f]),
			round(text.lineFloats[f + 1]),
			round(text.lineFloats[f + 2]),
			round(text.lineFloats[f + 3]),
			round(text.lineFloats[f + 4]),
			text.lineInts[i + LineInt.FLAGS] ?? 0,
			page.sourceRegion(source),
			JSON.stringify(content.slice(start, end))
		]
		const first = text.lineInts[i + LineInt.FIRST_FRAGMENT] ?? 0
		const fragments = text.lineInts[i + LineInt.FRAGMENT_COUNT] ?? 0

		for (let k = first; k < first + fragments; k++) {
			const style = styles.get(text.fragmentInts[k * FRAGMENT_INTS + 2] ?? 0)
			const from = (text.fragmentInts[k * FRAGMENT_INTS] ?? 0) - start
			const to = (text.fragmentInts[k * FRAGMENT_INTS + 1] ?? 0) - start

			parts.push(
				`[${round(text.fragmentFloats[k * FRAGMENT_FLOATS])} ${from}-${to} ${style.font.family}:${style.font.weight}:${String(style.font.italic)}/${style.size} ${text.fragmentInts[k * FRAGMENT_INTS + 3] ?? 0}]`
			)
		}

		out.push(parts.join(" "))
	}

	return out.join("\n")
}

function lineDifference(a: string, b: string): string {
	const x = a.split("\n")
	const y = b.split("\n")

	for (let i = 0; i < Math.max(x.length, y.length); i++)
		if (x[i] !== y[i]) return `  line ${i}\n  shown: ${(x[i] ?? "").slice(0, 400)}\n  fresh: ${(y[i] ?? "").slice(0, 400)}`

	return "?"
}

// The first difference between two page lists, or undefined when they are equal.
export function firstDifference(shown: readonly string[], fresh: readonly string[]): string | undefined {
	for (let i = 0; i < Math.min(shown.length, fresh.length); i++) {
		if (shown[i] === fresh[i]) continue

		const where =
			shown.length === fresh.length ? `page ${i}` : `page count ${shown.length} vs ${fresh.length}; first difference page ${i}`

		return `${where}:\n${lineDifference(shown[i] ?? "", fresh[i] ?? "")}`
	}

	return shown.length === fresh.length ? undefined : `page count ${shown.length} vs ${fresh.length}`
}

// ---- viewer digest ----

// A file's page count and a hash of its pages (pagedFixtures' digest), or why it does not lay out.
export async function viewerDigest(file: string): Promise<string> {
	try {
		const pages = pageDigest(await layOut(new Uint8Array(readFileSync(file))))

		return `${pages.length}:${createHash("sha1").update(pages.join("\n")).digest("hex")}`
	} catch (error) {
		return `ERR ${error instanceof Error ? error.message.slice(0, 80) : String(error)}`
	}
}
