import { copyFileSync, existsSync, mkdirSync, readFileSync, statSync, writeFileSync } from "node:fs"
import { basename, dirname, extname, join } from "node:path"
import { unzipSync } from "fflate"
import { CORPUS } from "../fuzz/config"
import { openEditable } from "../fuzz/harness"
import { summarizeDocument, type DocumentSummary, type ManifestEntry, type SaveSidecar } from "../fuzz/saves"
import { WITNESSES } from "./accepted"
import { summarizeFodt, type OdfSummary } from "./odfSummary"
import { ourPageCount, pdfPageCount, wordPageCount } from "./pdfText"
import { convert, outputPath, profileDir, removeProfile, type Conversion, type Format, type Office } from "./soffice"
import { TIERS, compareDeltas, normalizeLines, type DeltaComparison, type Tier } from "./textDelta"

// The LibreOffice interop oracle: edited saves the fuzzers and the scripted suite kept (fuzz/saves.ts),
// and the corpus files they came from, converted by LibreOffice to PDF, ODT, flat ODT, plain text and DOCX,
// and the ODT back to DOCX. Checks, each a difference between the original and the edited file so that
// LibreOffice's own reading of a document cancels out:
//   L1  every edited save converts to every format, within the timeout
//   L2  the text an edit added and removed reads the same in LibreOffice as in our model (changes accepted,
//       changes rejected, and changes shown in LibreOffice's plain text)
//   L3  the structure an edit added and removed (paragraphs, tables, notes, comments...) counts the same
//   L4  page counts, LibreOffice's against ours and Word's (a report, never a gate)
//   G7  LibreOffice's DOCX of an edited save, and its DOCX from the ODT, open in our engine with no gate reason
//       the original's do not have, and the edit reads the same in our text of LibreOffice's DOCX files as in
//       LibreOffice's own text of the original and the edited save

export const FORMATS_ALL: readonly Format[] = ["pdf", "odt", "fodt", "txt", "docx"]

// Structure counts compared in L3, our sidecar's names. Gated ones must match; the rest are reported.
export const L3_GATED: readonly string[] = ["bodyParagraphs", "tables", "rows", "comments", "footnotes", "endnotes", "toc", "equations"]
export const L3_REPORTED: readonly string[] = ["cells", "images", "insertions", "deletions", "bookmarks", "headers", "footers"]

export interface InteropOptions {
	readonly saves: string
	// From selectEntries, once for every office (the manifest grows while fuzzers run).
	readonly entries: readonly ManifestEntry[]
	readonly work: string
	readonly jobs: number
	readonly timeoutMs: number
	readonly batch: number
	// Edited saves this big or bigger are not laid out by us for L4.
	readonly layoutMaxBytes: number
}

export interface EntryResult {
	readonly path: string
	readonly source: string
	readonly family: string
	readonly l1: readonly string[]
	readonly l1Original: readonly string[]
	readonly l2: Readonly<Record<string, Readonly<Record<Tier, DeltaComparison>>>>
	readonly l3: Readonly<Record<string, { readonly theirs: number; readonly ours: number }>>
	readonly g7: { readonly gate?: string; readonly backGate?: string; readonly text?: DeltaComparison; readonly error?: string }
	// Our reading's witness counts (accepted.ts WITNESSES) of the original and the edited file.
	readonly witness: { readonly before: Readonly<Record<string, number>>; readonly after: Readonly<Record<string, number>> }
	// Pages of the original (LibreOffice, ours, Word's when it last saved) and of the edited save.
	readonly l4: {
		readonly loBefore?: number
		readonly loAfter?: number
		readonly oursBefore?: number
		readonly ours?: number
		readonly word?: number
	}
}

export interface OfficeReport {
	readonly office: string
	readonly entries: readonly EntryResult[]
	readonly originals: number
	readonly conversions: number
	readonly ms: number
}

function spread<T>(list: readonly T[], count: number): T[] {
	if (list.length <= count) return [...list]

	return Array.from({ length: count }, (_, k) => list[Math.floor((k * list.length) / count)] as T)
}

// The saves to test: every scripted one and `perFamily` of each other family, at most `limit` in all.
export function selectEntries(saves: string, perFamily: number, limit: number | undefined): ManifestEntry[] {
	const manifest = JSON.parse(readFileSync(join(saves, "manifest.json"), "utf8")) as { entries: ManifestEntry[] }
	const byFamily = new Map<string, ManifestEntry[]>()

	for (const entry of manifest.entries) {
		const list = byFamily.get(entry.family) ?? []

		list.push(entry)
		byFamily.set(entry.family, list)
	}

	const out: ManifestEntry[] = []

	for (const [family, list] of [...byFamily].sort(([a], [b]) => (a < b ? -1 : 1))) {
		// Evenly spread over the family's files (the manifest is sorted by name).
		out.push(...(family === "script" ? list : spread(list, perFamily)))
	}

	return limit === undefined ? out : spread(out, limit)
}

interface Task {
	readonly inputs: string[]
	readonly outdir: string
	readonly format: Format
}

// Output directories: one per input directory and format, so outputs never collide.
function outdirOf(work: string, office: Office, set: string, group: string, format: Format): string {
	return join(work, office.name, set, group, format)
}

async function runTasks(office: Office, tasks: readonly Task[], options: InteropOptions): Promise<Map<string, Conversion>> {
	const results = new Map<string, Conversion>()
	let next = 0
	const worker = async (index: number): Promise<void> => {
		const profile = profileDir(join(options.work, office.name), index)

		try {
			while (next < tasks.length) {
				const task = tasks[next++] as Task

				for (const conversion of await convert(office, profile, task.format, task.inputs, task.outdir, options.timeoutMs))
					results.set(`${task.outdir}\u0000${conversion.input}`, conversion)
			}
		} finally {
			removeProfile(profile)
		}
	}

	await Promise.all(Array.from({ length: options.jobs }, (_, index) => worker(index)))

	return results
}

function batches(inputs: readonly string[], size: number): string[][] {
	const out: string[][] = []

	for (let i = 0; i < inputs.length; i += size) out.push(inputs.slice(i, i + size))

	return out
}

function read(path: string): Uint8Array | undefined {
	return existsSync(path) ? new Uint8Array(readFileSync(path)) : undefined
}

const fodtCache = new Map<string, OdfSummary | undefined>()

function fodtSummary(path: string): OdfSummary | undefined {
	if (!fodtCache.has(path)) {
		const bytes = read(path)

		fodtCache.set(path, bytes === undefined ? undefined : summarizeFodt(bytes))
	}

	return fodtCache.get(path)
}

async function gateOf(bytes: Uint8Array | undefined): Promise<string> {
	if (bytes === undefined) return "no file"

	try {
		const opened = await openEditable(bytes)

		return "refused" in opened ? opened.refused : ""
	} catch (error) {
		return `threw (${error instanceof Error ? error.message.slice(0, 200) : String(error)})`
	}
}

function tiers(
	theirsBefore: string,
	theirsAfter: string,
	oursBefore: string,
	oursAfter: string,
	labels: boolean
): Record<Tier, DeltaComparison> {
	const tb = normalizeLines(theirsBefore, { labels })
	const ta = normalizeLines(theirsAfter, { labels })
	const ob = normalizeLines(oursBefore)
	const oa = normalizeLines(oursAfter)

	return Object.fromEntries(TIERS.map(tier => [tier, compareDeltas(tb, ta, ob, oa, tier)])) as Record<Tier, DeltaComparison>
}

function documentXml(bytes: Uint8Array): string {
	try {
		const entries = unzipSync(bytes, { filter: file => file.name === "word/document.xml" })
		const xml = entries["word/document.xml"]

		return xml === undefined ? "" : new TextDecoder().decode(xml)
	} catch {
		return ""
	}
}

export async function runOffice(office: Office, options: InteropOptions, log: (line: string) => void = () => {}): Promise<OfficeReport> {
	const started = performance.now()
	const entries = options.entries
	const sources = [...new Set(entries.map(entry => entry.source))].sort()
	const tasks: Task[] = []
	const groupOf = (path: string): string => (path.startsWith(CORPUS) ? extname(path).slice(1) : basename(dirname(path)))
	const setOf = (path: string): string => (path.startsWith(CORPUS) ? "orig" : "edit")
	const addTasks = (paths: readonly string[], formats: readonly Format[]): void => {
		const groups = new Map<string, string[]>()

		for (const path of paths) {
			const key = `${setOf(path)}/${groupOf(path)}`
			const list = groups.get(key) ?? []

			list.push(stage(path))
			groups.set(key, list)
		}

		for (const [key, list] of groups) {
			const [set, group] = key.split("/") as [string, string]

			for (const format of formats)
				for (const inputs of batches(list, options.batch))
					tasks.push({ inputs, outdir: outdirOf(options.work, office, set, group, format), format })
		}
	}
	// LibreOffice opens copies only: never a file of the corpus, the repository or the shared saves.
	const staged = (path: string): string => join(options.work, "inputs", setOf(path), groupOf(path), basename(path))
	const stage = (path: string): string => {
		const copy = staged(path)

		if (!existsSync(copy) || statSync(copy).mtimeMs < statSync(path).mtimeMs) {
			mkdirSync(dirname(copy), { recursive: true })
			copyFileSync(path, copy)
		}

		return copy
	}
	const originals = sources.map(source => join(CORPUS, source))
	const edited = entries.map(entry => join(options.saves, entry.path))

	addTasks([...originals, ...edited], FORMATS_ALL)
	log(`${office.name}: ${entries.length} saves of ${sources.length} originals, ${tasks.length} conversions`)

	const results = await runTasks(office, tasks, options)
	const conversion = (path: string, format: Format): { ok: boolean; output: string; error?: string } => {
		const outdir = outdirOf(options.work, office, setOf(path), groupOf(path), format)
		const found = results.get(`${outdir}\u0000${staged(path)}`)

		return {
			ok: found?.ok === true,
			output: outputPath(outdir, path, format),
			...(found?.error === undefined ? {} : { error: found.error })
		}
	}

	// And back: LibreOffice's ODT of each edited save, to DOCX again.
	const backInputs = edited
		.map(path => conversion(path, "odt"))
		.filter(item => item.ok)
		.map(item => item.output)
	const backTasks: Task[] = []
	const backGroups = new Map<string, string[]>()

	for (const input of backInputs) {
		const group = basename(dirname(dirname(input)))
		const list = backGroups.get(group) ?? []

		list.push(input)
		backGroups.set(group, list)
	}

	for (const [group, list] of backGroups)
		for (const inputs of batches(list, options.batch))
			backTasks.push({ inputs, outdir: outdirOf(options.work, office, "back", group, "docx"), format: "docx" })

	const backResults = await runTasks(office, backTasks, options)
	const out: EntryResult[] = []
	const originalGate = new Map<string, Promise<string>>()
	const originalResave = new Map<string, Promise<DocumentSummary | undefined>>()
	// Our model's reading of each file, from its bytes (the sidecars' copy may predate a fix to the reading).
	const summaries = new Map<string, Promise<DocumentSummary | undefined>>()
	const ourSummary = (file: string): Promise<DocumentSummary | undefined> => {
		let summary = summaries.get(file)

		if (summary === undefined) {
			summary = summarizeDocument(read(file) as Uint8Array)
			if (file.startsWith(CORPUS)) summaries.set(file, summary)
		}

		return summary
	}
	const wordPages = new Map<string, number | undefined>()
	const ourPages = new Map<string, Promise<number | undefined>>()
	const pagesOf = (file: string, bytes: number): Promise<number | undefined> => {
		if (bytes >= options.layoutMaxBytes) return Promise.resolve(undefined)

		return ourPageCount(read(file) as Uint8Array).catch(() => undefined)
	}

	for (const [index, entry] of entries.entries()) {
		const path = join(options.saves, entry.path)
		const original = join(CORPUS, entry.source)
		const sidecar = JSON.parse(readFileSync(join(options.saves, entry.sidecar), "utf8")) as SaveSidecar
		const oursBefore = await ourSummary(original)
		const oursAfter = await ourSummary(path)

		if (oursBefore === undefined || oursAfter === undefined) throw new Error(`${entry.path}: our model cannot read it`)

		const l1 = FORMATS_ALL.filter(format => !conversion(path, format).ok).map(
			format => `${format}: ${conversion(path, format).error ?? "failed"}`
		)
		const l1Original = FORMATS_ALL.filter(format => !conversion(original, format).ok).map(format => format)
		const l2: Record<string, Record<Tier, DeltaComparison>> = {}
		const l3: Record<string, { theirs: number; ours: number }> = {}
		const before = fodtSummary(conversion(original, "fodt").output)
		const after = conversion(path, "fodt").ok ? fodtSummary(conversion(path, "fodt").output) : undefined

		if (before !== undefined && after !== undefined && conversion(original, "fodt").ok) {
			l2["final"] = tiers(before.final, after.final, oursBefore.final, oursAfter.final, false)
			l2["original"] = tiers(before.original, after.original, oursBefore.original, oursAfter.original, false)

			for (const metric of [...L3_GATED, ...L3_REPORTED]) {
				const theirs = (after.counts[metric] ?? 0) - (before.counts[metric] ?? 0)
				const ours = (oursAfter.counts[metric] ?? 0) - (oursBefore.counts[metric] ?? 0)

				if (theirs !== ours) l3[metric] = { theirs, ours }
			}
		}

		const txtBefore = conversion(original, "txt")
		const txtAfter = conversion(path, "txt")

		if (txtBefore.ok && txtAfter.ok)
			l2["markup"] = tiers(
				readFileSync(txtBefore.output, "utf8"),
				readFileSync(txtAfter.output, "utf8"),
				oursBefore.markup,
				oursAfter.markup,
				true
			)

		// G7
		let g7: EntryResult["g7"] = {}

		try {
			const resaved = conversion(path, "docx")
			const originalResaved = conversion(original, "docx")
			const back = backResults.get(
				`${outdirOf(options.work, office, "back", entry.family, "docx")}\u0000${conversion(path, "odt").output}`
			)
			const backPath = outputPath(
				outdirOf(options.work, office, "back", entry.family, "docx"),
				conversion(path, "odt").output,
				"docx"
			)

			if (!originalGate.has(original))
				originalGate.set(original, originalResaved.ok ? gateOf(read(originalResaved.output)) : Promise.resolve("no file"))

			const was = await originalGate.get(original)
			const gate = resaved.ok ? await gateOf(read(resaved.output)) : "no file"
			const backGate = back?.ok === true ? await gateOf(read(backPath)) : "no file"
			const ours = resaved.ok ? await summarizeDocument(read(resaved.output) as Uint8Array) : undefined

			if (!originalResave.has(original))
				originalResave.set(
					original,
					originalResaved.ok ? summarizeDocument(read(originalResaved.output) as Uint8Array) : Promise.resolve(undefined)
				)

			const oursOriginal = await originalResave.get(original)
			const newGate = (value: string): string | undefined => (value === "" || value === was ? undefined : value)

			g7 = {
				...(newGate(gate) === undefined ? {} : { gate }),
				...(newGate(backGate) === undefined ? {} : { backGate }),
				// As a delta too: what LibreOffice writes in place of what (a frame for a framed paragraph, a field it
				// cannot resolve) is in both files.
				...(ours === undefined || oursOriginal === undefined || after === undefined || before === undefined
					? {}
					: {
							text: compareDeltas(
								normalizeLines(before.final),
								normalizeLines(after.final),
								normalizeLines(oursOriginal.final),
								normalizeLines(ours.final),
								"words"
							)
						})
			}
		} catch (error) {
			g7 = { error: error instanceof Error ? error.message.slice(0, 300) : String(error) }
		}

		// L4
		const l4: { loBefore?: number; loAfter?: number; oursBefore?: number; ours?: number; word?: number } = {}
		const pdfBefore = conversion(original, "pdf")
		const pdfAfter = conversion(path, "pdf")

		if (pdfBefore.ok) l4.loBefore = pdfPageCount(read(pdfBefore.output) as Uint8Array)
		if (pdfAfter.ok) l4.loAfter = pdfPageCount(read(pdfAfter.output) as Uint8Array)

		if (!wordPages.has(original)) {
			const xml = documentXml(read(original) as Uint8Array)

			wordPages.set(original, xml.includes("lastRenderedPageBreak") ? wordPageCount(xml) : undefined)
		}

		const word = wordPages.get(original)

		if (word !== undefined) l4.word = word

		if (!ourPages.has(original)) ourPages.set(original, pagesOf(original, (read(original) as Uint8Array).length))

		const pagesBefore = await ourPages.get(original)
		const ours = await pagesOf(path, sidecar.bytes)

		if (pagesBefore !== undefined) l4.oursBefore = pagesBefore
		if (ours !== undefined) l4.ours = ours

		const witness = {
			before: Object.fromEntries(WITNESSES.map(name => [name, oursBefore.counts[name] ?? 0])),
			after: Object.fromEntries(WITNESSES.map(name => [name, oursAfter.counts[name] ?? 0]))
		}

		out.push({ path: entry.path, source: entry.source, family: entry.family, l1, l1Original, l2, l3, g7, l4, witness })

		if ((index + 1) % 100 === 0) log(`${office.name}: ${index + 1}/${entries.length} analysed`)
	}

	const report: OfficeReport = {
		office: office.name,
		entries: out,
		originals: sources.length,
		conversions: results.size + backResults.size,
		ms: Math.round(performance.now() - started)
	}

	mkdirSync(options.work, { recursive: true })
	writeFileSync(join(options.work, `report-${office.name}.json`), JSON.stringify(report, null, 1))

	return report
}

export interface Tally {
	readonly saves: number
	readonly l1: number
	readonly l1Original: number
	readonly l2: Readonly<Record<string, Readonly<Record<Tier, number>>>>
	readonly l3: Readonly<Record<string, number>>
	readonly g7Gate: number
	readonly g7Text: number
	readonly g7Error: number
	// Edited saves LibreOffice and we both laid out, and how many agree; originals Word laid out, and how many
	// LibreOffice and we agree with.
	readonly l4: {
		readonly compared: number
		readonly loEqualsOurs: number
		readonly withWord: number
		readonly loEqualsWord: number
		readonly oursEqualsWord: number
	}
}

// Failures counted per check (an entry may fail several).
export function tally(report: OfficeReport): Tally {
	const l2: Record<string, Record<Tier, number>> = {}
	const l3: Record<string, number> = {}
	let l1 = 0
	let l1Original = 0
	let g7Gate = 0
	let g7Text = 0
	let g7Error = 0
	const l4 = { compared: 0, loEqualsOurs: 0, withWord: 0, loEqualsWord: 0, oursEqualsWord: 0 }
	const originals = new Set<string>()

	for (const entry of report.entries) {
		if (entry.l1.length > 0) l1++
		if (entry.l1Original.length > 0) l1Original++

		for (const [view, byTier] of Object.entries(entry.l2)) {
			l2[view] ??= { lines: 0, words: 0, chars: 0 }

			for (const tier of TIERS) if (!byTier[tier].equal) (l2[view] as Record<Tier, number>)[tier]++
		}

		for (const metric of Object.keys(entry.l3)) l3[metric] = (l3[metric] ?? 0) + 1

		if (entry.g7.gate !== undefined || entry.g7.backGate !== undefined) g7Gate++
		if (entry.g7.text !== undefined && !entry.g7.text.equal) g7Text++
		if (entry.g7.error !== undefined) g7Error++

		if (entry.l4.loAfter !== undefined && entry.l4.ours !== undefined) {
			l4.compared++

			if (entry.l4.loAfter === entry.l4.ours) l4.loEqualsOurs++
		}

		if (entry.l4.word !== undefined && !originals.has(entry.source)) {
			originals.add(entry.source)
			l4.withWord++

			if (entry.l4.loBefore === entry.l4.word) l4.loEqualsWord++
			if (entry.l4.oursBefore === entry.l4.word) l4.oursEqualsWord++
		}
	}

	return { saves: report.entries.length, l1, l1Original, l2, l3, g7Gate, g7Text, g7Error, l4 }
}
