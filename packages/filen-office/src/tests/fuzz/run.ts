import { readFileSync } from "node:fs"
import { basename } from "node:path"
import type { Transaction } from "prosemirror-state"
import { protectionOf } from "../../core/docx/edit/protection"
import { CORPUS, fuzzConfig, nameHash, rng, type FuzzConfig, type Rand } from "./config"
import { firstDifference, viewerDigest } from "./digest"
import {
	bodySession,
	documents,
	duplicateSources,
	editorExecute,
	editorMarkers,
	editorSignatures,
	freshPages,
	isSaveRefusal,
	isSuspended,
	markerCounts,
	memoisedProof,
	openEditable,
	openPaged,
	otherStories,
	reopenProblem,
	resume,
	savePaged,
	saveSessions,
	schemaOracle,
	sessionDescription,
	sessionPristine,
	sessionUnchanged,
	settle,
	shownPages,
	storySession,
	viewName,
	type DocxEditorCommand,
	type DocxEditorSession,
	type Editable,
	type PagedMode,
	type PagedRun,
	type ParagraphSignature
} from "./harness"
import * as Reader from "./independentReader"
import { OpPicker, type OpContext } from "./ops/index"
import { inlinePosition, undoBurst } from "./ops/base"

// One fuzz job in its own process; the fuzz tests run these through config.runJobs. Prints JSON lines.
//   paged <file> <seed> <sync|race|proof>   one paged run: after each settled step the pages shown must
//                                           equal a fresh layout of the saved file
//   damage <file>                           every damage sequence of one file, verified by the independent reader
//   viewer <file>...                        the viewer digests of unedited files
// Both read FILEN_OFFICE_FUZZ_* (harness.fuzzConfig) for runs, seeds, steps, ops, modes and view.

function print(value: Record<string, unknown>): void {
	process.stdout.write(`${JSON.stringify(value)}\n`)
}

function relative(file: string): string {
	return file.startsWith(CORPUS) ? file.slice(CORPUS.length + 1) : basename(file)
}

function errorText(error: unknown, length: number): string {
	return (error instanceof Error ? (error.stack ?? error.message) : String(error)).slice(0, length)
}

// ---- paged ----

async function paged(config: FuzzConfig, file: string, seed: number, mode: PagedMode): Promise<void> {
	const started = performance.now()
	const name = relative(file)
	const steps = config.steps ?? 20
	const rand = rng(seed)
	const result = (outcome: Record<string, unknown>): void =>
		print({ file: name, seed, mode, ...outcome, ms: Math.round(performance.now() - started) })
	const opened = await openPaged(new Uint8Array(readFileSync(file)), mode, config.view)

	if ("refused" in opened) return result({ refused: opened.refused })

	const run: PagedRun = opened
	const editor = run.editor
	const picker = new OpPicker("paged", seed, config.ops)
	const pending: Promise<unknown>[] = []
	const contexts = new Map<Rand, OpContext>()
	const contextOf = (stream: Rand): OpContext => {
		let context = contexts.get(stream)

		if (context === undefined) {
			context = {
				rand: stream,
				get session() {
					return editor.session
				},
				editor,
				numIds: [],
				styleIds: [],
				transact(_label, build) {
					const session = editor.session
					const tr = session.state.tr

					return String(session.dispatch(build(tr) ?? tr))
				},
				execute(_label, command) {
					const done = editorExecute(editor, command)

					if (typeof done === "boolean") return String(done)

					pending.push(done)

					return "async"
				}
			}
			contexts.set(stream, context)
		}

		return context
	}
	const log: string[] = []
	// After each settled step (and at the end) the pages shown must equal a fresh layout of the saved file.
	const check = async (step: number | "end"): Promise<Record<string, unknown> | undefined> => {
		await settle(editor)

		const shown = shownPages(editor)
		let saved: Uint8Array

		try {
			saved = await savePaged(run)
		} catch (error) {
			return { step, saveFail: errorText(error, 300), log }
		}

		const difference = firstDifference(shown, await freshPages(saved, editor))

		return difference === undefined ? undefined : { step, view: viewName(editor), diverge: difference, log }
	}

	let applied = 0

	for (let step = 0; step < steps; step++) {
		const { op, rand: stream } = picker.next(rand)
		const before = documents(editor)

		try {
			log.push(await op.run(contextOf(stream)))

			// A command waiting for layout (a field update) waits while the editor is suspended, until it resumes.
			if (!isSuspended(editor)) await Promise.all(pending.splice(0))
		} catch (error) {
			return result({ step, crash: errorText(error, 800), log })
		}

		const after = documents(editor)

		if (after.length !== before.length || after.some((doc, index) => doc !== before[index])) applied++

		// Racing: the next edit lands while layout is still running.
		if (mode === "race" && step < steps - 1 && rand() < 0.7) {
			await new Promise(resolve => setImmediate(resolve))
			continue
		}

		// Suspended: nothing is laid out until it resumes.
		if (isSuspended(editor)) continue

		const failure = await check(step)

		if (failure !== undefined) return result(failure)
	}

	if (isSuspended(editor)) {
		resume(editor)
		log.push("resume")
	}

	try {
		await Promise.all(pending.splice(0))
	} catch (error) {
		return result({ step: "end", crash: errorText(error, 800), log })
	}

	const failure = await check("end")

	if (failure !== undefined) return result(failure)

	result({ ok: true, steps, applied, pages: editor.pageCount, protection: protectionOf(editor.body.session.context).mode })
}

// ---- damage ----

interface SourceFacts {
	readonly entries: Map<string, Reader.ZipEntry>
	readonly mainPart: string
	readonly document: Uint8Array
	readonly relationships: Map<string, Reader.Relationship>
	readonly signatures: ParagraphSignature[]
	readonly numberingIds: Set<string>
	readonly violations: Reader.Counts
	readonly markers: Map<string, number>
	readonly parts: Map<string, { numbering: Set<string>; problems: string[]; styles: Set<string> | undefined; styleProblems: string[] }>
	schemaErrors: Map<string, number> | undefined
}

const BIG = 40e6
// The kinds harness.markerCounts counts, and those no edit takes out.
const MARKER_KINDS: ReadonlySet<string> = new Set([
	"bookmarkStart",
	"bookmarkEnd",
	"commentRangeStart",
	"commentRangeEnd",
	"commentReference",
	"permStart",
	"permEnd",
	"moveFromRangeStart",
	"moveToRangeStart"
])
const KEPT_KINDS: ReadonlySet<string> = new Set(["bookmarkStart", "bookmarkEnd", "permStart", "permEnd"])

function read(entries: Map<string, Reader.ZipEntry>, name: string): Uint8Array | undefined {
	const entry = entries.get(name)

	return entry === undefined ? undefined : Reader.entryData(entry)
}

function relationshipsName(part: string): string {
	return part.replace(/([^/]+)$/, "_rels/$1.rels")
}

function resolveTarget(main: string, target: string): string {
	return target.startsWith("/") ? target.slice(1) : main.replace(/[^/]+$/, "") + target
}

// Paragraph signatures as the reader sees a written story; list ids the source lacks count as "NEW".
function readerSignatures(document: Uint8Array, numbering: ReadonlySet<string>): ParagraphSignature[] {
	const body = Reader.bodyOf(Reader.parseXml(document))

	if (body === undefined) return []

	return Reader.storyParagraphs(body).map(paragraph => {
		const signature = Reader.paragraphSignature(paragraph)
		const props = JSON.parse(signature.props) as unknown[]

		if (props[1] !== null && props[1] !== 0 && !numbering.has(String(props[1]))) props[1] = "NEW"

		return { props: JSON.stringify(props), text: signature.text, fmt: signature.fmt }
	})
}

const signatureKey = (signature: ParagraphSignature): string => `${signature.props}\u0001${signature.text}\u0001${signature.fmt}`

function numberingAndStyles(entries: Map<string, Reader.ZipEntry>, main: string, relationships: Map<string, Reader.Relationship>) {
	const values = [...relationships.values()]
	const numberingPart = values.find(relationship => relationship.type.endsWith("/numbering"))
	const stylesPart = values.find(relationship => relationship.type.endsWith("/styles"))
	const numbering = Reader.numberingOf(numberingPart === undefined ? undefined : read(entries, resolveTarget(main, numberingPart.target)))
	const styles = Reader.stylesOf(stylesPart === undefined ? undefined : read(entries, resolveTarget(main, stylesPart.target)))

	return {
		numbering: numbering.ids,
		problems: numbering.problems,
		styles: stylesPart === undefined ? undefined : styles.ids,
		styleProblems: styles.problems
	}
}

function sourceFacts(input: Uint8Array, mainPart: string): SourceFacts {
	const entries = Reader.zipEntries(input)
	const main = mainPart.replace(/^\//, "")
	const document = read(entries, main) ?? new Uint8Array()
	const relationships = Reader.relationshipsOf(read(entries, relationshipsName(main)))
	const parts = numberingAndStyles(entries, main, relationships)
	const big = document.length > BIG

	return {
		entries,
		mainPart: main,
		document,
		relationships,
		signatures: big ? [] : readerSignatures(document, parts.numbering),
		numberingIds: parts.numbering,
		violations: big
			? new Map()
			: Reader.violations(Reader.parseXml(document), { numIds: parts.numbering, styleIds: parts.styles, relationships }).counts,
		markers: big ? new Map() : markerCounts(document),
		parts: new Map([[main, parts]]),
		schemaErrors: undefined
	}
}

// Values of the relationships-namespace attributes of a part.
function relationshipIdsUsed(bytes: Uint8Array): Set<string> {
	const out = new Set<string>()
	const root = Reader.parseXml(bytes).root

	if (root === undefined) return out

	for (const element of [root, ...Reader.descendants(root)])
		for (const attribute of element.attributes) if (attribute.ns === Reader.R) out.add(attribute.value)

	return out
}

interface Damage {
	readonly log: string[]
	finding(kind: string, data: Record<string, unknown>): void
	bump(key: string): void
}

// `document` without its first bookmark start in a text box, else its first one (selftest).
function withoutMarker(document: Uint8Array): Uint8Array {
	const text = new TextDecoder().decode(document)
	const box = text.indexOf("<w:txbxContent")
	const find = (from: number): number => (from < 0 ? -1 : text.indexOf("<w:bookmarkStart ", from))
	const start = find(box) >= 0 ? find(box) : find(0)

	if (start < 0) return document

	return new TextEncoder().encode(text.slice(0, start) + text.slice(text.indexOf(">", start) + 1))
}

// What the independent reader finds wrong with a saved file that Gate 2 let through.
async function verifySaved(
	damage: Damage,
	source: SourceFacts,
	saved: { bytes: Uint8Array; changedParts: readonly string[] },
	session: DocxEditorSession,
	schema: ((bytes: Uint8Array) => Map<string, number>) | undefined,
	selftest = false
): Promise<void> {
	const log = damage.log
	const output = Reader.zipEntries(saved.bytes)
	const changed = new Set(saved.changedParts.map(part => part.replace(/^\//, "")))

	// Untouched entries keep their raw bytes.
	for (const [entry, before] of source.entries) {
		const after = output.get(entry)

		if (after === undefined) damage.finding("entryLost", { entry, log })
		else if (!changed.has(entry) && !Reader.sameRaw(before, after)) damage.finding("untouchedChanged", { entry, log })
	}

	for (const entry of output.keys()) if (!source.entries.has(entry) && !changed.has(entry)) damage.finding("unexpectedEntry", { entry })

	for (const entry of changed) {
		const before = source.entries.get(entry)
		const after = output.get(entry)

		if (
			before !== undefined &&
			after !== undefined &&
			Buffer.from(Reader.entryData(before)).equals(Buffer.from(Reader.entryData(after)))
		)
			damage.bump("changedButSameContent")
	}

	const reopen = await reopenProblem(saved.bytes)

	if (reopen !== undefined) damage.finding("reopenFailed", { problem: reopen, log })

	const main = source.mainPart
	const document = read(output, main)

	if (document === undefined) return damage.finding("mainPartLost", { log })

	if (source.document.length > BIG) return damage.bump("bigSkipped")

	for (const entry of changed) {
		if (!/\.(xml|rels)$/.test(entry)) continue

		const errors = Reader.parseXml(read(output, entry) ?? new Uint8Array()).errors
		const before = source.entries.has(entry) ? Reader.parseXml(read(source.entries, entry) ?? new Uint8Array()).errors.length : 0

		if (errors.length > before) damage.finding("xmlErrors", { entry, errors: errors.slice(0, 3), log })
	}

	const relationshipsOut = Reader.relationshipsOf(read(output, relationshipsName(main)))

	for (const id of relationshipsOut.keys())
		if (id.endsWith("#DUP") && !source.relationships.has(id)) damage.finding("relsDuplicateId", { id })

	// An external hyperlink relationship goes when the save no longer uses one the source used (Word drops them).
	let used: { before: Set<string>; after: Set<string> } | undefined
	const dropped = (id: string, relationship: Reader.Relationship): boolean => {
		if (!relationship.external || !relationship.type.endsWith("/hyperlink")) return false

		used ??= { before: relationshipIdsUsed(source.document), after: relationshipIdsUsed(document) }

		return used.before.has(id) && !used.after.has(id)
	}

	for (const [id, relationship] of source.relationships) {
		const after = relationshipsOut.get(id)

		if (after === undefined ? !dropped(id, relationship) : after.target !== relationship.target || after.type !== relationship.type)
			damage.finding("relChanged", { id })
	}

	const before = source.parts.get(main)
	const after = numberingAndStyles(output, main, relationshipsOut)

	for (const problem of after.problems) if (!before?.problems.includes(problem)) damage.finding("numberingBad", { problem, log })
	for (const problem of after.styleProblems) if (!before?.styleProblems.includes(problem)) damage.finding("stylesBad", { problem, log })

	const types = Reader.contentTypes(read(output, "[Content_Types].xml") ?? new Uint8Array())

	for (const entry of changed) {
		if (source.entries.has(entry)) continue

		const extension = entry.slice(entry.lastIndexOf(".") + 1).toLowerCase()

		if (!types.overrides.has(`/${entry.toLowerCase()}`) && !types.defaults.has(extension)) damage.finding("newPartUntyped", { entry })
	}

	// Word-rejection counts of the main part: only new violations count.
	const violations = Reader.violations(Reader.parseXml(document), {
		numIds: after.numbering,
		styleIds: after.styles,
		relationships: relationshipsOut
	})

	for (const [kind, times] of violations.counts) {
		const was = source.violations.get(kind) ?? 0

		if (times > was) damage.finding(`word:${kind}`, { before: was, after: times, sample: violations.samples.get(kind), log })
	}

	if (schema !== undefined && changed.has(main)) {
		source.schemaErrors ??= schema(source.document)

		for (const [message, times] of schema(document)) {
			const was = source.schemaErrors.get(message) ?? 0

			// A value the source already had wrong is copied with the run it is in (a split repeats it).
			if (was > 0 && message.includes("[facet ")) continue
			if (times > was) damage.finding("xsd", { message, before: was, after: times, log })
		}
	}

	// The file holds the markers the editor holds: comment commands add and take out comment ranges and
	// references, resolving a tracked move takes its ranges out. Bookmarks and permission ranges survive every
	// edit, so those are also checked against the source.
	const markers = markerCounts(selftest ? withoutMarker(document) : document)
	const held = source.document.length > BIG ? source.markers : editorMarkers(session)
	const expected = new Map([...held].filter(([key]) => MARKER_KINDS.has(key.slice(0, key.indexOf(":")))))

	for (const [key, times] of source.markers) {
		if (KEPT_KINDS.has(key.slice(0, key.indexOf(":"))) && (expected.get(key) ?? 0) < times) damage.finding("markerLost", { key, log })
	}

	for (const [key, times] of expected) if ((markers.get(key) ?? 0) < times) damage.finding("markerLost", { key, log })
	for (const [key, times] of markers) if (times > (expected.get(key) ?? 0)) damage.finding("markerGained", { key, log })

	// The editor's paragraphs appear in order in the output; every other output paragraph is a source paragraph.
	const editor = editorSignatures(session.state.doc)
	const written = readerSignatures(document, source.numberingIds)
	const matched = new Set<number>()
	let from = 0

	for (let i = 0; i < editor.length; i++) {
		const key = signatureKey(editor[i] as ParagraphSignature)
		let found = -1

		for (let k = from; k < written.length; k++) {
			if (signatureKey(written[k] as ParagraphSignature) === key) {
				found = k
				break
			}
		}

		if (found < 0) {
			const cut = (signature: ParagraphSignature): ParagraphSignature => ({
				...signature,
				text: signature.text.slice(0, 200),
				fmt: signature.fmt.slice(0, 200)
			})

			return damage.finding("semantic:editorParagraphMissing", {
				index: i,
				editor: cut(editor[i] as ParagraphSignature),
				near: written.slice(from, from + 3).map(cut),
				log
			})
		}

		matched.add(found)
		from = found + 1
	}

	const pool = new Map<string, number>()

	for (const signature of source.signatures) pool.set(signatureKey(signature), (pool.get(signatureKey(signature)) ?? 0) + 1)

	for (let k = 0; k < written.length; k++) {
		if (matched.has(k)) continue

		const key = signatureKey(written[k] as ParagraphSignature)
		const left = pool.get(key) ?? 0

		if (left <= 0) return damage.finding("semantic:unexplainedParagraph", { index: k, text: written[k]?.text.slice(0, 200), log })

		pool.set(key, left - 1)
	}
}

const RANGE_MARKERS = new Set([
	"bookmarkStart",
	"bookmarkEnd",
	"commentRangeStart",
	"commentRangeEnd",
	"moveFromRangeStart",
	"moveFromRangeEnd",
	"moveToRangeStart",
	"moveToRangeEnd",
	"permStart",
	"permEnd"
])

interface BodyBlock {
	// Its bytes.
	readonly key: string
	readonly start: number
	// Range markers in it (itself included): name and id, and where.
	readonly markers: readonly { readonly marker: string; readonly start: number; readonly end: number }[]
}

// The body's top-level elements.
function bodyBlocks(bytes: Uint8Array, main: string): BodyBlock[] {
	const document = Reader.parseXml(read(Reader.zipEntries(bytes), main) ?? new Uint8Array())
	const body = Reader.bodyOf(document)

	if (body === undefined) return []

	return Reader.childElements(body).map(child => ({
		key: document.source.slice(child.start, child.end),
		start: child.start,
		markers: [child, ...Reader.descendants(child)]
			.filter(element => element.ns === Reader.W && RANGE_MARKERS.has(element.local))
			.map(element => ({
				marker: `${element.local}:${Reader.attributeNS(element, Reader.W, "id") ?? ""}`,
				start: element.start,
				end: element.end
			}))
	}))
}

// What the writer drops when it rewrites a paragraph: w14:textId (Word's hash of the text) and the layout hints
// w:lastRenderedPageBreak and w:proofErr.
function rewritten(key: string): string {
	return key.replace(/\s[\w.-]+:textId="[^"]*"/g, "").replace(/<w:(?:lastRenderedPageBreak|proofErr)\b[^>]*\/>/g, "")
}

// Top-level elements of the body that differ between `input` and `saved` though block moves keep every block's
// bytes. A range a move turned around (its end above its start) is put back in order by the editor, collapsed
// where its end went: the markers of ranges the save holds in other blocks than the source are left out of the
// comparison, and a block that held one, or holds one now, may come back rewritten.
function blockMoveDifferences(input: Uint8Array, saved: Uint8Array, main: string): number {
	const before = bodyBlocks(input, main)
	const after = bodyBlocks(saved, main)
	const holders = (blocks: readonly BodyBlock[]): Map<string, string[]> => {
		const out = new Map<string, string[]>()

		for (const block of blocks) {
			for (const { marker } of block.markers) {
				const list = out.get(marker)

				if (list === undefined) out.set(marker, [block.key])
				else list.push(block.key)
			}
		}

		return out
	}
	const was = holders(before)
	const is = holders(after)
	const moved = new Set<string>()

	for (const marker of new Set([...was.keys(), ...is.keys()])) {
		if ([...(was.get(marker) ?? [])].sort().join("\u0000") !== [...(is.get(marker) ?? [])].sort().join("\u0000")) moved.add(marker)
	}

	// A block's bytes without the markers of moved ranges ("" for a block that was only such a marker), and
	// whether it held one.
	const stripped = (block: BodyBlock): { key: string; held: boolean } => {
		let key = block.key
		let held = false

		for (const { marker, start, end } of [...block.markers].reverse()) {
			if (!moved.has(marker)) continue

			key = key.slice(0, start - block.start) + key.slice(end - block.start)
			held = true
		}

		return { key, held }
	}
	const a = new Map<string, number>()
	const b = new Map<string, number>()
	const add = (counts: Map<string, number>, key: string): void => {
		if (key !== "") counts.set(key, (counts.get(key) ?? 0) + 1)
	}
	// Stripped bytes of the source's blocks, kept as they were or (`held`) as the writer rewrites them: a block
	// that held such a marker, or that holds one now (one came into it).
	const kept = new Set<string>()
	const held = new Set<string>()
	const strippedAfter = after.map(stripped)
	const receivers = new Set(strippedAfter.filter(block => block.held).map(block => rewritten(block.key)))

	for (const block of before) {
		const stripping = stripped(block)
		const loose = stripping.held || receivers.has(rewritten(stripping.key))

		if (loose) held.add(stripping.key)
		else kept.add(stripping.key)

		add(a, loose ? rewritten(stripping.key) : stripping.key)
	}

	for (const { key, held: holds } of strippedAfter) add(b, !holds && kept.has(key) && !held.has(key) ? key : rewritten(key))

	let differences = 0

	for (const [key, times] of a) if ((b.get(key) ?? 0) !== times) differences++
	for (const key of b.keys()) if (!a.has(key)) differences++

	return differences
}

async function damage(config: FuzzConfig, file: string): Promise<void> {
	const started = performance.now()
	const name = relative(file)
	const input = new Uint8Array(readFileSync(file))
	const editable = await openEditable(input)
	const stats: Record<string, number> = {}
	const bump = (key: string): void => {
		stats[key] = (stats[key] ?? 0) + 1
	}

	if ("refused" in editable) return print({ file: name, refused: editable.refused })

	const modes = new Set(config.modes ?? ["stories"])
	const names = modes.has("blockmove") ? ["blockMove"] : config.ops
	const most = config.steps ?? 16
	const proof = memoisedProof(editable.context)
	const source = sourceFacts(input, editable.mainPart)
	const schema = modes.has("xsd") ? schemaOracle() : undefined
	const base = (Math.imul(config.seed, 100_003) ^ nameHash(name)) >>> 0
	// What ops change in the edit context the sequences share (tracking switched on, an author named): each
	// sequence starts from the file's.
	const { review, identity } = editable.context
	const opened = { tracking: review.tracking, persist: review.persist, author: identity.author, initials: identity.initials }
	// The body's top-level blocks (blockmove mode has nothing to move with fewer than two).
	let blocks = 0

	for (let q = 0; q < config.runs; q++) {
		const seed = (base + q) >>> 0

		review.tracking = opened.tracking
		review.persist = opened.persist
		identity.author = opened.author
		identity.initials = opened.initials

		const rand = rng(seed)
		const log: string[] = []
		const report: Damage = {
			log,
			finding(kind, data) {
				print({ finding: kind, file: name, seed, ...data })
			},
			bump
		}
		let sessions: DocxEditorSession[]

		try {
			sessions = [bodySession(editable, proof)]
			blocks = sessions[0]?.state.doc.childCount ?? 0

			if (modes.has("stories")) {
				for (const story of otherStories(editable)) {
					if (rand() < 0.5) {
						const session = storySession(editable, story, proof)

						sessions.push(session)
						log.push(sessionDescription(session))
					}
				}
			}
		} catch (error) {
			report.finding("sessionThrew", { error: errorText(error, 1500) })
			continue
		}

		// The first sequence of each file saves unedited: it must come back byte for byte.
		const count = q === 0 ? 0 : 1 + Math.floor(rand() * most)

		const picker = new OpPicker("damage", seed, names)

		for (const session of sessions) {
			const contexts = new Map<Rand, OpContext>()
			const contextOf = (stream: Rand): OpContext => {
				const context = contexts.get(stream) ?? damageContext(editable, session, stream, report)

				contexts.set(stream, context)

				return context
			}

			for (let i = 0; i < count; i++) {
				const before = session.state.doc

				try {
					if (modes.has("undo") && rand() < 0.35) log.push(undoBurst(contextOf(rand)))
					else {
						const { op, rand: stream } = picker.next(rand)

						log.push(await op.run(contextOf(stream)))
					}
				} catch (error) {
					report.finding("opThrew", { error: errorText(error, 1500), log })
				}

				if (session.state.doc !== before) bump("applied")
			}

			if (sessions.length > 1) log.push("--- next session")
		}

		bump("seq")

		const duplicates = sessions.flatMap(duplicateSources)
		let saved: { bytes: Uint8Array; changedParts: readonly string[] }

		if (duplicates.length > 0) bump("dupSrc")

		try {
			saved = await saveSessions(editable, sessions)
		} catch (error) {
			if (isSaveRefusal(error)) {
				bump(`refused.${error.reason}`)

				if (error.reason !== "readOnly") report.finding(`refused:${error.reason}`, { message: error.message.slice(0, 1500), log })
			} else {
				bump("saveThrew")
				report.finding("saveThrew", { error: errorText(error, 2000), log })
			}

			continue
		}

		bump("saved")

		if (modes.has("blockmove")) {
			const differences = blockMoveDifferences(input, saved.bytes, source.mainPart)

			if (differences > 0) report.finding("blockMoveChangedBytes", { differences, log })
		}

		if (duplicates.length > 0) report.finding("savedWithDuplicateSrc", { types: [...new Set(duplicates)], log })
		if (count === 0 && !Buffer.from(saved.bytes).equals(Buffer.from(input))) report.finding("uneditedNotIdentical", {})
		if (sessions.every(sessionPristine) && saved.changedParts.length > 0)
			report.finding("undoneButChanged", { parts: saved.changedParts, log })

		bump("verified")

		const body = sessions[0] as DocxEditorSession

		// Checks the checker: text typed after the save is in the editor but not the file, and a bookmark the
		// file loses (one in a text box first), so every verified sequence must report them.
		if (modes.has("selftest")) {
			const position = inlinePosition(rand, body.state.doc)

			if (position !== undefined) body.dispatch(body.state.tr.insertText("SELFTEST", position))
		}

		try {
			await verifySaved(report, source, saved, body, schema, modes.has("selftest"))
		} catch (error) {
			report.finding("verifierThrew", { error: errorText(error, 800) })
		}

		if (modes.has("undoall")) {
			for (const session of sessions) while (session.execute({ type: "history.undo" })) continue

			try {
				const again = await saveSessions(editable, sessions)
				const unchanged = sessions.every(sessionUnchanged)

				if (unchanged && !Buffer.from(again.bytes).equals(Buffer.from(input)))
					report.finding("undoAllNotIdentical", { parts: again.changedParts, log })
				if (!unchanged) bump("undoAllNotEq")
			} catch (error) {
				report.finding("undoAllSaveThrew", { error: errorText(error, 300), log })
			}
		}
	}

	print({ file: name, stats, blocks, protection: protectionOf(editable.context).mode, ms: Math.round(performance.now() - started) })
}

function damageContext(editable: Editable, session: DocxEditorSession, rand: Rand, report: Damage): OpContext {
	return {
		rand,
		session,
		editor: undefined,
		numIds: editable.numIds,
		styleIds: editable.styleIds,
		transact(label, build: (tr: Transaction) => Transaction | void) {
			let tr: Transaction

			try {
				tr = session.state.tr
				tr = build(tr) ?? tr
			} catch (error) {
				report.bump("op.buildThrew")

				return `(build threw ${String(error).slice(0, 60)})`
			}

			try {
				const applied = session.dispatch(tr)

				report.bump(applied ? "op.ok" : "op.refused")

				return String(applied)
			} catch (error) {
				report.finding("dispatchThrew", { label, error: errorText(error, 1500) })

				return "DISPATCH THREW"
			}
		},
		execute(label, command: DocxEditorCommand) {
			try {
				return String(session.execute(command))
			} catch (error) {
				report.finding("commandThrew", { command: label, error: errorText(error, 1500), log: report.log })

				return "THREW"
			}
		}
	}
}

const [command, file, seed, mode] = process.argv.slice(2)

if (command === "viewer") for (const path of process.argv.slice(3)) print({ file: relative(path), digest: await viewerDigest(path) })
else {
	const config = fuzzConfig()

	if (config === undefined) throw new Error("Set FILEN_OFFICE_FUZZ")

	if (command === "paged" && file !== undefined && (mode === "sync" || mode === "race" || mode === "proof"))
		await paged(config, file, Number(seed), mode)
	else if (command === "damage" && file !== undefined) await damage(config, file)
	else throw new Error("Usage: run.ts paged <file> <seed> <sync|race|proof> | damage <file> | viewer <file>...")
}
