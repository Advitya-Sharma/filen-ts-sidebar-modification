import {
	addRelationships,
	editContentTypes,
	partKey,
	relationshipsPartName,
	relativeTarget,
	sliceClock,
	crc32,
	BytePieces,
	type SliceClock,
	CONTENT_TYPES_PART,
	type ByteSink,
	type ByteSource,
	type NewRelationship,
	type PartBytes,
	type PartChange,
	type Slicing,
	type WriteZipOptions,
	type ZipPlan
} from "../../opc/index"
import { XmlError, type XmlDocument } from "../../xml/index"
import type { Node } from "prosemirror-model"
import type { DocxEditContext } from "../edit/context"
import type { NewList } from "../edit/formats"
import type { DocxEditorSession } from "../edit/session"
import type { StoryPart } from "../model/index"
import type { DocxPackage } from "../package"
import { DocxSaveError, DocxVerifyError } from "./failure"
import { NUMBERING_CONTENT_TYPE, planNumbering } from "./numbering"
import { DocxWriteError, StoryWriter, type StoryDraft, type StoryOutput, type WriteEnvironment } from "./story"
import { STYLES_CONTENT_TYPE, stylesWithDefinitions } from "./styles"
import { verifyDocxParts } from "./verify"
import { changedCrcs, zipDocxSave, type DocxPartChanges } from "./zip"

const RELATIONSHIP_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const RELATIONSHIPS_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml"

export interface DocxSaveInput {
	readonly pkg: DocxPackage
	readonly context: DocxEditContext
	// One session per story (the body, headers, notes...). Read-only sessions are not written.
	readonly sessions: readonly DocxEditorSession[]
}

export interface DocxSaveTimings {
	readonly serializeMs: number
	readonly verifyPartsMs: number
	readonly zipMs: number
	readonly verifyZipMs: number
}

// A story whose container was rewritten. `output.bytes` holds the very pieces of the part's new content.
export interface StoryChange {
	readonly session: DocxEditorSession
	// The session's document as written (read once as the save started).
	readonly doc: Node
	readonly part: StoryPart
	readonly output: StoryOutput
}

// What a save writes: the parts the edits changed (story parts, numbering, styles, relationships, content
// types), each built from the original part's bytes, plus parts that did not exist.
export interface DocxChanges extends DocxPartChanges {
	readonly numIds: ReadonlyMap<number, number>
	readonly stories: readonly StoryChange[]
}

export interface DocxPrepareOptions {
	// Write and check in slices (this thread stays responsive); a pause that rejects abandons the save with
	// that error. Without it the save runs in one go.
	readonly slicing?: Slicing
	// Compute each changed part's CRC-32 here, from the bytes the part checks read (PartChange.crc32), so the
	// zip half checks the archive against them wherever it runs (default true; a trial that zips nothing
	// can skip it).
	readonly checksums?: boolean
}

// Bytes checksummed between checks of the clock.
const CRC_STEP = 4 * 1024 * 1024

// The CRC-32 of a part's new content, in steps.
async function checksum(bytes: PartBytes, clock: SliceClock | undefined): Promise<number> {
	if (bytes instanceof Uint8Array) return crc32(bytes)
	if (!(bytes instanceof BytePieces)) {
		let crc = 0

		for await (const chunk of bytes.chunks()) crc = crc32(chunk, crc)

		return crc
	}

	let crc = 0

	for (let from = 0; from < bytes.length; from += CRC_STEP) {
		bytes.visit(from, Math.min(CRC_STEP, bytes.length - from), (piece, pieceFrom, count) => {
			crc = crc32(piece.subarray(pieceFrom, pieceFrom + count), crc)

			return true
		})

		if (clock?.due === true) await clock.pause()
	}

	return crc
}

// Relationships of one source part with new ones planned, ids allocated as addRelationships gives them.
class RelationshipPlan {
	private readonly taken = new Set<string>()
	private highest = 0
	readonly additions: NewRelationship[] = []
	readonly ids: string[] = []

	constructor(existing: readonly string[]) {
		for (const id of existing) {
			this.taken.add(id)

			const numbered = /^rId(\d+)$/.exec(id)

			if (numbered !== null) this.highest = Math.max(this.highest, Number(numbered[1]))
		}
	}

	add(relationship: NewRelationship): string {
		let id = `rId${++this.highest}`

		while (this.taken.has(id)) id = `rId${++this.highest}`

		this.taken.add(id)
		this.additions.push(relationship)
		this.ids.push(id)

		return id
	}
}

async function optionalXml(pkg: DocxPackage, name: string | undefined, slicing: Slicing | undefined): Promise<XmlDocument | undefined> {
	return name === undefined || !pkg.has(name) ? undefined : pkg.xml(name, slicing)
}

function freeName(pkg: DocxPackage, name: string): string {
	if (!pkg.has(name)) return name

	for (let n = 1; ; n++) {
		const candidate = name.replace(/\.xml$/, `${n}.xml`)

		if (!pkg.has(candidate)) return candidate
	}
}

interface DraftedStory {
	readonly session: DocxEditorSession
	readonly doc: Node
	readonly container: number
	readonly draft: StoryDraft
}

// Pieces gathered between looks at the clock.
const GATHER_STEP = 16_384

// A story part's new content: the source with each rewritten container replaced by its story as written,
// as pieces (views of the source and the new markup), never joined: a one-word edit in a long document
// costs the edit, not a copy of the part. Each story's output is a window of it.
async function assemble(
	part: StoryPart,
	drafted: readonly DraftedStory[],
	clock: SliceClock | undefined
): Promise<{ bytes: BytePieces; stories: StoryChange[] }> {
	const document = part.document
	const source = document.bytes
	const ordered = [...drafted].sort((a, b) => document.start(a.container) - document.start(b.container))
	const pieces: Uint8Array[] = []
	const windows: { readonly at: number; readonly size: number }[] = []
	const due = (): boolean => clock?.due === true
	let from = 0
	let at = 0

	for (const { container, draft } of ordered) {
		const start = document.start(container)

		if (start < from) throw new DocxWriteError("unsupported", `Rewritten stories of ${part.name} overlap`)

		pieces.push(source.subarray(from, start))
		at += start - from

		const views = draft.out.views()

		// One at a time: a story rewritten throughout has more pieces than a call takes arguments.
		for (let index = 0; index < views.length; index++) {
			pieces.push(views[index] as Uint8Array)

			if (clock !== undefined && index % GATHER_STEP === GATHER_STEP - 1 && due()) await clock.pause()
		}

		windows.push({ at, size: draft.out.size })
		at += draft.out.size
		from = document.end(container)
	}

	pieces.push(source.subarray(from))

	const bytes = clock === undefined ? BytePieces.of(pieces) : await BytePieces.build(pieces, due, clock.pause)
	const stories = ordered.map(({ session, doc, draft }, index) => {
		const window = windows[index] ?? { at: 0, size: 0 }

		return {
			session,
			doc,
			part,
			output: {
				bytes: bytes.window(window.at, window.size),
				trusted: draft.out.trusted,
				paragraphs: draft.paragraphs,
				containers: draft.containers
			}
		}
	})

	return { bytes, stories }
}

export async function serializeDocx(input: DocxSaveInput, options: DocxPrepareOptions = {}): Promise<DocxChanges> {
	const { pkg, context } = input
	const slicing = options.slicing
	const clock = slicing === undefined ? undefined : sliceClock(slicing)
	const sessions = input.sessions.filter(session => !session.readOnly)
	// What the save writes is the documents as they are now, whatever happens between slices.
	const docs = sessions.map(session => session.state.doc)
	const model = context.model
	let revision = model.maxIds.revision
	const relationships = new Map<string, RelationshipPlan>()
	const relationshipsOf = async (part: string): Promise<RelationshipPlan> => {
		const key = partKey(part)
		let plan = relationships.get(key)

		if (plan === undefined) {
			plan = new RelationshipPlan((await pkg.relationships(part)).all.map(relationship => relationship.id))
			relationships.set(key, plan)
		}

		return plan
	}
	const lists = new Map<number, NewList>()
	const missingStyles = new Map<string, string>()

	for (const session of sessions) {
		for (const [id, list] of session.newLists) lists.set(id, list)
		for (const [id, name] of session.missingStyles) missingStyles.set(id, name)
	}

	const numberingPart = pkg.part("numbering")
	const numbering = planNumbering(await optionalXml(pkg, numberingPart, clock), context.numbering, lists)

	// Relationship plans exist for every story part up front, so hyperlinks get their ids in order.
	for (const session of sessions) await relationshipsOf(session.story.part.name)

	const env: WriteEnvironment = {
		context,
		nextRevisionId: () => ++revision,
		numId(id) {
			const numId = numbering.numIds.get(id)

			if (numId === undefined) throw new DocxWriteError("missingSource", `Unknown new list ${id}`)

			return numId
		},
		hyperlink(part, target) {
			const plan = relationships.get(partKey(part))

			if (plan === undefined) throw new DocxWriteError("unsupported", `No relationships for ${part}`)

			return plan.add({ type: `${RELATIONSHIP_BASE}/hyperlink`, target, external: true })
		}
	}

	// Story containers per part, replaced in the part's own bytes; a story written exactly as its source is
	// left out (compared without joining it: its copies are views of the source).
	const drafted = new Map<StoryPart, DraftedStory[]>()

	for (const [index, session] of sessions.entries()) {
		const story = session.story
		const part = story.part
		const doc = docs[index] ?? session.state.doc
		const writer = new StoryWriter(env, story)
		const draft = clock === undefined ? writer.draftStorySync(doc) : await writer.draftStory(doc, clock)
		const container = part.node(story.story.item)

		if (draft.out.equals(part.document.bytes, part.document.start(container), part.document.end(container))) continue

		const list = drafted.get(part)

		if (list === undefined) drafted.set(part, [{ session, doc, container, draft }])
		else list.push({ session, doc, container, draft })
	}

	const assembled = new Map<string, BytePieces>()
	const stories: StoryChange[] = []

	for (const [part, list] of drafted) {
		const result = await assemble(part, list, clock)

		assembled.set(part.name, result.bytes)
		stories.push(...result.stories)

		if (clock?.due === true) await clock.pause()
	}

	const replace: PartChange[] = []
	const add: PartChange[] = []
	const overrides: { partName: string; contentType: string }[] = []

	for (const [name, part] of assembled) replace.push({ name, bytes: part })

	const numberingBytes = numbering.bytes()

	if (numberingBytes !== null) {
		if (numberingPart !== undefined) {
			replace.push({ name: numberingPart, bytes: numberingBytes })
		} else {
			const name = freeName(pkg, "word/numbering.xml")

			add.push({ name, bytes: numberingBytes })
			overrides.push({ partName: name, contentType: NUMBERING_CONTENT_TYPE })
			;(await relationshipsOf(pkg.mainPart)).add({
				type: `${RELATIONSHIP_BASE}/numbering`,
				target: relativeTarget(pkg.mainPart, name)
			})
		}
	}

	const stylesPart = pkg.part("styles")
	const stylesBytes = stylesWithDefinitions(await optionalXml(pkg, stylesPart, clock), context.styles, missingStyles)

	if (stylesBytes !== null) {
		if (stylesPart !== undefined) {
			replace.push({ name: stylesPart, bytes: stylesBytes })
		} else {
			const name = freeName(pkg, "word/styles.xml")

			add.push({ name, bytes: stylesBytes })
			overrides.push({ partName: name, contentType: STYLES_CONTENT_TYPE })
			;(await relationshipsOf(pkg.mainPart)).add({ type: `${RELATIONSHIP_BASE}/styles`, target: relativeTarget(pkg.mainPart, name) })
		}
	}

	const types = await pkg.opc.contentTypes()

	if (clock?.due === true) await clock.pause()

	for (const [key, plan] of relationships) {
		if (plan.additions.length === 0) continue

		const source = pkg.opc.entry(key)?.name ?? key
		const relsName = relationshipsPartName(source)
		const existing = pkg.has(relsName) ? await pkg.opc.read(relsName) : null
		const result = addRelationships(existing, plan.additions)

		if (result.ids.join() !== plan.ids.join())
			throw new DocxWriteError("unsupported", `Relationship ids of ${source} did not come out as planned`)

		if (existing === null) {
			add.push({ name: relsName, bytes: result.bytes })

			// A package with no relationships part typed by extension (malformed but readable) types the new one.
			if (types.typeOf(relsName) === undefined) overrides.push({ partName: relsName, contentType: RELATIONSHIPS_CONTENT_TYPE })
		} else {
			replace.push({ name: pkg.opc.entry(relsName)?.name ?? relsName, bytes: result.bytes })
		}
	}

	if (overrides.length > 0) {
		replace.push({ name: CONTENT_TYPES_PART, bytes: editContentTypes(await pkg.opc.read(CONTENT_TYPES_PART), { overrides }) })
	}

	return { replace, add, numIds: numbering.numIds, stories }
}

// The parts a save writes, checked by the part-level half of gate 2, with the time each step took (ms, slices
// and the turns between them included).
export interface DocxPreparedSave extends DocxPartChanges {
	readonly changedParts: readonly string[]
	readonly serializeMs: number
	readonly verifyPartsMs: number
}

// Gate 0 (the document may be edited at all), the lossless writer, then the part-level half of gate 2: what
// needs the editor and the model. The zip half (zipDocxSave) needs only the original archive and these parts.
export async function prepareDocxSave(input: DocxSaveInput, options: DocxPrepareOptions = {}): Promise<DocxPreparedSave> {
	const { pkg, context } = input

	if (context.gate.length > 0) throw new DocxSaveError("readOnly", `The document is view-only (${context.gate.join(", ")})`)

	// One clock for every step, so each continues the slice the one before started.
	const clock = options.slicing === undefined ? undefined : sliceClock(options.slicing)
	const sliced = clock === undefined ? {} : { slicing: clock }
	let started = Date.now()
	let changes: DocxChanges

	try {
		changes = await serializeDocx(input, sliced)
	} catch (error) {
		if (error instanceof DocxWriteError || error instanceof XmlError) throw new DocxSaveError("write", error.message)

		throw error
	}

	const serializeMs = Date.now() - started

	started = Date.now()

	try {
		await verifyDocxParts({ original: pkg, context, changes }, clock)
	} catch (error) {
		if (error instanceof DocxVerifyError) throw new DocxSaveError(error.reason, error.message)

		throw error
	}

	const withCrc = async (change: PartChange): Promise<PartChange> =>
		options.checksums === false ? change : { ...change, crc32: await checksum(change.bytes, clock) }
	const replace: PartChange[] = []
	const add: PartChange[] = []

	for (const change of changes.replace) replace.push(await withCrc(change))
	for (const change of changes.add) add.push(await withCrc(change))

	return {
		replace,
		add,
		changedParts: [...changes.replace, ...changes.add].map(change => change.name),
		serializeMs,
		verifyPartsMs: Date.now() - started
	}
}

// A save ready to be written: the zip plan (untouched entries copied raw), the parts it rewrites with the
// CRC each must have in the archive, and the timings so far. Gate 2 has checked the parts.
export interface DocxSavePlan {
	readonly plan: ZipPlan
	readonly changedParts: readonly string[]
	// partKey -> CRC-32 of the part's new bytes.
	readonly crcs: ReadonlyMap<string, number>
	readonly serializeMs: number
	readonly verifyPartsMs: number
}

// prepareDocxSave as a zip plan. Mobile can execute the plan outside the WebView (writeZip next to the file)
// and check the result there with verifyDocxZip.
export async function planDocxSave(input: DocxSaveInput, options: DocxPrepareOptions = {}): Promise<DocxSavePlan> {
	const prepared = await prepareDocxSave(input, options)

	return {
		plan: await input.pkg.opc.plan({ replace: prepared.replace, add: prepared.add }),
		changedParts: prepared.changedParts,
		crcs: await changedCrcs(prepared),
		serializeMs: prepared.serializeMs,
		verifyPartsMs: prepared.verifyPartsMs
	}
}

export interface DocxSaveOptions extends DocxPrepareOptions {
	readonly zip?: WriteZipOptions
	// Where the file goes (default: memory, returned as `bytes`).
	readonly sink?: ByteSink
	// Random access to what `sink` received, for the zip half of gate 2; required with a sink.
	readonly readBack?: () => Promise<ByteSource>
}

export interface DocxSaveResult {
	// The file, when no sink was given.
	readonly bytes: Uint8Array | undefined
	// Parts rewritten or added; every other entry was copied as it was, compressed bytes included.
	readonly changedParts: readonly string[]
	readonly timings: DocxSaveTimings
}

// Writes the edited document on this thread: prepareDocxSave, then the zip with every untouched entry copied
// raw and the zip half of gate 2 on the result (zipDocxSave). Any failure refuses the save with a reason; the
// editor stays dirty.
export async function saveDocx(input: DocxSaveInput, options: DocxSaveOptions = {}): Promise<DocxSaveResult> {
	const prepared = await prepareDocxSave(input, options)
	const zipped = await zipDocxSave(input.pkg.opc, prepared, {
		...(options.zip === undefined ? {} : { zip: options.zip }),
		...(options.sink === undefined ? {} : { sink: options.sink }),
		...(options.readBack === undefined ? {} : { readBack: options.readBack })
	})

	return {
		bytes: zipped.bytes,
		changedParts: zipped.changedParts,
		timings: {
			serializeMs: prepared.serializeMs,
			verifyPartsMs: prepared.verifyPartsMs,
			zipMs: zipped.zipMs,
			verifyZipMs: zipped.verifyZipMs
		}
	}
}
