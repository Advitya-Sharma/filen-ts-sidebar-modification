import {
	addRelationships,
	bytesSource,
	crc32,
	editContentTypes,
	memorySink,
	partKey,
	relationshipsPartName,
	relativeTarget,
	writeZip,
	CONTENT_TYPES_PART,
	type ByteSink,
	type ByteSource,
	type NewRelationship,
	type PartChange,
	type WriteZipOptions,
	type ZipPlan
} from "../../opc/index"
import { XmlEditor, XmlError, type XmlDocument } from "../../xml/index"
import type { DocxEditContext } from "../edit/context"
import type { NewList } from "../edit/formats"
import type { DocxEditorSession } from "../edit/session"
import type { StoryPart } from "../model/index"
import type { DocxPackage } from "../package"
import { NUMBERING_CONTENT_TYPE, planNumbering } from "./numbering"
import { DocxWriteError, StoryWriter, type StoryOutput, type WriteEnvironment } from "./story"
import { STYLES_CONTENT_TYPE, stylesWithDefinitions } from "./styles"
import { DocxVerifyError, verifyDocxParts, verifyDocxZip, type VerifyFailure } from "./verify"

const RELATIONSHIP_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const RELATIONSHIPS_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml"

export interface DocxSaveInput {
	readonly pkg: DocxPackage
	readonly context: DocxEditContext
	// One session per story (the body, headers, notes...). Read-only sessions are not written.
	readonly sessions: readonly DocxEditorSession[]
}

export type DocxSaveFailure = "readOnly" | "write" | VerifyFailure

export class DocxSaveError extends Error {
	readonly reason: DocxSaveFailure

	constructor(reason: DocxSaveFailure, message: string) {
		super(message)
		this.name = "DocxSaveError"
		this.reason = reason
	}
}

export interface DocxSaveTimings {
	readonly serializeMs: number
	readonly verifyPartsMs: number
	readonly zipMs: number
	readonly verifyZipMs: number
}

// A story whose container was rewritten.
export interface StoryChange {
	readonly session: DocxEditorSession
	readonly part: StoryPart
	readonly output: StoryOutput
}

// What a save writes: the parts the edits changed (story parts, numbering, styles, relationships, content
// types), each built from the original part's bytes, plus parts that did not exist.
export interface DocxChanges {
	readonly replace: readonly PartChange[]
	readonly add: readonly PartChange[]
	readonly numIds: ReadonlyMap<number, number>
	readonly stories: readonly StoryChange[]
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

async function optionalXml(pkg: DocxPackage, name: string | undefined): Promise<XmlDocument | undefined> {
	return name === undefined || !pkg.has(name) ? undefined : pkg.xml(name)
}

function freeName(pkg: DocxPackage, name: string): string {
	if (!pkg.has(name)) return name

	for (let n = 1; ; n++) {
		const candidate = name.replace(/\.xml$/, `${n}.xml`)

		if (!pkg.has(candidate)) return candidate
	}
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
	if (a.length !== b.length) return false

	for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false

	return true
}

export async function serializeDocx(input: DocxSaveInput): Promise<DocxChanges> {
	const { pkg, context } = input
	const sessions = input.sessions.filter(session => !session.readOnly)
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
	const numbering = planNumbering(await optionalXml(pkg, numberingPart), context.numbering, lists)

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

	// Story containers per part, replaced in the part's own bytes.
	const editors = new Map<string, XmlEditor>()
	const stories: StoryChange[] = []

	for (const session of sessions) {
		const story = session.story
		const part = story.part
		const output = new StoryWriter(env, story).writeStory(session.state.doc)
		const container = part.node(story.story.item)
		const original = part.document.bytes.subarray(part.document.start(container), part.document.end(container))

		if (sameBytes(output.bytes, original)) continue

		let editor = editors.get(part.name)

		if (editor === undefined) {
			editor = new XmlEditor(part.document)
			editors.set(part.name, editor)
		}

		editor.replace(container, output.bytes)
		stories.push({ session, part, output })
	}

	const replace: PartChange[] = []
	const add: PartChange[] = []
	const overrides: { partName: string; contentType: string }[] = []

	for (const [name, editor] of editors) replace.push({ name, bytes: editor.serialize() })

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
	const stylesBytes = stylesWithDefinitions(await optionalXml(pkg, stylesPart), context.styles, missingStyles)

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

// Gate 0 (the document may be edited at all), the lossless writer, then the part-level half of gate 2.
// Mobile can execute the plan outside the WebView (writeZip next to the file) and check the result there
// with verifyDocxZip.
export async function planDocxSave(input: DocxSaveInput): Promise<DocxSavePlan> {
	const { pkg, context } = input

	if (context.gate.length > 0) throw new DocxSaveError("readOnly", `The document is view-only (${context.gate.join(", ")})`)

	let started = Date.now()
	let changes: DocxChanges

	try {
		changes = await serializeDocx(input)
	} catch (error) {
		if (error instanceof DocxWriteError || error instanceof XmlError) throw new DocxSaveError("write", error.message)

		throw error
	}

	const serializeMs = Date.now() - started

	started = Date.now()

	try {
		await verifyDocxParts({ original: pkg, context, changes })
	} catch (error) {
		if (error instanceof DocxVerifyError) throw new DocxSaveError(error.reason, error.message)

		throw error
	}

	const verifyPartsMs = Date.now() - started
	const changed = [...changes.replace, ...changes.add]
	const plan = await pkg.opc.plan({ replace: changes.replace, add: changes.add })

	return {
		plan,
		changedParts: changed.map(change => change.name),
		crcs: new Map(changed.map(change => [partKey(change.name), crc32(change.bytes)])),
		serializeMs,
		verifyPartsMs
	}
}

export interface DocxSaveOptions {
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

interface DeflateStream {
	readonly readable: unknown
	readonly writable: { getWriter(): { write(chunk: Uint8Array): Promise<void>; close(): Promise<void> } }
}

// The platform's own raw deflate (CompressionStream "deflate-raw", in browsers, workers and Node), several
// times faster than fflate on large parts; undefined where there is none (Hermes).
export function nativeDeflate(): ((data: Uint8Array) => Promise<Uint8Array>) | undefined {
	const Compression = Reflect.get(globalThis, "CompressionStream") as (new (format: string) => DeflateStream) | undefined
	const Body = Reflect.get(globalThis, "Response") as (new (body: unknown) => { arrayBuffer(): Promise<ArrayBuffer> }) | undefined

	if (typeof Compression !== "function" || typeof Body !== "function") return undefined

	try {
		new Compression("deflate-raw")
	} catch {
		return undefined
	}

	return async data => {
		const stream = new Compression("deflate-raw")
		const writer = stream.writable.getWriter()
		const written = Promise.all([writer.write(data), writer.close()])
		const result = new Uint8Array(await new Body(stream.readable).arrayBuffer())

		await written

		return result
	}
}

// Writes the edited document: the plan (see planDocxSave), the zip with every untouched entry copied raw,
// then the zip half of gate 2 on the result. Any failure refuses the save with a reason; the editor stays
// dirty.
export async function saveDocx(input: DocxSaveInput, options: DocxSaveOptions = {}): Promise<DocxSaveResult> {
	const save = await planDocxSave(input)
	let started = Date.now()
	const memory = options.sink === undefined ? memorySink() : undefined
	const sink = options.sink ?? memory

	if (sink === undefined) throw new DocxSaveError("invalidPackage", "No sink")

	const deflate = options.zip?.deflate ?? nativeDeflate()

	await writeZip(save.plan, input.pkg.opc.zip, sink, { ...options.zip, ...(deflate !== undefined ? { deflate } : {}) })

	const zipMs = Date.now() - started
	const bytes = memory?.bytes()

	started = Date.now()

	const source = bytes !== undefined ? bytesSource(bytes) : await options.readBack?.()

	if (source === undefined) throw new DocxSaveError("invalidPackage", "The written file cannot be read back")

	try {
		await verifyDocxZip(input.pkg, source, save)
	} catch (error) {
		if (error instanceof DocxVerifyError) throw new DocxSaveError(error.reason, error.message)

		throw error
	}

	return {
		bytes,
		changedParts: save.changedParts,
		timings: { serializeMs: save.serializeMs, verifyPartsMs: save.verifyPartsMs, zipMs, verifyZipMs: Date.now() - started }
	}
}
