import { RevisionView } from "../../core/docx/index"
import {
	DocxEditorSession,
	openDocxEditContext,
	type BlockProof,
	type DocxEditContext,
	type DocxEditorOptions,
	type DocxMediaInfo,
	authorInitials
} from "../../core/docx/edit/index"
import { DocumentEdits, PagedEditor, type PagedEditorOptions } from "../../core/docx/paged/index"
import type { DocxPrepareOptions, DocxPreparedSave, DocxSaveInput, DocxSaveTimings } from "../../core/docx/write/save"
import type { Slicing } from "../../core/slicing"
import { HitTester } from "../../core/text/index"
import type { XmlDocument } from "../../core/xml/index"
import { DocxEngine, type DocxEngineOpen, type DocxEngineServices } from "./engine"
import type { DocxHost } from "./host"
import type { DocxOpenOptions } from "./protocol"
import { DocxClosedError, inThreadDocxZipper, type DocxZipper, type DocxZipperFactory } from "./save"

// Editing runs on the thread that shows the pages (web main thread, mobile WebView), with the model, the
// layout and the editing session in one realm: no model copy, no mirroring protocol, one code path for
// both platforms. The engine still sends pages as events (the same renderer draws them); edits lay out
// again in the keystroke's task as far as the pages in view, the rest in slices.

export interface DocxEditingHost {
	// Opens the document (pass to DocxDocumentStore as its source's host); one document per editing host.
	readonly host: DocxHost
	// The paged editor once the document is open for editing; rejects when it could not be.
	readonly editor: Promise<PagedEditor>
	// What saveDocx / planDocxSave write (every story's session) once the editor is open, else undefined.
	saveInput(): DocxSaveInput | undefined
	// Writes the document as `sessions` hold it (default every story's session; stand-ins from withReverted to
	// write some blocks as last saved): the lossless writer and the part half of gate 2 in slices on this
	// thread, then the zip and its half of gate 2 in the zipper (DocxEditingOptions.writer). The editor must
	// not change the documents meanwhile (they are read once, as the save starts). Rejects with a
	// DocxSaveError when the writer refuses, with a DocxClosedError when the document closed first.
	save(sessions?: readonly DocxEditorSession[]): Promise<DocxWritten>
	// The writer and the part half of gate 2 alone, in slices (nothing zipped or stored): whether, and what,
	// a save of `sessions` would write.
	prepare(sessions?: readonly DocxEditorSession[]): Promise<DocxPreparedSave>
	// Body blocks the proof locked as the document opened (0 before then, or without a proof).
	proofLocked(): number
	// The document author name for comments and tracked changes (the app asked the user; never filled in for
	// them); undefined clears it. Holds for the document from now on, also before it opened.
	setAuthor(name: string | undefined): void
	// Shows tracked changes another way (DocxOpenOptions.view), laid out again in place: no reopen. False
	// before the editor opened or when the document already shows that view.
	setView(view: "final" | "markup" | "original"): boolean
	// Registers bytes a command will use (a picture to insert): their handle, or undefined when over the size
	// limit or before the editor opened.
	addMedia(chunks: Uint8Array | readonly Uint8Array[], info?: DocxMediaInfo): number | undefined
}

export interface DocxWritten {
	readonly bytes: Uint8Array
	readonly changedParts: readonly string[]
	readonly timings: DocxSaveTimings
}

export interface DocxEditingOptions {
	readonly session?: Omit<DocxEditorOptions, "story">
	// The document author name (see DocxEditingHost.setAuthor) and the clock dates written into the file read.
	readonly author?: string
	readonly clock?: () => number
	readonly paged?: PagedEditorOptions
	// Gate 1 for the document, made from its editing context (writerProof, injected: the writer is not part
	// of this entry). The body is proved in slices before the editor opens; other stories when entered.
	readonly proof?: (context: DocxEditContext) => BlockProof
	// The writer for save() and prepare() (prepareDocxSave, injected like the proof), and where saves zip the
	// file (default: on this thread; web passes a worker's, docxZipWorker). Without it the host cannot save.
	readonly writer?: DocxEditingWriter
}

export interface DocxEditingWriter {
	readonly prepare: (input: DocxSaveInput, options: DocxPrepareOptions) => Promise<DocxPreparedSave>
	readonly zipper?: DocxZipperFactory
}

// Top-level blocks proved per step between checks of the clock.
const PROOF_CHUNK = 32

// Gate 1 over the body's blocks in slices (no long task however long the document): each step proves
// a run of top-level blocks, then the next run starts after them however the locks resized them. The
// locks are not edits: the session is marked saved again. Returns the blocks locked.
async function proveInSlices(session: DocxEditorSession, proof: BlockProof, slicing: Slicing, closed: () => boolean): Promise<number> {
	let locked = 0
	let index = 0
	let from = 0
	let started = slicing.now()

	for (;;) {
		const doc = session.state.doc

		if (index >= doc.childCount || closed()) break

		let to = from
		let end = index

		while (end < doc.childCount && end - index < PROOF_CHUNK) {
			to += doc.child(end).nodeSize
			end++
		}

		locked += session.prove(proof, from, to)
		from = to + session.state.doc.content.size - doc.content.size
		index = end

		if (slicing.now() - started >= slicing.sliceMs) {
			await slicing.pause()
			started = slicing.now()
		}
	}

	if (locked > 0) session.markSaved()

	return locked
}

// A styles or numbering part as opened, or undefined (missing or unreadable: the save writes a new one).
async function partXml(open: DocxEngineOpen, kind: string): Promise<XmlDocument | undefined> {
	const name = open.pkg.part(kind)

	if (name === undefined || !open.pkg.has(name)) return undefined

	try {
		return await open.pkg.xml(name)
	} catch {
		return undefined
	}
}

export function inThreadEditingHost(
	services: (options: DocxOpenOptions) => DocxEngineServices,
	editing: DocxEditingOptions = {}
): DocxEditingHost {
	let resolve: (editor: PagedEditor) => void = () => undefined
	let reject: (error: unknown) => void = () => undefined
	const editor = new Promise<PagedEditor>((done, fail) => {
		resolve = done
		reject = fail
	})
	let opened = false
	let saveInput: DocxSaveInput | undefined
	let proofLocked = 0
	let author = editing.author
	let current: PagedEditor | undefined
	// Set once the editor is open: how saves slice their work here, and the zipper.
	let writing: { readonly slicing: Slicing; readonly zipper: DocxZipper } | undefined
	// `checksums`: a trial (nothing zipped) skips the parts' CRCs.
	const prepare = (sessions: readonly DocxEditorSession[] | undefined, checksums: boolean): Promise<DocxPreparedSave> => {
		const input = saveInput

		if (editing.writer === undefined) return Promise.reject(new Error("This editing host has no writer"))
		if (input === undefined || writing === undefined) return Promise.reject(new DocxClosedError())

		return editing.writer.prepare(
			{ pkg: input.pkg, context: input.context, sessions: sessions ?? input.sessions },
			{ slicing: writing.slicing, checksums }
		)
	}
	const save = async (sessions?: readonly DocxEditorSession[]): Promise<DocxWritten> => {
		const prepared = await prepare(sessions, true)

		if (writing === undefined) throw new DocxClosedError()

		const zipped = await writing.zipper.zip(prepared)

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

	const host: DocxHost = (source, listener, options = {}) => {
		if (opened) throw new Error("An editing host opens one document")

		opened = true

		let closed = false
		let paged: PagedEditor | undefined
		const view =
			options.view === "original" ? RevisionView.Original : options.view === "markup" ? RevisionView.Markup : RevisionView.Final
		let edits: DocumentEdits | undefined
		const engineServices = services(options)
		const engine: DocxEngine = new DocxEngine(
			source,
			engineServices,
			event => {
				if (closed) return

				listener(event)

				if (event.type === "error") reject(new Error(event.message))
			},
			options,
			{
				prepare(model) {
					edits = new DocumentEdits(model, view)

					return { model: edits.model, edits }
				},
				opened(open: DocxEngineOpen) {
					void start(open).catch((error: unknown) => reject(error))
				}
			}
		)
		const start = async (open: DocxEngineOpen): Promise<void> => {
			const context = await openDocxEditContext(open.pkg, open.model, {
				store: open.store,
				...(open.styles === undefined ? {} : { styles: open.styles }),
				...(open.numbering === undefined ? {} : { numbering: open.numbering }),
				resolver: open.resolver,
				...(editing.clock === undefined ? {} : { clock: editing.clock })
			})

			// The name as it is now (setAuthor may have changed it since the host was made).
			context.identity.author = author
			context.identity.initials = author === undefined ? undefined : authorInitials(author)

			if (closed || edits === undefined) return

			edits.define({ context, styles: await partXml(open, "styles"), numbering: await partXml(open, "numbering") })

			const slicing: Slicing = {
				pause: engineServices.yield,
				sliceMs: engineServices.sliceMs ?? 12,
				now: engineServices.now ?? (() => Date.now())
			}
			const proof = editing.proof?.(context)
			// With a proof the body is proved below, in slices; other stories prove as their sessions start.
			const bodyOptions = proof === undefined ? (editing.session ?? {}) : { ...editing.session, proof: null }
			const storyOptions = proof === undefined ? editing.session : { ...editing.session, proof }
			// Built in slices: a long document opens without long tasks.
			const session = await DocxEditorSession.open(context, bodyOptions, slicing)

			if (closed) return

			if (proof !== undefined) proofLocked = await proveInSlices(session, proof, slicing, () => closed)

			if (closed) return

			const editor = new PagedEditor({
				session,
				edits,
				layout: open.layout,
				tester: new HitTester(open.layout.styles),
				...(editing.paged === undefined ? {} : { options: editing.paged }),
				...(storyOptions === undefined ? {} : { sessionOptions: storyOptions })
			})

			paged = editor
			current = editor
			writing = {
				// A save under way when the document closes stops at its next turn.
				slicing: {
					sliceMs: slicing.sliceMs,
					now: slicing.now,
					pause: async () => {
						await slicing.pause()

						if (closed) throw new DocxClosedError()
					}
				},
				zipper: (editing.writer?.zipper ?? inThreadDocxZipper)(open.pkg.opc)
			}
			saveInput = {
				pkg: open.pkg,
				context,
				get sessions() {
					return editor.sessions
				}
			}
			// Pages laid out again go to the renderer like any others.
			editor.subscribe(event => {
				if (event.type === "pages" && !closed) engine.shipPages(event.changed, event.final)
			})
			resolve(editor)
		}

		void engine.run()

		return {
			objects: ids => engine.objects(ids),
			media: partName => engine.media(partName),
			locate: anchor => engine.locate(anchor),
			find: (query, findOptions) => engine.find(query, findOptions),
			annotations: () => engine.annotations(),
			close() {
				closed = true
				current = undefined
				saveInput = undefined
				writing?.zipper.dispose()
				writing = undefined
				paged?.dispose()
				paged?.session.dispose()
				engine.close()
				reject(new Error("closed"))
			}
		}
	}

	return {
		host,
		editor,
		saveInput: () => saveInput,
		save,
		prepare: sessions => prepare(sessions, false),
		proofLocked: () => proofLocked,
		setAuthor(name) {
			author = name

			if (current !== undefined) {
				const identity = current.body.session.context.identity

				identity.author = name
				identity.initials = name === undefined ? undefined : authorInitials(name)
			}
		},
		setView(view) {
			return (
				current?.setView(
					view === "original" ? RevisionView.Original : view === "markup" ? RevisionView.Markup : RevisionView.Final
				) ?? false
			)
		},
		addMedia: (chunks, info) => current?.addMedia(chunks, info)
	}
}
