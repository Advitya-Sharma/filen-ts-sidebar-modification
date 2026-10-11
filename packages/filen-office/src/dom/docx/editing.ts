import { RevisionView } from "../../core/docx/index"
import { DocxEditorSession, openDocxEditContext, type DocxEditorOptions } from "../../core/docx/edit/index"
import { DocumentEdits, PagedEditor, type PagedEditorOptions } from "../../core/docx/paged/index"
import { HitTester } from "../../core/text/index"
import type { XmlDocument } from "../../core/xml/index"
import { DocxEngine, type DocxEngineOpen, type DocxEngineServices } from "./engine"
import type { DocxHost } from "./host"
import type { DocxOpenOptions } from "./protocol"

// Editing runs on the thread that shows the pages (web main thread, mobile WebView), with the model, the
// layout and the editing session in one realm: no model copy, no mirroring protocol, one code path for
// both platforms. The engine still sends pages as events (the same renderer draws them); edits lay out
// again in the keystroke's task as far as the pages in view, the rest in slices.

export interface DocxEditingHost {
	// Opens the document (pass to DocxDocumentStore as its source's host); one document per editing host.
	readonly host: DocxHost
	// The paged editor once the document is open for editing; rejects when it could not be.
	readonly editor: Promise<PagedEditor>
}

export interface DocxEditingOptions {
	readonly session?: Omit<DocxEditorOptions, "story">
	readonly paged?: PagedEditorOptions
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
				resolver: open.resolver
			})

			if (closed || edits === undefined) return

			edits.define({ context, styles: await partXml(open, "styles"), numbering: await partXml(open, "numbering") })

			// Built in slices: a long document opens without long tasks.
			const session = await DocxEditorSession.open(context, editing.session ?? {}, {
				pause: engineServices.yield,
				sliceMs: engineServices.sliceMs ?? 12,
				now: engineServices.now ?? (() => Date.now())
			})

			if (closed) return

			paged = new PagedEditor({
				session,
				edits,
				layout: open.layout,
				tester: new HitTester(open.layout.styles),
				...(editing.paged === undefined ? {} : { options: editing.paged }),
				...(editing.session === undefined ? {} : { sessionOptions: editing.session })
			})
			// Pages laid out again go to the renderer like any others.
			paged.subscribe(event => {
				if (event.type === "pages" && !closed) engine.shipPages(event.changed, event.final)
			})
			resolve(paged)
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
				paged?.dispose()
				paged?.session.dispose()
				engine.close()
				reject(new Error("closed"))
			}
		}
	}

	return { host, editor }
}
