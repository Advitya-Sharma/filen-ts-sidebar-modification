import { bytesSource } from "../../core/opc/index"
import type { Slicing } from "../../core/slicing"
import { isRtf } from "../../core/rtf/reader"
import { RtfWriteRefusal } from "../../core/rtf/write/out"
import { RtfAborted, openRtfSource, proveRtf, writeRtf, type RtfSource, type RtfTurns } from "../../core/rtf/write/writer"
import type { DocxEngineServices } from "./engine"
import { inThreadDocxHost, macrotaskYield } from "./host"
import { heldSource } from "./protection"
import type { DocxOpenOptions } from "./protocol"
import { DocxRtfRefusal, type DocxRtf, type DocxRtfEdits, type RtfLockedBlock } from "./rtf"

// The RTF documents held where the engine runs: each one's source, read once into its translation (the reader keeps
// every block's source range), proved on request and written back on each save. A document's requests run in turn,
// in slices: closing it ends the one running (RtfAborted) and those still queued.

export type DocxRtfOpened = { readonly ok: true; readonly size: number } | { readonly ok: false; readonly notRtf: boolean }

export type DocxRtfWritten =
	{ readonly ok: true; readonly bytes: Uint8Array } | { readonly ok: false; readonly reason: string; readonly message: string }

export interface DocxRtfServer {
	open(id: number, bytes: Uint8Array): Promise<DocxRtfOpened>
	// The translation of an open document, for an engine session opened on it.
	package(id: number): Uint8Array | undefined
	take(id: number): Uint8Array
	prove(id: number): Promise<readonly RtfLockedBlock[]>
	write(id: number, edits: DocxRtfEdits): Promise<DocxRtfWritten>
	// Drops a held document; false when `id` is not one.
	close(id: number): boolean
}

interface Held {
	source: RtfSource | undefined
	queue: Promise<unknown>
	readonly turns: RtfTurns
	closed: boolean
}

// One macrotask turn for every server of the realm (a channel each would stay open).
let realmSlicing: Slicing | undefined

// `slicing`: how proofs and writes give the thread back (default a macrotask turn every 12 ms).
export function docxRtfServer(slicing?: Slicing): DocxRtfServer {
	const held = new Map<number, Held>()
	const slices = slicing ?? (realmSlicing ??= { pause: macrotaskYield(), sliceMs: 12, now: () => Date.now() })
	const holding = (): Held => {
		const document: Held = {
			source: undefined,
			queue: Promise.resolve(),
			turns: {
				signal: {
					get aborted() {
						return document.closed
					}
				},
				slicing: slices
			},
			closed: false
		}

		return document
	}
	// Runs `task` after the document's earlier requests.
	const inTurn = <T>(id: number, task: (document: Held) => Promise<T>): Promise<T> => {
		const document = held.get(id)

		if (document === undefined) return Promise.reject(new Error("The document is not open"))

		const run = document.queue.then(() => (document.closed ? Promise.reject(new RtfAborted()) : task(document)))

		document.queue = run.catch(() => undefined)

		return run
	}
	const sourceOf = (document: Held): RtfSource => {
		if (document.source === undefined) throw new Error("The document did not open")

		return document.source
	}

	return {
		async open(id, bytes) {
			if (!isRtf(bytes)) return { ok: false, notRtf: true }

			const opening = holding()

			held.set(id, opening)

			let source: RtfSource | undefined

			try {
				source = await openRtfSource(bytes, opening.turns)
			} catch (error) {
				if (!(error instanceof RtfAborted)) throw error
			}

			// Closed while reading (the read ended at its next slice): nothing is kept.
			if (opening.closed || source === undefined) {
				if (held.get(id) === opening) held.delete(id)

				return { ok: false, notRtf: !opening.closed }
			}

			const document = opening

			document.source = source

			return { ok: true, size: source.read.docx.length }
		},
		package(id) {
			return held.get(id)?.source?.read.docx
		},
		take(id) {
			const document = held.get(id)

			if (document === undefined) throw new Error("The document is not open")

			// The reader's package is already deflated; the copy crosses to the page, the held one stays.
			return sourceOf(document).read.docx.slice()
		},
		prove(id) {
			return inTurn(id, async document =>
				(await proveRtf(sourceOf(document), document.turns)).locked.map(({ part, offset }) => ({ part, offset }))
			)
		},
		write(id, edits) {
			return inTurn(id, async document => {
				const source = sourceOf(document)

				try {
					const written = await writeRtf(
						source,
						{ parts: new Map(edits.parts), removed: edits.removed },
						{ turns: document.turns }
					)

					// An untouched file is the source itself: a copy goes, the source stays.
					return { ok: true, bytes: written.bytes === source.bytes ? written.bytes.slice() : written.bytes }
				} catch (error) {
					if (error instanceof RtfWriteRefusal) return { ok: false, reason: error.reason, message: error.message }

					throw error
				}
			})
		},
		close(id) {
			const document = held.get(id)

			if (document === undefined) return false

			document.closed = true

			return held.delete(id)
		}
	}
}

// The same on the calling thread (tests, a host without workers).
export function inThreadDocxRtf(services: (options: DocxOpenOptions) => DocxEngineServices): DocxRtf {
	const engines = inThreadDocxHost(services)

	return {
		async open(bytes, options = {}) {
			if (options.signal?.aborted === true) return { kind: "aborted" }

			const server = docxRtfServer()
			const opened = await server.open(0, bytes.slice())

			if (!opened.ok) return opened.notRtf ? { kind: "notRtf" } : { kind: "aborted" }

			return {
				kind: "opened",
				document: {
					host: (_source, listener, openOptions) =>
						engines(bytesSource(server.package(0) ?? new Uint8Array(0)), listener, openOptions),
					bytes: heldSource(opened.size),
					take: () => Promise.resolve(server.take(0)),
					prove: () => server.prove(0),
					async write(edits) {
						const written = await server.write(0, edits)

						if (!written.ok) throw new DocxRtfRefusal(written.reason, written.message)

						return written.bytes
					},
					release: () => {
						server.close(0)
					}
				}
			}
		}
	}
}
