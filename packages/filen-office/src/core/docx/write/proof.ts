import type { Node } from "prosemirror-model"
import { XmlError, encodeUtf8, parseXml } from "../../xml/index"
import type { DocxEditContext } from "../edit/context"
import type { StoryDoc } from "../edit/build"
import type { BlockProof } from "../edit/proof"
import { NO_SOURCE, docxSchema, type LockReason } from "../edit/schema"
import { canonicalElement, canonicalFragment } from "./canonical"
import type { Out } from "./markup"
import { DocxWriteError, StoryWriter, type WriteEnvironment } from "./story"

const OPEN = encodeUtf8("<proof>")
const CLOSE = encodeUtf8("</proof>")

// Gate 1: each source paragraph and table is written through the edit path (runs rebuilt from their bases
// and marks, properties spliced, containers rebuilt from their parts) and must come out with the source's
// XML meaning (see canonical.ts). A block that does not is locked; the rest of the document stays editable.
export function writerProof(context: DocxEditContext): BlockProof {
	let revision = context.model.maxIds.revision
	const env: WriteEnvironment = {
		context,
		nextRevisionId: () => ++revision,
		numId: id => id,
		hyperlink: () => {
			throw new DocxWriteError("unsupported", "A new hyperlink while proving")
		}
	}
	const writers = new WeakMap<StoryDoc, StoryWriter>()
	// Writers in the middle of a proof in steps: a check meanwhile takes a writer of its own.
	const busy = new WeakSet<StoryWriter>()

	function* prove(node: Node, story: StoryDoc, due: (() => boolean) | undefined): Generator<void, LockReason | undefined> {
		const src = node.attrs["src"] as number

		if (src === NO_SOURCE) return undefined

		let writer = writers.get(story)

		if (writer === undefined || busy.has(writer)) {
			writer = new StoryWriter(env, story)

			if (!writers.has(story)) writers.set(story, writer)
		}

		let out: Out

		try {
			if (due !== undefined) {
				busy.add(writer)
				out = yield* writer.writeBlockAloneSteps(node, due)
			} else {
				out = writer.writeBlockAlone(node)
			}
		} catch (error) {
			if (error instanceof DocxWriteError) return "proof"

			throw error
		} finally {
			busy.delete(writer)
		}

		const document = story.part.document
		const element = story.part.node(src)
		const table = node.type === docxSchema.nodes.table

		// Most blocks come out as their very bytes.
		if (out.equals(document.bytes, document.start(element), document.end(element))) return undefined

		const written = out.finish()
		const wrapped = new Uint8Array(OPEN.length + written.length + CLOSE.length)

		wrapped.set(OPEN, 0)
		wrapped.set(written, OPEN.length)
		wrapped.set(CLOSE, OPEN.length + written.length)

		let actual: string

		try {
			const fragment = parseXml(wrapped)

			actual = canonicalFragment(fragment, fragment.documentElement(), table)
		} catch (error) {
			if (error instanceof XmlError) return "proof"

			throw error
		}

		const source = canonicalElement(document, element, table)

		return actual === source ? undefined : "proof"
	}

	return {
		check(node: Node, story: StoryDoc) {
			const steps = prove(node, story, undefined)

			for (;;) {
				const step = steps.next()

				if (step.done === true) return step.value
			}
		},
		steps: (node, story, due) => prove(node, story, due)
	}
}
