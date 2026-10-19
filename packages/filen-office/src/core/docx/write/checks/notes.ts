import type { SliceClock } from "../../../opc/index"
import { NodeKind, type XmlDocument } from "../../../xml/index"
import { Ns } from "../../model/namespaces"
import { partReader } from "../../model/reader"
import { DocxVerifyError } from "../failure"
import type { DocxCheck } from "./index"

// Notes hold together: every new note reference resolves to a normal note (ids below 1 are Word's separators, and
// a separator or continuation notice of any id is no note a reference can show; ECMA-376 Part 1 §17.11,
// [MS-OI29500] 2.1.301 b) and is the only reference to it, and the separator notes a notes part has stay (never
// touched). Note references stay in the body, and floating drawings (wp:anchor) out of comments: Word makes
// neither, and LibreOffice opens no document with one, so a header, footer, notes or comments part may hold no
// more of them than its source did. Gate 2 counts note and reference ids itself (verify.ts), the normal notes here.

const OUTSIDE_BODY: ReadonlySet<string> = new Set(["hdr", "ftr", "footnotes", "endnotes", "comments"])

// Nodes looked at between looks at the clock.
const NODE_STEP = 4096

// A part's note references (by "footnoteReference:<id>") and floating drawings.
interface Census {
	readonly references: Map<string, number>
	anchors: number
}

const REFERENCE = 1
const ANCHOR = 2

async function census(document: XmlDocument, clock: SliceClock | undefined): Promise<Census> {
	const reader = partReader(document)
	// What each element name is (0 nothing, REFERENCE, ANCHOR), worked out once per name.
	const kinds = new Int8Array(document.names.length).fill(-1)
	const found: Census = { references: new Map(), anchors: 0 }

	for (let node = 0; node < document.count; node++) {
		if (node % NODE_STEP === NODE_STEP - 1 && clock?.due === true) await clock.pause()
		if (document.kinds[node] !== NodeKind.Element) continue

		const id = document.nameId(node)
		let kind = kinds[id] ?? 0

		if (kind < 0) {
			const local = document.localName(node)

			kind = local === "footnoteReference" || local === "endnoteReference" ? REFERENCE : local === "anchor" ? ANCHOR : 0
			kinds[id] = kind
		}

		if (kind === REFERENCE && reader.group(node) === Ns.W) {
			const key = `${document.localName(node)}:${reader.attr(node, Ns.W, "id") ?? ""}`

			found.references.set(key, (found.references.get(key) ?? 0) + 1)
		} else if (kind === ANCHOR && reader.group(node) === Ns.WP) {
			found.anchors++
		}
	}

	return found
}

// A source part's census is taken once (saves repeat).
const sources = new WeakMap<XmlDocument, Census>()

export const noteChecks: readonly DocxCheck[] = [
	{
		fact: {
			elements: ["footnote", "endnote"],
			id(reader, node) {
				const type = reader.attr(node, Ns.W, "type")
				const id = reader.attr(node, Ns.W, "id")

				return id === undefined || (type !== undefined && type !== "normal")
					? undefined
					: `${reader.document.localName(node)}Normal:${id}`
			}
		},
		// The part as written (a skeleton: its new markup alone) against its source, reference by reference.
		async part(input, checked) {
			const document = checked.document
			const root = document.documentElement()

			if (root < 0 || !OUTSIDE_BODY.has(document.localName(root)) || partReader(document).group(root) !== Ns.W) return

			const written = await census(document, input.clock)
			const original = checked.part?.document
			let source = original === undefined ? undefined : sources.get(original)

			if (original !== undefined && source === undefined) {
				source = await census(original, input.clock)
				sources.set(original, source)
			}

			for (const [key, count] of written.references) {
				if (count > (source?.references.get(key) ?? 0))
					throw new DocxVerifyError("content", `${checked.name}: a note reference outside the body`)
			}

			if (document.localName(root) === "comments" && written.anchors > (source?.anchors ?? 0))
				throw new DocxVerifyError("content", `${checked.name}: a floating drawing in a comment`)
		},
		done(input) {
			for (const kind of ["footnote", "endnote"]) {
				const added = input.added(`${kind}Reference`)

				for (const id of added) {
					if (!/^\d+$/.test(id) || Number(id) < 1 || input.ids(`${kind}Normal:${id}`) < 1)
						throw new DocxVerifyError("danglingReference", `A new reference to ${kind} ${id} has no note`)
					if (input.ids(`${kind}Reference:${id}`) > 1)
						throw new DocxVerifyError("duplicateId", `${kind} ${id} has two references`)
				}

				for (const separator of ["-1", "0"]) {
					const key = `${kind}:${separator}`

					if (input.ids(key) < Math.min(1, input.source(key)))
						throw new DocxVerifyError("content", `The save removed ${kind} separator ${separator}`)
				}
			}
		}
	}
]
