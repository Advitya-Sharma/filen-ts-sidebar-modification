import { NodeKind, decodeUtf8, encodeUtf8, parseXml, XmlError, type XmlDocument } from "../../../xml/index"
import type { StoryPart } from "../../model/index"
import { partReader, type PartReader } from "../../model/reader"
import { Ns } from "../../model/namespaces"
import { canonicalElement } from "../canonical"
import { DocxVerifyError } from "../failure"
import { writtenTextboxes } from "../nested"
import type { CheckInput, CheckedPart, DocxCheck } from "./index"

// Text boxes and mc:AlternateContent. A save never leaves an mc:AlternateContent without the mc:Fallback its
// source had ([MS-ODRAWXML] 2.2.7: Word's drawing extensions have one mc:Choice and one mc:Fallback; a reader
// that does not understand the choice would lose the content), and a text box the save rewrote reads the same
// in every branch it was written into.

function isMc(reader: PartReader, node: number, local: string): boolean {
	const document = reader.document

	return document.kinds[node] === NodeKind.Element && document.localName(node) === local && reader.group(node) === Ns.MC
}

function fallbackless(reader: PartReader, node: number): boolean {
	const document = reader.document

	if (!isMc(reader, node, "AlternateContent")) return false

	for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child))
		if (isMc(reader, child, "Fallback")) return false

	return true
}

// Byte offsets of the source's mc:AlternateContent elements without a fallback, ascending (read once).
const sourceFallbackless = new WeakMap<XmlDocument, Int32Array>()

function fallbacklessOffsets(reader: PartReader): Int32Array {
	const document = reader.document
	let found = sourceFallbackless.get(document)

	if (found === undefined) {
		const offsets: number[] = []

		for (let node = 0; node < document.count; node++) if (fallbackless(reader, node)) offsets.push(document.start(node))

		found = Int32Array.from(offsets)
		sourceFallbackless.set(document, found)
	}

	return found
}

// How many of `offsets` lie in [from, to).
function countIn(offsets: Int32Array, from: number, to: number): number {
	const lower = (value: number): number => {
		let low = 0
		let high = offsets.length

		while (low < high) {
			const middle = (low + high) >> 1

			if ((offsets[middle] ?? 0) < value) low = middle + 1
			else high = middle
		}

		return low
	}

	return lower(to) - lower(from)
}

function countWritten(reader: PartReader): number {
	let count = 0

	for (let node = 0; node < reader.document.count; node++) if (fallbackless(reader, node)) count++

	return count
}

function checkFallbacks(input: CheckInput, checked: CheckedPart, part: StoryPart, written: PartReader): void {
	const source = fallbacklessOffsets(part.reader)
	let allowed: number

	if (checked.skeleton) {
		// The rewritten stories: their new markup is the skeleton; their copies are source bytes.
		allowed = 0

		for (const story of input.changes.stories) {
			if (story.part !== part) continue

			const container = part.node(story.session.story.story.item)
			const from = story.range?.from ?? part.document.start(container)
			const to = story.range?.to ?? part.document.end(container)
			const trusted = story.output.trusted

			allowed += countIn(source, from, to)

			for (let k = 0; k < trusted.length; k += 4) allowed -= countIn(source, trusted[k + 2] ?? 0, trusted[k + 3] ?? 0)
		}
	} else {
		allowed = input.original.has(checked.name) ? source.length : 0
	}

	if (countWritten(written) > allowed)
		throw new DocxVerifyError("content", `${checked.name}: an mc:AlternateContent lost its mc:Fallback`)
}

// A rewritten text box's markup parsed inside its part's root (which binds its prefixes): its canonical form
// and its text.
function parsedBox(part: StoryPart, bytes: Uint8Array): Box | undefined {
	const document = part.document
	const root = document.documentElement()
	const open = decodeUtf8(document.bytes, document.start(root), document.openEnd(root))
	const close = encodeUtf8(`</${document.name(root)}>`)
	const head = encodeUtf8(document.isSelfClosing(root) ? `${open.replace(/\s*\/>$/, "")}>` : open)
	const wrapped = new Uint8Array(head.length + bytes.length + close.length)

	wrapped.set(head, 0)
	wrapped.set(bytes, head.length)
	wrapped.set(close, head.length + bytes.length)

	try {
		const parsed = parseXml(wrapped)
		const box = parsed.firstChildElement(parsed.documentElement())

		return box < 0 ? undefined : { canonical: canonicalElement(parsed, box), text: parsed.textContent(box) }
	} catch (error) {
		if (error instanceof XmlError) return undefined

		throw error
	}
}

interface Box {
	readonly canonical: string
	readonly text: string
}

// The text boxes (w:txbxContent not inside another) of each branch of an mc:AlternateContent.
function branchBoxes(reader: PartReader, alternate: number): Box[][] {
	const document = reader.document
	const out: Box[][] = []

	for (let branch = document.firstChildElement(alternate); branch >= 0; branch = document.nextSiblingElement(branch)) {
		if (reader.group(branch) !== Ns.MC) continue

		const boxes: Box[] = []
		const end = document.end(branch)

		for (let node = branch + 1; node < document.count && document.start(node) < end; node++) {
			if (document.kinds[node] !== NodeKind.Element || document.localName(node) !== "txbxContent" || reader.group(node) !== Ns.W)
				continue

			boxes.push({ canonical: canonicalElement(document, node), text: document.textContent(node) })

			const close = document.end(node)

			while (node + 1 < document.count && document.start(node + 1) < close) node++
		}

		out.push(boxes)
	}

	return out
}

// Each text box the save rewrote stands in every branch it was written into, as written there, and reads the
// same text in each.
function checkTextboxes(checked: CheckedPart, part: StoryPart, reader: PartReader): void {
	const written = writtenTextboxes(part)

	if (written.length === 0) return

	const document = checked.document
	// Canonical form of each version -> the text box it belongs to.
	const owner = new Map<string, number>()
	const found = written.map(() => 0)

	written.forEach((textbox, index) => {
		for (const splice of textbox.splices) {
			const box = parsedBox(part, splice.bytes)

			if (box === undefined) throw new DocxVerifyError("malformedPart", `${checked.name}: a rewritten text box does not parse`)

			owner.set(box.canonical, index)
		}
	})

	for (let node = 0; node < document.count; node++) {
		if (!isMc(reader, node, "AlternateContent")) continue

		// A branch with no text box at all (a picture of the shape) holds no copy of the text.
		const branches = branchBoxes(reader, node).filter(boxes => boxes.length > 0)
		const width = Math.max(0, ...branches.map(boxes => boxes.length))

		for (let k = 0; k < width; k++) {
			const versions = branches.map(boxes => boxes[k])
			const index = versions
				.map(version => (version === undefined ? undefined : owner.get(version.canonical)))
				.find(entry => entry !== undefined)

			if (index === undefined) continue
			if (
				versions.some(
					version => version === undefined || owner.get(version.canonical) !== index || version.text !== versions[0]?.text
				)
			)
				throw new DocxVerifyError("content", `${checked.name}: a rewritten text box reads differently in its mc:Fallback`)

			found[index] = (found[index] ?? 0) + versions.length
		}
	}

	written.forEach((textbox, index) => {
		if (textbox.used > 0 && (found[index] ?? 0) < textbox.splices.length && textbox.splices.length > 1)
			throw new DocxVerifyError("content", `${checked.name}: a rewritten text box is missing from a branch`)
	})
}

export const textboxChecks: readonly DocxCheck[] = [
	{
		part(input, checked) {
			const part = checked.part

			if (part === undefined) return

			const reader = partReader(checked.document)

			checkFallbacks(input, checked, part, reader)
			checkTextboxes(checked, part, reader)
		}
	}
]
