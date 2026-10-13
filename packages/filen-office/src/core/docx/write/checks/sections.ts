import { NodeKind, type XmlDocument } from "../../../xml/index"
import { Ns } from "../../model/namespaces"
import { partReader, type PartReader } from "../../model/reader"
import { DocxVerifyError } from "../failure"
import { ownedByCopy } from "../newStory"
import { joinBytes, type Relationship } from "../../../opc/index"
import type { CheckInput, CheckedPart, DocxCheck } from "./index"

// The section properties a save wrote (new or spliced w:sectPr; copies of the source's are not looked at
// again): each header and footer reference resolves to a part of its content type that the package has, a
// section refers to at most one of each type, page size and margins leave room for text, columns fit in it,
// line numbering counts from no negative value, and a recorded change has an author and holds no references.
// A header or footer part the save creates as a copy of another (stories/clone.ts) relates what it holds as
// the original does: each relationship it keeps, by id, of the same type to the same target, or for a part a
// copy owns (a chart, an embedded object: write/newStory.ts ownedByCopy) to a part the save added holding the
// same bytes of the same content type, whose own relationships read so in turn.

const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml"
const CONTENT_TYPES: Readonly<Record<string, string>> = { headerReference: `${WML}.header+xml`, footerReference: `${WML}.footer+xml` }
// ST_TwipsMeasure / ST_SignedTwipsMeasure as Word bounds a page (22 inches).
const MAX_PAGE = 31680

function fail(reason: "content" | "danglingReference" | "duplicateId", message: string): never {
	throw new DocxVerifyError(reason, message)
}

function number(reader: PartReader, node: number, local: string): number | undefined {
	const value = reader.attr(node, Ns.W, local)

	if (value === undefined) return undefined
	if (!/^-?\d{1,10}$/.test(value)) fail("content", `w:${reader.document.localName(node)} w:${local}="${value}" is no whole number`)

	return Number(value)
}

function children(document: XmlDocument, reader: PartReader, node: number, local: string): number[] {
	const out: number[] = []

	for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child))
		if (reader.is(child, Ns.W, local)) out.push(child)

	return out
}

async function checkSection(input: CheckInput, part: string, reader: PartReader, node: number, nested: boolean): Promise<void> {
	const document = reader.document
	const relationships = await input.relationshipsOf(part)
	const seen = new Set<string>()

	for (const local of ["headerReference", "footerReference"] as const) {
		for (const reference of children(document, reader, node, local)) {
			if (nested) fail("content", "A recorded section change holds a header or footer reference")

			const type = reader.attr(reference, Ns.W, "type") ?? "default"
			const id = reader.attr(reference, Ns.R, "id")
			const key = `${local}:${type}`

			if (seen.has(key)) fail("duplicateId", `A section has two ${type} ${local}s`)

			seen.add(key)

			const target = id === undefined ? undefined : relationships.byId(id)
			const name = target?.partName

			if (target === undefined || target.external || name === undefined || !input.exists(name))
				fail("danglingReference", `A section's ${local} ${id ?? ""} resolves to no part`)
			if (input.types.typeOf(name) !== CONTENT_TYPES[local])
				fail(
					"danglingReference",
					`A section's ${local} ${id ?? ""} is ${name}, no ${local === "headerReference" ? "header" : "footer"}`
				)
		}
	}

	const size = children(document, reader, node, "pgSz")[0]
	const width = size === undefined ? 12240 : (number(reader, size, "w") ?? 12240)
	const height = size === undefined ? 15840 : (number(reader, size, "h") ?? 15840)

	if (width <= 0 || height <= 0 || width > MAX_PAGE || height > MAX_PAGE)
		fail("content", `A section's page is ${width} x ${height} twips`)

	const margins = children(document, reader, node, "pgMar")[0]
	const margin = (local: string, fallback: number): number =>
		margins === undefined ? fallback : (number(reader, margins, local) ?? fallback)
	const text = width - margin("left", 1440) - margin("right", 1440) - margin("gutter", 0)

	if (text <= 0 || height - Math.abs(margin("top", 1440)) - Math.abs(margin("bottom", 1440)) <= 0)
		fail("content", "A section's margins leave no room for text")

	for (const columns of children(document, reader, node, "cols")) {
		const count = number(reader, columns, "num") ?? 1
		const space = number(reader, columns, "space") ?? 720

		if (count < 1 || count > 45) fail("content", `A section has ${count} columns`)
		if (children(document, reader, columns, "col").length === 0 && (count - 1) * space >= text)
			fail("content", `${count} columns ${space} twips apart do not fit a ${text} twips wide text`)
	}

	for (const lines of children(document, reader, node, "lnNumType")) {
		// ST_DecimalNumber; what a command writes Word bounds further (MS-OI29500 2.1.215), files have 0 (LibreOffice).
		if ((number(reader, lines, "countBy") ?? 0) < 0 || (number(reader, lines, "start") ?? 0) < 0)
			fail("content", "A section numbers lines from a negative value")
	}

	for (const change of children(document, reader, node, "sectPrChange")) {
		if ((reader.attr(change, Ns.W, "author") ?? "") === "") fail("content", "A recorded section change has no author")

		for (const before of children(document, reader, change, "sectPr")) await checkSection(input, part, reader, before, true)
	}
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
	if (a.length !== b.length) return false

	for (let k = 0; k < a.length; k++) if (a[k] !== b[k]) return false

	return true
}

// Whether part `copy` the save added is a copy of the original's part `source` (see above).
async function copiedPart(input: CheckInput, source: string, copy: string, seen: Set<string>): Promise<boolean> {
	const added = input.changes.add.find(change => change.name === copy)

	if (
		added === undefined ||
		input.original.has(copy) ||
		input.types.typeOf(copy) !== (await input.original.opc.contentTypes()).typeOf(source)
	)
		return false
	if (!sameBytes(await joinBytes(added.bytes), await input.original.opc.read(source))) return false
	if (seen.has(copy)) return true

	seen.add(copy)

	return await sameRelationships(input, (await input.original.relationships(source)).all, copy, seen)
}

// Whether part `name` as written relates as `kept` does (see above).
async function sameRelationships(input: CheckInput, kept: readonly Relationship[], name: string, seen: Set<string>): Promise<boolean> {
	const written = await input.relationshipsOf(name)

	for (const relationship of kept) {
		const now = written.byId(relationship.id)

		if (now === undefined || now.type !== relationship.type || now.external !== relationship.external) return false
		if (relationship.external) {
			if (now.target !== relationship.target) return false

			continue
		}

		if (relationship.partName === undefined) {
			if (now.target !== relationship.target) return false

			continue
		}

		const target = now.partName

		if (target === undefined || !input.exists(target)) return false
		if (target === relationship.partName) continue
		if (!ownedByCopy(relationship, input.original) || !(await copiedPart(input, relationship.partName, target, seen))) return false
	}

	return true
}

async function copiedRelationships(input: CheckInput, checked: CheckedPart): Promise<void> {
	const kept = checked.part?.relationships.all ?? []

	if (kept.length === 0 || input.original.has(checked.name)) return

	if (!(await sameRelationships(input, kept, checked.name, new Set())))
		fail("danglingReference", `${checked.name}: a relationship reads otherwise than in the part it was copied from`)
}

async function part(input: CheckInput, checked: CheckedPart): Promise<void> {
	await copiedRelationships(input, checked)

	if (checked.part === undefined || checked.part.name !== input.original.mainPart) return

	const document = checked.document
	const reader = partReader(document)

	for (let node = 0; node < document.count; node++) {
		if (document.kinds[node] !== NodeKind.Element || !reader.is(node, Ns.W, "sectPr")) continue
		// Those inside a recorded change are checked with theirs.
		if (reader.is(document.parent(node), Ns.W, "sectPrChange")) continue

		await checkSection(input, checked.part.name, reader, node, false)
	}
}

export const sectionChecks: readonly DocxCheck[] = [{ part }]
