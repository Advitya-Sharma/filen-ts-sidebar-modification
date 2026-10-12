import { NodeKind, type XmlDocument } from "../../../xml/index"
import { Ns } from "../../model/namespaces"
import { partReader, type PartReader } from "../../model/reader"
import { DocxVerifyError } from "../failure"
import type { CheckInput, CheckedPart, DocxCheck } from "./index"

// The section properties a save wrote (new or spliced w:sectPr; copies of the source's are not looked at
// again): each header and footer reference resolves to a part of its content type that the package has, a
// section refers to at most one of each type, page size and margins leave room for text, columns fit in it,
// line numbering counts from no negative value, and a recorded change has an author and holds no references.

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

async function part(input: CheckInput, checked: CheckedPart): Promise<void> {
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
