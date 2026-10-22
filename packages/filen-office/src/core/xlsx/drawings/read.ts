import { D, drawingReader, groupUris } from "../../drawingml/reader"
import { selectAlternateContent, type XmlDocument } from "../../xml/index"
import { childByLocal } from "../refs/markup"
import { readAnchor, type DrawingAnchor } from "./anchor"

// A sheet's drawing part (ECMA-376 Part 1 §20.5, CT_Drawing): its anchored objects in z-order (the first is drawn
// first), what each is, where it sits and what an edit needs to find in it again. Nothing here throws on an object
// that does not read: it is listed as "unknown" and stays as the file has it.

export type DrawingObjectKind = "picture" | "shape" | "connector" | "group" | "chart" | "chartex" | "graphic" | "unknown"

export interface DrawingObjectRecord {
	// Its index among the part's objects (z-order).
	readonly key: number
	// The wsDr child holding it: the anchor, or an mc:AlternateContent around it.
	readonly node: number
	// The anchor element.
	readonly anchorNode: number
	readonly anchor: DrawingAnchor | null
	readonly kind: DrawingObjectKind
	// The element drawn (sp, pic, grpSp, cxnSp, graphicFrame; an mc branch's), -1 for none.
	readonly content: number
	// Every element naming the object (the cNvPr of the content, in each branch of alternate content).
	readonly properties: readonly number[]
	readonly name: string
	// Its alternative text (descr), falling back to its title.
	readonly alt: string
	readonly hidden: boolean
	// Locks the file sets on the object (a:picLocks and its siblings).
	readonly noMove: boolean
	readonly noResize: boolean
	// The relationship of a chart frame's chart (c:chart or cx:chart r:id).
	readonly chartRelationship: string | null
	// Every relationship id the object's markup names (pictures, charts, hyperlinks).
	readonly relationships: readonly string[]
}

export interface DrawingPartRecord {
	readonly document: XmlDocument
	// The wsDr root.
	readonly root: number
	readonly objects: readonly DrawingObjectRecord[]
}

const ANCHORS: ReadonlySet<string> = new Set(["twoCellAnchor", "oneCellAnchor", "absoluteAnchor"])
const PLACEMENT: ReadonlySet<string> = new Set(["from", "to", "pos", "ext", "clientData"])
const CHART_URIS: ReadonlySet<string> = new Set([
	"http://schemas.openxmlformats.org/drawingml/2006/chart",
	"http://purl.oclc.org/ooxml/drawingml/chart"
])
const CHARTEX_URI = "http://schemas.microsoft.com/office/drawing/2014/chartex"
const RELATIONSHIP_URIS: ReadonlySet<string> = new Set([
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
	"http://purl.oclc.org/ooxml/officeDocument/relationships"
])
// The namespaces whose mc:Choice branches are drawn: DrawingML's own and Office 2010's drawing extensions.
const UNDERSTOOD: ReadonlySet<string> = new Set([
	...groupUris(D.A),
	...groupUris(D.Xdr),
	...groupUris(D.Pic),
	...groupUris(D.C),
	...groupUris(D.R),
	...groupUris(D.A14)
])
const LOCKS: ReadonlySet<string> = new Set(["picLocks", "spLocks", "grpSpLocks", "cxnSpLocks", "graphicFrameLocks"])

const KINDS: Readonly<Record<string, DrawingObjectKind>> = {
	pic: "picture",
	sp: "shape",
	cxnSp: "connector",
	grpSp: "group",
	graphicFrame: "graphic"
}

function elements(document: XmlDocument, node: number): number[] {
	const out: number[] = []

	for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) out.push(child)

	return out
}

// The anchor inside a wsDr-level mc:AlternateContent: the drawn branch's, else any branch's.
function anchorIn(document: XmlDocument, node: number, branch: number): number {
	for (const candidate of [branch, ...elements(document, node)]) {
		if (candidate < 0) continue

		for (const child of elements(document, candidate)) if (ANCHORS.has(document.localName(child))) return child
	}

	return -1
}

// The cNvPr of an object element (under its nv*Pr), -1 for none.
function cNvPrOf(document: XmlDocument, element: number): number {
	for (const child of elements(document, element)) {
		const local = document.localName(child)

		if (local === "cNvPr") return child
		if (local.startsWith("nv") && local.endsWith("Pr")) {
			const found = childByLocal(document, child, "cNvPr")

			if (found >= 0) return found
		}
	}

	return -1
}

function locks(document: XmlDocument, element: number): { noMove: boolean; noResize: boolean } {
	let noMove = false
	let noResize = false
	const flag = (node: number, name: string): boolean => /^(1|true)$/.test(document.attribute(node, name) ?? "")

	for (const nv of elements(document, element)) {
		if (!document.localName(nv).startsWith("nv")) continue

		for (const properties of elements(document, nv))
			for (const lock of elements(document, properties)) {
				if (!LOCKS.has(document.localName(lock))) continue

				noMove ||= flag(lock, "noMove")
				noResize ||= flag(lock, "noResize")
			}
	}

	return { noMove, noResize }
}

// The graphicData uri of a graphic frame, "" for none.
function graphicUri(document: XmlDocument, frame: number): { uri: string; data: number } {
	const graphic = childByLocal(document, frame, "graphic")
	const data = graphic < 0 ? -1 : childByLocal(document, graphic, "graphicData")

	return { uri: data < 0 ? "" : (document.attribute(data, "uri") ?? ""), data }
}

export function readDrawingPart(document: XmlDocument): DrawingPartRecord {
	const reader = drawingReader(document)
	const namespaces = reader.namespaces
	const root = document.documentElement()
	const objects: DrawingObjectRecord[] = []
	const understands = (uri: string): boolean => UNDERSTOOD.has(uri)

	const relationshipIds = (node: number, into: string[]): void => {
		for (const attribute of document.attributes(node)) {
			const colon = attribute.name.indexOf(":")

			if (colon > 0 && RELATIONSHIP_URIS.has(namespaces.uri(node, attribute.name.slice(0, colon)) ?? "")) into.push(attribute.value)
		}

		for (const child of elements(document, node)) relationshipIds(child, into)
	}

	for (const node of root < 0 ? [] : elements(document, root)) {
		const local = document.localName(node)
		let anchorNode = -1

		if (ANCHORS.has(local)) anchorNode = node
		else if (namespaces.is(node, "http://schemas.openxmlformats.org/markup-compatibility/2006", "AlternateContent"))
			anchorNode = anchorIn(document, node, selectAlternateContent(namespaces, node, understands))

		if (anchorNode < 0) continue

		// The content: the anchor's first child past its placement; an mc:AlternateContent there is drawn through the
		// branch Office 2010 would take, and named in every branch.
		let content = -1
		let wrapper = -1

		for (const child of elements(document, anchorNode)) {
			if (PLACEMENT.has(document.localName(child))) continue

			if (namespaces.is(child, "http://schemas.openxmlformats.org/markup-compatibility/2006", "AlternateContent")) {
				wrapper = child
				content = elements(document, selectAlternateContent(namespaces, child, understands))[0] ?? -1
			} else {
				content = child
			}

			break
		}

		const branches =
			wrapper < 0
				? content < 0
					? []
					: [content]
				: elements(document, wrapper).flatMap(branch => elements(document, branch).slice(0, 1))
		const properties = branches.map(element => cNvPrOf(document, element)).filter(found => found >= 0)
		const named = content >= 0 ? cNvPrOf(document, content) : (properties[0] ?? -1)
		let kind: DrawingObjectKind = content < 0 ? "unknown" : (KINDS[document.localName(content)] ?? "unknown")
		let chartRelationship: string | null = null

		// A chart frame, in whichever branch holds it (a chartex has a shape as its fallback).
		for (const element of branches) {
			if (document.localName(element) !== "graphicFrame") continue

			const { uri, data } = graphicUri(document, element)
			const chart = CHART_URIS.has(uri) || uri === CHARTEX_URI ? childByLocal(document, data, "chart") : -1

			if (chart < 0) continue

			kind = uri === CHARTEX_URI ? "chartex" : "chart"

			const ids: string[] = []

			relationshipIds(chart, ids)
			chartRelationship = ids[0] ?? null

			break
		}

		const relationships: string[] = []

		relationshipIds(anchorNode, relationships)

		const descr = named < 0 ? "" : (document.attribute(named, "descr") ?? "")
		const title = named < 0 ? "" : (document.attribute(named, "title") ?? "")

		objects.push({
			key: objects.length,
			node,
			anchorNode,
			anchor: readAnchor(document, anchorNode),
			kind,
			content,
			properties,
			name: named < 0 ? "" : (document.attribute(named, "name") ?? ""),
			alt: descr === "" ? title : descr,
			hidden: named >= 0 && /^(1|true)$/.test(document.attribute(named, "hidden") ?? ""),
			...(content < 0 ? { noMove: false, noResize: false } : locks(document, content)),
			chartRelationship,
			relationships: [...new Set(relationships)]
		})
	}

	return { document, root, objects }
}
