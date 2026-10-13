import type { Node } from "prosemirror-model"
import { parseFraction } from "../../../drawingml/reader"
import { NodeKind, decodeUtf8, escapeAttribute, type XmlDocument } from "../../../xml/index"
import type { DocxEditContext } from "../../edit/context"
import { NO_SOURCE, docxSchema } from "../../edit/schema"
import type { DrawingInfo } from "../../model/drawings"
import { Item, type StoryPart } from "../../model/index"
import { Ns } from "../../model/namespaces"
import type { DocxCrop } from "../../edit/commands/drawings"
import { drawingEdit, drawingState, graphicTransform, pictureFill, sameCrop, sourceDrawing, type DrawingState } from "../../objects/edit"
import { sniffImage } from "../../objects/images"
import type { Out } from "../markup"
import { mediaRelationship } from "../parts/media"
import { src } from "../source"
import type { StoryWriting } from "../story"
import { DocxWriteError } from "../values"
import type { AtomWriter } from "./index"

// Drawings a command changed and pictures a command inserted. A changed drawing is its source bytes with
// exactly the markup the change touches rewritten in place (extents, the shape's transform, the alt text,
// the picture's crop, the anchor's position and wrapping, an inline frame turned into an anchor or back), so
// everything the editor does not model stays as it was. A drawing inside mc:AlternateContent also has its VML
// fallback (the copy Word keeps for older readers) brought in line: its style's size, rotation, position and
// wrapping, its picture's crop and its alt text. A new picture is a DrawingML inline (or anchored) picture as Word writes one, its bytes a
// media part of the save (parts/media.ts).

const nodes = docxSchema.nodes
const WP_URI = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
const A_URI = "http://schemas.openxmlformats.org/drawingml/2006/main"
const PIC_URI = "http://schemas.openxmlformats.org/drawingml/2006/picture"
const R_URI = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const W10_URI = "urn:schemas-microsoft-com:office:word"
const SVG_URI = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
// The a:blip extension holding an SVG picture's own bytes ([MS-ODRAWXML] 2.3.10 svgBlip).
const SVG_EXTENSION = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
const EMU_PER_POINT = 12700
// Word's distance from text of a floating picture: 0.125 in left and right.
const FLOAT_DISTANCE = 114300
// Children an inline frame may have (CT_Inline); an anchor's others go when it becomes inline.
const INLINE_CHILDREN = new Set(["extent", "effectExtent", "docPr", "cNvGraphicFramePr", "graphic"])
const WRAP_ELEMENTS = new Set(["wrapNone", "wrapSquare", "wrapTight", "wrapThrough", "wrapTopAndBottom"])
// ST_RelFromH / ST_RelFromV as VML's mso-position-*-relative names them.
const VML_RELATIVE_H: Readonly<Record<string, string>> = {
	column: "text",
	margin: "margin",
	page: "page",
	character: "char",
	leftMargin: "left-margin-area",
	rightMargin: "right-margin-area",
	insideMargin: "inner-margin-area",
	outsideMargin: "outer-margin-area"
}
const VML_RELATIVE_V: Readonly<Record<string, string>> = {
	paragraph: "text",
	margin: "margin",
	page: "page",
	line: "line",
	topMargin: "top-margin-area",
	bottomMargin: "bottom-margin-area",
	insideMargin: "inner-margin-area",
	outsideMargin: "outer-margin-area"
}

// ---- edits on source bytes -----------------------------------------------------------------------------

// Source bytes [from, to) written as `text`.
interface Edit {
	readonly from: number
	readonly to: number
	readonly text: string
}

// Source bytes [from, to) of the story's part with `edits` made, stories rewritten inside spliced in.
function writeEdited(w: StoryWriting, out: Out, from: number, to: number, edits: readonly Edit[]): void {
	const ordered = edits.map((edit, index) => ({ edit, index })).sort((a, b) => a.edit.from - b.edit.from || a.index - b.index)
	let at = from

	for (const { edit } of ordered) {
		if (edit.from < at || edit.to < edit.from || edit.to > to) throw new DocxWriteError("unsupported", "Overlapping drawing changes")

		w.copySource(out, at, edit.from, false)
		out.text(edit.text)
		at = edit.to
	}

	w.copySource(out, at, to, false)
}

function isSpace(b: number | undefined): boolean {
	return b === 32 || b === 9 || b === 10 || b === 13
}

// Sets unprefixed attribute `name` of `node` to `value` (null: removes it).
function setAttribute(document: XmlDocument, node: number, name: string, value: string | null, edits: Edit[]): void {
	const bytes = document.bytes
	const spans = document.attributeSpans(node)

	for (let k = 0; k < spans.length; k += 4) {
		const nameStart = spans[k] ?? 0

		if (decodeUtf8(bytes, nameStart, spans[k + 1] ?? 0) !== name) continue

		if (value === null) {
			let from = nameStart

			while (from > document.start(node) && isSpace(bytes[from - 1])) from--

			edits.push({ from, to: (spans[k + 3] ?? 0) + 1, text: "" })
		} else {
			edits.push({ from: spans[k + 2] ?? 0, to: spans[k + 3] ?? 0, text: escapeAttribute(value) })
		}

		return
	}

	if (value === null) return

	const at = document.openEnd(node) - (document.isSelfClosing(node) ? 2 : 1)

	edits.push({ from: at, to: at, text: ` ${name}="${escapeAttribute(value)}"` })
}

function attributeOf(document: XmlDocument, node: number, name: string): string | undefined {
	const spans = document.attributeSpans(node)

	for (let k = 0; k < spans.length; k += 4) {
		if (decodeUtf8(document.bytes, spans[k] ?? 0, spans[k + 1] ?? 0) === name)
			return decodeUtf8(document.bytes, spans[k + 2] ?? 0, spans[k + 3] ?? 0)
	}

	return undefined
}

// The raw attributes of a start tag other than those named in `drop`, as written (" name='value'"...).
function keptAttributes(document: XmlDocument, node: number, drop: ReadonlySet<string>): string {
	const bytes = document.bytes
	const spans = document.attributeSpans(node)
	let out = ""

	for (let k = 0; k < spans.length; k += 4) {
		const name = decodeUtf8(bytes, spans[k] ?? 0, spans[k + 1] ?? 0)

		if (!drop.has(name)) out += ` ${decodeUtf8(bytes, spans[k] ?? 0, (spans[k + 3] ?? 0) + 1)}`
	}

	return out
}

function children(document: XmlDocument, node: number): number[] {
	const out: number[] = []

	for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) out.push(child)

	return out
}

// ---- values ----------------------------------------------------------------------------------------------

function int(value: number): string {
	return String(Math.round(value))
}

// A blip fill's a:srcRect for `crop` (sides of 0 left out, their default), as Word writes it.
function srcRectMarkup(a: string, crop: DocxCrop): string {
	const sides: [string, number][] = [
		["l", crop.left],
		["t", crop.top],
		["r", crop.right],
		["b", crop.bottom]
	]
	const attributes = sides
		.filter(([, value]) => value !== 0)
		.map(([name, value]) => ` ${name}="${int(value)}"`)
		.join("")

	return `<${a === "" ? "" : `${a}:`}srcRect${attributes}/>`
}

// A VML crop side (v:imagedata cropleft...): a fraction in 1/65536ths, as Word writes it; null when none.
function vmlCrop(value: number): string | null {
	return value === 0 ? null : `${int((value / 100000) * 65536)}f`
}

function points(emu: number): string {
	return `${String(Math.round((emu / EMU_PER_POINT) * 100) / 100)}pt`
}

// How far a box of cx x cy rotated by `degrees` reaches past its own extent on each side (the effect extent a
// rotated inline drawing needs): [left and right, top and bottom].
function rotationReach(cx: number, cy: number, degrees: number): [number, number] {
	if (degrees % 360 === 0) return [0, 0]

	const radians = (degrees * Math.PI) / 180
	const cos = Math.abs(Math.cos(radians))
	const sin = Math.abs(Math.sin(radians))

	return [Math.round((cx * cos + cy * sin - cx) / 2), Math.round((cx * sin + cy * cos - cy) / 2)]
}

// The z orders (wp:anchor relativeHeight) given to drawings made floating, by drawing (`key`), and the largest
// of the document's, read once: a drawing gets one above every other once, the same at every write after.
const zOrders = new WeakMap<DocxEditContext, { highest: number; readonly given: Map<string, number> }>()

function zOrder(context: DocxEditContext, key: string): number {
	let orders = zOrders.get(context)

	if (orders === undefined) {
		let highest = 251658240

		for (const drawing of context.model.drawings) highest = Math.max(highest, drawing?.anchor?.relativeHeight ?? 0)

		orders = { highest, given: new Map() }
		zOrders.set(context, orders)
	}

	let z = orders.given.get(key)

	if (z === undefined) {
		z = Math.min(orders.highest + 1024, 0x7fffffff)
		orders.highest = z
		orders.given.set(key, z)
	}

	return z
}

// A wrap element for `mode` (behind and in front of text: wrapNone), keeping the source's polygon of a tight or
// through wrapping where the mode still has one, and its wrapText side.
function wrapMarkup(part: StoryPart, wp: string, mode: string, source: number): string {
	const qualify = (local: string): string => (wp === "" ? local : `${wp}:${local}`)
	const document = part.document

	if (mode === "behind" || mode === "inFront") return `<${qualify("wrapNone")}/>`
	if (mode === "topAndBottom") return `<${qualify("wrapTopAndBottom")}/>`

	const side = source >= 0 ? (attributeOf(document, source, "wrapText") ?? "bothSides") : "bothSides"
	const local = mode === "tight" ? "wrapTight" : mode === "through" ? "wrapThrough" : "wrapSquare"

	if (local === "wrapSquare") return `<${qualify(local)} wrapText="${escapeAttribute(side)}"/>`

	const polygon = source < 0 ? -1 : part.reader.child(source, Ns.WP, "wrapPolygon")
	const shape =
		polygon >= 0
			? decodeUtf8(document.bytes, document.start(polygon), document.end(polygon))
			: `<${qualify("wrapPolygon")} edited="0"><${qualify("start")} x="0" y="0"/><${qualify("lineTo")} x="0" y="21600"/><${qualify("lineTo")} x="21600" y="21600"/><${qualify("lineTo")} x="21600" y="0"/><${qualify("lineTo")} x="0" y="0"/></${qualify("wrapPolygon")}>`

	return `<${qualify(local)} wrapText="${escapeAttribute(side)}">${shape}</${qualify(local)}>`
}

function positionMarkup(wp: string, axis: "H" | "V", relative: string, offset: number): string {
	const qualify = (local: string): string => (wp === "" ? local : `${wp}:${local}`)

	return `<${qualify(`position${axis}`)} relativeFrom="${escapeAttribute(relative)}"><${qualify("posOffset")}>${int(offset)}</${qualify("posOffset")}></${qualify(`position${axis}`)}>`
}

// A size relative to the page or margins (wp14:sizeRelH / sizeRelV holding a positive percentage, MS-ODRAWXML
// 2.18.1.3-4) is what Word lays that axis out by, not the extent: an axis a resize set turns absolute, written
// as Word writes an absolute size there (0: Word reads only positive values as relative, MS-ODRAWXML note <125>).
function absoluteSize(document: XmlDocument, frame: number, container: string, local: string, edits: Edit[]): void {
	for (const relative of children(document, frame)) {
		if (document.localName(relative) !== container) continue

		for (const percent of children(document, relative)) {
			if (document.localName(percent) !== local || document.isSelfClosing(percent)) continue

			// ST_Percentage: thousandths of a percent, or "NN%" (Strict).
			const value = parseFraction(decodeUtf8(document.bytes, document.openEnd(percent), document.contentEnd(percent)), Number.NaN)

			if (value > 0) edits.push({ from: document.openEnd(percent), to: document.contentEnd(percent), text: "0" })
		}
	}
}

// ---- a changed source drawing ------------------------------------------------------------------------------

// The edits a drawing's DrawingML frame needs to be `state` (as its source `drawing` is `base`).
function frameEdits(
	w: StoryWriting,
	part: StoryPart,
	node: Node,
	drawing: DrawingInfo,
	base: DrawingState,
	state: DrawingState,
	edits: Edit[]
): void {
	const document = part.document
	const reader = part.reader
	const frame = drawing.frame

	if (frame < 0) throw new DocxWriteError("unsupported", "A changed drawing without a frame")

	const wp = document.prefix(frame)
	const qualify = (local: string): string => (wp === "" ? local : `${wp}:${local}`)
	const extent = reader.child(frame, Ns.WP, "extent")
	const effect = reader.child(frame, Ns.WP, "effectExtent")
	const resized = state.cx !== base.cx || state.cy !== base.cy
	const turned = state.rot !== base.rot

	if (resized) {
		if (extent < 0) throw new DocxWriteError("unsupported", "A resized drawing without an extent")

		setAttribute(document, extent, "cx", int(state.cx), edits)
		setAttribute(document, extent, "cy", int(state.cy), edits)

		if (state.cx !== base.cx) absoluteSize(document, frame, "sizeRelH", "pctWidth", edits)
		if (state.cy !== base.cy) absoluteSize(document, frame, "sizeRelV", "pctHeight", edits)
	}

	const xfrm = graphicTransform(part, drawing)

	if (resized && xfrm >= 0) {
		const ext = reader.child(xfrm, Ns.A, "ext")

		if (ext >= 0) {
			setAttribute(document, ext, "cx", int(state.cx), edits)
			setAttribute(document, ext, "cy", int(state.cy), edits)
		}
	}

	if (turned) {
		if (xfrm < 0) throw new DocxWriteError("unsupported", "A rotated drawing without a transform")

		setAttribute(document, xfrm, "rot", state.rot === 0 ? null : int(state.rot * 60000), edits)
	}

	// The effect extent covers what a rotation reaches past the extent (Word lays inline drawings out by it).
	if ((resized || turned) && (state.rot !== 0 || base.sourceRot !== 0)) {
		const [oldH, oldV] = rotationReach(base.cx, base.cy, base.sourceRot)
		const [newH, newV] = rotationReach(state.cx, state.cy, state.rot)
		const values = [
			drawing.effectLeft - oldH + newH,
			drawing.effectTop - oldV + newV,
			drawing.effectRight - oldH + newH,
			drawing.effectBottom - oldV + newV
		]

		if (effect >= 0) {
			;["l", "t", "r", "b"].forEach((name, k) => setAttribute(document, effect, name, int(values[k] ?? 0), edits))
		} else if (extent >= 0) {
			edits.push({
				from: document.end(extent),
				to: document.end(extent),
				text: `<${qualify("effectExtent")} l="${int(values[0] ?? 0)}" t="${int(values[1] ?? 0)}" r="${int(values[2] ?? 0)}" b="${int(values[3] ?? 0)}"/>`
			})
		}
	}

	if (state.descr !== base.descr) {
		const value = state.descr === "" ? null : state.descr

		if (drawing.docPr < 0) throw new DocxWriteError("unsupported", "A described drawing without wp:docPr")

		setAttribute(document, drawing.docPr, "descr", value, edits)

		// A picture marked decorative has no alt text (Word disables it while marked): described, it is not.
		if (value !== null) {
			const end = document.end(drawing.docPr)

			for (let at = drawing.docPr + 1; at < document.count && document.start(at) < end; at++) {
				if (
					document.kinds[at] === NodeKind.Element &&
					document.localName(at) === "decorative" &&
					/^(1|true|on)$/.test(attributeOf(document, at, "val") ?? "")
				)
					setAttribute(document, at, "val", "0", edits)
			}
		}

		const picture = drawing.graphicData < 0 ? -1 : reader.descendant(drawing.graphicData, Ns.Pic, "cNvPr")

		if (picture >= 0) setAttribute(document, picture, "descr", value, edits)
	}

	// A picture swapped in place (pictureEdits) loses its a:srcRect with its blip.
	if (!sameCrop(state.crop, base.crop) && typeof drawingEdit(node)?.media !== "number") {
		const fill = pictureFill(part, drawing)
		const blip = fill < 0 ? -1 : reader.child(fill, Ns.A, "blip")

		if (blip < 0) throw new DocxWriteError("unsupported", "A cropped drawing without a picture fill")

		const rect = reader.child(fill, Ns.A, "srcRect")
		const markup = srcRectMarkup(document.prefix(blip), state.crop)

		// CT_BlipFillProperties: blip, srcRect, then the fill mode.
		if (rect >= 0) edits.push({ from: document.start(rect), to: document.end(rect), text: markup })
		else edits.push({ from: document.end(blip), to: document.end(blip), text: markup })
	}

	placementEdits(w, part, drawing, base, state, edits)
}

function placementEdits(
	w: StoryWriting,
	part: StoryPart,
	drawing: DrawingInfo,
	base: DrawingState,
	state: DrawingState,
	edits: Edit[]
): void {
	const document = part.document
	const reader = part.reader
	const frame = drawing.frame
	const wasInline = base.wrap === "inline"
	const inline = state.wrap === "inline"
	const moved = state.posX !== base.posX || state.posY !== base.posY || state.relH !== base.relH || state.relV !== base.relV
	const rewrapped = state.wrap !== base.wrap
	const kids = children(document, frame)
	const named = (local: string): number => kids.find(child => reader.group(child) === Ns.WP && document.localName(child) === local) ?? -1
	const wrapNode = kids.find(child => reader.group(child) === Ns.WP && WRAP_ELEMENTS.has(document.localName(child))) ?? -1
	const extent = named("extent")
	const docPr = named("docPr")
	const effect = named("effectExtent")
	const prefix = document.prefix(frame)
	const local = (name: string): string => (prefix === "" ? name : `${prefix}:${name}`)
	const position = (axis: "H" | "V"): string =>
		axis === "H" ? positionMarkup(prefix, "H", state.relH, state.posX ?? 0) : positionMarkup(prefix, "V", state.relV, state.posY ?? 0)

	if (wasInline && inline) return

	if (!wasInline && !inline) {
		if (rewrapped) {
			const markup = wrapMarkup(part, prefix, state.wrap, wrapNode)

			if (wrapNode >= 0) edits.push({ from: document.start(wrapNode), to: document.end(wrapNode), text: markup })
			else if (docPr >= 0) edits.push({ from: document.start(docPr), to: document.start(docPr), text: markup })
			else throw new DocxWriteError("unsupported", "An anchor without wp:docPr")

			const behind = state.wrap === "behind"

			if (behind !== (drawing.anchor?.behindDoc ?? false)) setAttribute(document, frame, "behindDoc", behind ? "1" : "0", edits)
		}

		if (moved) {
			const h = named("positionH")
			const v = named("positionV")

			if (h >= 0) edits.push({ from: document.start(h), to: document.end(h), text: position("H") })
			if (v >= 0) edits.push({ from: document.start(v), to: document.end(v), text: position("V") })
			if ((h < 0 || v < 0) && extent >= 0) {
				edits.push({
					from: document.start(extent),
					to: document.start(extent),
					text: (h < 0 ? position("H") : "") + (v < 0 ? position("V") : "")
				})
			}

			// Offsets count only while the anchor is not placed by simplePos.
			if (drawing.anchor?.simplePos === true) setAttribute(document, frame, "simplePos", "0", edits)
		}

		return
	}

	if (wasInline) {
		// Inline to floating: an anchor with Word's defaults, the inline's other attributes (wp14 ids) kept.
		const kept = keptAttributes(document, frame, new Set(["distT", "distB", "distL", "distR"]))
		const behind = state.wrap === "behind" ? "1" : "0"
		const z = zOrder(w.env.context, `source:${String(drawing.index)}`)

		edits.push({
			from: document.start(frame),
			to: document.openEnd(frame),
			text: `<${local("anchor")} distT="0" distB="0" distL="${String(FLOAT_DISTANCE)}" distR="${String(FLOAT_DISTANCE)}" simplePos="0" relativeHeight="${String(z)}" behindDoc="${behind}" locked="0" layoutInCell="1" allowOverlap="1"${kept}>`
		})

		if (extent < 0 || docPr < 0) throw new DocxWriteError("unsupported", "An inline drawing without wp:extent or wp:docPr")

		edits.push({
			from: document.start(extent),
			to: document.start(extent),
			text: `<${local("simplePos")} x="0" y="0"/>${position("H")}${position("V")}`
		})

		const after = effect >= 0 ? effect : extent

		edits.push({ from: document.end(after), to: document.end(after), text: wrapMarkup(part, prefix, state.wrap, -1) })
		edits.push({ from: document.contentEnd(frame), to: document.end(frame), text: `</${local("anchor")}>` })

		return
	}

	// Floating to inline: what CT_Inline has, nothing of the anchor's placement.
	const kept = keptAttributes(
		document,
		frame,
		new Set([
			"distT",
			"distB",
			"distL",
			"distR",
			"simplePos",
			"relativeHeight",
			"behindDoc",
			"locked",
			"layoutInCell",
			"allowOverlap",
			"hidden"
		])
	)

	edits.push({
		from: document.start(frame),
		to: document.openEnd(frame),
		text: `<${local("inline")} distT="0" distB="0" distL="0" distR="0"${kept}>`
	})

	for (const child of kids) {
		if (reader.group(child) === Ns.WP && INLINE_CHILDREN.has(document.localName(child))) continue
		if (reader.group(child) === Ns.A && document.localName(child) === "graphic") continue

		edits.push({ from: document.start(child), to: document.end(child), text: "" })
	}

	edits.push({ from: document.contentEnd(frame), to: document.end(frame), text: `</${local("inline")}>` })
}

// ---- the VML fallback of a drawing inside mc:AlternateContent ------------------------------------------------

// A VML style ("name:value;...") with properties set (null: removed), the others as they were.
export function editStyle(style: string, changes: ReadonlyMap<string, string | null>): string {
	const done = new Set<string>()
	const out: string[] = []

	for (const piece of style.split(";")) {
		const colon = piece.indexOf(":")
		const key = (colon < 0 ? piece : piece.slice(0, colon)).trim().toLowerCase()
		const change = changes.get(key)

		if (change === undefined) {
			if (piece.trim() !== "") out.push(piece)

			continue
		}

		// A property written twice: the first one carries the change, the rest go.
		if (done.has(key)) continue

		done.add(key)

		if (change !== null) out.push(`${key}:${change}`)
	}

	for (const [key, change] of changes) if (!done.has(key) && change !== null) out.push(`${key}:${change}`)

	return out.join(";")
}

// The value of property `key` of a VML style (the first, as editStyle changes it), or undefined.
function styleValue(style: string, key: string): string | undefined {
	for (const piece of style.split(";")) {
		const colon = piece.indexOf(":")

		if (colon >= 0 && piece.slice(0, colon).trim().toLowerCase() === key) return piece.slice(colon + 1).trim()
	}

	return undefined
}

// The VML shape of the mc:Fallback of AlternateContent `alternate`: the first v:* drawing element of its
// w:pict (not a shape type), or -1.
function fallbackShape(part: StoryPart, alternate: number): number {
	const document = part.document
	const reader = part.reader
	const fallback = reader.child(alternate, Ns.MC, "Fallback")
	const pict = fallback < 0 ? -1 : reader.child(fallback, Ns.W, "pict")

	if (pict < 0) return -1

	for (const child of children(document, pict)) {
		if (reader.group(child) === Ns.V && document.localName(child) !== "shapetype") return child
	}

	return -1
}

function fallbackEdits(
	part: StoryPart,
	alternate: number,
	base: DrawingState,
	state: DrawingState,
	edits: Edit[],
	swapped: string | undefined
): void {
	const shape = fallbackShape(part, alternate)

	// No VML copy to keep in line: the fallback is something else (a picture of a chart, say), kept as it is.
	if (shape < 0) return

	const document = part.document
	const reader = part.reader
	const changes = new Map<string, string | null>()

	if (state.cx !== base.cx || state.cy !== base.cy) {
		const style = attributeOf(document, shape, "style") ?? ""

		changes.set("width", points(state.cx))
		changes.set("height", points(state.cy))

		// Relative sizes (tenths of a percent) turn absolute as the DrawingML's do.
		for (const [key, changed] of [
			["mso-width-percent", state.cx !== base.cx],
			["mso-height-percent", state.cy !== base.cy]
		] as const) {
			if (changed && parseFraction(styleValue(style, key) ?? "", Number.NaN) > 0) changes.set(key, "0")
		}
	}

	if (state.rot !== base.rot) changes.set("rotation", state.rot === 0 ? null : String(state.rot))

	const moved = state.posX !== base.posX || state.posY !== base.posY || state.relH !== base.relH || state.relV !== base.relV

	if (state.wrap === "inline" && base.wrap !== "inline") {
		for (const key of ["position", "margin-left", "margin-top", "z-index", "mso-position-horizontal", "mso-position-vertical"])
			changes.set(key, null)
		changes.set("mso-position-horizontal-relative", "char")
		changes.set("mso-position-vertical-relative", "line")
	} else if (state.wrap !== "inline" && (moved || base.wrap === "inline" || state.wrap !== base.wrap)) {
		changes.set("position", "absolute")
		changes.set("margin-left", points(state.posX ?? 0))
		changes.set("margin-top", points(state.posY ?? 0))
		changes.set("mso-position-horizontal", "absolute")
		changes.set("mso-position-horizontal-relative", VML_RELATIVE_H[state.relH] ?? "text")
		changes.set("mso-position-vertical", "absolute")
		changes.set("mso-position-vertical-relative", VML_RELATIVE_V[state.relV] ?? "text")

		if (state.wrap !== base.wrap) changes.set("z-index", state.wrap === "behind" ? "-251658240" : "251659264")
	}

	if (changes.size > 0) setAttribute(document, shape, "style", editStyle(attributeOf(document, shape, "style") ?? "", changes), edits)

	if (!sameCrop(state.crop, base.crop) || swapped !== undefined) {
		const imageData =
			children(document, shape).find(child => reader.group(child) === Ns.V && document.localName(child) === "imagedata") ?? -1
		const r = imageData < 0 ? undefined : reader.namespaces.prefixFor(imageData, R_URI)

		if (imageData >= 0 && swapped !== undefined && r !== undefined && r !== "")
			setAttribute(document, imageData, `${r}:id`, swapped, edits)

		if (imageData >= 0) {
			setAttribute(document, imageData, "cropleft", vmlCrop(state.crop.left), edits)
			setAttribute(document, imageData, "croptop", vmlCrop(state.crop.top), edits)
			setAttribute(document, imageData, "cropright", vmlCrop(state.crop.right), edits)
			setAttribute(document, imageData, "cropbottom", vmlCrop(state.crop.bottom), edits)
		}
	}

	if (state.descr !== base.descr) setAttribute(document, shape, "alt", state.descr === "" ? null : state.descr, edits)

	// Its wrapping: w10:wrap for text that flows around it, none for inline, behind and in front.
	if (state.wrap !== base.wrap) {
		const wrap = children(document, shape).find(child => reader.group(child) === Ns.W10 && document.localName(child) === "wrap") ?? -1
		const type =
			state.wrap === "square" || state.wrap === "tight" || state.wrap === "through"
				? state.wrap
				: state.wrap === "topAndBottom"
					? "topAndBottom"
					: null

		if (type === null) {
			if (wrap >= 0) edits.push({ from: document.start(wrap), to: document.end(wrap), text: "" })
		} else if (wrap >= 0) {
			setAttribute(document, wrap, "type", type, edits)
		} else {
			const bound = reader.namespaces.prefixFor(shape, W10_URI)
			const element = bound === undefined || bound === "" ? `w10:wrap xmlns:w10="${W10_URI}"` : `${bound}:wrap`

			if (document.isSelfClosing(shape)) {
				const at = document.openEnd(shape) - 2

				edits.push({ from: at, to: document.openEnd(shape), text: `><${element} type="${type}"/></${document.name(shape)}>` })
			} else {
				edits.push({ from: document.contentEnd(shape), to: document.contentEnd(shape), text: `<${element} type="${type}"/>` })
			}
		}
	}
}

function writeChangedDrawing(w: StoryWriting, out: Out, node: Node): void {
	const context = w.env.context
	const part = w.part
	const drawing = sourceDrawing(context, node)
	const own = src(node)

	if (drawing === undefined || drawing.part !== part.index)
		throw new DocxWriteError("missingSource", "A changed drawing without its DrawingML")

	const base = drawingState(context, node.type.create({ ...node.attrs, overrides: null }))
	const state = drawingState(context, node)

	if (!base.placeable && (state.wrap !== base.wrap || state.posX !== base.posX || state.posY !== base.posY || state.descr !== base.descr))
		throw new DocxWriteError("unsupported", "A change to the placement of a drawing that has none to change")
	if (!base.resizable && (state.cx !== base.cx || state.cy !== base.cy))
		throw new DocxWriteError("unsupported", "A resized drawing that cannot be resized")
	if (!base.rotatable && state.rot !== base.rot) throw new DocxWriteError("unsupported", "A rotated drawing that cannot be rotated")
	if (!base.croppable && !sameCrop(state.crop, base.crop))
		throw new DocxWriteError("unsupported", "A cropped drawing that cannot be cropped")

	const document = part.document
	const element = part.node(own)
	const edits: Edit[] = []

	frameEdits(w, part, node, drawing, base, state, edits)

	const swapped = pictureEdits(w, part, node, drawing, edits)

	if (part.kinds[own] === Item.AlternateContent) fallbackEdits(part, element, base, state, edits, swapped)

	writeEdited(w, out, document.start(element), document.end(element), edits)
}

// The relationship id of the picture of media `handle` from the part being written (a placeholder the layout
// resolves while laying out edits, when there is no save to add parts to).
function pictureRelationship(w: StoryWriting, handle: number): string {
	const parts = w.env.parts

	return parts === undefined ? `filenMedia${String(handle)}` : mediaRelationship(parts, w.part.name, handle)
}

// Extensions of a blip about its picture, gone with it: the SVG it stands for, artistic effects of it.
const PICTURE_EXTENSIONS = new Set([SVG_EXTENSION, "{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}"])

// A source picture whose picture Change Picture swapped (its overrides' media): its a:blip anew (the new
// picture's relationship, an SVG's PNG with the SVG as Word writes it; the blip's own attributes, effects and
// other extensions kept) and no a:srcRect. The relationship the VML copy shows, or undefined.
function pictureEdits(w: StoryWriting, part: StoryPart, node: Node, drawing: DrawingInfo, edits: Edit[]): string | undefined {
	const edit = drawingEdit(node)
	const media = edit?.media

	if (typeof media !== "number") return undefined

	const document = part.document
	const reader = part.reader
	const fill = pictureFill(part, drawing)
	const blip = fill < 0 ? -1 : reader.child(fill, Ns.A, "blip")
	const bytes = w.env.context.media.bytes(media)
	const image = bytes === undefined ? undefined : sniffImage(bytes)

	if (blip < 0) throw new DocxWriteError("unsupported", "A swapped picture without a picture fill")
	if (image === undefined || typeof image === "string") throw new DocxWriteError("missingSource", `The picture of media ${String(media)}`)
	if (image.format === "svg" && typeof edit?.fallback !== "number")
		throw new DocxWriteError("missingSource", "An SVG picture without its PNG")

	const bound = reader.namespaces.prefixFor(blip, R_URI)
	const r = bound === undefined || bound === "" ? "r" : bound
	const declare = bound === undefined || bound === "" ? ` xmlns:r="${R_URI}"` : ""
	const a = document.prefix(blip)
	const qa = (local: string): string => (a === "" ? local : `${a}:${local}`)
	const shown = pictureRelationship(w, image.format === "svg" ? (edit?.fallback as number) : media)
	let inner = ""
	let extensions = ""

	for (const child of children(document, blip)) {
		if (document.localName(child) !== "extLst") {
			inner += decodeUtf8(document.bytes, document.start(child), document.end(child))

			continue
		}

		for (const ext of children(document, child)) {
			if (!PICTURE_EXTENSIONS.has(attributeOf(document, ext, "uri") ?? ""))
				extensions += decodeUtf8(document.bytes, document.start(ext), document.end(ext))
		}
	}

	if (image.format === "svg")
		extensions += `<${qa("ext")} uri="${SVG_EXTENSION}"><asvg:svgBlip xmlns:asvg="${SVG_URI}" ${r}:embed="${pictureRelationship(w, media)}"/></${qa("ext")}>`

	const kept = keptAttributes(document, blip, new Set(["r:embed", "r:link", `${r}:embed`, `${r}:link`]))
	const content = inner + (extensions === "" ? "" : `<${qa("extLst")}>${extensions}</${qa("extLst")}>`)

	edits.push({
		from: document.start(blip),
		to: document.end(blip),
		text: `<${qa("blip")} ${r}:embed="${shown}"${declare}${kept}${content === "" ? "/>" : `>${content}</${qa("blip")}>`}`
	})

	const rect = reader.child(fill, Ns.A, "srcRect")

	if (rect >= 0) edits.push({ from: document.start(rect), to: document.end(rect), text: "" })

	return shown
}

// ---- a new picture ------------------------------------------------------------------------------------------

// The prefix the part's root binds `uri` to, or `fallback` declared on the element written (`declare`).
function prefixOf(part: StoryPart, uri: string, fallback: string): { prefix: string; declare: string } {
	const root = part.document.documentElement()
	const bound = part.reader.namespaces.prefixFor(root, uri)

	return bound === undefined || bound === ""
		? { prefix: fallback, declare: ` xmlns:${fallback}="${uri}"` }
		: { prefix: bound, declare: "" }
}

// The wp:docPr ids new pictures were written with in a save (how often each), and the ids copies of one were
// given (per id and copy, the same at every save of a session).
const DOCPR_STATE = {}
const copyIds = new WeakMap<DocxEditContext, Map<string, number>>()

// The id a new picture is written with: its own, or for a copy of one written already in this save (pasted,
// copied by a drag), one of its own (wp:docPr ids are unique in a document, ECMA-376 Part 1 20.4.2.5).
function uniqueDocPr(w: StoryWriting, docPr: number): number {
	const parts = w.env.parts

	if (parts === undefined) return docPr

	const written = parts.state(DOCPR_STATE, () => new Map<number, number>())
	const copy = written.get(docPr) ?? 0

	written.set(docPr, copy + 1)

	if (copy === 0) return docPr

	const context = w.env.context
	let given = copyIds.get(context)

	if (given === undefined) {
		given = new Map()
		copyIds.set(context, given)
	}

	const key = `${String(docPr)}#${String(copy)}`
	let id = given.get(key)

	if (id === undefined) {
		id = context.ids.next("docPr")
		given.set(key, id)
	}

	return id
}

function writeNewPicture(w: StoryWriting, out: Out, node: Node): void {
	const context = w.env.context
	const edit = drawingEdit(node)
	const media = edit?.media
	const docPr = edit?.docPr

	if (edit === null || typeof media !== "number" || typeof docPr !== "number") throw new DocxWriteError("missingSource", "A new drawing")

	const bytes = context.media.bytes(media)
	const image = bytes === undefined ? undefined : sniffImage(bytes)

	if (image === undefined || typeof image === "string") throw new DocxWriteError("missingSource", `The picture of media ${String(media)}`)

	const svg = image.format === "svg"
	const fallback = edit.fallback

	if (svg && typeof fallback !== "number") throw new DocxWriteError("missingSource", "An SVG picture without its PNG")

	const state = drawingState(context, node)
	const part = w.part
	const parts = w.env.parts
	// While laying out edits there is no save to add parts to: the picture is drawn as its frame.
	const relationship = (handle: number): string =>
		parts === undefined ? `filenMedia${String(handle)}` : mediaRelationship(parts, part.name, handle)
	const wp = prefixOf(part, WP_URI, "wp")
	const r = prefixOf(part, R_URI, "r")
	const q = (local: string): string => `${wp.prefix}:${local}`
	// Copies keep the original's z order (layout writes, which see no save, give them its id).
	const id = uniqueDocPr(w, docPr)
	const name = `Picture ${String(id)}`
	const descr = state.descr === "" ? "" : ` descr="${escapeAttribute(state.descr)}"`
	const cx = int(state.cx)
	const cy = int(state.cy)
	const blip = svg
		? `<a:blip ${r.prefix}:embed="${relationship(fallback as number)}"><a:extLst><a:ext uri="${SVG_EXTENSION}"><asvg:svgBlip xmlns:asvg="${SVG_URI}" ${r.prefix}:embed="${relationship(media)}"/></a:ext></a:extLst></a:blip>`
		: `<a:blip ${r.prefix}:embed="${relationship(media)}"/>`
	const [reachH, reachV] = rotationReach(state.cx, state.cy, state.rot)
	const graphic =
		`<${q("extent")} cx="${cx}" cy="${cy}"/><${q("effectExtent")} l="${int(reachH)}" t="${int(reachV)}" r="${int(reachH)}" b="${int(reachV)}"/>` +
		`${state.wrap === "inline" ? "" : wrapMarkup(part, wp.prefix, state.wrap, -1)}` +
		`<${q("docPr")} id="${String(id)}" name="${name}"${descr}/>` +
		`<${q("cNvGraphicFramePr")}><a:graphicFrameLocks xmlns:a="${A_URI}" noChangeAspect="1"/></${q("cNvGraphicFramePr")}>` +
		`<a:graphic xmlns:a="${A_URI}"><a:graphicData uri="${PIC_URI}"><pic:pic xmlns:pic="${PIC_URI}">` +
		`<pic:nvPicPr><pic:cNvPr id="0" name="${name}"${descr}/><pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr></pic:nvPicPr>` +
		`<pic:blipFill>${blip}${state.crop.left === 0 && state.crop.top === 0 && state.crop.right === 0 && state.crop.bottom === 0 ? "" : srcRectMarkup("a", state.crop)}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
		`<pic:spPr bwMode="auto"><a:xfrm${state.rot === 0 ? "" : ` rot="${int(state.rot * 60000)}"`}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
		"</pic:pic></a:graphicData></a:graphic>"
	const frame =
		state.wrap === "inline"
			? `<${q("inline")} distT="0" distB="0" distL="0" distR="0"${wp.declare}>${graphic}</${q("inline")}>`
			: `<${q("anchor")} distT="0" distB="0" distL="${String(FLOAT_DISTANCE)}" distR="${String(FLOAT_DISTANCE)}" simplePos="0" relativeHeight="${String(zOrder(context, `new:${String(docPr)}`))}" behindDoc="${state.wrap === "behind" ? "1" : "0"}" locked="0" layoutInCell="1" allowOverlap="1"${wp.declare}>` +
				`<${q("simplePos")} x="0" y="0"/>${positionMarkup(wp.prefix, "H", state.relH, state.posX ?? 0)}${positionMarkup(wp.prefix, "V", state.relV, state.posY ?? 0)}` +
				// The wrap element goes after the effect extent (CT_Anchor), which the graphic markup starts with.
				graphic +
				`</${q("anchor")}>`

	out.text(`<${w.names.el("drawing")}${r.declare}>${frame}</${w.names.el("drawing")}>`)
}

export const drawingAtoms: readonly AtomWriter[] = [
	{
		type: nodes.drawing,
		edited: (node: Node) => (node.attrs["overrides"] ?? null) !== null,
		write(w, out, node) {
			if (src(node) === NO_SOURCE) writeNewPicture(w, out, node)
			else writeChangedDrawing(w, out, node)
		}
	}
]
