import { attr, clamp, hexColor } from "./xml"

// DrawingML for converted pictures and shapes (ECMA-376 Part 1 §20.4 wordprocessingDrawing, §20.2 picture,
// Word 2010 wordprocessingShape).

export const EMU_PER_TWIP = 635

export interface Crop {
	// Fractions of the picture cut from each side (0–1).
	readonly left: number
	readonly top: number
	readonly right: number
	readonly bottom: number
}

export interface DrawingName {
	readonly id: number
	readonly name: string
	readonly description: string | undefined
}

function docPr(name: DrawingName): string {
	return `<wp:docPr${attr("id", name.id)}${attr("name", name.name)}${name.description === undefined ? "" : attr("descr", name.description)}/>`
}

function emu(value: number): number {
	return clamp(value, 0, 27273042316900)
}

function srcRect(crop: Crop | undefined): string {
	if (crop === undefined) return ""

	const part = (value: number): number => clamp(value * 100000, -10000000, 10000000)
	const sides: [string, number][] = [
		["l", part(crop.left)],
		["t", part(crop.top)],
		["r", part(crop.right)],
		["b", part(crop.bottom)]
	]

	if (sides.every(([, value]) => value === 0)) return ""

	return `<a:srcRect${sides.map(([name, value]) => (value === 0 ? "" : attr(name, value))).join("")}/>`
}

export interface Transform {
	// 60000ths of a degree, clockwise.
	readonly rotation?: number
	readonly flipH?: boolean
	readonly flipV?: boolean
}

function xfrm(width: number, height: number, transform: Transform): string {
	return `<a:xfrm${transform.rotation === undefined || transform.rotation === 0 ? "" : attr("rot", transform.rotation)}${transform.flipH === true ? ' flipH="1"' : ""}${
		transform.flipV === true ? ' flipV="1"' : ""
	}><a:off x="0" y="0"/><a:ext${attr("cx", emu(width))}${attr("cy", emu(height))}/></a:xfrm>`
}

export function pictureGraphic(
	relationship: string,
	width: number,
	height: number,
	name: DrawingName,
	crop: Crop | undefined,
	transform: Transform = {}
): string {
	return (
		'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
		`<pic:nvPicPr><pic:cNvPr id="0"${attr("name", name.name)}/><pic:cNvPicPr/></pic:nvPicPr>` +
		`<pic:blipFill><a:blip${attr("r:embed", relationship)}/>${srcRect(crop)}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
		`<pic:spPr>${xfrm(width, height, transform)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
		"</pic:pic></a:graphicData></a:graphic>"
	)
}

export interface ShapeLook {
	readonly preset: string
	// RGB, or undefined for no fill / no line.
	readonly fill: number | undefined
	readonly line: number | undefined
	// EMUs.
	readonly lineWidth: number
	readonly textBox: boolean
	// Insets of the text in EMUs: left, top, right, bottom.
	readonly insets: readonly [number, number, number, number]
}

export function shapeGraphic(
	width: number,
	height: number,
	look: ShapeLook,
	content: string | undefined,
	transform: Transform = {}
): string {
	const fill = look.fill === undefined ? "<a:noFill/>" : `<a:solidFill><a:srgbClr${attr("val", hexColor(look.fill))}/></a:solidFill>`
	const line =
		look.line === undefined
			? "<a:ln><a:noFill/></a:ln>"
			: `<a:ln${attr("w", clamp(look.lineWidth, 0, 20116800))}><a:solidFill><a:srgbClr${attr("val", hexColor(look.line))}/></a:solidFill></a:ln>`
	const [left, top, right, bottom] = look.insets

	return (
		'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>' +
		(look.textBox ? '<wps:cNvSpPr txBox="1"/>' : "<wps:cNvSpPr/>") +
		`<wps:spPr>${xfrm(width, height, transform)}<a:prstGeom${attr("prst", look.preset)}><a:avLst/></a:prstGeom>${fill}${line}</wps:spPr>` +
		(content === undefined ? "" : `<wps:txbx><w:txbxContent>${content}</w:txbxContent></wps:txbx>`) +
		`<wps:bodyPr${attr("lIns", emu(left))}${attr("tIns", emu(top))}${attr("rIns", emu(right))}${attr("bIns", emu(bottom))}/>` +
		"</wps:wsp></a:graphicData></a:graphic>"
	)
}

export function inline(graphic: string, width: number, height: number, name: DrawingName): string {
	return (
		`<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent${attr("cx", emu(width))}${attr("cy", emu(height))}/>` +
		`${docPr(name)}<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>${graphic}</wp:inline></w:drawing>`
	)
}

export interface Placement {
	// FSPA bx: 0 margin, 1 page, 2 column; by: 0 margin, 1 page, 2 paragraph.
	readonly relativeH: number
	readonly relativeV: number
	// EMUs from the reference edge.
	readonly x: number
	readonly y: number
	// FSPA wr / wrk.
	readonly wrap: number
	readonly wrapSide: number
	readonly behindText: boolean
	// z-order among the anchored drawings.
	readonly order: number
	// Wrap distances in EMUs: top, bottom, left, right.
	readonly distances: readonly [number, number, number, number]
}

const WRAP_SIDES = ["bothSides", "left", "right", "largest"]

function wrapping(placement: Placement): string {
	const side = attr("wrapText", WRAP_SIDES[placement.wrapSide] ?? "bothSides")
	const polygon =
		'<wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon>'

	switch (placement.wrap) {
		case 1:
			return "<wp:wrapTopAndBottom/>"
		case 3:
			return "<wp:wrapNone/>"
		case 4:
			return `<wp:wrapTight${side}>${polygon}</wp:wrapTight>`
		case 5:
			return `<wp:wrapThrough${side}>${polygon}</wp:wrapThrough>`
		default:
			return `<wp:wrapSquare${side}/>`
	}
}

export function anchor(graphic: string, width: number, height: number, name: DrawingName, placement: Placement): string {
	const [top, bottom, left, right] = placement.distances
	const relativeH = placement.relativeH === 1 ? "page" : placement.relativeH === 2 ? "column" : "margin"
	const relativeV = placement.relativeV === 1 ? "page" : placement.relativeV === 2 ? "paragraph" : "margin"
	// ST_PositionOffset is a 32-bit integer.
	const offset = (value: number): number => clamp(value, -2147483648, 2147483647)

	return (
		`<w:drawing><wp:anchor${attr("distT", emu(top))}${attr("distB", emu(bottom))}${attr("distL", emu(left))}${attr("distR", emu(right))} simplePos="0"${attr(
			"relativeHeight",
			clamp(placement.order, 0, 0xfffffff)
		)}${attr("behindDoc", placement.behindText ? 1 : 0)} locked="0" layoutInCell="1" allowOverlap="1">` +
		'<wp:simplePos x="0" y="0"/>' +
		`<wp:positionH${attr("relativeFrom", relativeH)}><wp:posOffset>${String(offset(placement.x))}</wp:posOffset></wp:positionH>` +
		`<wp:positionV${attr("relativeFrom", relativeV)}><wp:posOffset>${String(offset(placement.y))}</wp:posOffset></wp:positionV>` +
		`<wp:extent${attr("cx", emu(width))}${attr("cy", emu(height))}/>${wrapping(placement)}${docPr(name)}<wp:cNvGraphicFramePr/>${graphic}</wp:anchor></w:drawing>`
	)
}
