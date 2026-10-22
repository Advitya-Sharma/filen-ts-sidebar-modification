import { drawingTheme, type DrawingTheme } from "../../drawingml/color"
import { DrawingList } from "../../drawingml/displayList"
import type { ImageResolver } from "../../drawingml/picture"
import { D, drawingReader } from "../../drawingml/reader"
import { convertGraphic, type ShapeContext } from "../../drawingml/shape"
import { MAX_METAFILE_BYTES, convertMetafile, expandMetafiles, type MetafileConversion } from "../../metafile/index"
import { parseXml, type XmlDocument } from "../../xml/index"
import { relationshipKind } from "../comments/parts"
import type { XlsxWorkbookModel } from "../model/types"
import { childByLocal } from "../refs/markup"
import { EMU_PER_PX, type DrawingAnchor } from "./anchor"
import type { DrawingObjectRecord } from "./read"
import type { DrawingStore } from "./store"

// One drawing object as a display list (core/drawingml), drawn into the frame of its own size: pictures, shapes,
// connectors and groups through the shared shape walker, EMF and WMF pictures converted into the list. Media are
// named by part, never read here (the renderer asks for them); charts and other graphic frames are not drawn here.

// What a renderer gets: a compact display list (every typed array owns its buffer).
export interface SheetDrawingData {
	readonly width: number
	readonly height: number
	readonly commands: Uint8Array
	readonly coords: Float32Array
	readonly itemCount: number
	readonly itemFloats: Float32Array
	readonly itemInts: Int32Array
	readonly paints: DrawingList["paints"]
	readonly strokes: DrawingList["strokes"]
	readonly images: DrawingList["images"]
	readonly texts: DrawingList["texts"]
	readonly shadows: DrawingList["shadows"]
	readonly clips?: DrawingList["clips"]
	readonly itemClips?: Int32Array
}

// The 12 scheme colours of SpreadsheetML's theme, in clrScheme order (the workbook's theme reads them so).
const SCHEME = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]
// SpreadsheetML has no colour map: the default one (ECMA-376 Part 1 §19.3.1.6 defaults, as Excel maps them).
const MAPPED: Readonly<Record<string, string>> = { bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2" }

function hex(value: number): string {
	return (value & 0xffffff).toString(16).padStart(6, "0").toUpperCase()
}

const themes = new WeakMap<XlsxWorkbookModel, Promise<DrawingTheme>>()

// The workbook's theme for drawings (scheme colours, fonts and the format scheme its style references use).
export function workbookDrawingTheme(workbook: XlsxWorkbookModel): Promise<DrawingTheme> {
	let theme = themes.get(workbook)

	if (theme === undefined) {
		theme = (async () => {
			const colors = workbook.theme?.colors ?? []
			const scheme = (name: string): string | undefined => {
				const index = SCHEME.indexOf(MAPPED[name] ?? name)
				const color = index < 0 ? undefined : colors[index]

				return color === undefined ? undefined : hex(color)
			}
			let document: XmlDocument | undefined

			try {
				const part = (await workbook.parts.relationships(workbook.workbookPart)).all.find(
					relationship => !relationship.external && relationshipKind(relationship.type) === "theme"
				)?.partName

				if (part !== undefined && workbook.parts.get(part)?.source !== null) document = parseXml(await workbook.parts.read(part))
			} catch {
				document = undefined
			}

			return drawingTheme(document, scheme)
		})()
		themes.set(workbook, theme)
	}

	return theme
}

// The size an object is drawn at, EMU: its own xfrm's (what the file laid it out for), else its anchor's extent.
function frameSize(document: XmlDocument, object: DrawingObjectRecord): { cx: number; cy: number } {
	const reader = drawingReader(document)
	const content = object.content

	if (content >= 0) {
		const properties =
			document.localName(content) === "graphicFrame"
				? content
				: childByLocal(document, content, document.localName(content) === "grpSp" ? "grpSpPr" : "spPr")
		const xfrm = properties < 0 ? -1 : childByLocal(document, properties, "xfrm")
		const ext = xfrm < 0 ? -1 : reader.child(xfrm, D.A, "ext")
		const cx = Number(document.attribute(ext, "cx") ?? Number.NaN)
		const cy = Number(document.attribute(ext, "cy") ?? Number.NaN)

		if (ext >= 0 && cx > 0 && cy > 0 && Number.isFinite(cx) && Number.isFinite(cy)) return { cx, cy }
	}

	return extentOf(object.anchor)
}

// An anchor's own extent (one-cell, absolute); a two-cell anchor's size depends on the cells: an inch square stands in.
function extentOf(anchor: DrawingAnchor | null): { cx: number; cy: number } {
	if (anchor !== null && anchor.mode !== "twoCell" && anchor.cx > 0 && anchor.cy > 0) return { cx: anchor.cx, cy: anchor.cy }

	return { cx: 96 * EMU_PER_PX, cy: 96 * EMU_PER_PX }
}

export function compact(list: DrawingList): SheetDrawingData {
	const out = list.compact()

	return {
		width: out.width,
		height: out.height,
		commands: out.commands,
		coords: out.coords,
		itemCount: out.itemCount,
		itemFloats: out.itemFloats,
		itemInts: out.itemInts,
		paints: out.paints,
		strokes: out.strokes,
		// Bitmaps a metafile carries are copied: the answer is transferred, and nothing here keeps them.
		images: out.images.map(image => (image.bytes === undefined ? image : { ...image, bytes: image.bytes.slice() })),
		texts: out.texts,
		shadows: out.shadows,
		...(out.itemClips === null ? {} : { clips: out.clips, itemClips: out.itemClips })
	}
}

export type DrawingRender =
	// live: it holds charts (in a group), whose data follows the cells: what it shows, asked again with this stamp as
	// `since` when the session's version moves.
	| { readonly kind: "drawing"; readonly drawing: SheetDrawingData; readonly live?: string }
	// A drawing holding charts still shows what the request's `since` names.
	| { readonly kind: "same"; readonly stamp: string }
	// chart: a chart frame (its own renderer draws it); graphic: a frame of another kind; empty: nothing to draw.
	| { readonly kind: "placeholder"; readonly reason: "chart" | "graphic" | "empty" | "failed" }

// Draws one object of a drawing part. Never throws: anything unexpected is a "failed" placeholder.
// `frame`: what draws the charts a group holds (the charts feature's nestedChartFrames).
export async function renderObject(
	store: DrawingStore,
	part: string,
	document: XmlDocument,
	object: DrawingObjectRecord,
	frame?: ShapeContext["frame"]
): Promise<DrawingRender> {
	if (object.kind === "chart" || object.kind === "chartex") return { kind: "placeholder", reason: "chart" }
	if (object.content < 0) return { kind: "placeholder", reason: "empty" }

	try {
		const relationships = store.relationships(part)
		const images: ImageResolver = id => {
			const relationship = relationships.find(candidate => candidate.id === id)
			const partName = relationship === undefined || relationship.external ? undefined : relationship.partName

			if (partName === undefined || store.bytes(partName) === null) return undefined
			if (store.bytes(partName) === undefined && store.workbook.parts.get(partName) === undefined) return undefined

			return { partName, contentType: store.workbook.parts.get(partName)?.contentType }
		}
		const theme = await workbookDrawingTheme(store.workbook)
		const size = frameSize(document, object)
		const parent = document.parent(object.content)
		const result = convertGraphic(
			{ list: new DrawingList(), theme, images, effects: false, ...(frame === undefined ? {} : { frame }) },
			document,
			// The walker draws the children of the node it is given: the content alone, through its parent.
			parent,
			size.cx,
			size.cy
		)

		if (object.kind === "graphic" && result.list.itemCount === 0) return { kind: "placeholder", reason: "graphic" }

		const list = await expandMetafiles(result.list, image => metafile(store, image.partName))

		return { kind: "drawing", drawing: compact(list) }
	} catch {
		return { kind: "placeholder", reason: "failed" }
	}
}

async function metafile(store: DrawingStore, partName: string | undefined): Promise<MetafileConversion | undefined> {
	if (partName === undefined) return undefined

	try {
		const part = store.workbook.parts.get(partName)

		if (part?.source !== null && part?.source !== undefined && part.source.uncompressedSize > MAX_METAFILE_BYTES) return undefined

		const bytes = await store.media(partName)

		return bytes === null || bytes.length > MAX_METAFILE_BYTES ? undefined : convertMetafile(bytes)
	} catch {
		return undefined
	}
}
