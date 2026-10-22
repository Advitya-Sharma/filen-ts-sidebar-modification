import { D, drawingReader } from "../../drawingml/reader"
import type { SheetId } from "../../sheet/protocol"
import { NodeKind, parseXml } from "../../xml/index"
import { heldOrSource, relationshipKind } from "../comments/parts"
import type { XlsxDocument } from "../edit/document"
import { partRelationships } from "./parts"

// The chart of a chartsheet (ECMA-376 Part 1 §18.3.1.18 chartsheet, §20.5.2.1 absoluteAnchor): the drawing its
// <drawing> names holds one graphic frame, whose chart this is, at the size the anchor gives it.

export interface ChartSheetChart {
	readonly part: string
	// EMU.
	readonly width: number
	readonly height: number
}

// Excel's default chartsheet chart when its anchor gives no size: a landscape page's printable area.
const DEFAULT_WIDTH = 8_670_925
const DEFAULT_HEIGHT = 6_292_850

export async function chartSheetChart(document: XlsxDocument, sheet: SheetId): Promise<ChartSheetChart | null> {
	const entry = document.entry(sheet)

	if (entry === undefined || entry.kind !== "chartsheet" || entry.part === "") return null

	const workbook = document.workbook

	try {
		const drawing = (await partRelationships(workbook, entry.part)).find(
			relationship => !relationship.external && relationshipKind(relationship.type) === "drawing"
		)?.partName

		if (drawing === undefined) return null

		const bytes =
			heldOrSource(workbook, drawing, new Map()) ??
			(workbook.parts.get(drawing)?.source === null ? null : await workbook.parts.read(drawing))

		if (bytes === null || bytes === undefined) return null

		const reader = drawingReader(parseXml(bytes))
		const document$ = reader.document
		let chart = -1

		for (let node = 0; node < document$.count && chart < 0; node++)
			if (document$.kind(node) === NodeKind.Element && reader.is(node, D.C, "chart")) chart = node

		const id = chart < 0 ? undefined : reader.attrNs(chart, D.R, "id")
		const part =
			id === undefined
				? undefined
				: (await partRelationships(workbook, drawing)).find(relationship => relationship.id === id && !relationship.external)
						?.partName

		if (part === undefined) return null

		// The anchor the frame is in: an absoluteAnchor's ext, else Excel's page-sized default.
		let anchor = document$.parent(chart)

		while (anchor >= 0 && !reader.is(anchor, D.Xdr, "absoluteAnchor")) anchor = document$.parent(anchor)

		const ext = reader.child(anchor, D.Xdr, "ext")
		const width = reader.number(ext, "cx", 0)
		const height = reader.number(ext, "cy", 0)

		return width > 0 && height > 0 ? { part, width, height } : { part, width: DEFAULT_WIDTH, height: DEFAULT_HEIGHT }
	} catch {
		return null
	}
}
