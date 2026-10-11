import type { XmlDocument } from "../xml/index"
import { D, drawingReader } from "./reader"
import { convertShapeTree, type ShapeContext, type ShapeConversion } from "./shape"

// A SmartArt diagram as drawn from the drawing Office caches next to the diagram data (dsp:drawing, the
// part named by dsp:dataModelExt@relId). The layout algorithms themselves are not run: without the cached
// drawing the host shows the diagram's fallback or a placeholder. Undefined when the part is not a
// dsp:drawing.
export function convertSmartArt(context: ShapeContext, document: XmlDocument, width: number, height: number): ShapeConversion | undefined {
	const reader = drawingReader(document)
	const root = document.documentElement()

	if (!reader.is(root, D.Dsp, "drawing")) return undefined

	const tree = reader.child(root, D.Dsp, "spTree")

	if (tree < 0) return undefined

	return convertShapeTree(context, document, tree, width, height)
}
