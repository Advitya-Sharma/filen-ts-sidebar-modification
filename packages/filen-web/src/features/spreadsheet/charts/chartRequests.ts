import type { ChartEditType } from "@filen/office/react/sheet/charts"
import type { SheetGridModel } from "@filen/office/react/sheet"

// A chart asked for from outside the charts feature (the toolbar's Insert menu): of the selected cells, of a type. The
// charts feature registers what inserts it while it is mounted; a request without one does nothing.

const handlers = new WeakMap<SheetGridModel, (type: ChartEditType) => void>()

export function registerChartInsert(model: SheetGridModel, handler: (type: ChartEditType) => void): () => void {
	handlers.set(model, handler)

	return () => {
		if (handlers.get(model) === handler) handlers.delete(model)
	}
}

export function requestChart(model: SheetGridModel, type: ChartEditType): void {
	handlers.get(model)?.(type)
}
