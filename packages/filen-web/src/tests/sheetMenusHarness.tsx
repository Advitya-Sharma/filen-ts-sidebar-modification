import type { ReactNode } from "react"
import { useSheetGrid, type SheetGridModel, type SheetGridState } from "@filen/office/react/sheet"

// Renders `children(grid)` with the model's live state.
export function GridHarness({ model, children }: { model: SheetGridModel; children: (grid: SheetGridState) => ReactNode }) {
	const grid = useSheetGrid(model)

	return <>{children(grid)}</>
}
