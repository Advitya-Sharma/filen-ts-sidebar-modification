import type { SpreadsheetKind } from "@/features/spreadsheet/lib/model"
import type { LayerKey } from "@/features/spreadsheet/lib/sizeLayer"

// Where the sizes on show live: in the file while it is a workbook this viewer may edit (a resize made
// before its proof waits for it, like any edit); beside it (lib/sizeLayer.ts) for everything else, whose
// layer an editable workbook then ignores.
export function sizesInFile(kind: SpreadsheetKind, canEdit: boolean): boolean {
	return kind === "xlsx" && canEdit
}

export function layerKeyFor(stable: string | undefined, documentKey: string): LayerKey {
	return stable === undefined ? { kind: "session", id: documentKey } : { kind: "stable", id: stable }
}
