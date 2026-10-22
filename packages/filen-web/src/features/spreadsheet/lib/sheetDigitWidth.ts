import type { SheetOpened } from "@filen/office/core/sheet"
import { DEFAULT_DIGIT_PX } from "@filen/office/react/sheet"
import type { SheetEventRelay } from "@/features/spreadsheet/lib/sheetEvents"
import { isTextSheetKind } from "@/features/spreadsheet/lib/sheetHosts"

// Column widths are counted in the widest digit of the workbook's default font (ECMA-376 Part 1 §18.3.1.13), which
// the session works out from that font's name and size once asked, the same on every device. Done once a workbook
// opens, before its grid shows: a width other than Calibri 11's 7 px remakes the layouts made before it (the active
// sheet's, those its events carry).
export async function withWorkbookDigit(opened: SheetOpened, events: SheetEventRelay): Promise<SheetOpened> {
	const { info, session } = opened

	if (isTextSheetKind(info.kind)) {
		return opened
	}

	const { px } = await session.query("digitWidth", {})

	if (px === DEFAULT_DIGIT_PX) {
		return opened
	}

	await events.relayout(sheet => session.layout(sheet))

	return { ...opened, info: { ...info, activeLayout: await session.layout(info.activeSheet) } }
}
