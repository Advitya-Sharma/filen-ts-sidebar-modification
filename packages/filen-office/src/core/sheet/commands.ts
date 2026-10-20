import type { ByteSource } from "../opc/io"
import { openXlsx, sheetOpenFailure, XlsxReadError } from "../xlsx/read/index"
import { XlsxDocument } from "../xlsx/edit/document"
import { clearCells, setCells } from "../xlsx/edit/cells"
import { copyCells, pasteCells } from "../xlsx/edit/clipboard"
import { applyFormat } from "../xlsx/edit/format"
import { fitSample, resizeAxis } from "../xlsx/edit/resize"
import { addSheet, moveAxis, renameSheet } from "../xlsx/edit/structure-basic"
import type { SheetEvent, SheetOpenFailure, SheetOpened, SheetSessionOptions } from "./protocol"
import { defineSheetFeature, sheetRegistry, type SheetFeature } from "./registry"
import { RenderStyleTable, SheetSession, sessionFeature } from "./session"

// The xlsx session's own features (cells, formats, sizes, rows and columns, sheets, the clipboard) and the open
// that puts a session together. Later features join the list in their own modules.

export const cellsFeature = defineSheetFeature<XlsxDocument>({
	id: "cells",
	commands: {
		"cells.set": {
			info: { write: true },
			run: (context, command) => setCells(context.doc, command.sheet, command.cells, context.changes)
		},
		"cells.clear": {
			info: { write: true },
			run: (context, command) => clearCells(context.doc, command.sheet, command.ranges, command.what, context.changes)
		}
	}
})

export const formatFeature = defineSheetFeature<XlsxDocument>({
	id: "format",
	commands: {
		"format.apply": {
			info: { write: true, protection: "formatCells" },
			run: (context, command) => applyFormat(context.doc, command.sheet, command.ranges, command.patch, context.changes)
		}
	}
})

export const axisFeature = defineSheetFeature<XlsxDocument>({
	id: "axis",
	commands: {
		"axis.insert": {
			info: { write: true },
			run: (context, command) =>
				moveAxis(context.doc, command.sheet, command.axis, "insert", command.at, command.count, context.changes, context.slicing)
		},
		"axis.delete": {
			info: { write: true },
			run: (context, command) =>
				moveAxis(context.doc, command.sheet, command.axis, "delete", command.at, command.count, context.changes, context.slicing)
		},
		"axis.resize": {
			info: { write: true },
			run: (context, command) => {
				const refusal = context.doc.protects(command.sheet, command.axis === "rows" ? "formatRows" : "formatColumns")

				return refusal !== null
					? { refused: refusal }
					: resizeAxis(context.doc, command.sheet, command.axis, command.indices, command.sizes, context.changes)
			}
		}
	},
	queries: {
		fitSample: { run: (context, request) => fitSample(context.doc, request.sheet, request.axis, request.index, request.limit) }
	}
})

export const sheetsFeature = defineSheetFeature<XlsxDocument>({
	id: "sheets",
	commands: {
		"sheet.add": {
			info: { write: true, structure: true },
			run: (context, command) => addSheet(context.doc, command.name, command.index, context.changes)
		},
		"sheet.rename": {
			info: { write: true, structure: true },
			run: (context, command) => renameSheet(context.doc, command.sheet, command.name, context.changes)
		}
	}
})

export const clipboardFeature = defineSheetFeature<XlsxDocument>({
	id: "clipboard",
	commands: {
		"clipboard.paste": {
			info: { write: true },
			run: (context, command) =>
				pasteCells(context.doc, command.sheet, command.at, command.text, command.clip, context.changes, context.slicing)
		}
	},
	queries: {
		copy: { run: (context, request) => copyCells(context.doc, request.sheet, request.range, request.cut) }
	}
})

export const XLSX_FEATURES: readonly SheetFeature<XlsxDocument>[] = [
	cellsFeature,
	formatFeature,
	axisFeature,
	sheetsFeature,
	clipboardFeature
]

export type XlsxSessionOpen =
	| ({ readonly kind: "opened" } & SheetOpened & { readonly session: SheetSession<XlsxDocument> })
	| { readonly kind: "failed"; readonly reason: SheetOpenFailure; readonly message: string }
	| { readonly kind: "aborted" }

// Opens an xlsx (or its decrypted package) as a session: the package read, the active sheet parsed, the rest in
// idle slices. `features` adds to the base list.
export async function openXlsxSession(
	source: ByteSource,
	options: SheetSessionOptions,
	onEvent: (event: SheetEvent) => void,
	features: readonly SheetFeature<XlsxDocument>[] = []
): Promise<XlsxSessionOpen> {
	const opened = await openXlsx(source, {
		...(options.slicing !== undefined && { slicing: options.slicing }),
		...(options.signal !== undefined && { signal: options.signal })
	})

	if (!opened.ok) {
		if (opened.reason === "aborted") return { kind: "aborted" }

		return {
			kind: "failed",
			reason: opened.reason === "encrypted" || opened.reason === "legacyBinary" ? "unsupported" : sheetOpenFailure(opened.reason),
			message: opened.message
		}
	}

	const styles = new RenderStyleTable()
	const document = new XlsxDocument(opened.workbook, styles, {
		kind: options.kind,
		...(options.settings !== undefined && { settings: options.settings }),
		...(options.writeUnlocked !== undefined && { writeUnlocked: options.writeUnlocked }),
		...(options.encryption !== undefined && { encryption: options.encryption })
	})
	const session = new SheetSession(
		document,
		styles,
		{ features: own => sheetRegistry([...XLSX_FEATURES, ...features, sessionFeature(own)]) },
		options,
		onEvent
	)

	try {
		const info = await session.start(options.signal)

		return { kind: "opened", session, info }
	} catch (error) {
		session.close()

		if (error instanceof XlsxReadError) {
			if (error.code === "aborted") return { kind: "aborted" }

			return { kind: "failed", reason: error.code === "tooLarge" ? "tooLarge" : "malformed", message: error.message }
		}

		throw error
	}
}
