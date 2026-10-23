import { odsToXlsx } from "./toXlsx"
import { xlsxToOds } from "./toOds"
import {
	aborted,
	Aborted,
	type SheetConversion,
	type SheetConversionFormat,
	type SheetConvertOptions,
	type SheetSourceFormat
} from "./types"

const FAMILIES: Readonly<Record<string, SheetSourceFormat>> = {
	xlsx: "xlsx",
	xlsm: "xlsx",
	xltx: "xlsx",
	xltm: "xlsx",
	ods: "ods",
	ots: "ods"
}

// Each family converts to the other.
export function canConvertSheet(from: SheetSourceFormat, to: SheetConversionFormat): boolean {
	return from !== to
}

// The family of a file name's extension, undefined for any other.
export function sheetSourceFormat(name: string): SheetSourceFormat | undefined {
	const dot = name.lastIndexOf(".")

	return dot < 0 ? undefined : FAMILIES[name.slice(dot + 1).toLowerCase()]
}

// A workbook of one family as one of the other. Never rejects: every failure is a typed refusal.
export async function convertSheet(
	source: Uint8Array,
	from: SheetSourceFormat,
	to: SheetConversionFormat,
	options: SheetConvertOptions = {}
): Promise<SheetConversion> {
	if (!canConvertSheet(from, to)) return { ok: false, reason: "unsupported", message: `${from} is not converted to ${to}` }
	if (aborted(options)) return { ok: false, reason: "aborted", message: "The conversion was aborted" }

	try {
		return from === "xlsx" ? await xlsxToOds(source, options) : await odsToXlsx(source, options)
	} catch (error) {
		if (error instanceof Aborted || aborted(options)) return { ok: false, reason: "aborted", message: "The conversion was aborted" }

		return { ok: false, reason: "unsupported", message: error instanceof Error ? `${error.name}: ${error.message}` : String(error) }
	}
}
