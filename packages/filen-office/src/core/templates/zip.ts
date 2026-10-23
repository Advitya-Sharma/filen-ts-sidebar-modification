import { encodeUtf8 } from "../xml/utf8"
import { memorySink } from "../opc/io"
import { BytePieces } from "../opc/pieces"
import { writeZip, type ZipPlanItem } from "../opc/zipWriter"

export const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

// A part: XML markup (written after the declaration, deflated), or with "stored" text written as it is, uncompressed
// (an ODF package's "mimetype").
export type TemplatePart = readonly [name: string, markup: string] | readonly [name: string, text: string, "stored"]

// A few small parts in the order given, with the zip writer's fixed timestamps: the same input always gives the
// same bytes.
export async function zipParts(parts: readonly TemplatePart[]): Promise<Uint8Array> {
	const items: ZipPlanItem[] = parts.map(([name, markup, stored]) => ({
		kind: "data",
		name,
		bytes: BytePieces.of([encodeUtf8(stored === undefined ? DECLARATION + markup : markup)]),
		compress: stored === undefined
	}))
	const sink = memorySink()

	await writeZip({ items }, null, sink)

	return sink.bytes()
}
