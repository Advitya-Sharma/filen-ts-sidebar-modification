import { encodeUtf8 } from "../xml/utf8"
import { memorySink } from "../opc/io"
import { BytePieces } from "../opc/pieces"
import { writeZip, type ZipPlanItem } from "../opc/zipWriter"

export const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

// A few small parts in the order given, deflated, with the zip writer's fixed timestamps: the same input
// always gives the same bytes.
export async function zipParts(parts: readonly (readonly [name: string, markup: string])[]): Promise<Uint8Array> {
	const items: ZipPlanItem[] = parts.map(([name, markup]) => ({
		kind: "data",
		name,
		bytes: BytePieces.of([encodeUtf8(DECLARATION + markup)]),
		compress: true
	}))
	const sink = memorySink()

	await writeZip({ items }, null, sink)

	return sink.bytes()
}
