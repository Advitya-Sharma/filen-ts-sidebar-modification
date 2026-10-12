import { joinBytes } from "../../../opc/index"
import { NodeKind } from "../../../xml/index"
import { partReader } from "../../model/reader"
import { Ns } from "../../model/namespaces"
import { sniffImage } from "../../objects/images"
import { DocxVerifyError } from "../failure"
import type { CheckInput, DocxCheck } from "./index"

// Drawings the save added (new drawing ids; gate 2 counts docPr ids for duplicates itself): extents and ids
// in range. Every picture relationship the save added points at a part typed as a picture, and every picture
// part it added holds the picture its content type says. Drawings copied from the source are left as they
// were, whatever their values.

// ST_PositiveCoordinate, ST_DrawingElementId.
const MAX_COORDINATE = 27273042316900
const MAX_DRAWING_ID = 4294967295
const INTEGER = /^\d+$/

function inRange(value: string | undefined, max: number): boolean {
	return value !== undefined && INTEGER.test(value) && Number(value) <= max
}

// Drawing ids (as written) whose drawing has values out of range, per save, for done() to match with the ids
// the save added.
const outOfRange = new WeakMap<CheckInput, Set<string>>()

export const drawingChecks: readonly DocxCheck[] = [
	{
		async part(input, checked) {
			// Copies of source nodes are what they were; only the save's own markup is looked at.
			if (checked.part === undefined || (!checked.skeleton && input.original.has(checked.name))) return

			const document = checked.document
			const reader = partReader(document)
			const added = new Set<string>()
			const source = input.original.has(checked.name) ? await input.original.relationships(checked.name) : undefined
			const written = await input.relationshipsOf(checked.name)

			for (let node = 0; node < document.count; node++) {
				if (document.kinds[node] !== NodeKind.Element) continue

				const group = reader.group(node)
				const local = document.localName(node)

				if (group === Ns.WP && (local === "inline" || local === "anchor")) {
					const extent = reader.child(node, Ns.WP, "extent")
					const docPr = reader.child(node, Ns.WP, "docPr")
					const id = docPr < 0 ? undefined : reader.attr(docPr, Ns.None, "id")
					const valid =
						extent >= 0 &&
						inRange(reader.attr(extent, Ns.None, "cx"), MAX_COORDINATE) &&
						inRange(reader.attr(extent, Ns.None, "cy"), MAX_COORDINATE) &&
						inRange(id, MAX_DRAWING_ID)

					if (!valid) {
						let found = outOfRange.get(input)

						if (found === undefined) {
							found = new Set()
							outOfRange.set(input, found)
						}

						found.add(id ?? "")
					}
				} else if ((group === Ns.A && local === "blip") || local === "svgBlip") {
					const id = reader.attr(node, Ns.R, "embed")

					if (id !== undefined && source?.byId(id) === undefined) added.add(id)
				}
			}

			for (const id of added) {
				const relationship = written.byId(id)
				const type = relationship?.partName === undefined ? undefined : input.types.typeOf(relationship.partName)

				if (relationship === undefined || relationship.external || type?.toLowerCase().startsWith("image/") !== true)
					throw new DocxVerifyError("relationship", `${checked.name}: picture ${id} does not point at a picture part`)
			}
		},
		async done(input) {
			const problems = outOfRange.get(input)

			if (problems !== undefined) {
				for (const id of input.added("docPr")) {
					if (problems.has(id)) throw new DocxVerifyError("content", `A new drawing (${id}) has an extent or id out of range`)
				}
			}

			for (const change of input.changes.add) {
				const type = input.types.typeOf(change.name)?.toLowerCase()

				if (type?.startsWith("image/") !== true) continue

				const image = sniffImage(await joinBytes(change.bytes))

				if (typeof image === "string" || image.contentType !== type)
					throw new DocxVerifyError("contentType", `${change.name} does not hold the ${type} picture it is typed as`)
			}
		}
	}
]
