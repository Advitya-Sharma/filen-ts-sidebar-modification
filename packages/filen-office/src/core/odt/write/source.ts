import { ZipArchive, bytesSource, memorySink, writeZip, type ZipPlanItem } from "../../opc/index"
import { jobOf, type RtfTurns } from "../../rtf/turns"
import type { Slicing } from "../../slicing"
import { translateOdt } from "../translate"
import type { OdtOpenFailure, OdtTranslation } from "../types"

// An ODF text document as the writer holds it: the file, its archive and its translation, plus the translation as
// the WordprocessingML package the editor and the engine sessions open.

export interface OdtSource {
	readonly bytes: Uint8Array
	readonly archive: ZipArchive
	readonly translation: OdtTranslation
	// The translation zipped (pictures stored, as they come): what an editor takes a copy of.
	readonly docx: Uint8Array
}

export type OdtSourceOpen =
	| { readonly ok: true; readonly source: OdtSource }
	| { readonly ok: false; readonly reason: OdtOpenFailure | "invalidPackage"; readonly message: string }

const STORED_EXTENSIONS = /\.(png|jpe?g|gif|webp|emf|wmf|tiff?|bmp|svg)$/iu

export async function translationPackage(translation: OdtTranslation, slicing?: Slicing): Promise<Uint8Array> {
	const items: ZipPlanItem[] = []

	for (const [name, part] of translation.parts) {
		const bytes = part instanceof Uint8Array ? part : await part.read()

		items.push({ kind: "data", name, bytes, compress: !STORED_EXTENSIONS.test(name) })
	}

	const sink = memorySink()

	await writeZip({ items }, null, sink, slicing === undefined ? { level: 1 } : { level: 1, slicing })

	return sink.bytes()
}

export async function openOdtSource(bytes: Uint8Array, turns?: RtfTurns): Promise<OdtSourceOpen> {
	const job = jobOf(turns)
	const translated = await translateOdt(bytesSource(bytes), job === undefined ? {} : { slicing: job.aborting })

	if (!translated.ok) return translated

	const archive = await ZipArchive.open(bytesSource(bytes))

	return {
		ok: true,
		source: {
			bytes,
			archive,
			translation: translated.translation,
			docx: await translationPackage(translated.translation, job?.aborting)
		}
	}
}
