import { openDocxParts, type DocxOpenResult } from "../docx/package"
import { OpcPackage, type ByteSource, type Slicing } from "../opc/index"
import type { OdfEncryption } from "../odf/index"
import { translateOdt, type OdtTranslateOptions } from "./translate"
import type { OdtOpenFailure, OdtTranslation } from "./types"

export type OdtDocxOpenResult =
	| (DocxOpenResult & { readonly ok: true; readonly translation: OdtTranslation })
	| {
			readonly ok: false
			readonly reason: OdtOpenFailure | "invalidPackage"
			readonly message: string
			readonly encryption?: OdfEncryption
	  }

// An ODF text document opened as the word-processing package it translates to, for the docx stack.
export async function openOdtAsDocx(
	source: ByteSource,
	options: OdtTranslateOptions & { readonly slicing?: Slicing } = {}
): Promise<OdtDocxOpenResult> {
	const translated = await translateOdt(source, options)

	if (!translated.ok) return translated

	const opened = await openDocxParts(OpcPackage.fromParts(translated.translation.parts), options.slicing)

	if (!opened.ok) return { ok: false, reason: "invalidPackage", message: opened.message }

	return { ...opened, translation: translated.translation }
}
