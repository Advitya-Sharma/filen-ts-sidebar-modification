import type { Slicing } from "../opc/index"
import type { OdfPackage } from "./package"
import { readContentPrologue, type ContentPrologue } from "./prologue"
import { readUntilBody } from "./stream"
import { OdfStyleSheet } from "./styles"

export interface LoadedStyles {
	readonly sheet: OdfStyleSheet
	// undefined for a package without content.xml (an image document).
	readonly prologue: ContentPrologue | undefined
}

// The package's styles: styles.xml whole, content.xml up to office:body only.
export async function readStyleSheet(odf: OdfPackage, slicing?: Slicing): Promise<LoadedStyles> {
	const styles = await odf.xml("styles.xml", slicing)
	const entry = odf.archive.get("content.xml")
	const prologue = entry === undefined ? undefined : readContentPrologue(await readUntilBody(odf.archive, entry, slicing))

	return { sheet: new OdfStyleSheet(styles, prologue), prologue }
}
