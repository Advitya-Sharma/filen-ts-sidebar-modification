import { ZipError, type SliceClock } from "../opc/index"
import { XmlError, type XmlDocument } from "../xml/index"
import type { OdfPackage } from "./package"

// The document's own title (meta.xml: office:document-meta > office:meta > dc:title, ODF 1.3 Part 3 §4.3.2.2),
// whitespace collapsed; blank or none: undefined. A damaged meta.xml names none.
export async function metaTitle(odf: OdfPackage, clock?: SliceClock): Promise<string | undefined> {
	let meta: XmlDocument | undefined

	try {
		meta = await odf.xml("meta.xml", clock)
	} catch (error) {
		if (!(error instanceof XmlError || error instanceof ZipError)) throw error
	}

	if (meta === undefined) return undefined

	for (let block = meta.firstChildElement(meta.documentElement()); block >= 0; block = meta.nextSiblingElement(block)) {
		if (meta.localName(block) !== "meta") continue

		for (let child = meta.firstChildElement(block); child >= 0; child = meta.nextSiblingElement(child)) {
			if (meta.localName(child) !== "title") continue

			const title = meta.textContent(child).replace(/\s+/g, " ").trim()

			return title.length > 0 ? title : undefined
		}
	}

	return undefined
}
