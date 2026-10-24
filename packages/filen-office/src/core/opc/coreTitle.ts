import type { OpcPackage } from "./package"

const CORE_PROPERTIES = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
const STRICT_CORE_PROPERTIES = "http://purl.oclc.org/ooxml/officeDocument/relationships/metadata/core-properties"

// The title a package's core properties give (ECMA-376 Part 2 §11.1, dc:title), whitespace collapsed; blank or none:
// undefined. A missing or malformed core part names none.
export async function coreTitle(opc: OpcPackage): Promise<string | undefined> {
	try {
		const relationships = await opc.relationships("")
		const core = relationships.all.find(
			relationship =>
				!relationship.external &&
				(relationship.type === CORE_PROPERTIES || relationship.type === STRICT_CORE_PROPERTIES) &&
				relationship.partName !== undefined
		)

		if (core?.partName === undefined || !opc.has(core.partName)) return undefined

		const xml = await opc.xml(core.partName)
		const root = xml.documentElement()

		for (let child = xml.firstChildElement(root); child >= 0; child = xml.nextSiblingElement(child)) {
			if (xml.localName(child) === "title") {
				const title = xml.textContent(child).replace(/\s+/g, " ").trim()

				return title.length > 0 ? title : undefined
			}
		}
	} catch {
		// A missing or malformed core part names no title.
	}

	return undefined
}
