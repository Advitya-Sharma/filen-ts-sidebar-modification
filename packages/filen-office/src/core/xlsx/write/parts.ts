import { RELATIONSHIPS_NAMESPACE, type Relationship, type Relationships } from "../../opc/index"
import { XmlEditor, encodeUtf8, parseXml } from "../../xml/index"
import { attr, qualify, rootTag, utf8Part } from "./markup"

// Relationship parts: source relationships removed by id and new ones appended with the ids the change log
// allocated; everything else as written.

const EMPTY_RELATIONSHIPS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}"></Relationships>`

function relationshipMarkup(relationship: Relationship, prefix: string): string {
	return `<${qualify(prefix, "Relationship")}${attr("Id", relationship.id)}${attr("Type", relationship.type)}${attr("Target", relationship.target)}${relationship.external ? attr("TargetMode", "External") : ""}/>`
}

export function editRelationships(source: Uint8Array | null, added: readonly Relationship[], removed: ReadonlySet<string>): Uint8Array {
	// A UTF-16 part is edited (and written) as UTF-8.
	const bytes = source === null ? encodeUtf8(EMPTY_RELATIONSHIPS) : utf8Part(source)
	const document = parseXml(bytes)
	const root = document.documentElement()
	const editor = new XmlEditor(document)
	const { prefix } = rootTag(bytes)

	for (let node = document.firstChildElement(root); node >= 0; node = document.nextSiblingElement(node)) {
		const id = document.attribute(node, "Id")

		if (id !== undefined && removed.has(id)) editor.remove(node)
	}

	if (added.length > 0) editor.append(root, added.map(relationship => relationshipMarkup(relationship, prefix)).join(""))

	return editor.changed || source === null ? editor.serialize() : source
}

// The next "rId<n>" free of `taken`.
export function freeRelationshipId(taken: Set<string>): string {
	let highest = 0

	for (const id of taken) {
		const numbered = /^rId(\d+)$/.exec(id)

		if (numbered !== null) highest = Math.max(highest, Number(numbered[1]))
	}

	const id = `rId${highest + 1}`

	taken.add(id)

	return id
}

export function relationshipIds(relationships: Relationships): Set<string> {
	return new Set(relationships.all.map(relationship => relationship.id))
}
