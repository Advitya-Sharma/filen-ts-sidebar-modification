import type { Namespaces } from "./namespaces"

export const MC_NAMESPACE = "http://schemas.openxmlformats.org/markup-compatibility/2006"

// The branch of an mc:AlternateContent element to render (ECMA-376 Part 3, §10.2): the first mc:Choice
// whose Requires prefixes all resolve to namespaces the caller understands, else mc:Fallback, else -1.
// The element itself is never rewritten: it stays an opaque island that is saved byte for byte.
export function selectAlternateContent(namespaces: Namespaces, node: number, understands: (uri: string) => boolean): number {
	const document = namespaces.document

	if (!namespaces.is(node, MC_NAMESPACE, "AlternateContent")) return -1

	let fallback = -1

	for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) {
		if (namespaces.is(child, MC_NAMESPACE, "Choice")) {
			const requires = (document.attribute(child, "Requires") ?? "").split(/\s+/).filter(prefix => prefix !== "")

			if (
				requires.length > 0 &&
				requires.every(prefix => {
					const uri = namespaces.uri(child, prefix)

					return uri !== undefined && understands(uri)
				})
			) {
				return child
			}
		} else if (fallback < 0 && namespaces.is(child, MC_NAMESPACE, "Fallback")) {
			fallback = child
		}
	}

	return fallback
}
