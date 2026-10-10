// OPC part names (ECMA-376 Part 2, §6.2.2) as zip item names: no leading "/", "/"-separated segments,
// percent-encoded as written (non-ASCII may also be stored raw). Every name this package takes or returns
// is in that one form: zip entry names, resolved relationship targets, names for has/entry/read/plan.
// Lookups compare keys that decode it exactly once and fold ASCII case, as OPC equivalence requires, so a
// literal "%25" in a name stays a literal "%".

function asciiLower(text: string): string {
	return text.replace(/[A-Z]+/g, upper => upper.toLowerCase())
}

function percentDecode(text: string): string {
	if (text.indexOf("%") < 0) return text

	try {
		return decodeURIComponent(text)
	} catch {
		return text
	}
}

export function partKey(name: string): string {
	return asciiLower(percentDecode(name.startsWith("/") ? name.slice(1) : name))
}

// Structural validity of a zip item name as a part name. Anything that could address outside the package
// or alias another part is refused: empty, "." or ".." segments, a trailing "." on a segment, backslashes,
// encoded separators, and control characters.
export function isValidPartName(name: string): boolean {
	if (name === "" || name.startsWith("/") || name.endsWith("/")) return false
	if (/%(2f|5c)/i.test(name)) return false

	for (let i = 0; i < name.length; i++) {
		const code = name.charCodeAt(i)

		if (code < 0x20 || code === 0x7f || code === 0x5c) return false
	}

	for (const segment of name.split("/")) {
		if (segment === "" || segment.endsWith(".")) return false
	}

	return true
}

// Resolve a relationship target against its source part (RFC 3986 reference resolution over part names).
// Each segment's decoded form decides dot segments, so an encoded "%2e%2e" cannot slip past the root and an
// encoded "/" or backslash cannot invent a segment; the segments themselves are kept as written. Returns the
// part name (no leading "/"), or undefined when the target leaves the package.
export function resolvePartName(sourcePart: string, target: string): string | undefined {
	const withoutFragment = target.split("#")[0] ?? ""

	if (withoutFragment === "" || /^[a-z][a-z0-9+.-]*:/i.test(withoutFragment)) return undefined

	const segments: string[] = []

	if (!withoutFragment.startsWith("/")) {
		const base = sourcePart.split("/")

		base.pop()
		segments.push(...base)
	}

	for (const segment of withoutFragment.split("/")) {
		const decoded = percentDecode(segment)

		if (decoded.includes("/") || decoded.includes("\\")) return undefined
		if (decoded === "" || decoded === ".") continue

		if (decoded === "..") {
			if (segments.length === 0) return undefined

			segments.pop()
		} else {
			segments.push(segment)
		}
	}

	return segments.length === 0 ? undefined : segments.join("/")
}

// Encodes what a URI path segment cannot hold (spaces, non-ASCII, "#", "?" and the like). Existing escapes
// are left alone, as names are already in their written form; "[" and "]" stay for "[Content_Types]".
function encodeSegment(segment: string): string {
	return segment.replace(/[^A-Za-z0-9\-._~!$&'()*+,;=:@%[\]]/gu, char => encodeURIComponent(char))
}

// A part name as content types write it: "/" first, nothing a URI cannot carry.
export function partNameUri(partName: string): string {
	return `/${partName.split("/").map(encodeSegment).join("/")}`
}

// The relative target from `sourcePart` ("" for the package) to `partName`, as Word writes relationships.
export function relativeTarget(sourcePart: string, partName: string): string {
	const from = sourcePart.split("/")
	const to = partName.split("/")

	from.pop()

	let common = 0

	while (common < from.length && common < to.length - 1 && from[common] === to[common]) common++

	return [...from.slice(common).map(() => ".."), ...to.slice(common).map(encodeSegment)].join("/")
}

// "word/document.xml" -> "word/_rels/document.xml.rels"; "" (the package) -> "_rels/.rels".
export function relationshipsPartName(sourcePart: string): string {
	const slash = sourcePart.lastIndexOf("/")

	return `${sourcePart.slice(0, slash + 1)}_rels/${sourcePart.slice(slash + 1)}.rels`
}
