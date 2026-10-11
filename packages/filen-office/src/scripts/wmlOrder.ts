import { createHash } from "node:crypto"
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs"
import { dirname } from "node:path"
import { fileURLToPath } from "node:url"
import { unzipSync } from "fflate"
import { format, resolveConfig } from "prettier"
import { NodeKind, parseXml, type XmlDocument } from "../core/xml/index"

// Generates src/core/docx/write/order.generated.ts: the child order of the WordprocessingML property
// containers the writer splices into and the enumerations of the values it writes, from the Transitional schema of ECMA-376 Part 4 (5th edition).
// Every archive layer is checked against its pinned SHA-256.
//
//   pnpm run generate:wml-order [<ECMA-376-4 zip | OfficeOpenXML-XMLSchema-Transitional.zip | wml.xsd>]
//   pnpm run generate:wml-order --fetch
//
// Without an argument the Part 4 archive is read from SCHEMA_CACHE, or downloaded from ecma-international.org.
// --fetch only fills SCHEMA_CACHE (for the schema validation tests).

export const OUTPUT = fileURLToPath(new URL("../core/docx/write/order.generated.ts", import.meta.url))

export const SOURCE = {
	url: "https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip",
	sha256: "bd25da1109f73762356596918bf5ff8b74a1331642dba5f1c1d1dfc6bed34ecd",
	schemas: "OfficeOpenXML-XMLSchema-Transitional.zip",
	schemasSha256: "d34187520749998af306faf1b730e568b0ca6d88ad24638a407c0a9bb4ca04fc",
	wml: "wml.xsd",
	wmlSha256: "0f7784b36ddc42950e5ff354cd608b83b60e4de39325d5568913f1225365029b"
} as const

// The Part 4 archive, verified, where the schema validation tests look for it.
export const SCHEMA_CACHE = fileURLToPath(
	new URL("../../node_modules/.cache/filen-office-wml/ECMA-376-4_5th_edition_december_2016.zip", import.meta.url)
)

// Output name -> the complex type (or group) whose element children it lists.
export const CONTAINERS: Readonly<Record<string, string>> = {
	RUN_PROPERTIES: "CT_RPr",
	PARAGRAPH_MARK_PROPERTIES: "CT_ParaRPr",
	PARAGRAPH_PROPERTIES: "CT_PPr",
	NUMBERING_PROPERTIES: "CT_NumPr",
	NUMBERING: "CT_Numbering",
	STYLES: "CT_Styles"
}

// Output name -> the simple type whose enumeration it lists.
export const ENUMERATIONS: Readonly<Record<string, string>> = {
	JUSTIFICATION: "ST_Jc",
	UNDERLINE: "ST_Underline",
	HIGHLIGHT: "ST_HighlightColor",
	SHADING: "ST_Shd",
	THEME_COLOR: "ST_ThemeColor"
}

function sha256(bytes: Uint8Array): string {
	return createHash("sha256").update(bytes).digest("hex")
}

function extract(archive: Uint8Array, name: string, expected: string): Uint8Array {
	const entry = unzipSync(archive, { filter: file => file.name === name })[name]

	if (entry === undefined) throw new Error(`${name} is missing from the archive`)
	if (sha256(entry) !== expected) throw new Error(`${name} does not match its pinned SHA-256`)

	return entry
}

// wml.xsd from whichever layer `bytes` is.
export function wmlSchema(bytes: Uint8Array): Uint8Array {
	const digest = sha256(bytes)

	if (digest === SOURCE.wmlSha256) return bytes
	if (digest === SOURCE.schemasSha256) return extract(bytes, SOURCE.wml, SOURCE.wmlSha256)
	if (digest === SOURCE.sha256) return extract(extract(bytes, SOURCE.schemas, SOURCE.schemasSha256), SOURCE.wml, SOURCE.wmlSha256)

	throw new Error("Input matches none of the pinned ECMA-376 files")
}

// Every schema of the Transitional set (wml.xsd and what it imports), from the Part 4 archive or the
// schemas archive.
export function transitionalSchemas(bytes: Uint8Array): Map<string, Uint8Array> {
	const digest = sha256(bytes)
	const schemas =
		digest === SOURCE.schemasSha256
			? bytes
			: digest === SOURCE.sha256
				? extract(bytes, SOURCE.schemas, SOURCE.schemasSha256)
				: undefined

	if (schemas === undefined) throw new Error("Input is neither pinned ECMA-376 archive")

	const files = new Map(Object.entries(unzipSync(schemas, { filter: file => file.name.endsWith(".xsd") })))
	const wml = files.get(SOURCE.wml)

	if (wml === undefined || sha256(wml) !== SOURCE.wmlSha256) throw new Error(`${SOURCE.wml} does not match its pinned SHA-256`)

	return files
}

async function download(): Promise<Uint8Array> {
	const response = await fetch(SOURCE.url)

	if (!response.ok) throw new Error(`${SOURCE.url}: HTTP ${response.status}`)

	const bytes = new Uint8Array(await response.arrayBuffer())

	if (sha256(bytes) !== SOURCE.sha256) throw new Error(`${SOURCE.url} does not match its pinned SHA-256`)

	return bytes
}

// The Part 4 archive from SCHEMA_CACHE, downloaded there first when it is missing or does not match its pin.
export async function cachedArchive(): Promise<Uint8Array> {
	if (existsSync(SCHEMA_CACHE)) {
		const bytes = new Uint8Array(readFileSync(SCHEMA_CACHE))

		if (sha256(bytes) === SOURCE.sha256) return bytes
	}

	const bytes = await download()

	mkdirSync(dirname(SCHEMA_CACHE), { recursive: true })
	writeFileSync(SCHEMA_CACHE, bytes)

	return bytes
}

function local(name: string): string {
	return name.slice(name.indexOf(":") + 1)
}

// Element names of a complex type or group in document order: sequences and choices are flattened,
// group references and extension bases expanded (a choice's alternatives keep their listed order).
export function elementOrder(document: XmlDocument, typeName: string): string[] {
	const definitions = new Map<string, number>()

	for (let node = 0; node < document.count; node++) {
		if (document.kinds[node] !== NodeKind.Element) continue

		const kind = local(document.name(node))
		const name = document.attribute(node, "name")

		if ((kind === "complexType" || kind === "group") && name !== undefined && document.parent(node) === document.documentElement()) {
			definitions.set(`${kind}:${name}`, node)
		}
	}

	const out: string[] = []
	const visit = (node: number): void => {
		for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) {
			const kind = local(document.name(child))

			if (kind === "element") {
				const name = document.attribute(child, "name") ?? local(document.attribute(child, "ref") ?? "")

				if (!out.includes(name)) out.push(name)
			} else if (kind === "group" && document.attribute(child, "ref") !== undefined) {
				const target = definitions.get(`group:${local(document.attribute(child, "ref") ?? "")}`)

				if (target === undefined) throw new Error(`Unknown group ${document.attribute(child, "ref")}`)

				visit(target)
			} else if (kind === "extension") {
				const base = definitions.get(`complexType:${local(document.attribute(child, "base") ?? "")}`)

				if (base !== undefined) visit(base)

				visit(child)
			} else if (kind !== "attribute" && kind !== "attributeGroup" && kind !== "annotation") {
				visit(child)
			}
		}
	}
	const root = definitions.get(`complexType:${typeName}`) ?? definitions.get(`group:${typeName}`)

	if (root === undefined) throw new Error(`No type ${typeName}`)

	visit(root)

	return out
}

// The enumerated values of a top-level simple type, in schema order.
export function enumeration(document: XmlDocument, typeName: string): string[] {
	const root = document.documentElement()

	for (let node = document.firstChildElement(root); node >= 0; node = document.nextSiblingElement(node)) {
		if (local(document.name(node)) !== "simpleType" || document.attribute(node, "name") !== typeName) continue

		const out: string[] = []

		for (let restriction = document.firstChildElement(node); restriction >= 0; restriction = document.nextSiblingElement(restriction)) {
			if (local(document.name(restriction)) !== "restriction") continue

			for (let child = document.firstChildElement(restriction); child >= 0; child = document.nextSiblingElement(child)) {
				const value = document.attribute(child, "value")

				if (local(document.name(child)) === "enumeration" && value !== undefined) out.push(value)
			}
		}

		if (out.length === 0) throw new Error(`${typeName} is not an enumeration`)

		return out
	}

	throw new Error(`No simple type ${typeName}`)
}

export async function renderOrder(schema: Uint8Array): Promise<string> {
	const document = parseXml(schema)
	const source = [
		`// Generated by \`pnpm run generate:wml-order\` from ECMA-376 Part 4 (5th edition) ${SOURCE.wml}`,
		`// (SHA-256 ${SOURCE.wmlSha256}), © Ecma International, see THIRD_PARTY_NOTICES.md.`,
		"// Do not edit. Child element local names in schema order, then enumerated values.",
		"",
		...Object.entries(CONTAINERS).map(
			([name, type]) => `// ${type}\nexport const ${name}: readonly string[] = ${JSON.stringify(elementOrder(document, type))}\n`
		),
		...Object.entries(ENUMERATIONS).map(
			([name, type]) =>
				`// ${type}\nexport const ${name}: ReadonlySet<string> = new Set(${JSON.stringify(enumeration(document, type))})\n`
		)
	].join("\n")

	return format(source, { ...(await resolveConfig(OUTPUT)), filepath: OUTPUT })
}

async function main(): Promise<void> {
	const argument = process.argv[2]

	if (argument === "--fetch") {
		await cachedArchive()

		return
	}

	const bytes = argument !== undefined ? new Uint8Array(readFileSync(argument)) : await cachedArchive()

	writeFileSync(OUTPUT, await renderOrder(wmlSchema(bytes)))
}

if (process.argv[1] === fileURLToPath(import.meta.url)) await main()
