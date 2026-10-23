import { escapeAttribute, encodeUtf8, Namespaces, parseXml } from "../xml/index"
import { OdfError } from "./errors"
import { OdfNs } from "./namespaces"
import { canonicalAttributes, elementName } from "./tree"

// How one entry is encrypted (ODF 1.3 Part 2 §4.8–§4.12). Kept so a refusal can say what was used.
export interface OdfEntryEncryption {
	readonly checksumType: string | undefined
	readonly algorithm: string | undefined
	readonly keyDerivation: string | undefined
	readonly startKeyGeneration: string | undefined
}

export interface OdfManifestEntry {
	readonly fullPath: string
	readonly mediaType: string
	readonly version: string | undefined
	readonly preferredViewMode: string | undefined
	// manifest:size: the plain size of an encrypted entry.
	readonly size: number | undefined
	readonly encryption: OdfEntryEncryption | undefined
}

export interface OdfManifest {
	readonly version: string | undefined
	readonly entries: readonly OdfManifestEntry[]
	// manifest:keyinfo (§4.15): the package is encrypted with OpenPGP keys.
	readonly keyInfo: boolean
}

export const MANIFEST_PATH = "META-INF/manifest.xml"

function intAttribute(value: string | undefined): number | undefined {
	if (value === undefined || !/^\d+$/.test(value)) return undefined

	const number = Number(value)

	return Number.isSafeInteger(number) ? number : undefined
}

// Reads META-INF/manifest.xml. Unknown elements and attributes are ignored, as §4 permits foreign markup.
export function parseManifest(bytes: Uint8Array): OdfManifest {
	const document = parseXml(bytes)
	const namespaces = new Namespaces(document)
	const root = document.documentElement()

	if (root < 0 || elementName(document, namespaces, root) !== "manifest:manifest") {
		throw new OdfError("invalidPackage", "META-INF/manifest.xml has no manifest:manifest root")
	}

	const entries: OdfManifestEntry[] = []
	let keyInfo = false

	for (let node = document.firstChildElement(root); node >= 0; node = document.nextSiblingElement(node)) {
		const name = elementName(document, namespaces, node)

		if (name === "manifest:keyinfo") {
			keyInfo = true

			continue
		}

		if (name !== "manifest:file-entry") continue

		const attributes = canonicalAttributes(document, namespaces, node)
		const fullPath = attributes.get("manifest:full-path")

		if (fullPath === undefined) continue

		let encryption: OdfEntryEncryption | undefined

		for (let child = document.firstChildElement(node); child >= 0; child = document.nextSiblingElement(child)) {
			if (elementName(document, namespaces, child) !== "manifest:encryption-data") continue

			const data = canonicalAttributes(document, namespaces, child)
			let algorithm: string | undefined
			let keyDerivation: string | undefined
			let startKeyGeneration: string | undefined

			for (let part = document.firstChildElement(child); part >= 0; part = document.nextSiblingElement(part)) {
				const partAttributes = canonicalAttributes(document, namespaces, part)

				switch (elementName(document, namespaces, part)) {
					case "manifest:algorithm":
						algorithm = partAttributes.get("manifest:algorithm-name")
						break
					case "manifest:key-derivation":
						keyDerivation = partAttributes.get("manifest:key-derivation-name")
						break
					case "manifest:start-key-generation":
						startKeyGeneration = partAttributes.get("manifest:start-key-generation-name")
						break
				}
			}

			encryption = { checksumType: data.get("manifest:checksum-type"), algorithm, keyDerivation, startKeyGeneration }
		}

		entries.push({
			fullPath,
			mediaType: attributes.get("manifest:media-type") ?? "",
			version: attributes.get("manifest:version"),
			preferredViewMode: attributes.get("manifest:preferred-view-mode"),
			size: intAttribute(attributes.get("manifest:size")),
			encryption
		})
	}

	return { version: canonicalAttributes(document, namespaces, root).get("manifest:version"), entries, keyInfo }
}

export interface ManifestEntryInit {
	readonly fullPath: string
	readonly mediaType: string
	// Written on the root entry ("/") only, as ODF 1.2+ producers do.
	readonly version?: string
	readonly preferredViewMode?: string
}

// META-INF/manifest.xml for an unencrypted package (Part 2 §4). Encrypted entries are never written.
export function writeManifest(version: string, entries: readonly ManifestEntryInit[]): Uint8Array {
	let xml = `<?xml version="1.0" encoding="UTF-8"?>\n<manifest:manifest xmlns:manifest="${OdfNs.manifest}" manifest:version="${escapeAttribute(version)}">\n`

	for (const entry of entries) {
		xml += ` <manifest:file-entry manifest:full-path="${escapeAttribute(entry.fullPath)}"`

		if (entry.version !== undefined) xml += ` manifest:version="${escapeAttribute(entry.version)}"`

		xml += ` manifest:media-type="${escapeAttribute(entry.mediaType)}"`

		if (entry.preferredViewMode !== undefined) xml += ` manifest:preferred-view-mode="${escapeAttribute(entry.preferredViewMode)}"`

		xml += "/>\n"
	}

	return encodeUtf8(`${xml}</manifest:manifest>\n`)
}
