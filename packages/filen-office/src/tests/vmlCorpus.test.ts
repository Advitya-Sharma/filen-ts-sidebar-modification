import { existsSync, readFileSync, readdirSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { ItemKind } from "../core/drawingml/displayList"
import type { ImageResolver } from "../core/drawingml/picture"
import { OpcPackage, bytesSource } from "../core/opc/index"
import { convertVml } from "../core/vml/index"
import { Namespaces, NodeKind, type XmlDocument } from "../core/xml/index"
import { CORPUS_DIR as CORPUS } from "./corpusDir"

// The licensed real-world corpus is kept out of the repository; this sweep runs only where it exists.
const EXTENSIONS = ["docx", "docm", "dotx", "dotm"]
const W_URIS = new Set(["http://schemas.openxmlformats.org/wordprocessingml/2006/main", "http://purl.oclc.org/ooxml/wordprocessingml/main"])

function corpusFiles(): string[] {
	const files: string[] = []

	for (const extension of EXTENSIONS) {
		const directory = join(CORPUS, extension)

		if (!existsSync(directory)) continue

		for (const name of readdirSync(directory).sort()) files.push(join(directory, name))
	}

	return files
}

describe.skipIf(!existsSync(CORPUS))("VML over the real-world corpus", () => {
	it("converts every w:pict and w:object without throwing", { timeout: 300_000 }, async () => {
		let found = 0
		let converted = 0
		let empty = 0
		let images = 0
		let missingImages = 0
		let textboxes = 0
		let textPaths = 0
		const items = [0, 0, 0]
		const failures: string[] = []
		const documents = new Set<string>()
		const decoder = new TextDecoder()

		for (const file of corpusFiles()) {
			const name = file.slice(CORPUS.length + 1)
			let opc: OpcPackage

			try {
				opc = await OpcPackage.open(bytesSource(new Uint8Array(readFileSync(file))))
			} catch {
				continue
			}

			const contentTypes = await opc.contentTypes().catch(() => undefined)

			for (const entry of opc.zip.entries) {
				if (!entry.name.startsWith("word/") || !entry.name.endsWith(".xml")) continue

				const text = decoder.decode(await opc.read(entry.name))

				if (!text.includes(":pict") && !text.includes(":object")) continue

				let document: XmlDocument

				try {
					document = await opc.xml(entry.name)
				} catch {
					continue
				}

				const relationships = await opc.relationships(entry.name)
				const resolver: ImageResolver = id => {
					const relationship = relationships.byId(id)

					if (
						relationship === undefined ||
						relationship.external ||
						relationship.partName === undefined ||
						!opc.has(relationship.partName)
					)
						return undefined

					return { partName: relationship.partName, contentType: contentTypes?.typeOf(relationship.partName) }
				}
				const namespaces = new Namespaces(document, { lenient: true })

				for (let node = 0; node < document.count; node++) {
					if (document.kinds[node] !== NodeKind.Element) continue

					const local = document.localName(node)

					if ((local !== "pict" && local !== "object") || !W_URIS.has(namespaces.elementUri(node) ?? "")) continue

					found++
					documents.add(name)

					try {
						const result = convertVml({ images: resolver }, document, node)

						if (result === undefined) {
							empty++

							continue
						}

						converted++
						textboxes += result.textboxes.length
						textPaths += result.textPaths.length

						for (let item = 0; item < result.list.itemCount; item++) {
							const kind = result.list.kind(item)

							items[kind] = (items[kind] ?? 0) + 1
						}

						for (const image of result.list.images) {
							images++

							if (image.partName === undefined) missingImages++
						}

						const bounds = result.list.bounds()

						if (!Array.from(bounds).every(Number.isFinite)) failures.push(`${name} ${entry.name}#${node}: non-finite bounds`)
					} catch (error) {
						failures.push(`${name} ${entry.name}#${node}: ${String(error)}`)
					}
				}
			}
		}

		console.log(
			`VML corpus: ${found} w:pict/w:object in ${documents.size} documents; ${converted} converted, ${empty} without a drawable shape, ${failures.length} failed; items: ${items[ItemKind.PATH]} paths, ${items[ItemKind.IMAGE]} images, ${items[ItemKind.TEXT]} texts; ${images} image refs (${missingImages} unresolved), ${textboxes} text boxes, ${textPaths} text paths`
		)

		expect(failures).toEqual([])
		expect(converted + empty).toBe(found)
	})
})
