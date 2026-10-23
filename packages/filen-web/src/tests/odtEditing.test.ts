import { describe, expect, it } from "vitest"
import { ZipArchive, bytesSource, memorySink, writeZip } from "@filen/office/core/opc"
import { buildDocxModel, openDocxPackage } from "@filen/office/core/docx"
import { DocxEditorSession, openDocxEditContext } from "@filen/office/core/docx/edit"
import { DocxSaveError } from "@filen/office/core/docx/write"
import { inThreadDocxRtf } from "@filen/office/dom/docx/rtfWorker"
import { rtfEditingBackend } from "@/features/preview/lib/rtfEditing"

// An OpenDocument text document saved through the docx editor's save path: Gate 1 from both writers, the docx writer's
// parts written back as ODT (prepare writes, the zipper hands out), refusals as the docx writer's own.

const NS =
	'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"'
const SECOND = '<text:p text:style-name="P1">Second <text:span text:style-name="T1">bold</text:span> paragraph</text:p>'

async function odt(mime: string): Promise<Uint8Array> {
	const text = (value: string): Uint8Array => new TextEncoder().encode(value)
	const sink = memorySink()

	await writeZip(
		{
			items: [
				{ kind: "data", name: "mimetype", bytes: text(mime), compress: false },
				{
					kind: "data",
					name: "content.xml",
					bytes: text(
						`<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS} office:version="1.3"><office:automatic-styles><style:style style:name="P1" style:family="paragraph"><style:paragraph-properties fo:text-align="center"/></style:style><style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style></office:automatic-styles><office:body><office:text><text:p>First paragraph</text:p>${SECOND}<text:p>Last <text:toc-mark text:string-value="Last" text:outline-level="1"/>paragraph</text:p></office:text></office:body></office:document-content>`
					),
					compress: true
				},
				{
					kind: "data",
					name: "styles.xml",
					bytes: text(
						`<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NS} office:version="1.3"><office:styles><style:style style:name="Standard" style:family="paragraph"/></office:styles></office:document-styles>`
					),
					compress: true
				},
				{
					kind: "data",
					name: "META-INF/manifest.xml",
					bytes: text(
						`<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="${mime}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/></manifest:manifest>`
					),
					compress: true
				}
			]
		},
		null,
		sink
	)

	return sink.bytes()
}

// The entries of a package by name, in the order they are stored.
async function unzip(bytes: Uint8Array): Promise<{ names: string[]; text: (name: string) => Promise<string> }> {
	const archive = await ZipArchive.open(bytesSource(bytes))
	const names = [...archive.entries].sort((a, b) => a.localHeaderOffset - b.localHeaderOffset).map(entry => entry.name)

	return {
		names,
		text: async name => {
			const entry = archive.get(name)

			return entry === undefined ? "" : new TextDecoder().decode(await archive.read(entry))
		}
	}
}

const ODT = await odt("application/vnd.oasis.opendocument.text")

async function editing(bytes: Uint8Array) {
	// Engine sessions are not opened here: no services.
	const opened = await inThreadDocxRtf(() => {
		throw new Error("no engine in this test")
	}, "odt").open(bytes)

	if (opened.kind !== "opened") throw new Error(opened.kind)

	const document = opened.document
	const backend = rtfEditingBackend(document, document.prove(), "ODT")
	const pkg = await openDocxPackage(bytesSource(await document.take()))

	if (!pkg.ok) throw new Error(pkg.reason)

	const context = await openDocxEditContext(pkg.package, await buildDocxModel(pkg.package))
	const made = backend.proof(context)
	const session = new DocxEditorSession(context, { proof: "later" in made ? (await made.later).full : made })
	const save = async (): Promise<Uint8Array> => {
		const prepared = await backend.writer.prepare({ pkg: pkg.package, context, sessions: [session] }, {})
		const zipper = backend.writer.zipper?.(pkg.package.opc)

		if (zipper === undefined) throw new Error("no zipper")

		return (await zipper.zip(prepared)).bytes
	}

	return { session, save, document, context }
}

function types(session: DocxEditorSession): string[] {
	const out: string[] = []

	session.state.doc.forEach(node => out.push(node.type.name))

	return out
}

describe("ODT editing backend", () => {
	it("saves an untouched document as its own bytes", async () => {
		const { save } = await editing(ODT)

		expect(await save()).toEqual(ODT)
	})

	it("saves an edit as ODT, mimetype first, the other paragraphs' markup kept", async () => {
		const { session, save } = await editing(ODT)

		session.dispatch(session.state.tr.insertText("Edited ", 1))

		const entries = await unzip(await save())
		const content = await entries.text("content.xml")

		expect(entries.names[0]).toBe("mimetype")
		expect(await entries.text("mimetype")).toBe("application/vnd.oasis.opendocument.text")
		expect(content).toContain("<text:p>Edited First paragraph</text:p>")
		expect(content).toContain(SECOND)
	})

	it("keeps a template a template", async () => {
		const template = await odt("application/vnd.oasis.opendocument.text-template")
		const { session, save } = await editing(template)

		session.dispatch(session.state.tr.insertText("Edited ", 1))

		expect(await (await unzip(await save())).text("mimetype")).toBe("application/vnd.oasis.opendocument.text-template")
	})

	it("locks the paragraph holding an index mark the writer cannot write, the others editable", async () => {
		const { session } = await editing(ODT)

		expect(types(session)).toEqual(["paragraph", "paragraph", "locked_block"])
	})

	it("refuses what it cannot write (a column break inside a paragraph) as the docx writer's own refusal", async () => {
		const { session, save } = await editing(ODT)
		const column = session.state.schema.nodes["break"]?.create({ type: "column" })

		if (column === undefined) throw new Error("no break node")

		// ODF breaks a column before a paragraph only (fo:break-before), never inside one.
		session.dispatch(session.state.tr.insert(6, column))

		const refusal = await save().then(
			() => undefined,
			(error: unknown) => error
		)

		expect(refusal).toBeInstanceOf(DocxSaveError)
		expect((refusal as DocxSaveError).reason).toBe("write")
		expect((refusal as DocxSaveError).message).toMatch(/^ODT: /u)
	})
})
