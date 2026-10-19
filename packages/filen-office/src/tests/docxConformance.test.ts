import { existsSync, readdirSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { AllSelection, TextSelection } from "prosemirror-state"
import { beforeAll, describe, expect, it } from "vitest"
import { DocxEditorSession } from "@filen/office/core/docx/edit"
import { resolveChanges, storyReviewIndex } from "../core/docx/review/accept"
import { saveDocx } from "@filen/office/core/docx/write"
import { bytesSource, OpcPackage } from "@filen/office/core/opc"
import { encodeUtf8 } from "@filen/office/core/xml"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import { CORPUS_DIR } from "./corpusDir"
import { openBytes, openSpec, STYLES_PART } from "./docxEditFixtures"
import { buildDocx, NAMESPACES, p, type SyntheticDocx } from "./docxSynthetic"
import { runEditScript } from "./docxWriteFixtures"
import { CATALOGUE, checkConformance, newViolations, type ConformanceReport } from "./oracles/conformance"
import { changedParts, editSuite, fuzzManifest, fuzzSaves } from "./oracles/suite"
import { XsdOracle, xsdOracleMissing } from "./oracles/xsd"
import { pagedEditor, settleAll } from "./pagedFixtures"

// The conformance catalogue (oracles/conformance.ts): for every item, a crafted
// package shows the check finds the violation, the writer's saves of the scenarios that make the construct add
// none, and (with the corpus) neither do the edit suite's saves or the fuzz saves (FILEN_OFFICE_FUZZ_SAVES).

const FIXTURES = join(import.meta.dirname, "fixtures")

// ---- crafted violations ----------------------------------------------------------------------------------

const DRAWING = (id: string): string =>
	`<w:r><w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/><wp:docPr id="${id}" name="Picture ${id}"/></wp:inline></w:drawing></w:r>`
const TABLE = (rows: string, grid = '<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>'): string => `<w:tbl><w:tblPr/>${grid}${rows}</w:tbl>`
const PROPERTIES = "application/vnd.openxmlformats-package.core-properties+xml"

const CRAFTED: Readonly<Record<string, { readonly spec: SyntheticDocx; readonly finds: readonly string[] }>> = {
	C01: {
		spec: { body: '<w:p><w:ins w:id="1" w:author="A" w:date="2026-01-01T10:00:00.5+02:00"><w:r><w:t>x</w:t></w:r></w:ins></w:p>' },
		finds: ["C01 annotation date not in Word's form"]
	},
	C02: {
		spec: { body: '<w:p><w:r><w:rPr><w:b w:val="on"/></w:rPr><w:t>x</w:t></w:r></w:p>' },
		finds: ["C02 ST_OnOff written as on/off"]
	},
	C03: { spec: { body: '<w:p><w:pPr><w:jc w:val="start"/></w:pPr></w:p>' }, finds: ["C03 jc start/end"] },
	C04: {
		spec: {
			body: TABLE(
				'<w:tr><w:tc><w:tcPr><w:tcW w:w="50%" w:type="pct"/></w:tcPr><w:p/></w:tc></w:tr>',
				'<w:tblGrid><w:gridCol w:w="1in"/></w:tblGrid>'
			)
		},
		finds: ["C04 pct width as a percent string", "C04 universal measure"]
	},
	C05: {
		spec: {
			body: TABLE("<w:tr><w:tc><w:p/></w:tc></w:tr>").replace(
				"<w:tblPr/>",
				'<w:tblPr><w:tblLook w:val="04A0" w:firstRow="0"/></w:tblPr>'
			)
		},
		finds: ["C05 tblLook val and attributes disagree"]
	},
	C06: {
		spec: { body: '<w:p><w:r><w:footnoteReference w:id="2"/></w:r></w:p>', footnotes: '<w:footnote w:id="1"><w:p/></w:footnote>' },
		finds: ["C06 footnotes part without separators", "C06 footnote reference unresolved"]
	},
	C07: {
		spec: {
			body: '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr></w:p>',
			parts: [
				{
					name: "word/numbering.xml",
					contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
					data: `<w:numbering ${NAMESPACES}><w:num w:numId="1"><w:abstractNumId w:val="5"/></w:num><w:abstractNum w:abstractNumId="-1"/></w:numbering>`
				}
			],
			relationships: [{ id: "rIdN", type: "numbering", target: "numbering.xml" }]
		},
		finds: [
			"C07 abstractNum after num",
			"C07 abstractNumId negative or not a number",
			"C07 num refers to a missing abstractNum",
			"C07 numId unresolved"
		]
	},
	C08: {
		spec: { body: `<w:p>${DRAWING("5")}${DRAWING("5")}${DRAWING("0")}</w:p>` },
		finds: ["C08 docPr id duplicated", "C08 docPr id out of range"]
	},
	C09: {
		spec: {
			body:
				'<w:p><w:commentRangeStart w:id="0"/><w:r><w:commentReference w:id="7"/></w:r><w:bookmarkStart w:id="3" w:name="b"/>' +
				'<w:del w:id="-1" w:author="A"><w:r><w:delText>x</w:delText></w:r></w:del></w:p>',
			comments: '<w:comment w:id="0" w:author="A"><w:p/></w:comment>'
		},
		finds: ["C09 comment range unpaired", "C09 comment reference unresolved", "C09 bookmark range unpaired", "C09 revision id -1"]
	},
	C10: {
		spec: {
			body: '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r></w:p><w:p><w:r><w:fldChar w:fldCharType="end"/><w:fldChar w:fldCharType="end"/></w:r></w:p>'
		},
		finds: ["C10 field end without begin"]
	},
	C11: {
		spec: {
			body:
				"<w:tbl><w:tblPr/></w:tbl>" +
				TABLE(
					'<w:tr><w:tc><w:tcPr><w:gridSpan w:val="3"/></w:tcPr><w:p/>' +
						TABLE("<w:tr><w:tc><w:p/></w:tc></w:tr>") +
						"</w:tc></w:tr>"
				) +
				p("after")
		},
		finds: ["C11 table without rows", "C11 table without tblGrid", "C11 cell not ending in a paragraph", "C11 row wider than its grid"]
	},
	C12: {
		spec: { body: TABLE("<w:tr><w:tc><w:p><w:pPr><w:sectPr/></w:pPr></w:p></w:tc></w:tr>") + "<w:sectPr/>" + p("after") },
		finds: ["C12 sectPr in a paragraph Word does not allow it in", "C12 body sectPr not last"]
	},
	C13: {
		spec: {
			body: '<w:p><w:hyperlink r:id="rIdI"><w:r><w:t>a</w:t></w:r></w:hyperlink><w:hyperlink r:id="rIdH"><w:r><w:t>b</w:t></w:r></w:hyperlink></w:p>',
			relationships: [
				{ id: "rIdI", type: "styles", target: "styles.xml" },
				{ id: "rIdH", type: "hyperlink", target: "https://example.com/" }
			],
			parts: [STYLES_PART("")]
		},
		finds: [
			"C13 hyperlink r:id not a hyperlink relationship",
			"C13 hyperlink relationship not External",
			"C13 hyperlink to a URL not External"
		]
	},
	C14: {
		spec: {
			body: `<w:p><w:bookmarkStart w:id="1" w:name="${"n".repeat(41)}"/><w:bookmarkEnd w:id="1"/><w:bookmarkStart w:id="2" w:name="a"/><w:bookmarkEnd w:id="2"/><w:bookmarkStart w:id="3" w:name="a"/><w:bookmarkEnd w:id="3"/></w:p>`
		},
		finds: ["C14 bookmark name over 40 characters", "C14 bookmark name duplicated"]
	},
	C15: {
		spec: { body: '<w:sdt><w:sdtPr><w:id w:val="4"/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>'.repeat(2) },
		finds: ["C15 sdt id duplicated"]
	},
	C16: {
		spec: { body: '<w:p w14:paraId="80000000"/><w:p w14:paraId="0000ABCD"/><w:p w14:paraId="0000ABCD"/>' },
		finds: ["C16 paraId out of range", "C16 paraId duplicated in its part"]
	},
	C17: {
		spec: {
			body: '<w:p><w:r><w:commentReference w:id="0"/></w:r></w:p>',
			comments: '<w:comment w:id="0" w:author="A"><w:p w14:paraId="00000001"/></w:comment>',
			commentsExtended: `<w15:commentsEx ${NAMESPACES}><w15:commentEx w15:paraId="00000009" w15:done="0"/></w15:commentsEx>`,
			people: `<w15:people ${NAMESPACES}><w15:person w15:author="Nobody"/></w15:people>`
		},
		finds: ["C17 commentEx paraId not a comment's last paragraph", "C17 person without a comment or revision"]
	},
	C18: {
		spec: {
			body: p("x"),
			settings: '<w:documentProtection w:edit="readOnly" w:enforcement="0" w:cryptSpinCount="20000000" w:hash="AA=="/>'
		},
		finds: ["C18 spin count out of range", "C18 incomplete hash attributes"]
	},
	C19: {
		spec: {
			document: `<w:document ${NAMESPACES.replace('mc:Ignorable="w14 w15 wp14"', 'mc:Ignorable="w14 zz"')}><w:body>${p("x")}</w:body></w:document>`,
			parts: [
				{
					name: "docProps/core.xml",
					contentType: PROPERTIES,
					data: '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="cp"/>'
				}
			],
			packageRelationships: [
				{
					id: "rIdP",
					type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
					target: "docProps/core.xml"
				}
			]
		},
		finds: ["C19 undeclared prefix in Ignorable", "C19 markup compatibility in a properties part"]
	},
	C20: {
		spec: { body: "<w:p><w:r><w:t>a&#1;b\rc</w:t></w:r></w:p>" },
		finds: ["C20 character not allowed in XML 1.0", "C20 carriage return in content"]
	},
	C21: { spec: { body: p("x"), parts: [{ name: "word/blob.bin", data: "x" }] }, finds: ["C21 part without a content type"] },
	C22: {
		spec: {
			body: '<w:p><w:hyperlink r:id="rIdNone"><w:r><w:t>a</w:t></w:r></w:hyperlink></w:p>',
			relationships: [
				{ id: "rIdM", type: "image", target: "media/missing.png" },
				{ id: "rIdM", type: "image", target: "media/missing.png" }
			]
		},
		finds: ["C22 internal target missing", "C22 relationship id duplicated", "C22 relationship reference unresolved"]
	},
	C23: {
		spec: {
			body: '<w:p><w:r><w:commentReference w:id="0"/></w:r></w:p>',
			comments: '<w:comment w:id="0" w:author="A" w:initials="ABCDEFGHIJ"><w:p/></w:comment>'
		},
		finds: ["C23 comment initials over 9 characters"]
	},
	C24: {
		spec: {
			body:
				'<w:p><w:del w:id="1" w:author="A" w:date="2026-01-01T10:00:00Z"><w:fldSimple w:instr=" PAGE "><w:r><w:delText>1</w:delText></w:r></w:fldSimple></w:del>' +
				'<w:ins w:id="2" w:author="A" w:date="2026-01-01T10:00:00Z"><w:hyperlink w:anchor="x"><w:r><w:t>a</w:t></w:r></w:hyperlink></w:ins></w:p>'
		},
		finds: ["C24 fldSimple inside a tracked change", "C24 hyperlink inside a tracked change"]
	},
	C25: {
		spec: { body: p("x") + '<w:sectPr><w:lnNumType w:countBy="0" w:start="40000"/></w:sectPr>' },
		finds: ["C25 lnNumType countBy outside 1..100", "C25 lnNumType start outside 0..32767"]
	}
}

// ---- the writer's scenarios ------------------------------------------------------------------------------

interface Saved {
	readonly name: string
	readonly source: Uint8Array
	readonly bytes: Uint8Array
}

type Editor = Awaited<ReturnType<typeof pagedEditor>>

const HEADING_STYLES = STYLES_PART(
	'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
		'<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>' +
		'<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:u w:val="single"/></w:rPr></w:style>'
)
const PLAIN = buildDocx({
	body:
		p("Alpha beta gamma delta") +
		'<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Results</w:t></w:r></w:p>' +
		p("Second paragraph here"),
	parts: [HEADING_STYLES],
	relationships: [{ id: "rIdS", type: "styles", target: "styles.xml" }]
})
const FIELDS = buildDocx({
	body:
		'<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r><w:fldSimple w:instr=" PAGE " w:dirty="true"><w:r><w:rPr><w:b/></w:rPr><w:t>7</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> now</w:t></w:r></w:p>' +
		p("after"),
	settings:
		'<w:documentProtection w:edit="trackedChanges" w:enforcement="0" w:cryptProviderType="rsaAES" w:cryptAlgorithmClass="hash" w:cryptAlgorithmType="typeAny" w:cryptAlgorithmSid="14" w:cryptSpinCount="100000" w:hash="AAAA" w:salt="AAAA"/>'
})

function select(editor: PagedEditor, from: number, to = from): void {
	const session = editor.session

	session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, from, to)))
}

function find(editor: PagedEditor, text: string): number {
	let found = -1

	editor.session.state.doc.descendants((node, pos) => {
		if (found < 0 && node.isText && (node.text ?? "").includes(text)) found = pos + (node.text ?? "").indexOf(text)

		return found < 0
	})

	if (found < 0) throw new Error(`No "${text}"`)

	return found
}

async function edited(name: string, source: Uint8Array, steps: (editor: PagedEditor) => Promise<void>): Promise<Saved> {
	const opened: Editor = await pagedEditor(source)

	opened.context.identity.author = "Ann Lee"
	opened.context.identity.initials = "AL"
	opened.context.identity.now = () => Date.UTC(2026, 9, 12, 10, 30)
	await steps(opened.editor)
	await settleAll(opened.editor)

	const result = await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: opened.editor.sessions })

	if (result.bytes === undefined) throw new Error(`${name}: not saved`)

	return { name, source, bytes: result.bytes }
}

async function run(editor: PagedEditor, command: Parameters<PagedEditor["execute"]>[0]): Promise<void> {
	if (!(await editor.execute(command))) throw new Error(`${command.type} did not apply`)
}

// Every scenario's save; each makes constructs of the items it names.
async function scenarios(): Promise<Saved[]> {
	const out: Saved[] = [
		await edited("comments", PLAIN, async editor => {
			select(editor, find(editor, "beta"), find(editor, "beta") + 4)
			await run(editor, { type: "comment.add", text: "First" })
			select(editor, find(editor, "Second"), find(editor, "Second") + 6)
			await run(editor, { type: "comment.add", text: "Second" })
			await run(editor, { type: "comment.reply", id: "0", text: "A reply" })
		}),
		await edited("notes", PLAIN, async editor => {
			select(editor, find(editor, "gamma"))
			await run(editor, { type: "note.insertFootnote" })
			editor.session.dispatch(editor.session.state.tr.insertText("Footnote text"))
		}),
		await edited("endnotes", PLAIN, async editor => {
			select(editor, find(editor, "gamma"))
			await run(editor, { type: "note.insertEndnote" })
		}),
		await edited("header", PLAIN, async editor => {
			select(editor, find(editor, "Alpha"))
			await run(editor, { type: "header.edit" })
			editor.session.dispatch(editor.session.state.tr.insertText("Header text"))
		}),
		await edited("tracked", FIELDS, async editor => {
			await run(editor, { type: "review.track", on: true })
			select(editor, find(editor, "after") + 5)
			editor.session.dispatch(editor.session.state.tr.insertText(" typed"))
			editor.session.dispatch(editor.session.state.tr.setSelection(new AllSelection(editor.session.state.doc)).deleteSelection())
		}),
		await edited("controls", PLAIN, async editor => {
			select(editor, find(editor, "delta") + 5)
			await run(editor, { type: "control.insert", kind: "plainText" })
			// (A control's placeholder says "here" too.)
			select(editor, find(editor, "Second paragraph here") + 21)
			await run(editor, { type: "control.insert", kind: "richText" })
		}),
		await edited("links", PLAIN, async editor => {
			select(editor, find(editor, "Alpha"), find(editor, "Alpha") + 5)
			await run(editor, { type: "link.set", target: "https://example.com/a b" })
			select(editor, find(editor, "here") + 4)
			await run(editor, { type: "link.set", heading: editor.session.state.doc.child(0).nodeSize })
		}),
		await edited("line numbers", PLAIN, async editor => {
			select(editor, find(editor, "Alpha"))
			await run(editor, { type: "page.setup", lineNumbers: { countBy: 5, start: 3 } })
		}),
		await edited("table", PLAIN, async editor => {
			select(editor, find(editor, "here") + 4)
			await run(editor, { type: "table.insert", rows: 2, cols: 3 })
			await run(editor, { type: "table.align", align: "end" })
		})
	]

	// The scripted edit suite over the committed fixtures.
	for (const name of readdirSync(FIXTURES)
		.filter(file => file.endsWith(".docx") && !file.startsWith("08-"))
		.sort()) {
		const source = new Uint8Array(readFileSync(join(FIXTURES, name)))
		const opened = await openBytes(source)

		if (opened.context.gate.length > 0) continue

		const script = runEditScript(opened)
		const result = await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: script.sessions })

		if (result.bytes === undefined) throw new Error(`${name}: not saved`)

		out.push({ name: `script ${name}`, source, bytes: result.bytes })
	}

	return out
}

interface Judged {
	readonly name: string
	readonly failures: readonly string[]
	readonly report: ConformanceReport
}

async function judge(saved: Saved, sourceReport?: ConformanceReport): Promise<Judged> {
	const before = sourceReport ?? (await checkConformance(saved.source))
	const changed = await changedParts(await OpcPackage.open(bytesSource(saved.source)), await OpcPackage.open(bytesSource(saved.bytes)))
	const after = await checkConformance(saved.bytes, new Set(changed))

	return { name: saved.name, failures: newViolations(before, after), report: after }
}

function ofItem(failures: readonly string[], id: string): string[] {
	return failures.filter(failure => failure.startsWith(`${id} `))
}

describe("the conformance catalogue", () => {
	it("has a crafted violation for every item, and every check it names", () => {
		expect(CATALOGUE.map(item => item.id)).toEqual(Object.keys(CRAFTED))
		expect(new Set(CATALOGUE.map(item => item.id)).size).toBe(CATALOGUE.length)

		for (const item of CATALOGUE) expect(item.source.length, item.id).toBeGreaterThan(10)
	})

	let judged: Judged[] = []

	beforeAll(async () => {
		judged = await Promise.all((await scenarios()).map(saved => judge(saved)))
	}, 300_000)

	for (const item of CATALOGUE) {
		describe(`${item.id} ${item.title}`, () => {
			it("finds a crafted violation", async () => {
				const crafted = CRAFTED[item.id]

				if (crafted === undefined) throw new Error(`No crafted violation for ${item.id}`)

				const report = await checkConformance(buildDocx(crafted.spec))

				for (const key of crafted.finds) expect(report.violations.get(key) ?? 0, key).toBeGreaterThan(0)
			})

			it("holds on the writer's saves", () => {
				const failures = judged.flatMap(entry => ofItem(entry.failures, item.id).map(failure => `${entry.name}: ${failure}`))
				const checked = judged.reduce((sum, entry) => sum + (entry.report.checked.get(item.id) ?? 0), 0)

				expect(failures).toEqual([])
				// Every item but the ones only a source can bring (a protected document's hash) checks constructs the
				// writer wrote.
				expect(checked, `${item.id} constructs in written parts`).toBeGreaterThan(0)
			})
		})
	}

	it.each([
		[
			"locked",
			'<w:fldSimple w:instr=" PAGE " w:fldLock="1"><w:r><w:t>1</w:t></w:r></w:fldSimple>',
			'<w:fldChar w:fldCharType="begin" w:fldLock="true"/>'
		],
		[
			"holding a bookmark",
			'<w:fldSimple w:instr=" PAGE "><w:bookmarkStart w:id="5" w:name="bm"/><w:r><w:t>1</w:t></w:r><w:bookmarkEnd w:id="5"/></w:fldSimple>',
			'<w:bookmarkStart w:id="5" w:name="bm"/>'
		]
	])("C24: a tracked deletion over a %s w:fldSimple writes it complex, and the save goes on", async (_name, field, kept) => {
		const opened = await openSpec({ body: p("Before") + `<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r>${field}</w:p>` })

		opened.context.identity.author = "Ann"
		opened.context.review.tracking = true

		const session = new DocxEditorSession(opened.context)

		session.dispatch(session.state.tr.setSelection(new AllSelection(session.state.doc)).deleteSelection())

		const saved = await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: [session] })
		const xml = new TextDecoder().decode(
			await (await OpcPackage.open(bytesSource(saved.bytes as Uint8Array))).read("word/document.xml")
		)

		expect(xml).not.toContain("<w:fldSimple")
		expect(xml).toContain('<w:delInstrText xml:space="preserve"> PAGE </w:delInstrText>')
		expect(xml).toContain(kept)
	})

	// A w:fldSimple holding a field of its own or a drawing has no complex form a w:del could hold: a tracked
	// deletion leaves it as it is (as a locked block), and a rejection gives the document back whole.
	it.each([
		[
			"holding a field of its own",
			'<w:fldSimple w:instr=" IF 1 = 1 x "><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:fldSimple>'
		],
		[
			"holding a picture",
			'<w:fldSimple w:instr=" PAGE "><w:r><w:drawing><wp:inline><wp:extent cx="10" cy="10"/><wp:docPr id="7" name="p"/><a:graphic>' +
				'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"/></a:graphic></wp:inline></w:drawing></w:r></w:fldSimple>'
		]
	])("C24: a tracked deletion leaves a w:fldSimple %s out, and rejected gives it back", async (_name, field) => {
		const opened = await openSpec({ body: p("Before") + `<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r>${field}</w:p>` })

		opened.context.identity.author = "Ann"
		opened.context.review.tracking = true

		const session = new DocxEditorSession(opened.context)
		const written = async (): Promise<string> =>
			new TextDecoder().decode(
				await (
					await OpcPackage.open(
						bytesSource((await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: [session] })).bytes as Uint8Array)
					)
				).read("word/document.xml")
			)

		session.dispatch(session.state.tr.setSelection(new AllSelection(session.state.doc)).deleteSelection())

		const deleted = await written()

		expect(deleted).toContain("<w:delText")
		expect(/<w:del\b[^>]*>(?:(?!<\/w:del>).)*<w:fldSimple/s.test(deleted)).toBe(false)
		expect(deleted).toContain(field)

		const rejected = resolveChanges(session, storyReviewIndex(session), { all: true }, false)

		session.dispatch(rejected.tr as never)
		expect(await written()).toContain(`<w:r><w:t xml:space="preserve">A </w:t></w:r>${field}`)
	})

	it("C24: a copy of a w:fldSimple pasted while tracking writes a complex field inside w:ins", { timeout: 120_000 }, async () => {
		const saved = await edited("fldSimple copy", FIELDS, async editor => {
			await run(editor, { type: "review.track", on: true })

			const state = editor.session.state
			const first = state.doc.child(0)

			editor.session.dispatch(state.tr.insert(state.doc.content.size, first.type.create(null, first.content)))
		})
		const xml = new TextDecoder().decode(await (await OpcPackage.open(bytesSource(saved.bytes))).read("word/document.xml"))

		expect(ofItem((await judge(saved)).failures, "C24")).toEqual([])
		const pasted = xml.slice(xml.lastIndexOf("<w:p>"))

		expect(pasted).toContain('<w:fldChar w:fldCharType="begin"/>')
		expect(pasted).toContain('<w:instrText xml:space="preserve"> PAGE </w:instrText>')
		expect(pasted).not.toContain("<w:fldSimple")
	})

	it("C24: a tracked deletion over a w:fldSimple writes a complex field inside w:del", { timeout: 120_000 }, async () => {
		const saved = await edited("fldSimple", FIELDS, async editor => {
			await run(editor, { type: "review.track", on: true })
			editor.session.dispatch(editor.session.state.tr.setSelection(new AllSelection(editor.session.state.doc)).deleteSelection())
		})
		const xml = new TextDecoder().decode(await (await OpcPackage.open(bytesSource(saved.bytes))).read("word/document.xml"))

		expect(ofItem((await judge(saved)).failures, "C24")).toEqual([])
		expect(xml).toContain('<w:fldChar w:fldCharType="begin" w:dirty="true"/>')
		expect(xml).toContain('<w:delInstrText xml:space="preserve"> PAGE </w:delInstrText>')

		if (xsdOracleMissing() === undefined) {
			const oracle = new XsdOracle()

			try {
				expect([...(oracle.validate([encodeUtf8(xml)])[0]?.keys() ?? [])]).toEqual([])
			} finally {
				oracle.dispose()
			}
		}

		// Rejecting brings the field back as it looked; accepting removes it.
		const rejected = await edited("rejected", saved.bytes, async editor => run(editor, { type: "review.rejectAll" }))
		const back = new TextDecoder().decode(await (await OpcPackage.open(bytesSource(rejected.bytes))).read("word/document.xml"))

		expect(back).toContain('<w:instrText xml:space="preserve"> PAGE </w:instrText>')
		expect(back).toContain("<w:r><w:rPr><w:b/></w:rPr><w:t>7</w:t></w:r>")

		const accepted = await edited("accepted", saved.bytes, async editor => run(editor, { type: "review.acceptAll" }))

		expect(
			new TextDecoder().decode(await (await OpcPackage.open(bytesSource(accepted.bytes))).read("word/document.xml"))
		).not.toContain("PAGE")
	})

	const manifest = fuzzManifest()

	it.skipIf(!existsSync(CORPUS_DIR))("no save of the edit suite or the fuzz saves adds a violation", { timeout: 7_200_000 }, async () => {
		const sources = new Map<string, ConformanceReport>()
		const failures: string[] = []
		const checked = new Map<string, number>()
		let saves = 0
		const take = async (save: { name: string; sourceName: string; source: Uint8Array; bytes: Uint8Array }): Promise<void> => {
			let before = sources.get(save.sourceName)

			if (before === undefined) {
				before = await checkConformance(save.source)
				sources.set(save.sourceName, before)
			}

			const result = await judge(save, before)

			saves++

			for (const failure of result.failures) failures.push(`${save.name}: ${failure}`)

			for (const [id, count] of result.report.checked) checked.set(id, (checked.get(id) ?? 0) + count)
		}

		for await (const save of editSuite()) await take(save)
		if (manifest !== undefined) for await (const save of fuzzSaves(manifest)) await take(save)

		console.log(
			`conformance: ${saves} saves; constructs checked in written parts ${JSON.stringify(Object.fromEntries([...checked].sort()))}`
		)
		expect(failures).toEqual([])
	})
})
