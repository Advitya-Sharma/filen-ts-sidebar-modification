import { unzipSync } from "fflate"
import { TextSelection, type Transaction } from "prosemirror-state"
import { describe, expect, it } from "vitest"
import { bytesSource } from "@filen/office/core/opc"
import { buildDocxModel, gateReasons, openDocxPackage, type DocxPackage, type Story } from "@filen/office/core/docx"
import { viewOnlyProtection } from "../core/docx/package"
import { DocxEditorSession, docxSchema, openDocxEditContext, type DocxEditContext } from "@filen/office/core/docx/edit"
import { saveDocx } from "@filen/office/core/docx/write"
import { FORMS_AFTER, FORMS_BEFORE } from "../core/docx/edit/handlers/forms"
import { REVIEW_TRANSFORMER } from "../core/docx/edit/handlers/review"
import { refusedForForms } from "../core/docx/edit/invariants/forms"
import { commandAllowed, protectionMode, protectionOf } from "../core/docx/edit/protection"
import { buildDocx, p, type SyntheticDocx } from "./docxSynthetic"
import { pagedEditor } from "./pagedFixtures"

// Document protection that allows editing within limits (ECMA-376 Part 1 §17.15.1.29): forms (content controls,
// legacy form fields and ranges open to everyone), comments (comments and open ranges), tracked changes (every
// edit recorded, the switch locked on). Read-only protection stays view-only.

const decoder = new TextDecoder()

function transformer(tr: Transaction, session: DocxEditorSession): Transaction {
	const before = FORMS_BEFORE(tr, session)

	return FORMS_AFTER(REVIEW_TRANSFORMER?.(before, session) ?? before, session)
}

interface Opened {
	readonly pkg: DocxPackage
	readonly context: DocxEditContext
	readonly session: DocxEditorSession
}

async function open(spec: SyntheticDocx, story?: (stories: readonly Story[]) => Story | undefined): Promise<Opened> {
	const opened = await openDocxPackage(bytesSource(buildDocx(spec)))

	if (!opened.ok) throw new Error(opened.reason)

	const model = await buildDocxModel(opened.package)
	const context = await openDocxEditContext(opened.package, model, { author: "Ann Lee", clock: () => Date.UTC(2026, 9, 12) })
	const chosen = story?.(model.stories)

	return {
		pkg: opened.package,
		context,
		session: new DocxEditorSession(context, { transformer, ...(chosen === undefined ? {} : { story: chosen }) })
	}
}

const protection = (edit: string, extra = ""): string => `<w:documentProtection w:edit="${edit}" w:enforcement="1"${extra}/>`

function find(session: DocxEditorSession, text: string): number {
	let at = -1

	session.state.doc.descendants((node, pos) => {
		if (at < 0 && node.isText && node.text?.includes(text) === true) at = pos + (node.text?.indexOf(text) ?? 0)

		return at < 0
	})

	if (at < 0) throw new Error(`no "${text}"`)

	return at
}

function type(session: DocxEditorSession, text: string, at: number): { applied: boolean; reason: string | undefined } {
	const tr = session.state.tr.insertText(text, at)
	const applied = session.dispatch(tr)

	return { applied, reason: refusedForForms(tr) }
}

function select(session: DocxEditorSession, from: number, to = from): void {
	session.dispatch(session.state.tr.setSelection(TextSelection.create(session.state.doc, from, to)))
}

const RICH =
	'<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="7"/><w:showingPlcHdr/></w:sdtPr><w:sdtContent>' +
	'<w:r><w:t>Click or tap here to enter text.</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p>'
const CHECK =
	'<w:p><w:sdt><w:sdtPr><w:id w:val="8"/><w14:checkbox><w14:checked w14:val="0"/></w14:checkbox></w:sdtPr><w:sdtContent>' +
	'<w:r><w:t>☐</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t xml:space="preserve"> I agree</w:t></w:r></w:p>'
const PERMITTED = (group: string): string =>
	`<w:p><w:permStart w:id="3" ${group}/><w:r><w:t>Open text</w:t></w:r><w:permEnd w:id="3"/></w:p>`

describe("document protection: gate 0", () => {
	it("keeps read-only protection view-only and opens the limited modes", () => {
		const facts = (edit: string | undefined, enforced = true) => ({
			variant: "docx" as const,
			strict: false,
			macros: false,
			template: false,
			signatures: 0,
			documentProtection: { edit, enforced, formatting: false, hashed: true },
			writeProtection: undefined,
			glossary: false,
			attachedTemplate: false,
			subDocuments: 0,
			malformed: false
		})

		expect(gateReasons(facts("readOnly"))).toEqual(["protected"])
		expect(gateReasons(facts("someFutureMode"))).toEqual(["protected"])

		for (const edit of ["forms", "comments", "trackedChanges", "none", undefined]) expect(gateReasons(facts(edit))).toEqual([])

		expect(gateReasons(facts("readOnly", false))).toEqual([])
		expect(viewOnlyProtection({ edit: "forms", enforced: true, formatting: false, hashed: false })).toBe(false)
		expect(protectionMode("forms", true)).toBe("forms")
		expect(protectionMode("forms", false)).toBe("none")
		expect(protectionMode("readOnly", true)).toBe("none")
	})

	it("reads the mode of a document, the hash kept", async () => {
		const forms = await open({
			body: p("x"),
			settings: protection("forms", ' w:cryptProviderType="rsaAES" w:hashValue="abc=" w:saltValue="def="')
		})

		expect(protectionOf(forms.context)).toEqual({ mode: "forms", forcesTracking: false, hashed: true })
		expect(forms.session.readOnly).toBe(false)

		const readOnly = await open({ body: p("x"), settings: protection("readOnly") })

		expect(readOnly.session.snapshot().readOnlyReasons).toEqual(["protected"])
		expect(protectionOf(readOnly.context).mode).toBe("none")
	})
})

describe("document protection: forms", () => {
	it("lets only content controls change, typing at a control's edge going into it", async () => {
		const opened = await open({ body: p("Fixed text") + RICH + CHECK, settings: protection("forms") })
		const session = opened.session

		expect(type(session, "x", find(session, "Fixed") + 2)).toEqual({ applied: false, reason: "protected" })
		expect(session.dispatch(session.state.tr.delete(1, 4))).toBe(false)

		const placeholder = find(session, "Click or tap")

		expect(type(session, "Ada", placeholder + 3).applied).toBe(true)
		expect(session.state.doc.textContent).toContain("Name: Ada end")

		// Right after the control's text: still the control's.
		const end = find(session, "Ada") + 3

		expect(type(session, " L", end).applied).toBe(true)
		expect(session.state.doc.textContent).toContain("Name: Ada L end")
		select(session, find(session, "Ada L") + 1)
		expect(session.snapshot().context.control).toMatchObject({ protection: "forms", placeholder: false })

		// A checkbox changes by its command only.
		const box = find(session, "☐")

		expect(session.dispatch(session.state.tr.insertText("x", box, box + 1))).toBe(false)
		select(session, box)
		expect(session.execute({ type: "control.check" })).toBe(true)
		expect(session.state.doc.textContent).toContain("☒ I agree")

		const xml = decoder.decode(
			unzipSync((await saveDocx({ pkg: opened.pkg, context: opened.context, sessions: [session] })).bytes as Uint8Array)[
				"word/settings.xml"
			] ?? new Uint8Array()
		)

		expect(xml === "" || xml.includes(protection("forms"))).toBe(true)
	})

	it("opens ranges an editing permission gives everyone, not one user's", async () => {
		const opened = await open({
			body: p("Fixed") + PERMITTED('w:edGrp="everyone"') + PERMITTED('w:ed="someone@example.com"').replace(/w:id="3"/g, 'w:id="4"'),
			settings: protection("forms")
		})
		const session = opened.session
		const positions: number[] = []

		session.state.doc.descendants((node, pos) => {
			if (node.isText && node.text === "Open text") positions.push(pos)
		})

		expect(type(session, "Z", (positions[0] ?? 0) + 2).applied).toBe(true)
		expect(type(session, "Z", (positions[1] ?? 0) + 2).applied).toBe(false)
		select(session, (positions[0] ?? 0) + 1)
		expect(commandAllowed(session, "format.bold")).toBe(true)
		select(session, 2)
		expect(commandAllowed(session, "format.bold")).toBe(false)
		expect(commandAllowed(session, "control.check")).toBe(true)
		expect(commandAllowed(session, "comment.add")).toBe(false)
		expect(commandAllowed(session, "history.undo")).toBe(true)
	})
})

describe("document protection: comments", () => {
	it("lets comments change and nothing else of the body outside open ranges", async () => {
		const spec: SyntheticDocx = {
			body: p("Body text") + PERMITTED('w:edGrp="everyone"'),
			comments: '<w:comment w:id="0" w:author="A"><w:p><w:r><w:t>A comment</w:t></w:r></w:p></w:comment>',
			settings: protection("comments")
		}
		const body = await open(spec)

		expect(type(body.session, "x", find(body.session, "Body") + 1)).toEqual({ applied: false, reason: "protected" })
		expect(type(body.session, "x", find(body.session, "Open") + 1).applied).toBe(true)
		expect(commandAllowed(body.session, "comment.add")).toBe(true)
		expect(commandAllowed(body.session, "format.bold")).toBe(false)
		expect(commandAllowed(body.session, "control.check")).toBe(false)

		const comment = await open(spec, stories => stories.find(story => story.kind === "comment"))

		expect(comment.session.story.story.kind).toBe("comment")
		expect(type(comment.session, "!", find(comment.session, "A comment") + 9).applied).toBe(true)
		expect(commandAllowed(comment.session, "format.bold")).toBe(true)
	})
})

describe("document protection: tracked changes", () => {
	it("records every edit and keeps the switch and accepting or rejecting changes locked", async () => {
		const opened = await open({ body: p("Some text"), settings: protection("trackedChanges") })
		const session = opened.session

		// With the integration, the switch is on from the start; without it, from the first edit.
		expect(type(session, "New ", 1).applied).toBe(true)
		expect(opened.context.review.tracking).toBe(true)

		const node = session.state.doc.nodeAt(1)

		expect(node?.marks.some(mark => mark.type === docxSchema.marks.insertion)).toBe(true)

		for (const command of ["review.track", "review.accept", "review.reject", "review.acceptAll", "review.rejectAll"] as const)
			expect(commandAllowed(session, command)).toBe(false)

		expect(commandAllowed(session, "format.bold")).toBe(true)
		expect(commandAllowed(session, "comment.add")).toBe(true)
	})
})

describe("document protection: commands", () => {
	it("disables what the mode does not allow in the snapshot and refuses it", async () => {
		const opened = await open({ body: p("Fixed text") + RICH, settings: protection("forms") })
		const session = opened.session

		select(session, find(session, "Fixed") + 1)
		expect(session.snapshot().enabled["format.bold"]).toBe(false)
		expect(session.execute({ type: "format.bold" })).toBe(false)

		const tracked = await open({ body: p("x"), settings: protection("trackedChanges") })

		expect(tracked.context.review.tracking).toBe(true)
	})

	it("routes document commands through the mode: comments in comments mode, none in forms mode", async () => {
		for (const [edit, allowed] of [
			["comments", true],
			["forms", false]
		] as const) {
			const { editor, context } = await pagedEditor(buildDocx({ body: p("Commented text"), settings: protection(edit) }))

			context.identity.author = "Ann Lee"
			editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1, 9)))
			expect(editor.snapshot().enabled["comment.add"]).toBe(allowed)
			expect(await editor.execute({ type: "comment.add", text: "Why?" })).toBe(allowed)
			editor.dispose()
		}
	})
})
