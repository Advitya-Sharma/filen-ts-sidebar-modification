import { describe, expect, expectTypeOf, it } from "vitest"
import { RevisionView } from "@filen/office/core/docx"
import { DOCX_EDITOR_COMMANDS, DOCX_PARAGRAPH_STYLES, NO_REVIEW, NO_VALUES, type DocxEditorSnapshot } from "@filen/office/core/docx/edit"
import {
	COMMAND_LABELS,
	DOCX_COMMAND_CATALOGUE,
	commandActive,
	commandFeature,
	commandRecordsAuthor,
	editorMarkupView,
	featureEditable,
	rowCommands,
	slotCommands,
	type DocxUncataloguedCommand,
	editorMarks,
	editorViewOnlyReasons,
	firstMatchFrom,
	lockReasonKey,
	featureLanded,
	markupRelayout,
	objectInteractions,
	shownMarkup,
	sameToolbarSnapshot,
	shortcutKeys,
	stepMatch,
	styleLabel,
	toolbarFit
} from "@/features/preview/components/docxEditor.logic"
import { viewOnlyReasonKey, viewOnlyReasons } from "@/features/preview/components/docxViewer.logic"
import { preview as en } from "@/locales/en/preview"

const VALUES = NO_VALUES
const CONTEXT: DocxEditorSnapshot["context"] = {
	story: { kind: "body", index: 0 },
	table: null,
	drawing: null,
	control: null,
	field: null,
	link: null,
	comment: null,
	equation: null
}
const REVIEW = NO_REVIEW

function snapshot(overrides: Partial<DocxEditorSnapshot> = {}): DocxEditorSnapshot {
	const enabled = Object.fromEntries(DOCX_EDITOR_COMMANDS.map(command => [command.type, true])) as DocxEditorSnapshot["enabled"]

	return {
		version: 1,
		readOnly: false,
		readOnlyReasons: [],
		dirty: false,
		canUndo: false,
		canRedo: false,
		selection: { empty: true, from: 1, to: 1, locked: null, paragraphs: 1 },
		formats: { bold: "on", italic: "mixed", underline: "off", strike: "off" },
		paragraph: { style: "normal", styleName: null, align: "center", list: "bullet", level: 0 },
		values: VALUES,
		context: CONTEXT,
		review: REVIEW,
		enabled,
		...overrides
	}
}

describe("docx editor toolbar logic", () => {
	it("keeps the toolbar's snapshot across keystrokes and changes it when anything it shows changes", () => {
		const shown = snapshot()

		// Typing moves the selection and bumps the version: nothing the toolbar shows.
		expect(
			sameToolbarSnapshot(
				shown,
				snapshot({ version: 2, dirty: true, selection: { empty: true, from: 9, to: 9, locked: null, paragraphs: 1 } })
			)
		).toBe(true)
		expect(sameToolbarSnapshot(shown, snapshot({ formats: { ...shown.formats, bold: "off" } }))).toBe(false)
		expect(sameToolbarSnapshot(shown, snapshot({ paragraph: { ...shown.paragraph, list: "none" } }))).toBe(false)
		expect(sameToolbarSnapshot(shown, snapshot({ enabled: { ...shown.enabled, "history.undo": false } }))).toBe(false)
		expect(sameToolbarSnapshot(shown, snapshot({ canUndo: true }))).toBe(false)
		expect(sameToolbarSnapshot(shown, snapshot({ values: { ...VALUES, size: 24 } }))).toBe(false)
		expect(sameToolbarSnapshot(shown, snapshot({ review: { ...REVIEW, tracking: true } }))).toBe(false)
		expect(sameToolbarSnapshot(shown, snapshot({ readOnly: true, readOnlyReasons: ["host"] }))).toBe(false)
		expect(sameToolbarSnapshot(shown, snapshot({ selection: { empty: false, from: 1, to: 5, locked: "proof", paragraphs: 1 } }))).toBe(
			false
		)
		expect(sameToolbarSnapshot(undefined, shown)).toBe(false)
		expect(sameToolbarSnapshot(undefined, undefined)).toBe(true)
	})

	it("labels every base command and every paragraph style with a real key", () => {
		for (const command of DOCX_EDITOR_COMMANDS) {
			const label = COMMAND_LABELS[command.type]

			if (commandFeature(command.type) === "base") {
				expect(label, command.type).toBeDefined()
			}

			if (label !== undefined) {
				expect(en).toHaveProperty(label)
			}
		}

		for (const style of DOCX_PARAGRAPH_STYLES) {
			expect(en).toHaveProperty(styleLabel(style).key)
		}

		expect(styleLabel("heading3")).toEqual({ key: "previewDocxStyleHeading", level: 3 })
		expect(styleLabel("quote")).toEqual({ key: "previewDocxStyleQuote" })
	})

	it("shows toggles on only when the whole selection has them, and the alignment and list in effect", () => {
		const state = snapshot()

		expect(commandActive(state, "format.bold")).toBe(true)
		expect(commandActive(state, "format.italic")).toBe(false)
		expect(commandActive(state, "align.center")).toBe(true)
		expect(commandActive(state, "align.start")).toBe(false)
		expect(commandActive(state, "list.bullet")).toBe(true)
		expect(commandActive(state, "list.numbered")).toBe(false)
		expect(commandActive(state, "history.undo")).toBe(false)
	})

	it("shows shortcuts with the platform's modifier names", () => {
		expect(shortcutKeys("mod+shift+z", true)).toEqual(["⌘", "⇧", "Z"])
		expect(shortcutKeys("mod+alt+1", false)).toEqual(["Ctrl", "Alt", "1"])
		expect(shortcutKeys("mod+space", false)).toEqual(["Ctrl", "Space"])
	})

	it("moves groups into the menu as the row narrows, in priority order", () => {
		expect(toolbarFit(2000, 0)).toEqual({ style: true, lists: true, align: true, clear: true })
		expect(toolbarFit(300, 0)).toEqual({ style: false, lists: false, align: false, clear: false })

		const middle = toolbarFit(640, 0)

		expect(middle.style).toBe(true)
		expect(middle.clear).toBe(false)
		// Badges take room too.
		expect(toolbarFit(640, 300).style).toBe(false)
	})

	it("explains read-only editing with keys that exist", () => {
		expect(editorViewOnlyReasons(["host", "protected", "linkedTextbox"], true, false)).toEqual(["protected", "linkedTextbox"])
		expect(editorViewOnlyReasons([], false, false)).toEqual(["notWritable"])
		expect(editorViewOnlyReasons([], false, true)).toEqual(["renamed"])

		for (const reason of ["renamed", "notWritable", "linkedTextbox", "textboxFallback", "lockedTextbox", "protected"] as const) {
			expect(en).toHaveProperty(viewOnlyReasonKey(reason))
		}

		for (const reason of ["contentLocked", "proof", "tableStructure", "fieldSpansBlocks"] as const) {
			expect(en).toHaveProperty(lockReasonKey(reason))
		}
	})
})

describe("docx command catalogue", () => {
	it("places every package command in exactly one feature, and lists only package commands", () => {
		expectTypeOf<DocxUncataloguedCommand>().toEqualTypeOf<never>()

		for (const command of DOCX_EDITOR_COMMANDS) {
			expect(commandFeature(command.type), command.type).toBeDefined()
		}

		const ids: string[] = Object.values(DOCX_COMMAND_CATALOGUE).flat()

		expect(new Set(ids).size).toBe(ids.length)
		expect(ids.toSorted()).toEqual(DOCX_EDITOR_COMMANDS.map(command => command.type).toSorted())
	})

	it("shows the row's groups from base commands only, and a slot only its labelled commands", () => {
		expect(rowCommands("history").map(command => command.info.type)).toEqual(["history.undo", "history.redo"])
		expect(rowCommands("style")).toEqual([])
		expect(rowCommands("align").map(command => command.label)).toEqual([
			"previewDocxAlignStart",
			"previewDocxAlignCenter",
			"previewDocxAlignEnd",
			"previewDocxAlignJustify"
		])

		for (const slot of ["insert", "review", "table", "format", "view"] as const) {
			for (const command of slotCommands(slot)) {
				expect(commandFeature(command.info.type)).not.toBe("base")
				expect(COMMAND_LABELS[command.info.type]).toBe(command.label)
			}
		}
	})

	it("asks for the author only for commands that will record one", () => {
		for (const command of DOCX_EDITOR_COMMANDS) {
			if (command.author !== true) {
				expect(commandRecordsAuthor({ type: command.type } as never, false), command.type).toBe(false)
			}
		}

		expect(commandRecordsAuthor({ type: "comment.add", text: "x" }, false)).toBe(true)
		expect(commandRecordsAuthor({ type: "comment.reply", id: "1", text: "x" }, true)).toBe(true)
		expect(commandRecordsAuthor({ type: "comment.edit", id: "1", text: "x" }, true)).toBe(true)
		expect(commandRecordsAuthor({ type: "comment.delete", id: "1" }, true)).toBe(false)
		// Tracking can always be turned off without a name.
		expect(commandRecordsAuthor({ type: "review.track", on: true }, false)).toBe(true)
		expect(commandRecordsAuthor({ type: "review.track", on: false }, true)).toBe(false)
		expect(commandRecordsAuthor({ type: "review.track" }, false)).toBe(true)
		expect(commandRecordsAuthor({ type: "review.track" }, true)).toBe(false)
	})
})

describe("docx editor find logic", () => {
	const matches = [
		{ from: 5, to: 8 },
		{ from: 20, to: 23 },
		{ from: 40, to: 43 }
	]

	it("starts at the first match from the caret, wrapping to the first", () => {
		expect(firstMatchFrom(matches, 0)).toBe(0)
		expect(firstMatchFrom(matches, 20)).toBe(1)
		expect(firstMatchFrom(matches, 21)).toBe(2)
		expect(firstMatchFrom(matches, 41)).toBe(0)
		expect(firstMatchFrom([], 3)).toBe(-1)
	})

	it("steps through matches, wrapping both ways", () => {
		expect(stepMatch(2, 3, 1)).toBe(0)
		expect(stepMatch(0, 3, -1)).toBe(2)
		expect(stepMatch(-1, 3, 1)).toBe(0)
		expect(stepMatch(-1, 3, -1)).toBe(2)
		expect(stepMatch(0, 0, 1)).toBe(-1)
	})

	it("merges spelling marks and matches in document order, the current match apart", () => {
		const spelling = [
			{ from: 1, to: 3, kind: "spelling" as const },
			{ from: 21, to: 22, kind: "spelling" as const },
			{ from: 50, to: 55, kind: "repeated" as const }
		]
		const marks = editorMarks(spelling, matches, 1)

		expect(marks.map(mark => `${String(mark.from)}:${mark.kind}`)).toEqual([
			"1:spelling",
			"5:find",
			"20:find-active",
			"21:spelling",
			"40:find",
			"50:repeated"
		])
		expect(editorMarks(spelling, [], -1)).toBe(spelling)
		expect(editorMarks([], [], -1)).toEqual([])
	})
})

describe("markup modes while editing", () => {
	it("lays the document out again only when the pages show another view", () => {
		expect(markupRelayout("final", "none")).toBeNull()
		expect(markupRelayout("final", "all")).toBe("markup")
		expect(markupRelayout("markup", "original")).toBe("original")
		expect(markupRelayout("original", "simple")).toBe("final")
	})

	it("shows the mode the editor's layout has, the user's where it fits", () => {
		expect(editorMarkupView(RevisionView.Markup)).toBe("markup")
		expect(editorMarkupView(RevisionView.Final)).toBe("final")
		expect(editorMarkupView(RevisionView.Original)).toBe("original")
		expect(shownMarkup("none", "final")).toBe("none")
		expect(shownMarkup("simple", "final")).toBe("simple")
		// Another view chosen elsewhere (a command): the menu follows the pages.
		expect(shownMarkup("simple", "markup")).toBe("all")
		expect(shownMarkup("all", "original")).toBe("original")
		expect(shownMarkup("original", "final")).toBe("simple")
	})
})

describe("object presses", () => {
	const press = { hit: { kind: "control", story: 0, pos: 1 } } as never

	it("takes presses only for objects of features offered", () => {
		const landed = (["drawings", "forms", "comments", "fields", "notes", "math"] as const).some(featureLanded)

		expect(
			objectInteractions(
				() => undefined,
				() => true
			) === undefined
		).toBe(!landed)
	})

	it("claims a press only while its feature can change the document", () => {
		const state = snapshot()

		expect(featureEditable(state, false, "base")).toBe(true)
		expect(featureEditable(state, true, "base")).toBe(false)
		expect(featureEditable({ ...state, readOnly: true }, false, "base")).toBe(false)
		expect(featureEditable(undefined, false, "base")).toBe(false)
		expect(
			featureEditable(
				snapshot({
					enabled: { ...state.enabled, ...Object.fromEntries(DOCX_COMMAND_CATALOGUE.forms.map(type => [type, false])) }
				}),
				false,
				"forms"
			)
		).toBe(false)

		const pressed: unknown[] = []
		const interactions = objectInteractions(
			found => pressed.push(found),
			() => false
		)

		expect(interactions?.control?.(press) ?? false).toBe(false)
		expect(interactions?.link).toBeUndefined()
		expect(pressed).toEqual([])
	})
})

describe("view-only paths for templates and macro-enabled documents", () => {
	it("keeps a template read-only where every document is: a public link, an archive, another's link, shared", () => {
		const where = { anon: false, archive: false, linked: false, variant: undefined }

		// No reason of its own any more (the package no longer gates templates).
		expect(viewOnlyReasons([], { ...where, variant: "drive" })).toEqual([])
		expect(viewOnlyReasons([], { ...where, anon: true })).toEqual(["publicLink"])
		expect(viewOnlyReasons([], { ...where, archive: true })).toEqual(["archive"])
		expect(viewOnlyReasons([], { ...where, linked: true })).toEqual(["linkedFile"])
		expect(viewOnlyReasons([], { ...where, variant: "sharedIn" })).toEqual(["sharedIn"])
		expect(viewOnlyReasons(["protected"], { ...where, anon: true })).toEqual(["publicLink", "protected"])
	})
})
