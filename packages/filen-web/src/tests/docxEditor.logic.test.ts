import { describe, expect, it } from "vitest"
import { DOCX_EDITOR_COMMANDS, DOCX_PARAGRAPH_STYLES, type DocxEditorSnapshot } from "@filen/office/core/docx/edit"
import {
	COMMAND_LABELS,
	commandActive,
	editorMarks,
	editorViewOnlyReasons,
	firstMatchFrom,
	lockReasonKey,
	markupSwitch,
	shortcutKeys,
	stepMatch,
	styleLabel,
	toolbarFit
} from "@/features/preview/components/docxEditor.logic"
import { viewOnlyReasonKey, viewOnlyReasons } from "@/features/preview/components/docxViewer.logic"
import { preview as en } from "@/locales/en/preview"

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
		enabled,
		...overrides
	}
}

describe("docx editor toolbar logic", () => {
	it("labels every command and every paragraph style with a real key", () => {
		for (const command of DOCX_EDITOR_COMMANDS) {
			expect(en).toHaveProperty(COMMAND_LABELS[command.type])
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
		expect(editorViewOnlyReasons(["host", "protected", "textbox"], true, false)).toEqual(["protected", "textbox"])
		expect(editorViewOnlyReasons([], false, false)).toEqual(["notWritable"])
		expect(editorViewOnlyReasons([], false, true)).toEqual(["renamed"])

		for (const reason of ["renamed", "notWritable", "textbox", "protected"] as const) {
			expect(en).toHaveProperty(viewOnlyReasonKey(reason))
		}

		for (const reason of ["contentLocked", "proof", "tableStructure", "fieldSpansBlocks"] as const) {
			expect(en).toHaveProperty(lockReasonKey(reason))
		}
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

describe("markupSwitch", () => {
	it("keeps the open document between modes laid out the same, reopens for another only while unchanged", () => {
		expect(markupSwitch("final", "none", false)).toEqual({ reopen: false })
		expect(markupSwitch("final", "simple", true)).toEqual({ reopen: false })
		expect(markupSwitch("final", "all", true)).toEqual({ reopen: true })
		expect(markupSwitch("final", "original", false)).toBeNull()
		expect(markupSwitch("markup", "all", false)).toEqual({ reopen: false })
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
