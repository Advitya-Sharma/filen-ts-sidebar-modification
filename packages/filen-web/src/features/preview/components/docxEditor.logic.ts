import {
	DOCX_EDITOR_COMMANDS,
	type DocxEditorCommand,
	type DocxEditorCommandGroup,
	type DocxEditorCommandInfo,
	type DocxEditorCommandType,
	type DocxEditorSnapshot,
	type DocxParagraphStyle,
	type DocxReadOnlyReason,
	type LockReason,
	type StoryMatch
} from "@filen/office/core/docx/edit"
import { RevisionView } from "@filen/office/core/docx"
import { DEFAULT_MARKUP_MODE, markupInfo, type DocxMarkupInfo, type DocxMarkupMode } from "@filen/office/dom/docx"
import type { DocxEditorInteractions, DocxEditorMark, DocxObjectPress } from "@filen/office/react/docx/editor"
import type { ObjectHitKind } from "@filen/office/core/docx/paged"
import type { SpellingMark } from "@filen/office/core/docx/spelling"
import type { DocxViewOnlyReason } from "@/features/preview/components/docxViewer.logic"
import type { PreviewKey } from "@/lib/i18n"
import { BASE_LABELS } from "@/features/preview/components/docxCommandLabels/base"
import { FORMATTING_LABELS } from "@/features/preview/components/docxCommandLabels/formatting"
import { INSERT_LABELS } from "@/features/preview/components/docxCommandLabels/insert"
import { COMMENTS_LABELS } from "@/features/preview/components/docxCommandLabels/comments"
import { REVIEW_LABELS } from "@/features/preview/components/docxCommandLabels/review"
import { TABLES_LABELS } from "@/features/preview/components/docxCommandLabels/tables"
import { SECTIONS_LABELS } from "@/features/preview/components/docxCommandLabels/sections"
import { NOTES_LABELS } from "@/features/preview/components/docxCommandLabels/notes"
import { FIELDS_LABELS } from "@/features/preview/components/docxCommandLabels/fields"
import { FORMS_LABELS } from "@/features/preview/components/docxCommandLabels/forms"
import { DRAWINGS_LABELS } from "@/features/preview/components/docxCommandLabels/drawings"
import { MATH_LABELS } from "@/features/preview/components/docxCommandLabels/math"
import { VIEW_LABELS } from "@/features/preview/components/docxCommandLabels/view"

// Idle autosave: this long after the last edit while there are unsaved changes (one version per pause).
export const AUTOSAVE_IDLE_MS = 30_000

// The package's commands by feature: each feature labels its own in docxCommandLabels/<feature>.ts and fills
// its toolbar slot. An id the package lacks fails `satisfies`, a package id missing here the test's type check.
export const DOCX_COMMAND_CATALOGUE = {
	base: [
		"format.bold",
		"format.italic",
		"format.underline",
		"format.strike",
		"format.clear",
		"style.set",
		"align.start",
		"align.center",
		"align.end",
		"align.justify",
		"list.bullet",
		"list.numbered",
		"list.indent",
		"list.outdent",
		"history.undo",
		"history.redo"
	],
	formatting: [
		"format.superscript",
		"format.subscript",
		"format.doubleStrike",
		"format.smallCaps",
		"format.allCaps",
		"format.clearAll",
		"font.family",
		"font.size",
		"font.grow",
		"font.shrink",
		"font.spacing",
		"color.text",
		"color.highlight",
		"paragraph.lineSpacing",
		"paragraph.spacing",
		"paragraph.indent",
		"format.copy",
		"format.paste"
	],
	insert: ["link.set", "link.remove", "insert.pageBreak", "insert.lineBreak", "insert.horizontalLine", "insert.symbol"],
	comments: [
		"comment.add",
		"comment.reply",
		"comment.edit",
		"comment.resolve",
		"comment.reopen",
		"comment.delete",
		"comment.next",
		"comment.prev"
	],
	review: [
		"review.track",
		"review.accept",
		"review.reject",
		"review.acceptAll",
		"review.rejectAll",
		"review.next",
		"review.prev",
		"view.markup"
	],
	tables: [
		"table.insert",
		"table.delete",
		"table.row.insertAbove",
		"table.row.insertBelow",
		"table.row.delete",
		"table.column.insertLeft",
		"table.column.insertRight",
		"table.column.delete",
		"table.cells.merge",
		"table.cell.split",
		"table.row.header",
		"table.style",
		"table.align",
		"table.column.width",
		"table.cell.shading",
		"table.cell.vAlign",
		"table.distributeColumns"
	],
	sections: [
		"insert.sectionBreak",
		"header.edit",
		"footer.edit",
		"header.differentFirst",
		"header.differentOddEven",
		"header.linkToPrevious",
		"header.remove",
		"footer.remove",
		"section.deleteBreak",
		"page.setup"
	],
	notes: ["note.insertFootnote", "note.insertEndnote", "note.goto", "note.delete", "note.convert", "note.convertAll", "note.options"],
	fields: [
		"field.update",
		"field.updateAll",
		"toc.insert",
		"toc.update",
		"insert.pageNumber",
		"insert.pageCount",
		"insert.date",
		"bookmark.add",
		"insert.field",
		"bookmark.delete",
		"bookmark.rename"
	],
	forms: [
		"control.check",
		"control.choose",
		"control.date",
		"control.clear",
		"control.text",
		"control.insert",
		"control.item.add",
		"control.item.remove"
	],
	drawings: ["image.insert", "drawing.resize", "drawing.position", "drawing.wrap", "drawing.altText", "drawing.rotate", "drawing.delete"],
	math: ["equation.insert", "equation.edit"],
	view: ["view.reflow"]
} as const satisfies Record<string, readonly DocxEditorCommandType[]>

export type DocxCommandFeature = keyof typeof DOCX_COMMAND_CATALOGUE

type CatalogueCommand<F extends DocxCommandFeature> = (typeof DOCX_COMMAND_CATALOGUE)[F][number]

// Package commands the catalogue does not place: must stay never.
export type DocxUncataloguedCommand = Exclude<DocxEditorCommandType, CatalogueCommand<DocxCommandFeature>>

// A feature's labels: a command is offered once it has one.
export type DocxCommandLabels<F extends DocxCommandFeature> = Readonly<Partial<Record<CatalogueCommand<F>, PreviewKey>>>

export const COMMAND_LABELS: Readonly<Partial<Record<DocxEditorCommandType, PreviewKey>>> = {
	...BASE_LABELS,
	...FORMATTING_LABELS,
	...INSERT_LABELS,
	...COMMENTS_LABELS,
	...REVIEW_LABELS,
	...TABLES_LABELS,
	...SECTIONS_LABELS,
	...NOTES_LABELS,
	...FIELDS_LABELS,
	...FORMS_LABELS,
	...DRAWINGS_LABELS,
	...MATH_LABELS,
	...VIEW_LABELS
}

const FEATURES: ReadonlyMap<string, DocxCommandFeature> = new Map(
	(Object.keys(DOCX_COMMAND_CATALOGUE) as DocxCommandFeature[]).flatMap(feature =>
		DOCX_COMMAND_CATALOGUE[feature].map(type => [type, feature] as const)
	)
)

export function commandFeature(type: DocxEditorCommandType): DocxCommandFeature | undefined {
	return FEATURES.get(type)
}

// Where the toolbar offers a feature's commands: the row's own groups (base), a slot component, or no
// toolbar place (form controls are used where they are, in the document).
export type DocxToolbarSlot = "insert" | "review" | "table" | "format" | "view"

const FEATURE_SLOTS: Readonly<Record<DocxCommandFeature, DocxToolbarSlot | null>> = {
	base: null,
	formatting: "format",
	insert: "insert",
	comments: "review",
	review: "review",
	tables: "table",
	sections: "insert",
	notes: "insert",
	fields: "insert",
	forms: null,
	drawings: "insert",
	math: "insert",
	view: "view"
}

export interface DocxToolbarCommand {
	readonly info: DocxEditorCommandInfo
	readonly label: PreviewKey
}

function toolbarCommands(keep: (info: DocxEditorCommandInfo, feature: DocxCommandFeature) => boolean): DocxToolbarCommand[] {
	const out: DocxToolbarCommand[] = []

	for (const info of DOCX_EDITOR_COMMANDS) {
		const feature = commandFeature(info.type)
		const label = COMMAND_LABELS[info.type]

		if (feature !== undefined && label !== undefined && keep(info, feature)) {
			out.push({ info, label })
		}
	}

	return out
}

// The row's buttons of a base group (style.set has its own picker).
export function rowCommands(group: DocxEditorCommandGroup): DocxToolbarCommand[] {
	return toolbarCommands((info, feature) => feature === "base" && info.group === group && info.type !== "style.set")
}

// The labelled commands a toolbar slot offers, in catalogue order.
export function slotCommands(slot: DocxToolbarSlot): DocxToolbarCommand[] {
	return toolbarCommands((_info, feature) => FEATURE_SLOTS[feature] === slot)
}

// What a toolbar slot component gets (docxInsertMenu, docxReviewMenu, docxTableMenu, docxFormatControls, the
// reflow toggle): mounted only once it has commands.
export interface DocxToolbarSlotProps {
	commands: readonly DocxToolbarCommand[]
	snapshot: DocxEditorSnapshot | undefined
	readOnly: boolean
	apple: boolean
	execute: (command: DocxEditorCommand) => void
}

// Commands that can record an author: the host has the user's name before running one that will.
const AUTHOR_COMMANDS: ReadonlySet<string> = new Set(DOCX_EDITOR_COMMANDS.filter(info => info.author === true).map(info => info.type))

// Whether running `command` records the author: turning tracking off never does (it must stay possible without
// a name).
export function commandRecordsAuthor(command: DocxEditorCommand, tracking: boolean): boolean {
	if (!AUTHOR_COMMANDS.has(command.type)) {
		return false
	}

	return command.type === "review.track" ? (command.on ?? !tracking) : true
}

// Runs `run` with the user's author name, asking for it first when none is set; never when the user declines.
export type DocxAuthorGate = (run: (author: string) => void) => void

// The feature whose commands act on each kind of object a press on the pages can hit; links keep the pages'
// own handling.
const OBJECT_FEATURES: Readonly<Record<ObjectHitKind, DocxCommandFeature | null>> = {
	drawing: "drawings",
	control: "forms",
	commentMarker: "comments",
	field: "fields",
	link: null,
	noteRef: "notes",
	equation: "math",
	tableBorder: "tables"
}

// A feature is offered once any of its commands is labelled.
export function featureLanded(feature: DocxCommandFeature): boolean {
	return DOCX_COMMAND_CATALOGUE[feature].some(type => COMMAND_LABELS[type] !== undefined)
}

// Whether a feature can change the document now: writable, and one of its commands enabled.
export function featureEditable(snapshot: DocxEditorSnapshot | undefined, readOnly: boolean, feature: DocxCommandFeature): boolean {
	return !readOnly && snapshot !== undefined && !snapshot.readOnly && DOCX_COMMAND_CATALOGUE[feature].some(type => snapshot.enabled[type])
}

// Presses on objects of the features offered go to `onPress` (no caret is placed) while `editable` says the
// feature can act; others reach the pages as before. Undefined while no feature is offered.
export function objectInteractions(
	onPress: (press: DocxObjectPress) => void,
	editable: (feature: DocxCommandFeature) => boolean
): DocxEditorInteractions | undefined {
	const entries: [ObjectHitKind, (press: DocxObjectPress) => boolean][] = []

	for (const kind of Object.keys(OBJECT_FEATURES) as ObjectHitKind[]) {
		const feature = OBJECT_FEATURES[kind]

		if (feature !== null && featureLanded(feature)) {
			entries.push([
				kind,
				press => {
					if (!editable(feature)) {
						return false
					}

					onPress(press)

					return true
				}
			])
		}
	}

	return entries.length === 0 ? undefined : Object.fromEntries(entries)
}

export type StyleLabel =
	| { key: "previewDocxStyleHeading"; level: number }
	| { key: "previewDocxStyleNormal" | "previewDocxStyleTitle" | "previewDocxStyleSubtitle" | "previewDocxStyleQuote" }

// The label of a paragraph style in the picker, with the heading level where it has one.
export function styleLabel(style: DocxParagraphStyle): StyleLabel {
	const heading = /^heading(\d)$/.exec(style)?.[1]

	if (heading !== undefined) {
		return { key: "previewDocxStyleHeading", level: Number(heading) }
	}

	switch (style) {
		case "normal":
			return { key: "previewDocxStyleNormal" }
		case "title":
			return { key: "previewDocxStyleTitle" }
		case "subtitle":
			return { key: "previewDocxStyleSubtitle" }
		default:
			return { key: "previewDocxStyleQuote" }
	}
}

// Whether a command is shown as switched on for the selection (a toggle or a radio among alignments/lists).
export function commandActive(snapshot: DocxEditorSnapshot, type: DocxEditorCommandType): boolean {
	switch (type) {
		case "format.bold":
			return snapshot.formats.bold === "on"
		case "format.italic":
			return snapshot.formats.italic === "on"
		case "format.underline":
			return snapshot.formats.underline === "on"
		case "format.strike":
			return snapshot.formats.strike === "on"
		case "align.start":
			return snapshot.paragraph.align === "start"
		case "align.center":
			return snapshot.paragraph.align === "center"
		case "align.end":
			return snapshot.paragraph.align === "end"
		case "align.justify":
			return snapshot.paragraph.align === "justify"
		case "list.bullet":
			return snapshot.paragraph.list === "bullet"
		case "list.numbered":
			return snapshot.paragraph.list === "numbered"
		default:
			return false
	}
}

// The command's platform-neutral shortcut ("mod+shift+z") as keys to show; Cmd on Apple platforms.
export function shortcutKeys(combo: string, apple: boolean): string[] {
	return combo.split("+").map(key => {
		switch (key) {
			case "mod":
				return apple ? "⌘" : "Ctrl"
			case "shift":
				return apple ? "⇧" : "Shift"
			case "alt":
				return apple ? "⌥" : "Alt"
			case "space":
				return "Space"
			default:
				return key.length === 1 ? key.toUpperCase() : key
		}
	})
}

// Why the editor shows the document view-only: the file's own gate reasons (the viewer's), a text box being
// edited, or a file the overlay can no longer save to.
export function editorViewOnlyReasons(reasons: readonly DocxReadOnlyReason[], writable: boolean, renamed: boolean): DocxViewOnlyReason[] {
	const out: DocxViewOnlyReason[] = []

	if (renamed) {
		out.push("renamed")
	} else if (!writable) {
		out.push("notWritable")
	}

	for (const reason of reasons) {
		if (reason !== "host" && !out.includes(reason)) {
			out.push(reason)
		}
	}

	return out
}

export function lockReasonKey(reason: LockReason): PreviewKey {
	switch (reason) {
		case "contentLocked":
			return "previewDocxLockedContent"
		case "proof":
		case "unmodelled":
		case "opaque":
		case "malformed":
			return "previewDocxLockedProof"
		default:
			return "previewDocxLockedStructure"
	}
}

// The match to make current for a new search: the first at or after the caret (as Word continues from it),
// else the first.
export function firstMatchFrom(matches: readonly StoryMatch[], pos: number): number {
	if (matches.length === 0) {
		return -1
	}

	let low = 0
	let high = matches.length

	while (low < high) {
		const middle = (low + high) >> 1

		if ((matches[middle]?.from ?? 0) < pos) {
			low = middle + 1
		} else {
			high = middle
		}
	}

	return low < matches.length ? low : 0
}

// The next (or previous) match, wrapping around.
export function stepMatch(active: number, count: number, direction: 1 | -1): number {
	if (count === 0) {
		return -1
	}

	if (active < 0) {
		return direction > 0 ? 0 : count - 1
	}

	return (active + direction + count) % count
}

const NO_MARKS: readonly DocxEditorMark[] = []

// Whether two snapshots show the same in the toolbar (formatting, enablement, undo, read-only, the lock at
// the selection, the values, context and review state slots show): the version, the selection's place and the
// dirty flag are not shown there.
export function sameToolbarSnapshot(a: DocxEditorSnapshot | undefined, b: DocxEditorSnapshot | undefined): boolean {
	if (a === b) {
		return true
	}

	if (a === undefined || b === undefined) {
		return false
	}

	return (
		a.readOnly === b.readOnly &&
		a.canUndo === b.canUndo &&
		a.canRedo === b.canRedo &&
		a.selection.locked === b.selection.locked &&
		a.readOnlyReasons.length === b.readOnlyReasons.length &&
		a.readOnlyReasons.every((reason, index) => b.readOnlyReasons[index] === reason) &&
		sameRecord(a.formats, b.formats) &&
		sameRecord(a.paragraph, b.paragraph) &&
		sameRecord(a.enabled, b.enabled) &&
		// The session keeps these objects while their answer is the same.
		a.values === b.values &&
		a.context === b.context &&
		a.review === b.review
	)
}

function sameRecord<T extends object>(a: T, b: T): boolean {
	if (a === b) {
		return true
	}

	const keys = Object.keys(a) as (keyof T)[]

	return keys.length === Object.keys(b).length && keys.every(key => a[key] === b[key])
}

// Spelling marks and find results as one list sorted by `from` (both come sorted), the current match marked
// apart. Find results draw over spelling.
export function editorMarks(spelling: readonly SpellingMark[], matches: readonly StoryMatch[], active: number): readonly DocxEditorMark[] {
	if (matches.length === 0) {
		return spelling.length === 0 ? NO_MARKS : spelling
	}

	const out: DocxEditorMark[] = []
	let s = 0

	for (let k = 0; k < matches.length; k++) {
		const match = matches[k]

		if (match === undefined) {
			continue
		}

		while (s < spelling.length && (spelling[s]?.from ?? 0) <= match.from) {
			const mark = spelling[s]

			if (mark !== undefined) {
				out.push(mark)
			}

			s++
		}

		out.push({ from: match.from, to: match.to, kind: k === active ? "find-active" : "find" })
	}

	for (; s < spelling.length; s++) {
		const mark = spelling[s]

		if (mark !== undefined) {
			out.push(mark)
		}
	}

	return out
}

// Groups that leave the row for the ⋯ menu, last first, as the row narrows: each needs this much width
// (px) for itself to stay. History and character formatting always stay.
const BUTTON_WIDTH = 34
const STYLE_WIDTH = 150
const LIST_BUTTONS = 4
const ALIGN_BUTTONS = 4
// History, character formatting, find, the menu and the zoom percentage.
const RESERVED_WIDTH = 7 * BUTTON_WIDTH + 2 * BUTTON_WIDTH + 64

export interface ToolbarFit {
	style: boolean
	lists: boolean
	align: boolean
	clear: boolean
}

// Which groups the row has room for at `width` (px), filled in priority order.
export function toolbarFit(width: number, extra: number): ToolbarFit {
	let room = width - RESERVED_WIDTH - extra
	const take = (need: number): boolean => {
		if (room < need) {
			room = -1

			return false
		}

		room -= need

		return true
	}

	const style = take(STYLE_WIDTH)
	const lists = take(LIST_BUTTONS * BUTTON_WIDTH + 8)
	const align = take(ALIGN_BUTTONS * BUTTON_WIDTH + 8)
	const clear = take(BUTTON_WIDTH + 8)

	return { style, lists, align, clear }
}

type MarkupView = DocxMarkupInfo["view"]

// How the open document is laid out (the editor's view is what the pages show).
export function editorMarkupView(view: RevisionView): MarkupView {
	return view === RevisionView.Markup ? "markup" : view === RevisionView.Original ? "original" : "final"
}

// The mode the markup menu shows: the user's, while the pages are laid out as it needs; else the default mode of
// the layout the editor shows (a view another surface or command chose).
export function shownMarkup(mode: DocxMarkupMode, view: MarkupView): DocxMarkupMode {
	if (markupInfo(mode).view === view) {
		return mode
	}

	return view === "markup" ? "all" : view === "original" ? "original" : DEFAULT_MARKUP_MODE
}

// The layout switching to `next` needs while the pages show `view`: null when they already do (only what is
// drawn over them changes).
export function markupRelayout(view: MarkupView, next: DocxMarkupMode): MarkupView | null {
	const needed = markupInfo(next).view

	return needed === view ? null : needed
}
