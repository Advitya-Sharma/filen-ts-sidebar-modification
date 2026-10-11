import type {
	DocxEditorCommandType,
	DocxEditorSnapshot,
	DocxParagraphStyle,
	DocxReadOnlyReason,
	LockReason,
	StoryMatch
} from "@filen/office/core/docx/edit"
import { markupInfo, type DocxMarkupInfo, type DocxMarkupMode } from "@filen/office/dom/docx"
import type { DocxEditorMark } from "@filen/office/react/docx/editor"
import type { SpellingMark } from "@filen/office/core/docx/spelling"
import type { DocxViewOnlyReason } from "@/features/preview/components/docxViewer.logic"
import type { PreviewKey } from "@/lib/i18n"

// Idle autosave: this long after the last edit while there are unsaved changes (one version per pause).
export const AUTOSAVE_IDLE_MS = 30_000

// The toolbar's label for each editor command (the package's labelKey is a neutral id an app maps).
export const COMMAND_LABELS: Readonly<Record<DocxEditorCommandType, PreviewKey>> = {
	"format.bold": "previewDocxBold",
	"format.italic": "previewDocxItalic",
	"format.underline": "previewDocxUnderline",
	"format.strike": "previewDocxStrikethrough",
	"format.clear": "previewDocxClearFormatting",
	"style.set": "previewDocxParagraphStyle",
	"align.start": "previewDocxAlignStart",
	"align.center": "previewDocxAlignCenter",
	"align.end": "previewDocxAlignEnd",
	"align.justify": "previewDocxAlignJustify",
	"list.bullet": "previewDocxBulletList",
	"list.numbered": "previewDocxNumberedList",
	"list.indent": "previewDocxIndent",
	"list.outdent": "previewDocxOutdent",
	"history.undo": "previewDocxUndo",
	"history.redo": "previewDocxRedo"
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

// Switching the markup menu to `next` while the document is open laid out as `view`: a mode with the same view
// keeps the open document (only what is drawn over it changes); another view opens it again, which only an
// unchanged document can. Null when it cannot switch now.
export function markupSwitch(view: DocxMarkupInfo["view"], next: DocxMarkupMode, unchanged: boolean): { reopen: boolean } | null {
	if (markupInfo(next).view === view) {
		return { reopen: false }
	}

	return unchanged ? { reopen: true } : null
}
