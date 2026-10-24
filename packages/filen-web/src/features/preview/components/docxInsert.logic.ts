import type { DocxEditorCommandType, DocxEditorSnapshot } from "@filen/office/core/docx/edit"
import { useRef } from "react"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import type { DocxObjectPress } from "@filen/office/react/docx/editor"
import type { PreviewKey } from "@/lib/i18n"

// What the Insert menu's dialogs offer and how they read what is typed: link addresses, date pictures,
// symbols, the field catalogue's names; and what the chip at a field offers.

// The commands whose input the insert dialogs take (from the Insert menu, the context menu, the Picture group).
export const INSERT_DIALOG_INPUTS: ReadonlySet<DocxEditorCommandType> = new Set([
	"link.set",
	"insert.symbol",
	"insert.date",
	"toc.insert",
	"insert.field",
	"bookmark.add",
	"bookmark.rename",
	"bookmark.delete",
	"note.options",
	"table.insert",
	"drawing.altText"
])

// A scheme the address names itself: one with an authority ("https://", "ftp://"), or one of those written
// without ("mailto:", "tel:"). "host:port" is no scheme.
const AUTHORITY = /^[a-z][a-z0-9+.-]*:\/\//i
const OPAQUE = /^(mailto|tel|callto|sms|news|urn):/i
const DRIVE_PATH = /^[a-z]:[\\/]/i
const SHARE_PATH = /^\\\\[^\\]+\\/
const EMAIL = /^[^\s@/:]+@[^\s@/]+\.[^\s@/]+$/

// A typed address as a link target: as written with a scheme, an email address as mailto:, anything else as
// https:// (Word adds the scheme the same way), a Windows path as a file: URL. Null when there is nothing to link to.
export function linkTarget(address: string): string | null {
	const trimmed = address.trim()

	// A path on a drive, as Word writes one: file:/// and the path with its backslashes, only spaces and
	// non-ASCII characters percent-encoded.
	if (DRIVE_PATH.test(trimmed)) {
		return `file:///${trimmed.replace(/\s|[\u0080-\u{10ffff}]/gu, char => encodeURIComponent(char))}`
	}

	// A path on a share: a file: URL.
	if (SHARE_PATH.test(trimmed)) {
		return encodeURI(`file:${trimmed.replace(/\\/g, "/")}`)
	}

	if (trimmed === "" || /\s/.test(trimmed)) {
		return null
	}

	if (AUTHORITY.test(trimmed) || OPAQUE.test(trimmed)) {
		return trimmed
	}

	return EMAIL.test(trimmed) ? `mailto:${trimmed}` : `https://${trimmed}`
}

// The language of the text at the caret (its w:lang, Latin script), as a DATE field there reads it.
export function caretLanguage(editor: PagedEditor): string | undefined {
	const session = editor.session
	const { $from } = session.state.selection

	if (!$from.parent.isTextblock) {
		return undefined
	}

	const props = session.formats.runProps(session.state.storedMarks ?? $from.marks())

	return session.context.resolver.run(session.formats.paragraph($from.parent), props).lang.latin ?? undefined
}

// Characters the symbol dialog names (the grid shows the rest as they are).
export const NAMED_SYMBOLS: Readonly<Record<string, PreviewKey>> = {
	" ": "previewDocxSymbolNonBreakingSpace",
	"‑": "previewDocxSymbolNonBreakingHyphen",
	"­": "previewDocxSymbolOptionalHyphen"
}

// The symbol dialog's grid: Word's special characters, then common symbols, arrows, math and Greek.
export const SYMBOLS: readonly string[] = [
	" ",
	"‑",
	"­",
	"—",
	"–",
	"…",
	"©",
	"®",
	"™",
	"§",
	"¶",
	"†",
	"‡",
	"•",
	"·",
	"°",
	"‘",
	"’",
	"“",
	"”",
	"«",
	"»",
	"€",
	"£",
	"¥",
	"¢",
	"±",
	"×",
	"÷",
	"≈",
	"≠",
	"≤",
	"≥",
	"∞",
	"√",
	"∑",
	"∫",
	"∂",
	"∆",
	"µ",
	"π",
	"Ω",
	"α",
	"β",
	"γ",
	"δ",
	"θ",
	"λ",
	"σ",
	"φ",
	"←",
	"↑",
	"→",
	"↓",
	"↔",
	"⇒",
	"✓",
	"✗",
	"★",
	"☆",
	"♥",
	"☐",
	"☑",
	"½",
	"¼",
	"¾"
]

// The field catalogue's types by name (FIELD_CATALOGUE).
const FIELD_NAMES: Readonly<Record<string, PreviewKey>> = {
	PAGE: "previewDocxFieldPage",
	NUMPAGES: "previewDocxFieldNumPages",
	SECTIONPAGES: "previewDocxFieldSectionPages",
	SECTION: "previewDocxFieldSection",
	DATE: "previewDocxFieldDate",
	TIME: "previewDocxFieldTime",
	REF: "previewDocxFieldRef",
	PAGEREF: "previewDocxFieldPageRef",
	SEQ: "previewDocxFieldSeq"
}

export function fieldNameKey(type: string): PreviewKey | undefined {
	return FIELD_NAMES[type]
}

// A catalogue field's example code naming `bookmark` where it names one (REF, PAGEREF).
export function fieldExample(example: string, bookmark: string | undefined): string {
	return bookmark === undefined ? example : example.replace("_Ref1", bookmark)
}

// The Insert ▸ Table grid: columns × rows.
export const TABLE_GRID = { cols: 10, rows: 8 } as const

// ---- field chip ------------------------------------------------------------------------------------------

export type DocxFieldChipAction = "field" | "toc" | "note"

// What the chip offers for the press with the toolbar's snapshot, or null.
export function fieldChipAction(
	press: DocxObjectPress | null,
	snapshot: DocxEditorSnapshot | undefined,
	readOnly: boolean
): DocxFieldChipAction | null {
	if (press === null || snapshot === undefined) {
		return null
	}

	if (press.hit.kind === "noteRef") {
		return snapshot.enabled["note.goto"] ? "note" : null
	}

	if (press.hit.kind !== "field" || readOnly || snapshot.readOnly) {
		return null
	}

	const field = snapshot.context.field

	if (press.hit.part === "toc") {
		return field?.toc === true && field.refusal === null && snapshot.enabled["toc.update"] ? "toc" : null
	}

	return field?.updatable === true && snapshot.enabled["field.update"] ? "field" : null
}

// A menu whose items run commands: the host gives the document the focus as one runs, so the menu gives it back
// there when it closes after one (its trigger gets it otherwise). `ran` after running; `kept` after an item
// that leaves focus alone (a setting, the zoom): focus goes back to where it was when the menu opened (the
// document, as toolbar triggers keep it on a press). `opened` to the menu's onOpenChange and `finalFocus` to
// its popup. An item that replaces the document's view (Mobile view) runs `afterClose`, with `closed` on the
// menu's onOpenChangeComplete: the document takes focus back first, so the view shown in its place keeps it.
// A select gives focus back to nothing once an option is picked (it is left on the page): `settled` on its
// onOpenChangeComplete gives it to the document.
export function useCommandMenuFocus(): {
	ran: () => void
	kept: () => void
	afterClose: (action: () => void) => void
	opened: (open: boolean) => void
	finalFocus: () => HTMLElement | boolean
	closed: (open: boolean) => void
	settled: (open: boolean) => void
} {
	const focusedRef = useRef<HTMLElement | null>(null)
	const beforeRef = useRef<HTMLElement | null>(null)
	const pendingRef = useRef<(() => void) | null>(null)
	// The pending action ran: focus is the new view's.
	const handedRef = useRef(false)

	return {
		ran: () => {
			focusedRef.current = activeElement()
		},
		kept: () => {
			focusedRef.current = beforeRef.current
		},
		afterClose: action => {
			focusedRef.current = beforeRef.current
			pendingRef.current = action
		},
		opened: open => {
			if (open) {
				beforeRef.current = activeElement()
				focusedRef.current = null
				pendingRef.current = null
				handedRef.current = false
			}
		},
		finalFocus: () => {
			const focused = focusedRef.current

			focusedRef.current = null

			if (handedRef.current) {
				handedRef.current = false

				return false
			}

			return focused?.isConnected === true ? focused : true
		},
		closed: open => {
			const pending = pendingRef.current

			if (open || pending === null) {
				return
			}

			pendingRef.current = null
			handedRef.current = true

			const before = beforeRef.current

			if (before?.isConnected === true) {
				before.focus({ preventScroll: true })
			}

			pending()
		},
		settled: open => {
			const focused = focusedRef.current

			focusedRef.current = null

			if (!open && focused?.isConnected === true) {
				focused.focus({ preventScroll: true })
			}
		}
	}
}

function activeElement(): HTMLElement | null {
	return document.activeElement instanceof HTMLElement ? document.activeElement : null
}
