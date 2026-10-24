// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from "vitest"
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react"
import {
	DOCX_EDITOR_COMMANDS,
	NO_REVIEW,
	NO_VALUES,
	type DocxEditorSnapshot,
	type DocxFontChoice,
	type DocxTableContext
} from "@filen/office/core/docx/edit"
import type { DocxEditToolbarProps } from "@/features/preview/components/docxEditToolbar"
import { themeSwatches } from "@/features/preview/components/docxEditor.logic"

// After a toolbar menu or select, focus is back in the document, so what is typed next lands there: after a
// command (the host focuses the document as it runs), after a setting or the zoom (where it was when the
// menu opened), and in the view shown after a switch to Mobile view.

const size = vi.hoisted(() => ({ width: 2000 }))

vi.mock("@/lib/useElementSize", () => ({ useElementSize: () => ({ width: size.width, height: 40 }) }))

const { DocxEditToolbar } = await import("@/features/preview/components/docxEditToolbar")

const TABLE: DocxTableContext = {
	rows: 2,
	columns: 3,
	fromRow: 0,
	toRow: 1,
	fromColumn: 1,
	toColumn: 2,
	mergeable: false,
	splittable: true,
	headerRow: true,
	styleId: "TableGrid",
	depth: 0,
	widths: [2000, 2000, 2000],
	rtl: false,
	align: "center",
	vAlign: null,
	shading: "FF0000",
	styles: [{ id: "TableGrid", name: "Table Grid" }]
}

const FONTS: readonly DocxFontChoice[] = [
	{ family: "+headings", name: "Calibri Light", kind: "theme", role: "headings" },
	{ family: "+body", name: "Calibri", kind: "theme", role: "body" },
	{ family: "Georgia", name: "Georgia", kind: "document" },
	{ family: "Arial", name: "Arial", kind: "common" }
]
const THEME: Record<string, string> = {
	background1: "FFFFFF",
	text1: "000000",
	background2: "E7E6E6",
	text2: "44546A",
	accent1: "4472C4",
	accent2: "ED7D31",
	accent3: "A5A5A5",
	accent4: "FFC000",
	accent5: "5B9BD5",
	accent6: "70AD47"
}

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
		formats: { bold: "off", italic: "off", underline: "off", strike: "off" },
		paragraph: { style: "normal", styleName: null, align: "start", list: "none", level: null },
		values: { ...NO_VALUES, font: "Georgia", size: 22, color: "auto", colorRgb: "auto" },
		context: {
			story: { kind: "body", index: 0 },
			table: null,
			drawing: null,
			control: null,
			field: null,
			link: null,
			comment: null,
			equation: null
		},
		review: NO_REVIEW,
		enabled,
		...overrides
	}
}

function renderToolbar(props: Partial<DocxEditToolbarProps> = {}) {
	const execute = vi.fn()
	const onPainterCancel = vi.fn()
	const requestInput = vi.fn()

	render(
		<DocxEditToolbar
			snapshot={snapshot()}
			readOnly={false}
			apple={false}
			execute={execute}
			viewOnly={[]}
			find={{
				open: false,
				replace: false,
				focus: 0,
				controls: {
					query: "",
					caseSensitive: false,
					count: 0,
					capped: false,
					active: -1,
					search: vi.fn(),
					step: vi.fn(),
					close: vi.fn(),
					replace: { disabled: true, one: vi.fn(), all: vi.fn() }
				}
			}}
			onFindOpen={vi.fn()}
			onFindClosed={vi.fn()}
			markup={null}
			onMarkupChange={vi.fn()}
			zoom={1}
			onZoom={vi.fn()}
			autosave
			onAutosaveChange={vi.fn()}
			reflow={false}
			onReflowChange={vi.fn()}
			surface="pages"
			styles={["normal", "heading1", "heading2"]}
			fonts={FONTS}
			theme={themeSwatches(name => THEME[name])}
			painter="off"
			onPainterCancel={onPainterCancel}
			fieldCodes={false}
			inputRequest={null}
			requestInput={requestInput}
			onInputDone={vi.fn()}
			{...props}
		/>
	)

	return { execute, onPainterCancel, requestInput }
}

// The document's editing area, focused when the test starts and by every command (as the host does).
let documentArea: HTMLTextAreaElement

function renderFocused(props: Partial<DocxEditToolbarProps> = {}) {
	documentArea = document.createElement("textarea")
	document.body.append(documentArea)
	documentArea.focus()

	const result = renderToolbar(props)

	result.execute.mockImplementation(() => {
		documentArea.focus()
	})

	return result
}

async function closed(): Promise<void> {
	await waitFor(() => {
		expect(screen.queryByRole("menu")).toBeNull()
		expect(screen.queryByRole("listbox")).toBeNull()
	})
	await new Promise(resolve => setTimeout(resolve, 0))
}

// Picks an option of a select the way a pointer does.
async function choose(label: string, option: string): Promise<void> {
	const trigger = screen.getByRole("combobox", { name: label })

	fireEvent.pointerDown(trigger, { pointerType: "mouse", button: 0 })
	fireEvent.mouseDown(trigger, { button: 0 })
	fireEvent.click(trigger)

	const item = await screen.findByRole("option", { name: option })

	fireEvent.pointerDown(item, { pointerType: "mouse", button: 0 })
	fireEvent.mouseUp(item, { button: 0 })
	fireEvent.click(item)
}

// Opens `menu` and clicks through `path`: submenus by name, the last item as [role, name].
async function pick(menu: string | RegExp, path: readonly string[], [role, name]: readonly [string, string]): Promise<void> {
	fireEvent.click(screen.getByRole("button", { name: menu }))

	for (const submenu of path) fireEvent.click(await screen.findByRole("menuitem", { name: submenu }))

	fireEvent.click(await screen.findByRole(role, { name }))
}

afterEach(() => {
	cleanup()
	documentArea.remove()
	size.width = 2000
})

describe("Focus after the editing toolbar's menus and selects", () => {
	it("is back in the document after More formatting ▸ Superscript", async () => {
		const { execute } = renderFocused()

		await pick("More formatting", [], ["menuitemcheckbox", "Superscript"])
		expect(execute).toHaveBeenCalledWith({ type: "format.superscript" })
		await closed()
		expect(document.activeElement).toBe(documentArea)
	})

	it("is back in the document after a font size is picked", async () => {
		const { execute } = renderFocused()

		await choose("Font size", "14")
		expect(execute).toHaveBeenCalledWith({ type: "font.size", halfPoints: 28 })
		await closed()
		expect(document.activeElement).toBe(documentArea)
	})

	it("is back in the document after a paragraph style is picked", async () => {
		const { execute } = renderFocused()

		await choose("Paragraph style", "Heading 1")
		expect(execute).toHaveBeenCalledWith({ type: "style.set", style: "heading1" })
		await closed()
		expect(document.activeElement).toBe(documentArea)
	})

	it("is back in the document after Table ▸ Insert row below", async () => {
		const base = snapshot()
		const { execute } = renderFocused({ snapshot: { ...base, context: { ...base.context, table: TABLE } } })

		await pick(/^Table/, [], ["menuitem", "Insert row below"])
		expect(execute).toHaveBeenCalledWith({ type: "table.row.insertBelow" })
		await closed()
		expect(document.activeElement).toBe(documentArea)
	})

	it("is back in the document after More tools ▸ Show field codes", async () => {
		const { execute } = renderFocused()

		await pick("More tools", [], ["menuitemcheckbox", "Show field codes"])
		expect(execute).toHaveBeenCalledWith({ type: "view.fieldCodes" })
		await closed()
		expect(document.activeElement).toBe(documentArea)
	})

	it("goes back to the document after More tools ▸ Zoom ▸ Fit to width, which runs no command", async () => {
		const onZoom = vi.fn()

		renderFocused({ onZoom })
		await pick("More tools", ["Zoom"], ["menuitem", "Fit to width"])
		expect(onZoom).toHaveBeenCalledWith("fit")
		await closed()
		expect(document.activeElement).toBe(documentArea)
	})

	it("switches to Mobile view once the menu has given focus back to the document", async () => {
		const focusedAtSwitch: (Element | null)[] = []
		const onReflowChange = vi.fn(() => {
			focusedAtSwitch.push(document.activeElement)
		})

		renderFocused({ onReflowChange })
		await pick("More tools", [], ["menuitemcheckbox", "Mobile view"])
		await waitFor(() => {
			expect(onReflowChange).toHaveBeenCalledWith(true)
		})
		expect(focusedAtSwitch).toEqual([documentArea])
	})
})
