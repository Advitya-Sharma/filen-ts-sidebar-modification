// @vitest-environment happy-dom

import { afterEach, describe, expect, it, vi } from "vitest"
import { cleanup, fireEvent, render, screen, waitFor, within } from "@testing-library/react"
import {
	DOCX_EDITOR_COMMANDS,
	NO_REVIEW,
	NO_VALUES,
	type DocxEditorSnapshot,
	type DocxFontChoice,
	type DocxTableContext
} from "@filen/office/core/docx/edit"
import type { DocxEditToolbarProps } from "@/features/preview/components/docxEditToolbar"
import { slotsWidth, themeSwatches, toolbarFit } from "@/features/preview/components/docxEditor.logic"

// The editing toolbar on its own: the paragraph styles protection leaves, the formatting controls showing the
// selection's values (one, several, none) and turning off with their commands, the colour pickers and the
// format painter's clicks.

const size = vi.hoisted(() => ({ width: 2000 }))

vi.mock("@/lib/useElementSize", () => ({ useElementSize: () => ({ width: size.width, height: 40 }) }))

const { DocxEditToolbar } = await import("@/features/preview/components/docxEditToolbar")

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

afterEach(() => {
	cleanup()
	size.width = 2000
})

describe("DocxEditToolbar", () => {
	it("offers only the paragraph styles the document's protection leaves open", async () => {
		renderToolbar()

		fireEvent.click(screen.getByRole("combobox", { name: "Paragraph style" }))

		const options = (await screen.findAllByRole("option")).map(option => option.textContent)

		expect(options).toEqual(["Normal", "Heading 1", "Heading 2"])
		expect(options).not.toContain("Title")
	})

	it("shows the selection's font and size, and nothing for several", () => {
		renderToolbar()

		expect(screen.getByRole("combobox", { name: "Font" }).textContent).toContain("Georgia")
		expect(screen.getByRole("combobox", { name: "Font size" }).textContent).toContain("11")
		cleanup()
		renderToolbar({ snapshot: snapshot({ values: { ...NO_VALUES, font: "mixed", size: "mixed" } }) })
		expect(screen.getByRole("combobox", { name: "Font" }).textContent).toContain("Several")
		expect(screen.getByRole("combobox", { name: "Font size" }).textContent).not.toMatch(/\d/)
		cleanup()
		// The theme's body font, by its role.
		renderToolbar({ snapshot: snapshot({ values: { ...NO_VALUES, font: "Calibri", size: 21 } }) })
		expect(screen.getByRole("combobox", { name: "Font" }).textContent).toContain("Calibri (Body)")
		expect(screen.getByRole("combobox", { name: "Font size" }).textContent).toMatch(/10[.,]5/)
	})

	it("turns each control off with its command (direct formatting under formatting restrictions), and when read-only", () => {
		const enabled = snapshot().enabled

		renderToolbar({ snapshot: snapshot({ enabled: { ...enabled, "font.family": false, "font.size": false, "color.text": false } }) })
		expect(screen.getByRole("combobox", { name: "Font" }).hasAttribute("data-disabled")).toBe(true)
		expect(screen.getByRole("combobox", { name: "Font size" }).hasAttribute("data-disabled")).toBe(true)
		expect(screen.getByRole("button", { name: "Font color" }).hasAttribute("disabled")).toBe(true)
		expect(screen.getByRole("button", { name: "Highlight color" }).hasAttribute("disabled")).toBe(false)
		cleanup()
		renderToolbar({ readOnly: true })
		expect(screen.getByRole("button", { name: "Highlight color" }).hasAttribute("disabled")).toBe(true)
		expect(screen.getByRole("button", { name: "Format painter" }).hasAttribute("disabled")).toBe(true)
	})

	it("sets a theme, standard or automatic text colour and a highlight from the shared pickers", async () => {
		const { execute } = renderToolbar()

		fireEvent.click(screen.getByRole("button", { name: "Font color" }))
		fireEvent.click(await screen.findByRole("button", { name: "Accent 1, lighter 60%" }))
		expect(execute).toHaveBeenLastCalledWith({ type: "color.text", color: { kind: "theme", theme: "accent1", tint: 102 } })

		fireEvent.click(screen.getByRole("button", { name: "Font color" }))
		fireEvent.click(await screen.findByRole("button", { name: "Dark red" }))
		expect(execute).toHaveBeenLastCalledWith({ type: "color.text", color: { kind: "rgb", value: "C00000" } })

		fireEvent.click(screen.getByRole("button", { name: "Font color" }))
		// Automatic is the colour in effect.
		const automatic = await screen.findByRole("button", { name: "Automatic" })

		fireEvent.click(automatic)
		expect(execute).toHaveBeenLastCalledWith({ type: "color.text", color: { kind: "auto" } })

		fireEvent.click(screen.getByRole("button", { name: "Highlight color" }))
		fireEvent.click(await screen.findByRole("button", { name: "Turquoise" }))
		expect(execute).toHaveBeenLastCalledWith({ type: "color.highlight", name: "cyan" })
	})

	it("picks up formatting with a click, keeps it with a double click, and a click while on turns it off", () => {
		const { execute, onPainterCancel } = renderToolbar()
		const painter = screen.getByRole("button", { name: "Format painter" })

		fireEvent.click(painter, { detail: 1 })
		expect(execute).toHaveBeenLastCalledWith({ type: "format.copy" })
		fireEvent.click(painter, { detail: 2 })
		expect(execute).toHaveBeenLastCalledWith({ type: "format.copy", sticky: true })
		cleanup()

		const on = renderToolbar({ painter: "sticky" })

		expect(screen.getByRole("button", { name: "Format painter (on until clicked again)" }).getAttribute("aria-pressed")).toBe("true")
		fireEvent.click(screen.getByRole("button", { name: "Format painter (on until clicked again)" }), { detail: 1 })
		expect(on.onPainterCancel).toHaveBeenCalledTimes(1)
		expect(on.execute).not.toHaveBeenCalled()
		expect(onPainterCancel).not.toHaveBeenCalled()
	})

	it("moves the font and colour pickers into the formatting menu when the row is narrow", async () => {
		size.width = 700

		const { execute } = renderToolbar({ snapshot: snapshot({ values: { ...NO_VALUES, vertAlign: "superscript", font: "Arial" } }) })

		expect(screen.queryByRole("combobox", { name: "Font" })).toBeNull()
		fireEvent.click(screen.getByRole("button", { name: "More formatting" }))

		const superscript = await screen.findByRole("menuitemcheckbox", { name: "Superscript" })

		expect(superscript.getAttribute("aria-checked")).toBe("true")
		expect(screen.getByRole("menuitem", { name: "Font" })).toBeTruthy()
		expect(screen.getByRole("menuitem", { name: "Font color" })).toBeTruthy()
		fireEvent.click(screen.getByRole("menuitemcheckbox", { name: "Small caps" }))
		expect(execute).toHaveBeenLastCalledWith({ type: "format.smallCaps" })
		// The menu closes with the format applied, in every browser (not by focus leaving it).
		await waitFor(() => {
			expect(screen.queryByRole("menuitemcheckbox", { name: "Small caps" })).toBeNull()
		})
	})

	it("moves groups into the ⋯ menu to make room for the Picture group, so the row never overflows", () => {
		const drawing = { kind: "picture", resizable: true, rotation: 0 } as unknown as NonNullable<
			DocxEditorSnapshot["context"]["drawing"]
		>
		const menus = slotsWidth({ insert: true, review: true, table: false })
		let width = 300

		while (!toolbarFit(width, menus, true, 2).lists) {
			width += 10
		}

		size.width = width
		renderToolbar()
		expect(screen.getByRole("button", { name: "Bulleted list" })).toBeTruthy()
		cleanup()

		const plain = snapshot()

		renderToolbar({ snapshot: { ...plain, context: { ...plain.context, drawing } } })
		expect(screen.queryByRole("button", { name: "Bulleted list" })).toBeNull()
	})

	it("keeps the ⋯ menu in the row at a phone's width (360 and 390 px), what does not fit in it", async () => {
		// A phone's window: the slot menus show as icons.
		const media = vi.spyOn(window, "matchMedia").mockImplementation(
			query =>
				({
					matches: false,
					media: query,
					addEventListener: () => undefined,
					removeEventListener: () => undefined
				}) as unknown as MediaQueryList
		)
		const plain = snapshot()
		const table: DocxTableContext = {
			rows: 2,
			columns: 2,
			fromRow: 0,
			toRow: 0,
			fromColumn: 0,
			toColumn: 0,
			mergeable: false,
			splittable: false,
			headerRow: false,
			styleId: null,
			depth: 0,
			widths: [2000, 2000],
			rtl: false,
			align: "start",
			vAlign: null,
			shading: null,
			styles: []
		}
		const inTable = { ...plain, context: { ...plain.context, table } }

		try {
			for (const window of [360, 390]) {
				for (const inside of [false, true]) {
					// The row's width inside its padding.
					size.width = window - 16

					const onFindOpen = vi.fn()
					const { execute } = renderToolbar({ snapshot: inside ? inTable : plain, onFindOpen })
					const toolbar = screen.getByRole("toolbar", { name: "Formatting" })
					const more = screen.getByRole("button", { name: "More tools" })

					// Never inside the groups, which scroll when they still overflow.
					expect(more.parentElement).toBe(toolbar)
					expect(screen.getByRole("button", { name: "Bold" })).toBeTruthy()

					const strikeShown = !(window === 360 && inside)
					const findShown = window === 390 && !inside

					expect(screen.queryByRole("button", { name: "Strikethrough" }) !== null).toBe(strikeShown)
					expect(screen.queryByRole("button", { name: "Find in document" }) !== null).toBe(findShown)
					expect(screen.queryByRole("button", { name: "Format painter" })).toBeNull()

					fireEvent.click(more)

					if (strikeShown) {
						expect(screen.queryByRole("menuitem", { name: "Strikethrough" })).toBeNull()
					} else {
						fireEvent.click(await screen.findByRole("menuitem", { name: "Strikethrough" }))
						expect(execute).toHaveBeenLastCalledWith({ type: "format.strike" })
						fireEvent.click(more)
					}

					if (findShown) {
						expect(screen.queryByRole("menuitem", { name: "Find in document" })).toBeNull()
					} else {
						fireEvent.click(await screen.findByRole("menuitem", { name: "Find in document" }))
						expect(onFindOpen).toHaveBeenLastCalledWith(false)
					}

					// The format painter, from the formatting menu.
					fireEvent.click(screen.getByRole("button", { name: "More formatting" }))
					fireEvent.click(await screen.findByRole("menuitemcheckbox", { name: "Format painter" }))
					expect(execute).toHaveBeenLastCalledWith({ type: "format.copy" })
					cleanup()
				}
			}
		} finally {
			media.mockRestore()
		}
	})

	it("keeps the Picture group's first tools in a phone's row and the rest in the ⋯ menu", async () => {
		const media = vi.spyOn(window, "matchMedia").mockImplementation(
			query =>
				({
					matches: false,
					media: query,
					addEventListener: () => undefined,
					removeEventListener: () => undefined
				}) as unknown as MediaQueryList
		)
		const plain = snapshot()
		const drawing = {
			kind: "picture",
			anchored: false,
			wrap: "inline",
			pos: 5,
			rotation: 0,
			resizable: true,
			rotatable: true,
			placeable: true,
			croppable: true
		} as unknown as NonNullable<DocxEditorSnapshot["context"]["drawing"]>

		try {
			// A 360 px window: Wrap text and Rotate stay.
			size.width = 360 - 16

			const { execute, requestInput } = renderToolbar({ snapshot: { ...plain, context: { ...plain.context, drawing } } })
			const group = screen.getByRole("group", { name: "Picture" })

			expect(within(group).getByRole("button", { name: "Wrap text" })).toBeTruthy()
			expect(within(group).getByRole("button", { name: "Rotate" })).toBeTruthy()
			expect(within(group).queryByRole("button", { name: "Delete picture" })).toBeNull()
			expect(screen.queryByRole("button", { name: "Alt text…" })).toBeNull()

			fireEvent.click(screen.getByRole("button", { name: "More tools" }))
			expect(await screen.findByRole("menuitemcheckbox", { name: "Crop" })).toBeTruthy()
			fireEvent.click(screen.getByRole("menuitem", { name: "Alt text…" }))
			expect(requestInput).toHaveBeenLastCalledWith("drawing.altText")
			fireEvent.click(screen.getByRole("button", { name: "More tools" }))
			fireEvent.click(await screen.findByRole("menuitem", { name: "Delete picture" }))
			expect(execute).toHaveBeenLastCalledWith({ type: "drawing.delete" })
		} finally {
			media.mockRestore()
		}
	})

	it("closes the Review menu once track changes is switched", async () => {
		const { execute } = renderToolbar()

		fireEvent.click(screen.getByRole("button", { name: "Review" }))
		fireEvent.click(await screen.findByRole("menuitemcheckbox", { name: "Track changes" }))
		expect(execute).toHaveBeenLastCalledWith({ type: "review.track", on: true })
		await waitFor(() => {
			expect(screen.queryByRole("menuitemcheckbox", { name: "Track changes" })).toBeNull()
		})
	})

	it("closes the ⋯ menu once a command it holds as a checked item runs", async () => {
		// Narrow: the alignment is in the ⋯ menu, "Align left" checked.
		size.width = 300

		const { execute } = renderToolbar()

		fireEvent.click(screen.getByRole("button", { name: "More tools" }))
		fireEvent.click(await screen.findByRole("menuitemcheckbox", { name: "Align left" }))
		expect(execute).toHaveBeenCalledTimes(1)
		await waitFor(() => {
			expect(screen.queryByRole("menuitemcheckbox", { name: "Align left" })).toBeNull()
		})
	})

	it("offers no header or footer editing in Mobile view, which shows neither", async () => {
		for (const surface of ["pages", "reflow"] as const) {
			renderToolbar({ surface, reflow: surface === "reflow" })
			fireEvent.click(screen.getByRole("button", { name: "Insert" }))
			await screen.findByRole("menu")
			expect(screen.queryByRole("menuitem", { name: "Header" }) !== null).toBe(surface === "pages")
			expect(screen.queryByRole("menuitem", { name: "Footer" }) !== null).toBe(surface === "pages")
			cleanup()
		}
	})

	it("opens the Paragraph and Page setup dialogs from the ⋯ menu and switches field codes", async () => {
		const { execute, requestInput } = renderToolbar()

		fireEvent.click(screen.getByRole("button", { name: "More tools" }))
		fireEvent.click(await screen.findByRole("menuitem", { name: "Page setup…" }))
		expect(requestInput).toHaveBeenLastCalledWith("page.setup")

		fireEvent.click(screen.getByRole("button", { name: "More tools" }))
		fireEvent.click(await screen.findByRole("menuitem", { name: "Paragraph…" }))
		expect(requestInput).toHaveBeenLastCalledWith("paragraph.indent")

		fireEvent.click(screen.getByRole("button", { name: "More tools" }))
		fireEvent.click(await screen.findByRole("menuitemcheckbox", { name: "Show field codes" }))
		expect(execute).toHaveBeenLastCalledWith({ type: "view.fieldCodes" })
		// A toggle closes the menu, as a command does.
		await waitFor(() => {
			expect(screen.queryByRole("menuitemcheckbox", { name: "Show field codes" })).toBeNull()
		})
	})
})
