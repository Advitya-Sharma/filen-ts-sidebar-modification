// @vitest-environment happy-dom

import { afterEach, describe, expect, it, vi } from "vitest"
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react"
import { DOCX_EDITOR_COMMANDS, NO_REVIEW, NO_VALUES, type DocxControlContext, type DocxEditorSnapshot } from "@filen/office/core/docx/edit"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import type { DocxObjectPress } from "@filen/office/react/docx/editor"
import { DocxFormControls } from "@/features/preview/components/docxFormControls"
import "@/lib/i18n"

// Content controls filled in where they are: a checkbox ticked by a press, a dropdown's items and a date
// picker in a popover at the control, why a control cannot change, and a new dropdown list's items.

function control(overrides: Partial<DocxControlContext> = {}): DocxControlContext {
	return {
		kind: "control",
		type: "dropDownList",
		title: "Colour",
		tag: null,
		checked: null,
		items: [
			{ text: "Red", value: "r" },
			{ text: "Green", value: "g" }
		],
		value: "r",
		date: null,
		dateFormat: null,
		placeholder: false,
		locked: null,
		undeletable: false,
		protection: "none",
		textType: null,
		maxLength: null,
		...overrides
	}
}

function snapshot(at: DocxControlContext | null): DocxEditorSnapshot {
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
		values: NO_VALUES,
		context: {
			story: { kind: "body", index: 0 },
			table: null,
			drawing: null,
			control: at,
			field: null,
			link: null,
			comment: null,
			equation: null
		},
		review: NO_REVIEW,
		enabled: Object.fromEntries(DOCX_EDITOR_COMMANDS.map(command => [command.type, true])) as DocxEditorSnapshot["enabled"]
	}
}

function press(part: string): DocxObjectPress {
	return {
		hit: { kind: "control", story: 0, pos: 5, part },
		page: 0,
		x: 100,
		y: 100,
		clientX: 120,
		clientY: 140,
		pointerType: "mouse",
		pointerId: 1,
		clicks: 1,
		shiftKey: false
	}
}

function renderControls(options: {
	at: DocxControlContext | null
	part?: string
	readOnly?: boolean
	inputRequest?: { type: "control.insert"; detail: string; nonce: number }
}) {
	const execute = vi.fn()
	const onPressDone = vi.fn()
	const onInputDone = vi.fn()
	const state = snapshot(options.at)
	const editor = {
		hitTest: vi.fn(() => ({ pos: 6, affinity: 1, page: 0, line: 0, atom: false })),
		select: vi.fn(),
		snapshot: () => state
	}
	// No container: the popover is placed at the press in the window, nothing measured while rendering.
	const container = null

	render(
		<DocxFormControls
			editor={editor as unknown as PagedEditor}
			snapshot={state}
			surface="pages"
			container={container}
			zoom={1}
			readOnly={options.readOnly ?? false}
			execute={execute}
			press={options.part === undefined ? null : press(options.part)}
			onPressDone={onPressDone}
			inputRequest={options.inputRequest ?? null}
			requestInput={vi.fn()}
			onInputDone={onInputDone}
		/>
	)

	return { execute, onPressDone, onInputDone, editor }
}

afterEach(() => {
	cleanup()
})

describe("DocxFormControls", () => {
	it("ticks a checkbox at once, the caret placed in it first", async () => {
		const { execute, onPressDone, editor } = renderControls({
			at: control({ type: "checkbox", items: [], checked: false }),
			part: "checkbox"
		})

		await waitFor(() => {
			expect(execute).toHaveBeenCalledWith({ type: "control.check" })
		})
		expect(editor.hitTest).toHaveBeenCalledWith(0, 100, 100)
		expect(editor.select).toHaveBeenCalledTimes(1)
		expect(onPressDone).toHaveBeenCalledTimes(1)
		expect(screen.queryByRole("dialog")).toBeNull()
	})

	it("tells why a locked checkbox does not change", async () => {
		const { execute } = renderControls({
			at: control({ type: "checkbox", items: [], checked: false, locked: "contentLocked" }),
			part: "checkbox"
		})

		expect((await screen.findByRole("status")).textContent).toBe("This control's content is locked.")
		expect(execute).not.toHaveBeenCalled()
	})

	it("offers a dropdown's items at the control, the chosen one marked, and clears it", async () => {
		const { execute, onPressDone } = renderControls({ at: control(), part: "list" })
		const options = await screen.findAllByRole("option")

		expect(options.map(option => option.textContent)).toEqual(["Red", "Green"])
		expect(options[0]?.getAttribute("aria-selected")).toBe("true")
		const green = options[1]

		if (green === undefined) {
			throw new Error("no second item")
		}

		fireEvent.click(green)
		expect(execute).toHaveBeenCalledWith({ type: "control.choose", value: "g" })
		expect(onPressDone).toHaveBeenCalled()
		cleanup()

		const cleared = renderControls({ at: control(), part: "list" })

		fireEvent.click(await screen.findByRole("button", { name: "Clear" }))
		expect(cleared.execute).toHaveBeenCalledWith({ type: "control.clear" })
	})

	it("sets a date control's date from the picker", async () => {
		const { execute } = renderControls({
			at: control({ type: "date", items: [], value: null, date: "2026-01-01", dateFormat: "M/d/yyyy" }),
			part: "date"
		})
		const field = await screen.findByLabelText("Date")

		expect((field as HTMLInputElement).value).toBe("2026-01-01")
		fireEvent.change(field, { target: { value: "2026-10-13" } })
		fireEvent.click(screen.getByRole("button", { name: "Set date" }))
		expect(execute).toHaveBeenCalledWith({ type: "control.date", iso: "2026-10-13" })
	})

	it("shows why a bound control or a read-only document cannot change, offering nothing", async () => {
		renderControls({ at: control({ locked: "bound" }), part: "list" })
		expect((await screen.findByRole("status")).textContent).toBe("This control is linked to data that can't be changed here.")
		expect(screen.getByRole("listbox").getAttribute("aria-disabled")).toBe("true")
		expect(screen.queryByRole("button", { name: "Clear" })).toBeNull()
		cleanup()

		renderControls({ at: control(), part: "list", readOnly: true })
		expect((await screen.findByRole("status")).textContent).toBe("This document can only be viewed.")
	})

	it("asks for a new dropdown list's items, added and removed, before inserting it", () => {
		const { execute, onInputDone } = renderControls({
			at: null,
			inputRequest: { type: "control.insert", detail: "dropDownList", nonce: 3 }
		})
		const field = screen.getByRole("textbox", { name: "New item" })
		const insert = screen.getByRole("button", { name: "Insert" })

		expect(insert.hasAttribute("disabled")).toBe(true)

		for (const item of ["Small", "Large", "Medium"]) {
			fireEvent.change(field, { target: { value: item } })
			fireEvent.keyDown(field, { key: "Enter" })
		}

		fireEvent.change(field, { target: { value: "Large" } })
		expect(screen.getByText("That item is already in the list.")).toBeTruthy()
		fireEvent.click(screen.getByRole("button", { name: "Remove Large" }))
		fireEvent.click(insert)
		expect(execute).toHaveBeenCalledWith({
			type: "control.insert",
			kind: "dropDownList",
			items: [
				{ text: "Small", value: "Small" },
				{ text: "Medium", value: "Medium" }
			]
		})
		expect(onInputDone).toHaveBeenCalledWith(3)
	})
})

describe("DocxFormControls' dropdown popover", () => {
	it("is a listbox the arrow keys move through and Enter chooses from", async () => {
		const { execute } = renderControls({ at: control(), part: "list" })
		const list = await screen.findByRole("listbox", { name: "Items" })
		const options = screen.getAllByRole("option")

		// The chosen entry is the active one; Down moves to the next.
		expect(list.getAttribute("aria-activedescendant")).toBe(options[0]?.id)
		fireEvent.keyDown(list, { key: "ArrowDown" })
		expect(list.getAttribute("aria-activedescendant")).toBe(options[1]?.id)
		fireEvent.keyDown(list, { key: "Home" })
		fireEvent.keyDown(list, { key: "End" })
		fireEvent.keyDown(list, { key: "Enter" })
		expect(execute).toHaveBeenCalledWith({ type: "control.choose", value: "g" })
	})

	it("adds and removes the list's entries there, each value kept", async () => {
		const { execute } = renderControls({ at: control(), part: "list" })

		fireEvent.click(await screen.findByRole("button", { name: "Edit items" }))
		fireEvent.change(screen.getByRole("textbox", { name: "New item" }), { target: { value: "Blue" } })
		fireEvent.click(screen.getByRole("button", { name: "Add" }))
		expect(execute).toHaveBeenLastCalledWith({
			type: "control.items",
			items: [
				{ text: "Red", value: "r" },
				{ text: "Green", value: "g" },
				{ text: "Blue", value: "Blue" }
			]
		})
		fireEvent.click(screen.getByRole("button", { name: "Remove Red" }))
		expect(execute).toHaveBeenLastCalledWith({ type: "control.items", items: [{ text: "Green", value: "g" }] })
	})

	it("offers no entry editing where the list cannot change", async () => {
		const state = snapshot(control())

		renderControls({ at: control({ locked: "contentLocked" }), part: "list" })
		expect(await screen.findByRole("listbox")).toBeTruthy()
		expect(screen.queryByRole("button", { name: "Edit items" })).toBeNull()
		expect(state.enabled["control.items"]).toBe(true)
	})
})

describe("DocxFormControls' dropdown popover, more", () => {
	it("moves to the next entry starting with a letter typed, scrolling it into view", async () => {
		const scrolled = vi.fn()

		Element.prototype.scrollIntoView = scrolled
		renderControls({
			at: control({
				items: [
					{ text: "Red", value: "r" },
					{ text: "Green", value: "g" },
					{ text: "Grey", value: "y" }
				]
			}),
			part: "list"
		})

		const list = await screen.findByRole("listbox")
		const options = screen.getAllByRole("option")

		fireEvent.keyDown(list, { key: "g" })
		expect(list.getAttribute("aria-activedescendant")).toBe(options[1]?.id)
		fireEvent.keyDown(list, { key: "g" })
		expect(list.getAttribute("aria-activedescendant")).toBe(options[2]?.id)
		expect(scrolled).toHaveBeenCalled()
	})

	it("adds characters typed in quick succession to the type-ahead, a new one after a pause", async () => {
		Element.prototype.scrollIntoView = vi.fn()

		const now = vi.spyOn(Date, "now")
		const { execute } = renderControls({
			at: control({
				items: [
					{ text: "Red", value: "r" },
					{ text: "Green", value: "g" },
					{ text: "Grey", value: "y" },
					{ text: "Grey blue", value: "b" }
				]
			}),
			part: "list"
		})
		const list = await screen.findByRole("listbox")
		const options = screen.getAllByRole("option")
		const type = (key: string, at: number): void => {
			now.mockReturnValue(at)
			fireEvent.keyDown(list, { key })
		}

		try {
			// "gre": Green; "grey": Grey, the active entry kept while it still matches; "grey " then "b": Grey blue.
			type("g", 1000)
			type("r", 1100)
			type("e", 1200)
			expect(list.getAttribute("aria-activedescendant")).toBe(options[1]?.id)
			type("y", 1300)
			expect(list.getAttribute("aria-activedescendant")).toBe(options[2]?.id)
			// A space while typing is part of the text, not a choice.
			type(" ", 1400)
			type("b", 1500)
			expect(list.getAttribute("aria-activedescendant")).toBe(options[3]?.id)
			expect(execute).not.toHaveBeenCalled()
			// After a pause: a new search, "r" alone.
			type("r", 2600)
			expect(list.getAttribute("aria-activedescendant")).toBe(options[0]?.id)
		} finally {
			now.mockRestore()
		}
	})

	it("refuses an entry whose text is another's value, keeping the list keyed by place", async () => {
		const { execute } = renderControls({ at: control(), part: "list" })

		fireEvent.click(await screen.findByRole("button", { name: "Edit items" }))
		// "g" is Green's value.
		fireEvent.change(screen.getByRole("textbox", { name: "New item" }), { target: { value: "g" } })
		expect(screen.getByText("That item is already in the list.")).toBeTruthy()
		expect(screen.getByRole("button", { name: "Add" }).hasAttribute("disabled")).toBe(true)
		expect(execute).not.toHaveBeenCalled()
	})
})
