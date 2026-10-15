// @vitest-environment happy-dom

import { createRef } from "react"
import { afterEach, describe, expect, it, vi } from "vitest"
import { act, cleanup, fireEvent, render, renderHook, screen } from "@testing-library/react"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import type { DocxDocument } from "@filen/office/react/docx"
import type { DocxEditorHandle, DocxReflowProps } from "@filen/office/react/docx/editor"
import { DropdownMenu, DropdownMenuContent } from "@/components/ui/dropdown-menu"
import { slotCommands } from "@/features/preview/components/docxEditor.logic"
import { INITIAL_REFLOW, reflowChosen, reflowCommands, reflowForWidth, useReflowMode } from "@/features/preview/components/docxReflow.logic"
import "@/lib/i18n"

// Mobile view: on by itself below 600 px of editor width, following width crossings until the user chooses;
// the ⋯ menu's switch; the surface hands the editor's props and handle to the reflowed view.

const reflowed = vi.hoisted(() => ({ props: [] as DocxReflowProps[] }))

vi.mock("@filen/office/react/docx/editor", async () => ({
	...(await vi.importActual<typeof import("@filen/office/react/docx/editor")>("@filen/office/react/docx/editor")),
	DocxReflow: (props: DocxReflowProps) => {
		reflowed.props.push(props)

		return <div data-testid="reflow" />
	}
}))

const { DocxReflowMenuItem, DocxReflowSurface } = await import("@/features/preview/components/docxReflowSurface")

afterEach(() => {
	cleanup()
	reflowed.props.length = 0
})

describe("Mobile view mode", () => {
	it("turns on below 600 px, follows crossings, and keeps the user's choice", () => {
		let state = reflowForWidth(INITIAL_REFLOW, 0)

		// Not measured yet: nothing decided.
		expect(state).toBe(INITIAL_REFLOW)

		state = reflowForWidth(state, 420)
		expect(state.reflow).toBe(true)

		// Resizing on the same side changes nothing.
		expect(reflowForWidth(state, 500)).toBe(state)

		state = reflowForWidth(state, 900)
		expect(state.reflow).toBe(false)

		state = reflowForWidth(state, 599)
		expect(state.reflow).toBe(true)

		// The user turns it off: crossings no longer decide.
		state = reflowChosen(state, false)
		expect(state).toMatchObject({ reflow: false, chosen: true })

		state = reflowForWidth(reflowForWidth(state, 900), 300)
		expect(state.reflow).toBe(false)

		expect(reflowChosen(state, false)).toBe(state)
		expect(reflowChosen(state, true).reflow).toBe(true)
	})

	it("starts wide in pages and switches as the width crosses, until set", () => {
		const { result, rerender } = renderHook(({ width }: { width: number }) => useReflowMode(width), { initialProps: { width: 0 } })

		expect(result.current.reflow).toBe(false)

		rerender({ width: 1200 })
		expect(result.current.reflow).toBe(false)

		rerender({ width: 480 })
		expect(result.current.reflow).toBe(true)

		act(() => {
			result.current.setReflow(false)
		})
		expect(result.current.reflow).toBe(false)

		rerender({ width: 1200 })
		rerender({ width: 400 })
		expect(result.current.reflow).toBe(false)
	})
})

describe("Mobile view commands", () => {
	it("offers no header or footer editing in Mobile view, all of them on the pages", () => {
		const sections = slotCommands("insert")
		const types = (reflow: boolean): string[] => reflowCommands(sections, reflow).map(command => command.info.type)

		expect(types(false)).toEqual(expect.arrayContaining(["header.edit", "footer.edit", "insert.pageBreak"]))
		expect(types(true)).not.toContain("header.edit")
		expect(types(true)).not.toContain("footer.edit")
		expect(types(true)).toContain("insert.pageBreak")
		expect(reflowCommands(sections, false)).toBe(sections)
	})
})

describe("Mobile view switch", () => {
	const props = (reflow: boolean, onReflowChange: (reflow: boolean) => void) => ({
		commands: slotCommands("view"),
		snapshot: undefined,
		readOnly: false,
		apple: false,
		execute: vi.fn(),
		surface: reflow ? ("reflow" as const) : ("pages" as const),
		inputRequest: null,
		requestInput: vi.fn(),
		onInputDone: vi.fn(),
		reflow,
		onReflowChange
	})

	it("is a checked item of the ⋯ menu that switches the view", () => {
		const onReflowChange = vi.fn()

		render(
			<DropdownMenu open>
				<DropdownMenuContent>
					<DocxReflowMenuItem {...props(false, onReflowChange)} />
				</DropdownMenuContent>
			</DropdownMenu>
		)

		const item = screen.getByRole("menuitemcheckbox", { name: "Mobile view" })

		expect(item.getAttribute("aria-checked")).toBe("false")

		fireEvent.click(item)

		expect(onReflowChange).toHaveBeenCalledWith(true)
	})

	it("closes the ⋯ menu once the view is switched", () => {
		const onOpenChange = vi.fn()

		render(
			<DropdownMenu
				open
				onOpenChange={onOpenChange}
			>
				<DropdownMenuContent>
					<DocxReflowMenuItem {...props(false, vi.fn())} />
				</DropdownMenuContent>
			</DropdownMenu>
		)

		fireEvent.click(screen.getByRole("menuitemcheckbox", { name: "Mobile view" }))
		expect(onOpenChange.mock.calls[0]?.[0]).toBe(false)
	})

	it("shows checked in Mobile view and nothing without its command", () => {
		const { rerender } = render(
			<DropdownMenu open>
				<DropdownMenuContent>
					<DocxReflowMenuItem {...props(true, vi.fn())} />
				</DropdownMenuContent>
			</DropdownMenu>
		)

		expect(screen.getByRole("menuitemcheckbox", { name: "Mobile view" }).getAttribute("aria-checked")).toBe("true")

		rerender(
			<DropdownMenu open>
				<DropdownMenuContent>
					<DocxReflowMenuItem
						{...props(true, vi.fn())}
						commands={[]}
					/>
				</DropdownMenuContent>
			</DropdownMenu>
		)

		expect(screen.queryByRole("menuitemcheckbox")).toBeNull()
	})
})

describe("Mobile view surface", () => {
	it("hands the reflowed view the editor, its handle and what the pages get", () => {
		const editor = {} as PagedEditor
		const document = { status: "ready" } as DocxDocument
		const handleRef = createRef<DocxEditorHandle>()
		const onFind = vi.fn()
		const commentMarker = vi.fn(() => true)
		const marks = [{ from: 1, to: 4, kind: "spelling" }]

		render(
			<DocxReflowSurface
				editor={editor}
				document={document}
				handleRef={handleRef}
				label="Report.docx"
				readOnly
				apple
				marks={marks}
				onFind={onFind}
				interactions={{ commentMarker }}
			/>
		)

		expect(screen.getByTestId("reflow")).toBeTruthy()

		const given = reflowed.props.at(-1)

		expect(given).toMatchObject({ editor, document, label: "Report.docx", readOnly: true, apple: true, marks, onFind })
		expect(given?.ref).toBe(handleRef)
		expect(given?.interactions?.commentMarker).toBe(commentMarker)
		expect([
			given?.storyLabel?.("footnote", "3"),
			given?.storyLabel?.("endnote", "iv"),
			given?.storyLabel?.("footnote", ""),
			given?.storyLabel?.("textbox", "")
		]).toEqual(["Footnote 3", "Endnote iv", "Footnote", "Text box"])
	})
})
