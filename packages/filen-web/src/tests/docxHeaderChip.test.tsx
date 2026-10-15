// @vitest-environment happy-dom

import { afterEach, describe, expect, it, vi } from "vitest"
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react"
import { DOCX_EDITOR_COMMANDS, type DocxEditorSnapshot } from "@filen/office/core/docx/edit"
import "@/lib/i18n"
import { DocxHeaderChip } from "@/features/preview/components/docxHeaderChip"

// The header chip: which header it is (its section only once there are several), Close named for it, and Link
// to previous asked first where it deletes the section's own header.

type Header = NonNullable<DocxEditorSnapshot["context"]["story"]["header"]>

const ENABLED = Object.fromEntries(DOCX_EDITOR_COMMANDS.map(command => [command.type, true])) as DocxEditorSnapshot["enabled"]

function header(overrides: Partial<Header> = {}): Header {
	return {
		kind: "header",
		section: 1,
		sections: 2,
		type: "default",
		differentFirst: false,
		differentOddEven: false,
		linkedToPrevious: false,
		canLink: true,
		...overrides
	}
}

function renderChip(value: Header) {
	const execute = vi.fn()
	const onClose = vi.fn()

	render(
		<DocxHeaderChip
			header={value}
			enabled={ENABLED}
			readOnly={false}
			execute={execute}
			onClose={onClose}
		/>
	)

	return { execute, onClose }
}

afterEach(() => {
	cleanup()
})

describe("DocxHeaderChip", () => {
	it("names the section only in a document with several, and its Close for the header or footer", () => {
		const { onClose } = renderChip(header())

		expect(screen.getByRole("group", { name: "Header — Section 2" })).toBeTruthy()
		fireEvent.click(screen.getByRole("button", { name: "Close header" }))
		expect(onClose).toHaveBeenCalledTimes(1)
		cleanup()

		renderChip(header({ kind: "footer", section: 0, sections: 1, type: "first" }))
		expect(screen.getByRole("group", { name: "First page footer" })).toBeTruthy()
		expect(screen.getByRole("button", { name: "Close footer" })).toBeTruthy()
	})

	it("closes the options menu once a toggle is switched", async () => {
		const { execute } = renderChip(header())

		fireEvent.click(screen.getByRole("button", { name: "Header and footer options" }))
		fireEvent.click(await screen.findByRole("menuitemcheckbox", { name: "Different first page" }))
		expect(execute).toHaveBeenLastCalledWith({ type: "header.differentFirst" })
		await waitFor(() => {
			expect(screen.queryByRole("menuitemcheckbox", { name: "Different first page" })).toBeNull()
		})
	})

	it("asks before Link to previous deletes the section's own header, and links once confirmed", async () => {
		const { execute } = renderChip(header())

		fireEvent.click(screen.getByRole("button", { name: "Header and footer options" }))
		fireEvent.click(await screen.findByRole("menuitemcheckbox", { name: "Link to previous" }))
		expect(execute).not.toHaveBeenCalled()

		const dialog = await screen.findByRole("alertdialog", { name: "Link to the previous section?" })

		expect(dialog.textContent).toContain("own header is deleted")
		fireEvent.click(screen.getByRole("button", { name: "Link" }))
		expect(execute).toHaveBeenCalledWith({ type: "header.linkToPrevious" })
	})

	it("says a footer's own is deleted, and gives focus back to the chip's options once answered", async () => {
		const { execute } = renderChip(header({ kind: "footer" }))
		const options = screen.getByRole("button", { name: "Header and footer options" })

		fireEvent.click(options)
		fireEvent.click(await screen.findByRole("menuitemcheckbox", { name: "Link to previous" }))

		const dialog = await screen.findByRole("alertdialog", { name: "Link to the previous section?" })

		expect(dialog.textContent).toContain("own footer is deleted")
		fireEvent.click(screen.getByRole("button", { name: "Cancel" }))
		await waitFor(() => {
			expect(document.activeElement).toBe(options)
		})
		expect(execute).not.toHaveBeenCalled()
	})

	it("unlinks at once: nothing is deleted then", async () => {
		const { execute } = renderChip(header({ linkedToPrevious: true }))

		fireEvent.click(screen.getByRole("button", { name: "Header and footer options" }))
		fireEvent.click(await screen.findByRole("menuitemcheckbox", { name: "Link to previous" }))
		expect(execute).toHaveBeenCalledWith({ type: "header.linkToPrevious" })
		expect(screen.queryByRole("alertdialog")).toBeNull()
	})
})
