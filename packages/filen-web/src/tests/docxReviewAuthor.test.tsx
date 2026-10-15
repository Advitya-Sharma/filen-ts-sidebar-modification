// @vitest-environment happy-dom

import { afterEach, describe, expect, it, vi } from "vitest"
import { act, cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react"
import { NO_REVIEW, type DocxEditorCommand, type DocxRefusalReason } from "@filen/office/core/docx/edit"
import type { PagedEditor } from "@filen/office/core/docx/paged"
import "@/lib/i18n"

// The author name gate (track changes without a name): an edit the editor refuses asks for the name, offering to
// turn tracking off; the edit is never made later, nor with a stand-in name.

const settings = vi.hoisted(() => ({ asks: 0, signals: [] as AbortSignal[] }))
const toasts = vi.hoisted(() => ({ error: vi.fn() }))

vi.mock("sonner", () => ({ toast: toasts }))

vi.mock("@/features/office/lib/officeSettings", () => ({
	ensureAuthorName: (ask: () => Promise<string | null>, signal?: AbortSignal) => {
		settings.asks++

		if (signal !== undefined) {
			settings.signals.push(signal)
		}

		return ask()
	}
}))

const { useDocxAuthorGate } = await import("@/features/preview/components/docxAuthorGate")

// An editor that refuses edits when told to, tracking without a name.
function fakeEditor(trackable = true): { editor: PagedEditor; refuse: (reason?: DocxRefusalReason) => void } {
	const listeners = new Set<(reason: DocxRefusalReason) => void>()
	const editor = {
		onRefused: (listener: (reason: DocxRefusalReason) => void) => {
			listeners.add(listener)

			return () => {
				listeners.delete(listener)
			}
		},
		snapshot: () => ({ review: { ...NO_REVIEW, tracking: true, authorNeeded: true }, enabled: { "review.track": trackable } })
	} as unknown as PagedEditor

	return {
		editor,
		refuse: (reason = "authorNeeded") => {
			act(() => {
				for (const listener of listeners) listener(reason)
			})
		}
	}
}

function Gate({
	editor,
	execute,
	onReady
}: {
	editor: PagedEditor
	execute: (command: DocxEditorCommand) => void
	onReady: (gate: ReturnType<typeof useDocxAuthorGate>["withAuthor"]) => void
}) {
	const { withAuthor, dialog } = useDocxAuthorGate({ authorName: null, editor, execute })

	onReady(withAuthor)

	return dialog
}

function mount(trackable = true): {
	refuse: (reason?: DocxRefusalReason) => void
	execute: ReturnType<typeof vi.fn>
	gate: () => ReturnType<typeof useDocxAuthorGate>["withAuthor"]
	unmount: () => void
} {
	const { editor, refuse } = fakeEditor(trackable)
	const execute = vi.fn()
	let gate: ReturnType<typeof useDocxAuthorGate>["withAuthor"] | undefined
	const view = render(
		<Gate
			editor={editor}
			execute={execute}
			onReady={withAuthor => {
				gate = withAuthor
			}}
		/>
	)

	return {
		refuse,
		execute,
		gate: () => {
			if (gate === undefined) {
				throw new Error("no gate")
			}

			return gate
		},
		unmount: view.unmount
	}
}

afterEach(() => {
	cleanup()
	vi.useRealTimers()
	toasts.error.mockReset()
	settings.asks = 0
	settings.signals = []
})

describe("author name gate", () => {
	it("asks once for refused edits, keeps the name given and makes no edit", async () => {
		const { refuse, execute } = mount()

		refuse()

		const field = await screen.findByRole("textbox", { name: "Name" })

		// Refusals while the prompt is open ask nothing more.
		refuse()
		refuse()
		expect(settings.asks).toBe(1)
		expect(screen.getByRole("button", { name: "Turn off track changes" })).toBeTruthy()
		fireEvent.change(field, { target: { value: "Jane Roe" } })
		fireEvent.click(screen.getByRole("button", { name: "Save" }))
		await waitFor(() => {
			expect(screen.queryByRole("textbox", { name: "Name" })).toBeNull()
		})
		expect(execute).not.toHaveBeenCalled()
	})

	it("turns track changes off instead of giving a name", async () => {
		const { refuse, execute } = mount()

		refuse()
		fireEvent.click(await screen.findByRole("button", { name: "Turn off track changes" }))
		await waitFor(() => {
			expect(screen.queryByRole("textbox", { name: "Name" })).toBeNull()
		})
		expect(execute).toHaveBeenCalledWith({ type: "review.track", on: false })
	})

	it("offers to turn track changes off only where that can be done", async () => {
		const { refuse } = mount(false)

		refuse()
		expect(await screen.findByRole("textbox", { name: "Name" })).toBeTruthy()
		expect(screen.queryByRole("button", { name: "Turn off track changes" })).toBeNull()
	})

	it("does nothing on a dismiss and asks again at the next refused edit", async () => {
		const { refuse, execute } = mount()

		refuse()
		fireEvent.keyDown(await screen.findByRole("textbox", { name: "Name" }), { key: "Escape" })
		await waitFor(() => {
			expect(screen.queryByRole("textbox", { name: "Name" })).toBeNull()
		})
		expect(execute).not.toHaveBeenCalled()
		refuse()
		expect(await screen.findByRole("textbox", { name: "Name" })).toBeTruthy()
		expect(settings.asks).toBe(2)
	})

	it("runs a command with the name given, never when declined", async () => {
		const { gate } = mount()
		const run = vi.fn()

		act(() => {
			gate()(run)
		})
		fireEvent.change(await screen.findByRole("textbox", { name: "Name" }), { target: { value: "Jo" } })
		fireEvent.click(screen.getByRole("button", { name: "Save" }))
		await waitFor(() => {
			expect(run).toHaveBeenCalledWith("Jo")
		})

		const declined = vi.fn()

		act(() => {
			gate()(declined)
		})
		fireEvent.keyDown(await screen.findByRole("textbox", { name: "Name" }), { key: "Escape" })
		await waitFor(() => {
			expect(screen.queryByRole("textbox", { name: "Name" })).toBeNull()
		})
		expect(declined).not.toHaveBeenCalled()
	})

	it("drops a question still open when the editor closes", async () => {
		const { refuse, unmount } = mount()

		refuse()
		await screen.findByRole("textbox", { name: "Name" })
		unmount()
		expect(settings.signals.at(-1)?.aborted).toBe(true)
	})

	it("tells about an edit that cannot be saved in a toast, once per burst, and asks no name for it", () => {
		vi.useFakeTimers({ toFake: ["Date"] })

		const { refuse } = mount()

		refuse("unsupported")
		vi.advanceTimersByTime(500)
		refuse("unsupported")
		vi.advanceTimersByTime(1500)
		refuse("unsupported")

		expect(toasts.error).toHaveBeenCalledTimes(1)
		expect(toasts.error).toHaveBeenCalledWith("This edit can't be made here", {
			description: "Part of what it changes can't be saved as a tracked change."
		})
		expect(settings.asks).toBe(0)
		expect(screen.queryByRole("textbox", { name: "Name" })).toBeNull()

		// After a quiet spell, a new burst.
		vi.advanceTimersByTime(2500)
		refuse("unsupported")
		expect(toasts.error).toHaveBeenCalledTimes(2)
	})

	it("asks for the name on a refusal for want of it, with no toast", async () => {
		const { refuse } = mount()

		refuse("authorNeeded")
		expect(await screen.findByRole("textbox", { name: "Name" })).toBeTruthy()
		expect(toasts.error).not.toHaveBeenCalled()
	})
})
