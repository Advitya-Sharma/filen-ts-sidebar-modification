// @vitest-environment jsdom

import { afterEach, describe, expect, it, vi } from "vitest"
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react"
import "@/lib/i18n"
import { DocxPasswordPrompt, type PasswordVerdict } from "@/features/preview/components/docxPasswordPrompt"

// The generic document password prompt: masked with a show/hide toggle, Escape stops a running check (the prompt
// stays), and dismisses it otherwise.

afterEach(() => {
	cleanup()
})

function prompt(verify: (password: string, signal: AbortSignal) => Promise<PasswordVerdict>, onOpenChange = vi.fn()) {
	render(
		<DocxPasswordPrompt
			open
			title="Document password"
			body="Enter it."
			wrongMessage="Wrong."
			label="Password"
			submitLabel="Open"
			progressLabel="Checking…"
			verify={verify}
			onOpenChange={onOpenChange}
		/>
	)

	return onOpenChange
}

function field(): HTMLInputElement {
	return screen.getByLabelText("Password", { selector: "input" })
}

describe("DocxPasswordPrompt", () => {
	it("masks the password, and shows it on request", async () => {
		prompt(() => Promise.resolve("accepted"))

		// The dialog focuses it on the next frame.
		await waitFor(() => {
			expect(window.document.activeElement).toBe(field())
		})
		expect(field().className).toContain("[-webkit-text-security:disc]")
		expect(field().getAttribute("autocomplete")).toBe("off")

		fireEvent.click(screen.getByRole("button", { name: "Show password" }))

		expect(field().className).not.toContain("[-webkit-text-security:disc]")
		expect(screen.getByRole("button", { name: "Hide password" }).getAttribute("aria-pressed")).toBe("true")
	})

	it("stops a running check on Escape and stays open; closes on Escape otherwise", async () => {
		let signal: AbortSignal | undefined
		const onOpenChange = prompt(
			(_password, given) =>
				new Promise(resolve => {
					signal = given
					given.addEventListener("abort", () => {
						resolve("aborted")
					})
				})
		)

		fireEvent.change(field(), { target: { value: "pw" } })
		fireEvent.click(screen.getByRole("button", { name: "Open" }))
		expect(screen.getByRole("button", { name: "Checking…" })).toBeTruthy()

		fireEvent.keyDown(window.document.activeElement ?? window.document.body, { key: "Escape" })

		await waitFor(() => {
			expect(signal?.aborted).toBe(true)
		})
		expect(onOpenChange).not.toHaveBeenCalled()
		expect(field().disabled).toBe(false)

		fireEvent.keyDown(field(), { key: "Escape" })

		await waitFor(() => {
			expect(onOpenChange).toHaveBeenCalledWith(false)
		})
	})

	it("drops the typed password once checked", async () => {
		const verify = vi.fn(() => Promise.resolve<PasswordVerdict>("wrong"))

		prompt(verify)
		fireEvent.change(field(), { target: { value: "pw" } })
		fireEvent.click(screen.getByRole("button", { name: "Open" }))

		expect((await screen.findByRole("alert")).textContent).toBe("Wrong.")
		expect(verify).toHaveBeenCalledWith("pw", expect.any(AbortSignal))
		expect(field().value).toBe("")
	})
})
