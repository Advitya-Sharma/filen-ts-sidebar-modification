// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { render as renderUi, screen, waitFor } from "@testing-library/react"
import type { ReactElement } from "react"
import { createTestQueryClient, queryClientWrapper } from "@/tests/testQueryClient"
import "@/lib/i18n"

// Links open through the trusted-domain query (useTrustedLinkGate).
function render(ui: ReactElement) {
	return renderUi(ui, { wrapper: queryClientWrapper(createTestQueryClient()) })
}

// docx-preview (now only the fallback when the own engine fails) mints an object URL per embedded image
// and font and never revokes one; the fallback owns that, once the render it belongs to is gone.

const { renderAsync } = vi.hoisted(() => ({
	renderAsync: vi.fn<(bytes: Uint8Array, host: HTMLElement) => Promise<void>>()
}))

vi.mock("docx-preview", () => ({ renderAsync }))

const { DocxPreviewFallback } = await import("@/features/preview/components/docxPreviewFallback")

const bytes = new Uint8Array([1])

const revokeObjectURL = vi.fn<(url: string) => void>()

beforeEach(() => {
	URL.revokeObjectURL = revokeObjectURL
})

afterEach(() => {
	revokeObjectURL.mockReset()
})

function paint(host: HTMLElement): void {
	host.innerHTML = `<style>@font-face { src: url(blob:https://app/font) }</style><img src="blob:https://app/img">`
}

describe("DocxPreviewFallback object URLs", () => {
	it("keeps them while the document is shown and revokes them on unmount", async () => {
		renderAsync.mockImplementation((_bytes, host) => {
			paint(host)

			return Promise.resolve()
		})

		const view = render(
			<DocxPreviewFallback
				bytes={bytes}
				alt="doc.docx"
			/>
		)

		await waitFor(() => {
			expect(view.container.querySelector("img")).not.toBeNull()
		})

		expect(revokeObjectURL).not.toHaveBeenCalled()

		view.unmount()

		expect(revokeObjectURL.mock.calls.map(([url]) => url).sort()).toEqual(["blob:https://app/font", "blob:https://app/img"])
	})

	it("revokes right away the URLs of a render that finishes after unmount", async () => {
		let finish: () => void = () => undefined
		renderAsync.mockImplementation(
			(_bytes, host) =>
				new Promise<void>(resolve => {
					finish = () => {
						paint(host)
						resolve()
					}
				})
		)

		const view = render(
			<DocxPreviewFallback
				bytes={bytes}
				alt="doc.docx"
			/>
		)

		await waitFor(() => {
			expect(renderAsync).toHaveBeenCalled()
		})

		view.unmount()
		finish()

		await waitFor(() => {
			expect(revokeObjectURL).toHaveBeenCalledTimes(2)
		})
	})
})

describe("DocxPreviewFallback links", () => {
	it("asks before a web link opens a domain not trusted yet", async () => {
		renderAsync.mockImplementation((_bytes, host) => {
			host.innerHTML = `<p><a href="https://example.com/x">Web</a></p>`

			return Promise.resolve()
		})

		const view = render(
			<DocxPreviewFallback
				bytes={bytes}
				alt="doc.docx"
			/>
		)
		const anchor = await screen.findByText("Web")
		const click = new MouseEvent("click", { bubbles: true, cancelable: true })

		anchor.dispatchEvent(click)

		expect(click.defaultPrevented).toBe(true)
		expect(await screen.findByRole("alertdialog", { name: "Open external link?" })).toBeTruthy()
		view.unmount()
	})

	it("asks before a middle-click opens it too, and leaves other buttons alone", async () => {
		renderAsync.mockImplementation((_bytes, host) => {
			host.innerHTML = `<p><a href="https://example.com/x">Web</a></p>`

			return Promise.resolve()
		})

		const view = render(
			<DocxPreviewFallback
				bytes={bytes}
				alt="doc.docx"
			/>
		)
		const anchor = await screen.findByText("Web")
		const right = new MouseEvent("auxclick", { bubbles: true, cancelable: true, button: 2 })

		anchor.dispatchEvent(right)
		expect(right.defaultPrevented).toBe(false)
		expect(screen.queryByRole("alertdialog")).toBeNull()

		const middle = new MouseEvent("auxclick", { bubbles: true, cancelable: true, button: 1 })

		anchor.dispatchEvent(middle)
		expect(middle.defaultPrevented).toBe(true)
		expect(await screen.findByRole("alertdialog", { name: "Open external link?" })).toBeTruthy()
		view.unmount()
	})
})
