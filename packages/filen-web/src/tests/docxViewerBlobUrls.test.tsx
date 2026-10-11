// @vitest-environment jsdom

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { render, waitFor } from "@testing-library/react"
import "@/lib/i18n"

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
