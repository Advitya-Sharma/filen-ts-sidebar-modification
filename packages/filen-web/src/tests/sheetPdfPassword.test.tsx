// @vitest-environment jsdom

import { describe, expect, it, vi } from "vitest"
import { fireEvent, render, screen, waitFor } from "@testing-library/react"
import { toast } from "sonner"
import type { SheetSessionApi } from "@filen/office/core/sheet"
import type { SheetPdfRequest } from "@filen/office/react/sheet/print"
import { useSheetPdfExport, type SheetPrintRequest } from "@/features/spreadsheet/save/useSheetPdfExport"
import { preview } from "@/locales/en/preview"
import "@/lib/i18n"

// A protected workbook's Download as PDF: its password is asked, checked in the worker against the file as stored,
// and goes with that export only; a workbook Excel's default password opens is not asked about.

const state = vi.hoisted(() => ({
	password: "secret",
	defaulted: false,
	checks: [] as string[],
	exports: [] as (string | undefined)[]
}))

vi.mock("@filen/office/dom/sheet", () => ({
	checkSheetPassword: (_session: unknown, _bytes: Uint8Array, password: string) => {
		state.checks.push(password)

		return Promise.resolve(state.defaulted ? "default" : password === state.password ? "accepted" : "wrong")
	}
}))
vi.mock("@/features/spreadsheet/save/sheetPdf", () => ({
	downloadSheetPdf: (_session: unknown, _request: unknown, _name: string, password?: string) => {
		state.exports.push(password)

		return Promise.resolve()
	}
}))

const REQUEST = {} as SheetPdfRequest
const PRINT: SheetPrintRequest = () => Promise.resolve(REQUEST)

function Harness({ encrypted }: { encrypted: boolean }) {
	const pdf = useSheetPdfExport({
		session: {} as SheetSessionApi,
		fileName: "Budget.xlsx",
		encryptedFile: encrypted ? new Uint8Array([1, 2, 3]) : undefined,
		offered: true
	})

	return (
		<>
			<button
				type="button"
				onClick={() => {
					pdf.start?.(PRINT)
				}}
			>
				Export
			</button>
			{pdf.prompt}
		</>
	)
}

function reset(): void {
	state.checks.length = 0
	state.exports.length = 0
}

describe("a spreadsheet's PDF of a password-protected workbook", () => {
	it("exports a plain workbook at once, without a password", async () => {
		reset()
		render(<Harness encrypted={false} />)
		fireEvent.click(screen.getByRole("button", { name: "Export" }))
		await waitFor(() => {
			expect(state.exports).toEqual([undefined])
		})
		expect(state.checks).toEqual([])
	})

	it("asks the password, says when it is wrong, and exports with the right one", async () => {
		reset()
		state.defaulted = false
		render(<Harness encrypted />)
		fireEvent.click(screen.getByRole("button", { name: "Export" }))

		const field = await screen.findByLabelText(preview.previewOfficePdfPasswordLabel)

		fireEvent.change(field, { target: { value: "nope" } })
		fireEvent.click(screen.getByRole("button", { name: preview.previewOfficePdfPasswordSubmit }))
		expect(await screen.findByText(preview.previewOfficePdfPasswordWrong)).toBeTruthy()
		expect(state.exports).toEqual([])

		fireEvent.change(screen.getByLabelText(preview.previewOfficePdfPasswordLabel), { target: { value: "secret" } })
		fireEvent.click(screen.getByRole("button", { name: preview.previewOfficePdfPasswordSubmit }))
		await waitFor(() => {
			expect(state.exports).toEqual(["secret"])
		})
		expect(state.checks).toEqual(["", "nope", "secret"])
		// The prompt closes as the export starts, so it cannot start a second one.
		await waitFor(() => {
			expect(screen.queryByRole("dialog")).toBeNull()
		})
		expect(state.exports).toEqual(["secret"])
	})

	it("does not ask about a workbook Excel's default password opens, and leaves its PDF unprotected", async () => {
		reset()
		state.defaulted = true
		render(<Harness encrypted />)
		fireEvent.click(screen.getByRole("button", { name: "Export" }))
		await waitFor(() => {
			expect(state.exports).toEqual([undefined])
		})
		expect(screen.queryByLabelText(preview.previewOfficePdfPasswordLabel)).toBeNull()
	})

	it("refuses a password the PDF cannot be protected with in the prompt, before the worker checks it", async () => {
		reset()
		state.defaulted = false

		const error = vi.spyOn(toast, "error").mockImplementation(() => "")

		render(<Harness encrypted />)
		fireEvent.click(screen.getByRole("button", { name: "Export" }))
		// SASLprep prohibits control characters (RFC 4013 §2.3, RFC 3454 C.2.1).
		fireEvent.change(await screen.findByLabelText(preview.previewOfficePdfPasswordLabel), { target: { value: "secret\u0007" } })
		fireEvent.click(screen.getByRole("button", { name: preview.previewOfficePdfPasswordSubmit }))

		await waitFor(() => {
			expect(error).toHaveBeenCalledWith(preview.previewSheetPdfPasswordUnsupported)
		})
		await waitFor(() => {
			expect(screen.queryByRole("dialog")).toBeNull()
		})
		// Only the default-password probe before the prompt reached the worker.
		expect(state.checks).toEqual([""])
		expect(state.exports).toEqual([])
		error.mockRestore()
	})
})
