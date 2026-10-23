// @vitest-environment jsdom

import { describe, expect, it, vi } from "vitest"
import { fireEvent, render, screen, waitFor } from "@testing-library/react"
import { toast } from "sonner"
import { TooltipProvider } from "@/components/ui/tooltip"
import { DocxFileMenu } from "@/features/preview/components/docxFileMenu"
import { useDocxPdfExport, type DocxPdfExportTarget } from "@/features/preview/hooks/useDocxPdfExport"
import { PreviewDownloadableProvider } from "@/features/preview/lib/accessMode"
import "@/lib/i18n"

const exports = vi.hoisted(() => ({
	calls: [] as { kind: string; name: string; password: string | undefined }[],
	unlocks: [] as string[]
}))

vi.mock("@/features/preview/lib/docxPdf", () => ({
	downloadDocxPdf: (source: { kind: string }, name: string, password?: string) => {
		exports.calls.push({ kind: source.kind, name, password })

		return Promise.resolve()
	},
	unlockForPdf: (_bytes: Uint8Array, password: string) => {
		exports.unlocks.push(password)

		return Promise.resolve(
			password.startsWith("right") ? { verdict: "accepted", document: { release: () => undefined } } : { verdict: "wrong" }
		)
	}
}))

const target: DocxPdfExportTarget = { fileName: "Report.docx", source: () => ({ kind: "file", bytes: new Uint8Array([1]) }) }

function Harness({ exportTarget }: { exportTarget: DocxPdfExportTarget | null }) {
	const pdf = useDocxPdfExport(exportTarget)

	return (
		<TooltipProvider>
			<DocxFileMenu onDownloadPdf={pdf.onDownloadPdf} />
			{pdf.prompt}
		</TooltipProvider>
	)
}

describe("the Word document's File menu", () => {
	it("downloads the document as a PDF from File ▸ Download as ▸ PDF", async () => {
		exports.calls.length = 0
		render(
			<PreviewDownloadableProvider
				downloadable
				onDownload={() => undefined}
			>
				<Harness exportTarget={target} />
			</PreviewDownloadableProvider>
		)

		fireEvent.click(screen.getByRole("button", { name: "File" }))
		fireEvent.click(await screen.findByRole("menuitem", { name: "Download as" }))
		fireEvent.click(await screen.findByRole("menuitem", { name: "PDF document (.pdf)" }))

		await waitFor(() => {
			expect(exports.calls).toEqual([{ kind: "file", name: "Report.docx", password: undefined }])
		})
	})

	it("is not offered where the file may not be downloaded, nor in the trash", () => {
		const { rerender } = render(
			<PreviewDownloadableProvider downloadable={false}>
				<Harness exportTarget={target} />
			</PreviewDownloadableProvider>
		)

		expect(screen.queryByRole("button", { name: "File" })).toBeNull()

		// In the trash the overlay offers no Download: its handler is null.
		rerender(
			<PreviewDownloadableProvider
				downloadable
				onDownload={null}
			>
				<Harness exportTarget={target} />
			</PreviewDownloadableProvider>
		)

		expect(screen.queryByRole("button", { name: "File" })).toBeNull()
	})

	it("asks a password-protected document's password first and checks it", async () => {
		exports.calls.length = 0
		render(
			<PreviewDownloadableProvider
				downloadable
				onDownload={() => undefined}
			>
				<Harness exportTarget={{ ...target, encryptedFile: new Uint8Array([0xd0, 0xcf]) }} />
			</PreviewDownloadableProvider>
		)

		fireEvent.click(screen.getByRole("button", { name: "File" }))
		fireEvent.click(await screen.findByRole("menuitem", { name: "Download as" }))
		fireEvent.click(await screen.findByRole("menuitem", { name: "PDF document (.pdf)" }))

		const field = await screen.findByLabelText("Document password")

		fireEvent.change(field, { target: { value: "not it" } })
		fireEvent.click(screen.getByRole("button", { name: "Download PDF" }))

		expect(await screen.findByText("That is not the document's password")).toBeTruthy()
		expect(exports.calls).toEqual([])

		fireEvent.change(screen.getByLabelText("Document password"), { target: { value: "right" } })
		fireEvent.click(screen.getByRole("button", { name: "Download PDF" }))

		// The prompt closes as the export starts, so it cannot start a second one.
		await waitFor(() => {
			expect(screen.queryByRole("dialog")).toBeNull()
		})
		expect(exports.calls).toEqual([{ kind: "unlocked", name: "Report.docx", password: "right" }])
	})

	it("refuses a password the PDF cannot be protected with in the prompt, before the document is unlocked or laid out", async () => {
		exports.calls.length = 0
		exports.unlocks.length = 0

		const error = vi.spyOn(toast, "error").mockImplementation(() => "")

		render(
			<PreviewDownloadableProvider
				downloadable
				onDownload={() => undefined}
			>
				<Harness exportTarget={{ ...target, encryptedFile: new Uint8Array([0xd0, 0xcf]) }} />
			</PreviewDownloadableProvider>
		)

		fireEvent.click(screen.getByRole("button", { name: "File" }))
		fireEvent.click(await screen.findByRole("menuitem", { name: "Download as" }))
		fireEvent.click(await screen.findByRole("menuitem", { name: "PDF document (.pdf)" }))
		// SASLprep prohibits control characters (RFC 4013 §2.3, RFC 3454 C.2.1).
		fireEvent.change(await screen.findByLabelText("Document password"), { target: { value: "right\u0007" } })
		fireEvent.click(screen.getByRole("button", { name: "Download PDF" }))

		await waitFor(() => {
			expect(error).toHaveBeenCalledWith("The PDF could not be made: this document's password cannot protect a PDF")
		})
		await waitFor(() => {
			expect(screen.queryByRole("dialog")).toBeNull()
		})
		expect(exports.unlocks).toEqual([])
		expect(exports.calls).toEqual([])
		error.mockRestore()
	})
})
