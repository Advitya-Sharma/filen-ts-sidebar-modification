import { serveDocxZipWorker } from "@filen/office/dom/docx/save"

// The zip half of a DOCX save (lib/docxEditing.ts): the checked parts deflated into the original archive and
// the result checked, off the page's thread. One worker per document being edited; the original file's
// bytes stay on the page, which answers the ranges this worker copies.
serveDocxZipWorker({
	post: (message, transfer) => {
		self.postMessage(message, { transfer: [...transfer] })
	},
	listen: onMessage => {
		self.addEventListener("message", event => {
			onMessage(event.data)
		})
	}
})
