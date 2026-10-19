import { parentPort } from "node:worker_threads"
import { serveDocxZipWorker } from "@filen/office/dom/docx/save"

// The zip half of a save off the main thread, as the web's docxZip.worker.ts runs it (bench: OpenOptions.zipWorker).
const port = parentPort

if (port === null) throw new Error("Not a worker")

serveDocxZipWorker({
	post: (message, transfer) => {
		port.postMessage(message, [...transfer])
	},
	listen: onMessage => {
		port.on("message", onMessage)
	}
})
