import { docxWorkerPool } from "@filen/office/dom/docx"
import { OFFICE_FONT_URLS } from "virtual:office-fonts"
import DocxWorker from "@/features/preview/workers/docx.worker.ts?worker"
import { recoverIfNewerBuild } from "@/lib/appUpdate"

// One docx worker for the session, shared by every open document and terminated a minute after the last
// one closes: the next document skips the worker's start-up and the fonts it already fetched. A closed
// document's engine drops everything it held; none of the engine is in the main bundle.
const WORKER_IDLE_MS = 60_000

const pool = docxWorkerPool(() => {
	const worker = new DocxWorker()

	return {
		post(message, transfer) {
			worker.postMessage(message, [...transfer])
		},
		listen(onMessage, onError) {
			worker.addEventListener("message", event => {
				onMessage(event.data)
			})
			worker.addEventListener("error", () => {
				// A deploy that changed this worker removes its old file (appUpdate.ts).
				recoverIfNewerBuild()
				onError()
			})
		},
		terminate() {
			worker.terminate()
		}
	}
}, WORKER_IDLE_MS)

export const docxHost = pool.host

// Password-protected files: probed, unlocked and held in the same worker, so neither the password spin nor the
// decrypted package is on the page (an editor takes its own copy).
export const docxProtection = pool.protection

// Starts the worker while a document still downloads, so its script is parsed by the time the bytes arrive.
export function prewarmDocxWorker(): void {
	pool.prewarm()
}

// Where the worker fetches each bundled font (content-hashed, same origin).
export const DOCX_FONT_URLS = OFFICE_FONT_URLS
