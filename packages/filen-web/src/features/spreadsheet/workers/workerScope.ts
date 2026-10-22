import type { SheetWorkerScope } from "@filen/office/dom/sheet/text"

// The worker's end of its channel to the page.
export function workerScope(): SheetWorkerScope {
	return {
		post: (message, transfer) => {
			self.postMessage(message, { transfer: [...transfer] })
		},
		listen: onMessage => {
			self.addEventListener("message", event => {
				onMessage(event.data)
			})
		}
	}
}
