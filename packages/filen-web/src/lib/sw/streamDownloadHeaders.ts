import { contentDispositionAttachment } from "@/lib/filename"

// The response headers of a download the page streams through the service worker: a forced attachment, with its
// length when the page knows it (null for bytes made as they stream, which the browser then reads to the end).
export function streamDownloadHeaders(name: string, size: number | null): Record<string, string> {
	return {
		"Content-Type": "application/octet-stream",
		"Content-Disposition": contentDispositionAttachment(name),
		...(size === null ? {} : { "Content-Length": String(size) }),
		"X-Content-Type-Options": "nosniff"
	}
}
