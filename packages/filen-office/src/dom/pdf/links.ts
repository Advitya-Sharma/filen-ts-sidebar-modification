import { classifyLink } from "@filen/shared/linkPolicy"
import type { PdfDocument } from "./types"

// A page's Link annotations reduced to what a viewer renders: the rectangle in PDF space and either an external
// URL the shared link policy allows or an in-document destination. pdf.js's own link service is never used:
// it assigns hrefs for schemes the policy refuses. Everything else (javascript: and data: URIs, launch, named
// and attachment actions, a URL with no host) is dropped before it reaches the DOM.

export type PdfRect = readonly [number, number, number, number]

export type PdfLink =
	| { readonly kind: "external"; readonly rect: PdfRect; readonly url: string }
	// A named destination (string) or an explicit one (an array whose first item is a page reference).
	| { readonly kind: "internal"; readonly rect: PdfRect; readonly dest: string | readonly unknown[] }

function isRecord(value: unknown): value is Readonly<Record<string, unknown>> {
	return typeof value === "object" && value !== null
}

function asRect(value: unknown): PdfRect | null {
	const rect: readonly unknown[] | null = Array.isArray(value) ? (value as readonly unknown[]) : null

	if (rect?.length !== 4) return null

	const [x1, y1, x2, y2] = rect

	if (typeof x1 !== "number" || typeof y1 !== "number" || typeof x2 !== "number" || typeof y2 !== "number") return null
	if (!Number.isFinite(x1) || !Number.isFinite(y1) || !Number.isFinite(x2) || !Number.isFinite(y2)) return null

	return [x1, y1, x2, y2]
}

// `allow` narrows the policy further (a host that must also parse the URL); it never widens it.
export function pdfLinks(annotations: readonly unknown[], allow?: (url: string) => boolean): PdfLink[] {
	const links: PdfLink[] = []

	for (const entry of annotations) {
		if (!isRecord(entry) || entry["subtype"] !== "Link") continue

		const rect = asRect(entry["rect"])

		if (rect === null) continue

		// pdf.js leaves `url` unset for schemes its own list lacks (sms:, geo:, ...); the raw value is judged by
		// the shared policy instead, which is the stricter of the two.
		const url = typeof entry["url"] === "string" ? entry["url"] : entry["unsafeUrl"]

		if (typeof url === "string") {
			const verdict = classifyLink(url)

			if (verdict.kind === "external" && (allow === undefined || allow(verdict.url)))
				links.push({ kind: "external", rect, url: verdict.url })

			continue
		}

		const dest = entry["dest"]

		if ((typeof dest === "string" && dest.length > 0) || (Array.isArray(dest) && dest.length > 0)) {
			links.push({ kind: "internal", rect, dest: dest as string | readonly unknown[] })
		}
	}

	return links
}

// The 1-based page a destination points at, or null for one that points nowhere (a broken document).
export async function resolveDestination(document: PdfDocument, dest: string | readonly unknown[]): Promise<number | null> {
	try {
		const explicit = typeof dest === "string" ? await document.getDestination(dest) : dest
		const target: unknown = Array.isArray(explicit) ? explicit[0] : undefined

		if (target === undefined || target === null) return null

		// A page index instead of a reference: some producers write one (PDF 32000-1 §12.3.2.2 allows it only
		// for remote go-to actions, pdf.js accepts it everywhere).
		const index = typeof target === "number" ? target : await document.getPageIndex(target)

		return Number.isInteger(index) && index >= 0 && index < document.numPages ? index + 1 : null
	} catch {
		return null
	}
}
