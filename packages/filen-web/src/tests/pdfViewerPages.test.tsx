// @vitest-environment jsdom

import { beforeEach, describe, expect, it, vi } from "vitest"
import { act, fireEvent, render, screen, waitFor } from "@testing-library/react"
import "@/lib/i18n"
import { registerAllActions } from "@/lib/keymap/actions"
import { linkedFileItem } from "@/tests/fixtures/sdk"
import { createTestQueryClient, queryClientWrapper } from "@/tests/testQueryClient"

// The preview's PDF viewer on @filen/office's pages against a stand-in pdf.js: pages, the link gate and find.

interface FakePage {
	getViewport: ReturnType<typeof vi.fn>
	render: ReturnType<typeof vi.fn>
	cleanup: ReturnType<typeof vi.fn>
	streamTextContent: () => unknown
	getTextContent: () => Promise<{ items: unknown[] }>
	getAnnotations: ReturnType<typeof vi.fn>
	userUnit: number
}

const { pages } = vi.hoisted(() => ({ pages: new Map<number, FakePage>() }))

// The link gate reads the trusted link domains.
vi.mock("@/lib/storage/adapter", () => ({ kvGetJson: () => Promise.resolve(null), kvSetJson: () => Promise.resolve() }))

function fakePage(text: string[] = []): FakePage {
	const items = text.map(str => ({ str, hasEOL: false }))

	return {
		getViewport: vi.fn(() => ({ width: 100, height: 100, convertToViewportPoint: (x: number, y: number) => [x, y] })),
		render: vi.fn(() => ({ promise: Promise.resolve(), cancel: vi.fn() })),
		cleanup: vi.fn(() => true),
		streamTextContent: () => items,
		getTextContent: () => Promise.resolve({ items }),
		getAnnotations: vi.fn(() => Promise.resolve([])),
		userUnit: 1
	}
}

vi.mock("pdfjs-dist", () => ({
	GlobalWorkerOptions: { workerSrc: "" },
	// Only a ranged source constructs one.
	PDFDataRangeTransport: vi.fn(),
	TextLayer: class {
		readonly textDivs: HTMLElement[] = []
		readonly textContentItemsStr: string[] = []
		private readonly params: { textContentSource: { str: string }[]; container: HTMLElement }

		constructor(params: { textContentSource: { str: string }[]; container: HTMLElement }) {
			this.params = params
		}

		render(): Promise<void> {
			for (const { str } of this.params.textContentSource) {
				const span = document.createElement("span")

				span.textContent = str
				this.textDivs.push(span)
				this.textContentItemsStr.push(str)
				this.params.container.append(span)
			}

			return Promise.resolve()
		}

		cancel(): void {
			// Renders at once.
		}
	},
	getDocument: () => ({
		promise: Promise.resolve({
			numPages: 3,
			getPage: (n: number) => Promise.resolve(pages.get(n)),
			getDestination: () => Promise.resolve(null),
			getPageIndex: () => Promise.resolve(0)
		}),
		destroy: () => Promise.resolve()
	})
}))

vi.mock("@/features/preview/hooks/usePreviewBytes", () => ({
	usePreviewBytes: () => ({ status: "success", bytes: new Uint8Array([1]), refetch: () => undefined })
}))

interface Observed {
	callback: IntersectionObserverCallback
	options: IntersectionObserverInit | undefined
	elements: Set<Element>
}

const observers: Observed[] = []

class FakeIntersectionObserver {
	readonly entry: Observed

	constructor(callback: IntersectionObserverCallback, options?: IntersectionObserverInit) {
		this.entry = { callback, options, elements: new Set() }
		observers.push(this.entry)
	}

	observe(element: Element): void {
		this.entry.elements.add(element)
	}

	unobserve(element: Element): void {
		this.entry.elements.delete(element)
	}

	disconnect(): void {
		this.entry.elements.clear()
	}
}

const { PdfViewer } = await import("@/features/preview/components/pdfViewer")

registerAllActions()

const item = linkedFileItem("doc.pdf", { uuid: "aaaaaaaa-0000-0000-0000-000000000001", mime: { Decrypted: "application/pdf" }, size: 1n })

function pageElement(n: number): Element {
	const element = document.querySelector(`[data-page="${String(n)}"]`)

	if (element === null) {
		throw new Error(`no page ${String(n)}`)
	}

	return element
}

// Brings a page into both margins, as the view would on scrolling to it.
function show(n: number, ratio = 1): void {
	act(() => {
		for (const observed of observers) {
			observed.callback(
				[{ target: pageElement(n), isIntersecting: true, intersectionRatio: ratio } as unknown as IntersectionObserverEntry],
				{} as IntersectionObserver
			)
		}
	})
}

beforeEach(() => {
	observers.length = 0
	pages.clear()

	for (const n of [1, 2, 3]) {
		pages.set(n, fakePage())
	}

	vi.stubGlobal("IntersectionObserver", FakeIntersectionObserver)
	// jsdom lays nothing out and has no scrollIntoView.
	Object.defineProperty(Element.prototype, "scrollIntoView", { configurable: true, value: vi.fn() })
})

async function mount(): Promise<void> {
	render(
		<PdfViewer
			item={item}
			alt="doc.pdf"
		/>,
		{ wrapper: queryClientWrapper(createTestQueryClient()) }
	)

	await waitFor(() => {
		for (const n of [1, 2, 3]) {
			expect(pages.get(n)?.getViewport).toHaveBeenCalled()
		}
	})
}

describe("PdfViewer", () => {
	it("labels each page's canvas and follows the page most in view", async () => {
		await mount()

		expect(screen.getByLabelText("doc.pdf — Page 2 of 3")).toBeTruthy()
		expect(screen.getByText("Page 1 of 3")).toBeTruthy()

		show(2)

		expect(await screen.findByText("Page 2 of 3")).toBeTruthy()
	})

	it("asks before a link annotation opens a domain not trusted yet", async () => {
		pages.set(1, {
			...fakePage(),
			getAnnotations: vi.fn(() => Promise.resolve([{ subtype: "Link", url: "https://example.com/x", rect: [0, 0, 10, 10] }]))
		})
		await mount()
		show(1)

		const link = await screen.findByRole("link", { name: "https://example.com/x" })
		const click = new MouseEvent("click", { bubbles: true, cancelable: true })

		link.dispatchEvent(click)

		expect(click.defaultPrevented).toBe(true)
		expect(await screen.findByRole("alertdialog", { name: "Open external link?" })).toBeTruthy()
	})

	it("finds text from the toolbar and highlights the current match", async () => {
		pages.set(1, fakePage(["alpha beta"]))
		pages.set(3, fakePage(["beta"]))
		await mount()
		show(1)

		fireEvent.click(screen.getByRole("button", { name: "Find in document" }))
		fireEvent.change(await screen.findByRole("searchbox", { name: "Find" }), { target: { value: "beta" } })

		expect(await screen.findByText("1/2")).toBeTruthy()
		await waitFor(() => {
			expect(pageElement(1).querySelector(".highlight.selected")?.textContent).toBe("beta")
		})

		fireEvent.click(screen.getByRole("button", { name: "Close find" }))

		await waitFor(() => {
			expect(document.querySelector(".highlight")).toBeNull()
		})
		expect(screen.getByText("Page 1 of 3")).toBeTruthy()
	})
})
