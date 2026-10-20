import { useEffect, useRef, useState } from "react"
import {
	mostVisiblePage,
	pdfStepZoomScale,
	pdfWheelZoomScale,
	pdfZoomPercent,
	clampPdfScale,
	PdfFinder,
	PDF_BASE_SCALE,
	PDF_FIND_INTS,
	type PdfFindResult,
	type PdfViewerCommand,
	type PdfViewerState
} from "../../dom/pdf/index"
import type { Slicing } from "../../core/slicing"
import type { PdfPagesBinding } from "./PdfPages"
import type { PdfDomDocument } from "./types"

// The viewer over an open document: page tracking, navigation, zoom and find, driven by PdfViewerCommands.
// Render `<PdfPages {...viewer.pages} />` with it.

export interface PdfViewer {
	readonly state: PdfViewerState
	readonly dispatch: (command: PdfViewerCommand) => void
	readonly pages: PdfPagesBinding
}

interface FindState {
	readonly open: boolean
	readonly query: string
	readonly caseSensitive: boolean
	readonly result: PdfFindResult | null
	readonly active: number
	// The page in view when the query was searched: the first match from it on becomes active.
	readonly from: number
	readonly reveal: number
}

const CLOSED: FindState = { open: false, query: "", caseSensitive: false, result: null, active: -1, from: 1, reveal: 0 }

// The search over the document's text gives the page a turn every slice.
const FIND_SLICING: Slicing = {
	pause: () =>
		new Promise(resolve => {
			setTimeout(resolve, 0)
		}),
	sliceMs: 8,
	now: () => performance.now()
}

function reducedMotion(): boolean {
	return typeof matchMedia === "function" && matchMedia("(prefers-reduced-motion: reduce)").matches
}

export function usePdfViewer(document: PdfDomDocument): PdfViewer {
	const scrollerRef = useRef<HTMLDivElement | null>(null)
	const elements = useRef(new Map<number, HTMLElement>())
	const ratios = useRef(new Map<number, number>())
	const [page, setPage] = useState(1)
	const [scale, setScale] = useState(PDF_BASE_SCALE)
	const [find, setFind] = useState<FindState>(CLOSED)
	const [finder, setFinder] = useState(() => ({ document, finder: new PdfFinder(document, FIND_SLICING) }))
	const pageCount = document.numPages
	const searching = find.open && find.query.length > 0 ? find.query : null
	const caseSensitive = find.caseSensitive

	if (finder.document !== document) setFinder({ document, finder: new PdfFinder(document, FIND_SLICING) })

	// Ctrl/Cmd+wheel zooms: a native, non-passive listener, as React's wheel handler is passive and could not stop
	// the browser's own page zoom. A plain wheel scrolls.
	useEffect(() => {
		const scroller = scrollerRef.current

		if (scroller === null) return

		const onWheel = (event: WheelEvent): void => {
			if (!event.ctrlKey && !event.metaKey) return

			event.preventDefault()
			setScale(previous => pdfWheelZoomScale(previous, event.deltaY))
		}

		scroller.addEventListener("wheel", onWheel, { passive: false })

		return () => {
			scroller.removeEventListener("wheel", onWheel)
		}
	}, [])

	// A newer query (or closing) aborts the search under way.
	useEffect(() => {
		if (searching === null) return

		const signal = { aborted: false }
		const current = finder.finder

		void current.search(searching, caseSensitive, signal).then(result => {
			if (result === null || signal.aborted) return

			setFind(previous => {
				if (!previous.open || previous.query !== searching || previous.caseSensitive !== caseSensitive) return previous

				let active = -1

				for (let match = 0; match < result.matches.length / PDF_FIND_INTS; match++) {
					if ((result.matches[match * PDF_FIND_INTS] ?? 0) >= previous.from) {
						active = match

						break
					}
				}

				if (active < 0 && result.matches.length > 0) active = 0

				return { ...previous, result, active, reveal: active < 0 ? previous.reveal : previous.reveal + 1 }
			})
		})

		return () => {
			signal.aborted = true
		}
	}, [finder, searching, caseSensitive])

	const goToPage = (target: number, smooth: boolean): void => {
		const clamped = Math.min(Math.max(1, Math.round(target)), pageCount)

		elements.current.get(clamped)?.scrollIntoView({ behavior: smooth && !reducedMotion() ? "smooth" : "auto", block: "start" })
	}

	const dispatch = (command: PdfViewerCommand): void => {
		switch (command.type) {
			case "page.previous":
				goToPage(page - 1, true)
				break
			case "page.next":
				goToPage(page + 1, true)
				break
			case "page.goto":
				goToPage(command.page, true)
				break
			case "zoom.in":
				setScale(previous => pdfStepZoomScale(previous, 1))
				break
			case "zoom.out":
				setScale(previous => pdfStepZoomScale(previous, -1))
				break
			case "zoom.set":
				setScale(clampPdfScale(command.scale))
				break
			case "zoom.by":
				setScale(previous => clampPdfScale(previous + command.delta))
				break
			case "find.open":
				setFind(previous => (previous.open ? previous : { ...CLOSED, open: true, caseSensitive: previous.caseSensitive }))
				break
			case "find.close":
				setFind(previous => ({ ...CLOSED, caseSensitive: previous.caseSensitive, reveal: previous.reveal }))
				break
			case "find.query": {
				const nextCase = command.caseSensitive ?? find.caseSensitive

				if (command.query === find.query && nextCase === find.caseSensitive && find.open) break

				setFind(previous => ({
					...previous,
					open: true,
					query: command.query,
					caseSensitive: nextCase,
					result: null,
					active: -1,
					from: page
				}))
				break
			}
			case "find.next":
			case "find.previous": {
				const step = command.type === "find.next" ? 1 : -1

				setFind(previous => {
					const count = previous.result === null ? 0 : previous.result.matches.length / PDF_FIND_INTS

					if (count === 0) return previous

					return { ...previous, active: (previous.active + step + count) % count, reveal: previous.reveal + 1 }
				})
				break
			}
		}
	}

	const onVisible = (visiblePage: number, ratio: number): void => {
		ratios.current.set(visiblePage, ratio)

		const visible = mostVisiblePage(ratios.current)

		if (visible !== null) setPage(visible)
	}

	const registerPage = (registered: number, element: HTMLElement | null): void => {
		if (element === null) elements.current.delete(registered)
		else elements.current.set(registered, element)
	}

	const onNavigate = (target: number): void => {
		goToPage(target, false)
	}

	const count = find.result === null ? 0 : find.result.matches.length / PDF_FIND_INTS
	const state: PdfViewerState = {
		page: Math.min(page, pageCount),
		pageCount,
		scale,
		zoom: pdfZoomPercent(scale),
		find: {
			open: find.open,
			query: find.query,
			caseSensitive: find.caseSensitive,
			done: find.query.length === 0 || find.result !== null,
			count,
			capped: find.result?.capped ?? false,
			active: find.active
		}
	}

	return {
		state,
		dispatch,
		pages: {
			document,
			scale,
			find: { matches: find.open ? (find.result?.matches ?? null) : null, active: find.active, reveal: find.reveal },
			scrollerRef,
			onVisible,
			registerPage,
			onNavigate
		}
	}
}
