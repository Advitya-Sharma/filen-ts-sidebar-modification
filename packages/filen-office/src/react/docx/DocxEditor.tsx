import { useEffect, useImperativeHandle, useLayoutEffect, useRef, useState, type MouseEvent, type PointerEvent, type Ref } from "react"
import { Region } from "../../core/docx/layout/index"
import type { HitResult, PagedEditor } from "../../core/docx/paged/index"
import { EditSink } from "./edit/sink"
import { EditSurface, type DocxEditorMark } from "./edit/surface"
import type { DocxDocument } from "./store"

// Paged editing on the rendered pages, unstyled: the pages (only those in view mounted, patched line by line
// as edits lay them out again), the caret, selection and composition drawn over them, pointer and touch
// selection mapped through the layout's geometry, and the hidden input view that holds focus. State shows as
// data-* attributes, the look as --fo-* variables; toolbars come from useDocxEditor.

export interface DocxEditorHandle {
	focus(): void
	scrollToCaret(): void
	scrollToPage(index: number, y?: number): void
	// The selection's two ends on screen (client px; bottom of each end's line), for touch handles and a menu
	// placed next to the selection; undefined where an end is not drawn.
	selectionPoints(): { readonly start: ClientPoint | undefined; readonly end: ClientPoint | undefined }
	// Moves one end of the selection to a point on screen (a touch handle being dragged): the other end stays
	// where it was when the drag began, whichever side of it this one goes.
	dragSelectionEnd(end: "start" | "end", clientX: number, clientY: number): void
	// The handle was let go.
	endSelectionDrag(): void
	// The position under a point on screen in the story being edited (where a context menu was asked for).
	hitTest(clientX: number, clientY: number): HitResult | undefined
	// The selection as a copy writes it, for an app's own Cut and Copy (see EditSink.clipboardContent).
	clipboardContent(): { readonly html: string; readonly text: string } | undefined
	deleteSelection(): boolean
	// Pastes what the app read from the system clipboard (see EditSink.paste).
	paste(content: { readonly html?: string; readonly text?: string }): boolean
	selectAll(): boolean
	// The element that scrolls the pages (overlays an app places over them follow its scroll events).
	scrollElement(): HTMLElement | null
}

export interface ClientPoint {
	readonly x: number
	readonly y: number
	readonly height: number
}

export interface DocxEditorProps {
	readonly editor: PagedEditor
	// The renderer's side of the same document (pages, tables, objects, media).
	readonly document: DocxDocument
	// 1 = 100 % (96 px per inch).
	readonly zoom: number
	readonly gap?: number
	readonly overscan?: number
	readonly className?: string
	// The accessible name of the editing area.
	readonly label: string
	readonly readOnly?: boolean
	// Cmd as the shortcut modifier (Apple platforms).
	readonly apple?: boolean
	// Find requested (Mod+F): the app's own find.
	readonly onFind?: () => void
	// Ranges to mark (spelling, search results), drawn as data-fo-mark boxes; sorted by `from` costs least.
	readonly marks?: readonly DocxEditorMark[]
	// Top-level blocks the input view renders on each side of the selection.
	readonly window?: number
	readonly ref?: Ref<DocxEditorHandle>
}

const EMPTY_MARKS: readonly DocxEditorMark[] = []
const REGION_NAMES: Readonly<Record<number, string>> = {
	[Region.BODY]: "body",
	[Region.HEADER]: "header",
	[Region.FOOTER]: "footer",
	[Region.FOOTNOTE]: "footnote",
	[Region.ENDNOTE]: "endnote"
}
const DEFAULT_PAGE_WIDTH = 816
const DEFAULT_PAGE_HEIGHT = 1056
// A press held this long without moving selects the word (touch).
const LONG_PRESS_MS = 500
const TAP_SLOP = 10
// Presses this close in time (ms) and place (px) count as one double or triple click (pointerdown has no
// click count).
const CLICK_MS = 500
const CLICK_SLOP = 4
// Distance from the scroll area's edge where dragging scrolls; the speed (px per ms) per px past its start,
// and the most.
const AUTO_SCROLL_EDGE = 32
const AUTO_SCROLL_RATE = 0.02
const AUTO_SCROLL_MAX = 8

interface Layout {
	readonly tops: Float64Array
	readonly heights: Float64Array
	readonly widths: Float64Array
	readonly total: number
	readonly contentWidth: number
	readonly zoom: number
	readonly gap: number
}

// A page's size on screen; one not laid out yet is sized like the last that is.
function pageWidth(document: DocxDocument, index: number, zoom: number): number {
	const page = document.pages[index] ?? document.pages[document.pages.length - 1]

	return (page?.width ?? document.info?.pageWidth ?? DEFAULT_PAGE_WIDTH) * zoom
}

function pageHeight(document: DocxDocument, index: number, zoom: number): number {
	const page = document.pages[index] ?? document.pages[document.pages.length - 1]

	return (page?.height ?? document.info?.pageHeight ?? DEFAULT_PAGE_HEIGHT) * zoom
}

// The pages' places, the same object while they do not change (pages laid out again keep their size as a
// rule): what is drawn from it is not drawn again on a keystroke.
class PageLayouts {
	private last: Layout | undefined = undefined

	layout(document: DocxDocument, count: number, zoom: number, gap: number): Layout {
		const last = this.last

		if (last !== undefined && last.tops.length === count && last.zoom === zoom && last.gap === gap) {
			let same = true

			for (let i = 0; i < count && same; i++)
				same = pageWidth(document, i, zoom) === last.widths[i] && pageHeight(document, i, zoom) === last.heights[i]

			if (same) return last
		}

		const tops = new Float64Array(count)
		const heights = new Float64Array(count)
		const widths = new Float64Array(count)
		let y = gap
		let widest = 0

		for (let i = 0; i < count; i++) {
			const width = pageWidth(document, i, zoom)
			const height = pageHeight(document, i, zoom)

			tops[i] = y
			heights[i] = height
			widths[i] = width
			widest = Math.max(widest, width)
			y += height + gap
		}

		this.last = { tops, heights, widths, total: y, contentWidth: widest + 2 * gap, zoom, gap }

		return this.last
	}
}

// The first page whose bottom is below `y`.
function pageAt(layout: Layout, y: number): number {
	let low = 0
	let high = layout.tops.length - 1

	while (low < high) {
		const middle = (low + high) >> 1

		if ((layout.tops[middle] ?? 0) + (layout.heights[middle] ?? 0) < y) low = middle + 1
		else high = middle
	}

	return low
}

interface Drag {
	readonly pointer: number
	readonly anchor: number
	clientX: number
	clientY: number
	frame: number
	// The last auto-scroll frame's time (0: none yet).
	time: number
}

interface Click {
	time: number
	x: number
	y: number
	count: number
}

interface Touch {
	readonly pointer: number
	readonly x: number
	readonly y: number
	timer: ReturnType<typeof setTimeout> | undefined
	handled: boolean
}

export function DocxEditor(props: DocxEditorProps) {
	const { editor, document, zoom, ref } = props
	const gap = props.gap ?? 16
	const overscan = props.overscan ?? 1
	const containerRef = useRef<HTMLDivElement | null>(null)
	const contentRef = useRef<HTMLDivElement | null>(null)
	const sinkHostRef = useRef<HTMLDivElement | null>(null)
	const sinkRef = useRef<EditSink | null>(null)
	const dragRef = useRef<Drag | null>(null)
	const touchRef = useRef<Touch | null>(null)
	const clickRef = useRef<Click>({ time: -Infinity, x: 0, y: 0, count: 0 })
	// The end a touch handle drag keeps fixed.
	const handleDragRef = useRef<{ end: "start" | "end"; fixed: number } | null>(null)
	const [surface] = useState(() => new EditSurface(editor))
	const [layouts] = useState(() => new PageLayouts())
	const [viewport, setViewport] = useState({ top: 0, height: 0, width: 0 })
	const [count, setCount] = useState(() => Math.max(editor.pageCount, document.pages.length))
	const [focused, setFocused] = useState(false)
	const [composing, setComposing] = useState(false)
	const [region, setRegion] = useState(() => editor.activeRegion)
	const readOnly = props.readOnly === true || editor.session.readOnly
	const layout = layouts.layout(document, count, zoom, gap)
	const visible =
		count === 0
			? { first: 0, last: -1 }
			: {
					first: Math.max(0, pageAt(layout, viewport.top) - overscan),
					last: Math.min(count - 1, pageAt(layout, viewport.top + viewport.height) + overscan)
				}
	const leftOf = (index: number): number =>
		Math.max(gap, ((Math.max(viewport.width, layout.contentWidth) - (layout.widths[index] ?? 0)) / 2) | 0)

	// The painted caret on screen and its column's width, for the input view to put its own caret there.
	const caretClientRect = (): { left: number; top: number; height: number; width: number } | undefined => {
		const caret = surface.caret()
		const element = containerRef.current?.querySelector<HTMLElement>(`[data-fo-page="${String(caret?.index ?? -1)}"]`)
		const page = caret === undefined ? undefined : editor.page(caret.index)

		if (caret === undefined || page === undefined || element === null || element === undefined) return undefined

		const rect = element.getBoundingClientRect()
		const column = editor.geometry.columnBox(page, caret.box.line)

		return {
			left: rect.left + caret.box.x * zoom,
			top: rect.top + caret.box.top * zoom,
			height: caret.box.height * zoom,
			width: (column.right - column.left) * zoom
		}
	}

	const scrollToCaret = (): void => {
		const container = containerRef.current
		const caret = editor.caret()

		if (container === null || caret === undefined) return

		const top = (layout.tops[caret.page] ?? 0) + caret.top * zoom
		const bottom = top + caret.height * zoom
		const margin = Math.min(40, container.clientHeight / 4)

		if (top < container.scrollTop + margin) container.scrollTop = Math.max(0, top - margin)
		else if (bottom > container.scrollTop + container.clientHeight - margin)
			container.scrollTop = bottom - container.clientHeight + margin
	}

	// The page nearest to a point on screen (by y: between pages, the closer one) and the point in its
	// coordinates, inside the page.
	const pointAt = (clientX: number, clientY: number): { index: number; x: number; y: number } | undefined => {
		const content = contentRef.current

		if (content === null || count === 0) return undefined

		const rect = content.getBoundingClientRect()
		const y = clientY - rect.top
		let index = Math.min(count - 1, pageAt(layout, y))

		if (index > 0 && y < (layout.tops[index] ?? 0)) {
			const above = (layout.tops[index - 1] ?? 0) + (layout.heights[index - 1] ?? 0)

			if (y - above < (layout.tops[index] ?? 0) - y) index--
		}

		const width = (layout.widths[index] ?? 0) / zoom
		const height = (layout.heights[index] ?? 0) / zoom

		return {
			index,
			x: Math.max(0, Math.min(width, (clientX - rect.left - leftOf(index)) / zoom)),
			y: Math.max(0, Math.min(height, (y - (layout.tops[index] ?? 0)) / zoom))
		}
	}

	// The position under a point on screen, in the story being edited.
	const hitAt = (clientX: number, clientY: number): HitResult | undefined => {
		const point = pointAt(clientX, clientY)

		return point === undefined ? undefined : editor.hitTest(point.index, point.x, point.y)
	}

	// Double-clicking a header or footer edits it; double-clicking the body goes back to it (as Word does).
	const switchStory = (clientX: number, clientY: number): void => {
		const point = pointAt(clientX, clientY)

		if (point === undefined) return

		const story = editor.storyAt(point.index, point.x, point.y)

		if (story >= 0 && story !== editor.activeStory) editor.enterStory(story, point.index)
		else if (story < 0 && editor.activeStory !== editor.body.session.story.story.index) editor.exitStory()
		else if (story >= 0) editor.geometry.preferPage(point.index)
	}

	const pageMove = (dir: 1 | -1, extend: boolean): boolean => {
		const container = containerRef.current
		const caret = editor.caret()

		if (container === null || caret === undefined) return false

		const step = container.clientHeight * 0.9
		const y = (layout.tops[caret.page] ?? 0) + (caret.top + caret.height / 2) * zoom + dir * step
		const index = Math.max(0, Math.min(count - 1, pageAt(layout, y)))
		const goal = editor.goalX() ?? caret.x
		const hit = editor.hitTest(index, goal, (y - (layout.tops[index] ?? 0)) / zoom)

		container.scrollTop = Math.max(0, container.scrollTop + dir * step)

		return hit === undefined ? false : editor.moveTo(hit, extend, true)
	}

	useImperativeHandle(ref, () => ({
		focus: () => sinkRef.current?.focus(),
		scrollToCaret,
		scrollToPage(index, y = 0) {
			const container = containerRef.current

			if (container !== null && index >= 0 && index < count)
				container.scrollTop = Math.max(0, (layout.tops[index] ?? 0) + y * zoom - gap)
		},
		selectionPoints() {
			const content = contentRef.current?.getBoundingClientRect()
			const ends = editor.ends()
			const point = (box: typeof ends.start): ClientPoint | undefined =>
				box === undefined || content === undefined
					? undefined
					: {
							x: content.left + leftOf(box.page) + box.x * zoom,
							y: content.top + (layout.tops[box.page] ?? 0) + (box.top + box.height) * zoom,
							height: box.height * zoom
						}

			return { start: point(ends.start), end: point(ends.end) }
		},
		dragSelectionEnd(end, clientX, clientY) {
			const hit = hitAt(clientX, clientY)
			let drag = handleDragRef.current

			if (drag === null || drag.end !== end) {
				const selection = editor.selection()

				drag = { end, fixed: end === "start" ? selection.to : selection.from }
				handleDragRef.current = drag
			}

			if (hit !== undefined) editor.selectRange(drag.fixed, hit)
		},
		endSelectionDrag() {
			handleDragRef.current = null
		},
		hitTest: hitAt,
		clipboardContent: () => sinkRef.current?.clipboardContent(),
		deleteSelection: () => sinkRef.current?.deleteSelection() ?? false,
		paste: content => sinkRef.current?.paste(content) ?? false,
		selectAll: () => sinkRef.current?.selectAll() ?? false,
		scrollElement: () => containerRef.current
	}))

	// Handlers the input view calls back read the current render's layout.
	const readOnlyRef = useRef(props.readOnly === true)
	const caretClientRectRef = useRef(caretClientRect)
	const pageMoveRef = useRef(pageMove)
	const onFindRef = useRef(props.onFind)
	const scrollToCaretRef = useRef(scrollToCaret)

	// The view's size and scroll position, read once per frame.
	useEffect(() => {
		const container = containerRef.current

		if (container === null) return

		let frame = 0
		const read = (): void => {
			frame = 0
			setViewport(previous =>
				previous.top === container.scrollTop &&
				previous.height === container.clientHeight &&
				previous.width === container.clientWidth
					? previous
					: { top: container.scrollTop, height: container.clientHeight, width: container.clientWidth }
			)
		}
		const schedule = (): void => {
			if (frame === 0) frame = requestAnimationFrame(read)
		}
		const observer = typeof ResizeObserver === "function" ? new ResizeObserver(schedule) : undefined

		read()
		container.addEventListener("scroll", schedule, { passive: true })
		observer?.observe(container)

		return () => {
			container.removeEventListener("scroll", schedule)
			observer?.disconnect()

			if (frame !== 0) cancelAnimationFrame(frame)
		}
	}, [])

	// The input view, for the editor's lifetime.
	useEffect(() => {
		const host = sinkHostRef.current

		if (host === null) return

		const sink = new EditSink(editor, host, {
			label: props.label,
			readOnly: readOnlyRef.current,
			apple: props.apple === true,
			...(props.window === undefined ? {} : { window: props.window }),
			caretClientRect: () => caretClientRectRef.current(),
			pageMove: (dir, extend) => pageMoveRef.current(dir, extend),
			onFind: () => onFindRef.current?.()
		})
		const onFocus = (): void => setFocused(true)
		const onBlur = (): void => setFocused(false)
		const onComposition = (): void => setComposing(sink.composing)

		sinkRef.current = sink
		sink.view.dom.addEventListener("focus", onFocus)
		sink.view.dom.addEventListener("blur", onBlur)
		sink.view.dom.addEventListener("compositionstart", onComposition)
		sink.view.dom.addEventListener("compositionend", onComposition)

		return () => {
			sink.view.dom.removeEventListener("focus", onFocus)
			sink.view.dom.removeEventListener("blur", onBlur)
			sink.view.dom.removeEventListener("compositionstart", onComposition)
			sink.view.dom.removeEventListener("compositionend", onComposition)
			sink.destroy()
			sinkRef.current = null
		}
		// The input view lives as long as the editor; what it calls back reads the current render through refs.
	}, [editor])

	useLayoutEffect(() => {
		readOnlyRef.current = props.readOnly === true
		sinkRef.current?.setReadOnly(props.readOnly === true)
		caretClientRectRef.current = caretClientRect
		pageMoveRef.current = pageMove
		onFindRef.current = props.onFind
		scrollToCaretRef.current = scrollToCaret
	})

	// Overlays, scrolling to the caret and the input view's alignment happen once per frame, before it is
	// painted (after the pages were patched), whatever triggered them.
	const frameRef = useRef<{ id: number; draw: boolean; scroll: boolean; align: boolean }>({
		id: 0,
		draw: false,
		scroll: false,
		align: false
	})
	const requestFrame = (work: { draw?: boolean; scroll?: boolean; align?: boolean }): void => {
		const frame = frameRef.current

		// Plain assignments: the React Compiler does not lower `||=` (it would skip this component).
		if (work.draw === true) frame.draw = true
		if (work.scroll === true) frame.scroll = true
		if (work.align === true) frame.align = true

		if (frame.id !== 0) return

		frame.id = requestAnimationFrame(() => {
			const current = frameRef.current

			current.id = 0

			if (current.draw) surface.draw()
			if (current.scroll && sinkRef.current?.focused === true) scrollToCaretRef.current()
			if (current.align || current.draw) sinkRef.current?.align()

			current.draw = false
			current.scroll = false
			current.align = false
		})
	}
	const requestFrameRef = useRef(requestFrame)

	useLayoutEffect(() => {
		requestFrameRef.current = requestFrame
	})

	// Layout and selection changes from the editor.
	useEffect(
		() =>
			editor.subscribe(event => {
				if (event.type === "pages") setCount(previous => (previous === event.count ? previous : event.count))
				else {
					if (event.type === "story") setRegion(editor.activeRegion)

					requestFrameRef.current({ draw: true, scroll: true })
				}
			}),
		[editor]
	)

	useEffect(
		() => () => {
			if (frameRef.current.id !== 0) cancelAnimationFrame(frameRef.current.id)

			surface.dispose()
		},
		[surface]
	)

	// Scrolling moves the painted caret on screen: the input view follows (IME windows anchor to it).
	useEffect(() => {
		if (focused) requestFrame({ align: true })
	}, [viewport.top, viewport.width, focused, zoom])

	// Pages painted from the renderer's current data, the view's pages reported; overlays redrawn when
	// anything they are drawn over changed.
	useLayoutEffect(() => {
		const session = document.session
		const media = document.media

		editor.setVisible(visible.first, visible.last)

		const repainted =
			session !== undefined && media !== undefined && surface.update(document.pages, { session, media, tables: document.tables })
		const marks = props.marks ?? EMPTY_MARKS
		const marked = surface.setMarks(marks)
		const changed = surface.setFocused(focused) || marked

		if (repainted || changed) requestFrame({ draw: true })
	})

	const stopAutoScroll = (): void => {
		const drag = dragRef.current

		if (drag !== null && drag.frame !== 0) cancelAnimationFrame(drag.frame)
	}

	// Dragging near or past the scroll area's edge scrolls, faster the farther past it.
	const autoScroll = (time: number): void => {
		const drag = dragRef.current
		const container = containerRef.current

		if (drag === null || container === null) return

		const rect = container.getBoundingClientRect()
		const past =
			drag.clientY < rect.top + AUTO_SCROLL_EDGE
				? drag.clientY - rect.top - AUTO_SCROLL_EDGE
				: drag.clientY > rect.bottom - AUTO_SCROLL_EDGE
					? drag.clientY - rect.bottom + AUTO_SCROLL_EDGE
					: 0
		const elapsed = drag.time === 0 ? 16 : Math.min(50, Math.max(0, time - drag.time))

		drag.frame = 0

		if (past === 0) {
			drag.time = 0

			return
		}

		drag.time = time

		const speed = Math.min(AUTO_SCROLL_MAX, Math.abs(past) * AUTO_SCROLL_RATE)

		container.scrollTop += Math.sign(past) * Math.max(1, Math.round(speed * elapsed))

		const hit = hitAt(drag.clientX, drag.clientY)

		if (hit !== undefined) editor.selectRange(drag.anchor, hit)

		drag.frame = requestAnimationFrame(autoScroll)
	}

	// A press on the pages or the room around them (not the scroll bar).
	const onContent = (event: PointerEvent<HTMLDivElement> | MouseEvent<HTMLDivElement>): boolean => {
		const container = containerRef.current

		if (event.button !== 0 || container === null || !(event.target instanceof Node) || !container.contains(event.target)) return false
		if (event.target !== container) return true

		const rect = container.getBoundingClientRect()

		return event.clientX - rect.left < container.clientWidth && event.clientY - rect.top < container.clientHeight
	}

	// The browser's own handling of a press (focus moving to the pages, a native selection) never runs: on
	// touch it is the compatibility mousedown after a tap (the pointer events left scrolling alone).
	const onMouseDown = (event: MouseEvent<HTMLDivElement>): void => {
		if (onContent(event)) event.preventDefault()
	}

	const onPointerDown = (event: PointerEvent<HTMLDivElement>): void => {
		if (!onContent(event)) return

		if (event.pointerType === "touch") {
			// Touch scrolls; a tap places the caret, a long press selects the word.
			const touch: Touch = { pointer: event.pointerId, x: event.clientX, y: event.clientY, timer: undefined, handled: false }

			touch.timer = setTimeout(() => {
				touch.timer = undefined

				const hit = hitAt(touch.x, touch.y)

				if (touchRef.current !== touch || hit === undefined) return

				touch.handled = true
				editor.selectWord(hit.pos)
				sinkRef.current?.focus()
			}, LONG_PRESS_MS)
			touchRef.current = touch

			return
		}

		event.preventDefault()

		// Presses at one point of the document (scrolling in between moves it), Shift+click extending.
		const click = clickRef.current
		const content = contentRef.current?.getBoundingClientRect()
		const x = event.clientX - (content?.left ?? 0)
		const y = event.clientY - (content?.top ?? 0)
		const again = !event.shiftKey && event.timeStamp - click.time <= CLICK_MS && Math.hypot(x - click.x, y - click.y) <= CLICK_SLOP

		click.count = event.shiftKey ? 1 : event.detail > 0 ? event.detail : again ? click.count + 1 : 1
		click.time = event.timeStamp
		click.x = x
		click.y = y

		const clicks = click.count
		const onPage = event.target instanceof Element && event.target.closest("[data-fo-page]") !== null

		if (clicks === 2 && onPage) switchStory(event.clientX, event.clientY)
		else if (editor.activeStory !== editor.body.session.story.story.index) {
			// Clicking another page's copy of the header being edited moves the caret there.
			const point = pointAt(event.clientX, event.clientY)

			if (point !== undefined && editor.storyAt(point.index, point.x, point.y) === editor.activeStory)
				editor.geometry.preferPage(point.index)
		}

		const hit = hitAt(event.clientX, event.clientY)

		if (hit === undefined) return

		if (clicks >= 3) editor.selectParagraph(hit.pos)
		else if (clicks === 2) editor.selectWord(hit.pos)
		else editor.select(hit, event.shiftKey)

		sinkRef.current?.focus()
		dragRef.current = {
			pointer: event.pointerId,
			anchor: editor.selection().anchor,
			clientX: event.clientX,
			clientY: event.clientY,
			frame: 0,
			time: 0
		}
		event.currentTarget.setPointerCapture(event.pointerId)
	}

	const onPointerMove = (event: PointerEvent<HTMLDivElement>): void => {
		const touch = touchRef.current

		if (touch !== null && touch.pointer === event.pointerId) {
			if (Math.hypot(event.clientX - touch.x, event.clientY - touch.y) > TAP_SLOP && touch.timer !== undefined) {
				clearTimeout(touch.timer)
				touch.timer = undefined
				touch.handled = true
			}

			return
		}

		const drag = dragRef.current

		if (drag === null || drag.pointer !== event.pointerId) return

		drag.clientX = event.clientX
		drag.clientY = event.clientY

		const hit = hitAt(event.clientX, event.clientY)

		if (hit !== undefined) editor.selectRange(drag.anchor, hit)

		if (drag.frame === 0) drag.frame = requestAnimationFrame(autoScroll)
	}

	const endDrag = (pointer: number): void => {
		if (dragRef.current?.pointer === pointer) {
			stopAutoScroll()
			dragRef.current = null
		}
	}

	// A gesture the browser took over (scrolling, a system gesture): nothing is selected or focused.
	const onPointerCancel = (event: PointerEvent<HTMLDivElement>): void => {
		const touch = touchRef.current

		if (touch !== null && touch.pointer === event.pointerId) {
			if (touch.timer !== undefined) clearTimeout(touch.timer)

			touchRef.current = null
		}

		endDrag(event.pointerId)
	}

	const onPointerUp = (event: PointerEvent<HTMLDivElement>): void => {
		const touch = touchRef.current

		if (touch !== null && touch.pointer === event.pointerId) {
			if (touch.timer !== undefined) clearTimeout(touch.timer)

			touchRef.current = null

			if (!touch.handled) {
				const hit = hitAt(event.clientX, event.clientY)

				if (hit !== undefined) editor.select(hit)

				// Focus in the gesture itself: mobile keyboards only open from one.
				sinkRef.current?.focus()
			}

			return
		}

		endDrag(event.pointerId)
	}

	const pages = []

	for (let index = visible.first; index <= visible.last; index++) {
		pages.push(
			<EditorPage
				key={index}
				index={index}
				surface={surface}
				zoom={zoom}
				top={layout.tops[index] ?? 0}
				left={leftOf(index)}
				width={layout.widths[index] ?? 0}
				height={layout.heights[index] ?? 0}
				pending={document.pages[index] === undefined}
			/>
		)
	}

	return (
		<div
			className={props.className === undefined ? "fo-docx-editor" : `fo-docx-editor ${props.className}`}
			data-fo-editing=""
			data-fo-status={document.status}
			data-fo-story={REGION_NAMES[region] ?? "other"}
			{...(focused ? { "data-fo-focused": "" } : {})}
			{...(composing ? { "data-fo-composing": "" } : {})}
			{...(readOnly ? { "data-fo-read-only": "" } : {})}
		>
			<div
				ref={containerRef}
				className="fo-docx-pages"
				// The input view is what assistive technology reads; the pages are its picture.
				aria-hidden="true"
				onMouseDown={onMouseDown}
				onPointerDown={onPointerDown}
				onPointerMove={onPointerMove}
				onPointerUp={onPointerUp}
				onPointerCancel={onPointerCancel}
			>
				<div
					ref={contentRef}
					className="fo-docx-pages-content"
					style={{ height: layout.total, width: layout.contentWidth }}
				>
					{pages}
				</div>
			</div>
			<div
				ref={sinkHostRef}
				className="fo-docx-sink-host"
			/>
		</div>
	)
}

interface EditorPageProps {
	readonly index: number
	readonly surface: EditSurface
	readonly zoom: number
	readonly top: number
	readonly left: number
	readonly width: number
	readonly height: number
	readonly pending: boolean
}

function EditorPage({ index, surface, zoom, top, left, width, height, pending }: EditorPageProps) {
	const hostRef = useRef<HTMLDivElement | null>(null)
	const layerRef = useRef<HTMLDivElement | null>(null)

	useLayoutEffect(() => {
		const host = hostRef.current
		const layer = layerRef.current

		if (host === null || layer === null) return

		surface.mount(index, host, layer)

		return () => surface.unmount(index)
	}, [index, surface])

	return (
		<div
			className="fo-docx-page"
			data-fo-page={index}
			{...(pending ? { "data-fo-pending": "" } : {})}
			style={{ top, left, width, height }}
		>
			<div
				className="fo-docx-page-scale"
				style={{ transform: zoom === 1 ? undefined : `scale(${zoom})` }}
			>
				<div ref={hostRef} />
				<div
					ref={layerRef}
					className="fo-docx-edit-layer"
				/>
			</div>
		</div>
	)
}
